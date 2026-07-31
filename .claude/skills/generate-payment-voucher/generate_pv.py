"""Generate a Payment Voucher draft (README §8).

DRY-RUN ONLY. This script cannot issue a document. It never increments the
counter, never writes to ``output/`` outside ``output/dryrun/``, and never
executes, schedules or releases a payment (§2). Live issuance is a separate,
reviewed change; ``--live`` exists solely to fail with an explanation.

Everything happens inside one Excel session: mirrors are refreshed from the
master register, inputs are written, formulas are rebuilt, the result is
verified against independently-computed values, and only then is the PDF
exported. Launching Excel costs ~50s, so doing it once rather than five times
is what makes the tool usable.
"""

from __future__ import annotations

import argparse
import shutil
import sys
import tempfile
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(REPO_ROOT / "scripts"))

import counters  # noqa: E402
from amount_in_words import CURRENCIES, words_for_cell  # noqa: E402
from excel_engine import excel_app, export_worksheet, open_workbook  # noqa: E402
from fill_template import (  # noqa: E402
    CP_GUARD,
    PRINTED_REFERENCE,
    SHEET_NAME,
    TOTAL_CELL,
    LineItem,
    fill_payment_voucher,
    validate_line_items,
)
from ref_and_filename import (  # noqa: E402
    build_filename,
    build_reference,
    check_path_length,
    counterparty_token,
    load_vendor_index,
)
from registers import load_registers, sync_currency, sync_mirrors  # noqa: E402

TEMPLATE = REPO_ROOT / "templates" / "NEK_Document_Templates.xlsx"
REGISTER = REPO_ROOT / "registers" / "NEK_Master_Registers.xlsx"
DRYRUN_DIR = REPO_ROOT / "output" / "dryrun"

DOCTYPE = "PV"
DRAFT_PREFIX = "DRAFT_"
DRAFT_HEADER = "DRAFT - NOT ISSUED - no counter consumed"

# Whatever the wording table covers. Driven from one place so the template's
# printed label and the script's written words cannot disagree.
SUPPORTED_CURRENCIES = set(CURRENCIES)

# Rounding tolerance when comparing Excel's total to Python's. Excel returns a
# float; the line items are Decimals.
TOTAL_TOLERANCE = Decimal("0.005")


class GenerationError(RuntimeError):
    """Raised when a voucher must not be produced."""


def parse_line(spec: str) -> LineItem:
    """``description|amount|account_code`` (account code optional)."""
    parts = spec.split("|")
    if len(parts) not in (2, 3):
        raise GenerationError(
            f"line {spec!r} must be 'description|amount' or "
            "'description|amount|account_code'"
        )
    try:
        amount = Decimal(parts[1].strip().replace(",", ""))
    except InvalidOperation as exc:
        raise GenerationError(f"line {spec!r} has an unreadable amount") from exc
    return LineItem(
        description=parts[0].strip(),
        amount=amount,
        account_code=parts[2].strip() if len(parts) == 3 else "",
    )


def validate(
    entity_code: str,
    bank_code: str,
    doc_date: date,
    line_items: list[LineItem],
    entities: dict,
    accounts: dict,
):
    """Every check that can fail runs before Excel is launched."""
    entity_code = entity_code.upper()
    bank_code = bank_code.upper()

    if entity_code not in entities:
        raise GenerationError(
            f"entity {entity_code!r} is not in the register. "
            f"Known: {', '.join(sorted(entities))}"
        )

    key = f"{entity_code}|{bank_code}"
    if key not in accounts:
        linked = sorted(a.bank_code for a in accounts.values() if a.entity_code == entity_code)
        raise GenerationError(
            f"bank {bank_code!r} is not linked to entity {entity_code!r} "
            f"(§7 mismatch guard). Accounts on file for {entity_code}: "
            f"{', '.join(linked) if linked else 'none'}"
        )

    account = accounts[key]
    if account.is_placeholder:
        # Independent of the currency restriction below, so it still holds once
        # non-MYR support lands. WT is the live case: its bank is recorded as
        # TBC and it is an alerts-only entity that issues no documents at all.
        raise GenerationError(
            f"bank code {account.bank_code!r} for {entity_code} is a "
            "placeholder, not a confirmed account. A reference built on it "
            "would be meaningless. Confirm the account in 09 Bank Accounts "
            "first."
        )

    # The template's currency label (E11) and words label (A20) now derive from
    # the account currency, so any currency the wording table covers is safe.
    # An unknown one still fails, in amount_in_words, before anything is drawn.
    if account.currency.upper() not in SUPPORTED_CURRENCIES:
        raise GenerationError(
            f"account {key} is denominated in {account.currency}, which has no "
            f"wording defined. Known: {', '.join(sorted(SUPPORTED_CURRENCIES))}. "
            "Add it to CURRENCIES in scripts/amount_in_words.py -- the template "
            "picks it up automatically."
        )

    validate_line_items(line_items)
    counters.assert_on_or_after_cutover(doc_date)
    return entities[entity_code], account


def generate(args: argparse.Namespace) -> Path:
    doc_date = datetime.strptime(args.date, "%Y-%m-%d").date()
    line_items = [parse_line(spec) for spec in args.line]

    # Read the register exactly once. Loading it again for the mirror sync
    # would let the file change in between, so a voucher could be validated
    # against one version and built from another.
    entities, accounts = load_registers(REGISTER)
    entity, account = validate(
        args.entity, args.bank, doc_date, line_items, entities, accounts
    )

    total = sum((item.amount for item in line_items), Decimal("0"))
    words = words_for_cell(total, account.currency)

    sequence = counters.peek_next(DOCTYPE, entity.code, doc_date)
    reference = build_reference(DOCTYPE, entity.code, account.bank_code, doc_date, sequence)

    vendor_index = load_vendor_index(REGISTER)
    token = counterparty_token(args.pay_to, vendor_index)
    filename = build_filename(
        doc_date, entity.code, DOCTYPE, token, reference, "pdf", prefix=DRAFT_PREFIX
    )

    DRYRUN_DIR.mkdir(parents=True, exist_ok=True)
    destination = DRYRUN_DIR / filename
    check_path_length(destination)

    print(f"  entity      {entity.code}  {entity.legal_name}")
    print(f"  account     {account.key}  {account.bank_name} ({account.currency})")
    print(f"  reference   {reference}   [peeked, counter untouched]")
    print(f"  total       {total}")
    print(f"  in words    {words}")
    print(f"  counterpty  {token}")
    print(f"  output      {destination.relative_to(REPO_ROOT)}")

    with tempfile.TemporaryDirectory(prefix="nek-pv-") as tmp:
        working = Path(tmp) / "working.xlsx"
        shutil.copy2(TEMPLATE, working)

        with excel_app() as app, open_workbook(app, working) as wb:
            # Master register wins, every run, before anything is filled.
            sync_mirrors(wb, entities, accounts)
            sync_currency(wb)

            ws = wb.Worksheets(SHEET_NAME)
            fill_payment_voucher(
                ws,
                entity_code=entity.code,
                bank_code=account.bank_code,
                sequence=sequence,
                doc_date=doc_date,
                pay_to=args.pay_to,
                tt_cheque=args.tt_cheque,
                line_items=line_items,
                amount_words=words,
                prepared_by=args.prepared_by,
                issued_by=args.issued_by,
                approved_by=args.approved_by,
            )

            app.CalculateFullRebuild()
            _verify(ws, reference, total)

            export_worksheet(
                app, wb, SHEET_NAME, destination, draft_header=DRAFT_HEADER
            )

            if args.keep_xlsx:
                wb.SaveCopyAs(str(destination.with_suffix(".xlsx")))

    return destination


def _verify(ws, reference: str, total: Decimal) -> None:
    """A green recalc proves formulas evaluate, not that they are right (§7)."""
    guard = ws.Range(CP_GUARD).Value
    if guard:
        raise GenerationError(f"template mismatch guard fired: {guard}")

    printed = ws.Range(PRINTED_REFERENCE).Value
    if str(printed or "").strip() != reference:
        raise GenerationError(
            f"printed reference {printed!r} does not match the expected "
            f"{reference!r}; the document would carry the wrong number"
        )

    excel_total = ws.Range(TOTAL_CELL).Value
    if excel_total is None:
        raise GenerationError(f"{TOTAL_CELL} is empty after recalculation")
    if abs(Decimal(str(excel_total)) - total) > TOTAL_TOLERANCE:
        raise GenerationError(
            f"total mismatch: sheet says {excel_total}, line items sum to {total}"
        )


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="generate_pv",
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p.add_argument("--entity", required=True, help="entity code, e.g. NEK")
    p.add_argument("--bank", required=True, help="bank code, e.g. BOC")
    p.add_argument("--date", required=True, help="document date, YYYY-MM-DD")
    p.add_argument("--pay-to", required=True, help="payee name")
    p.add_argument("--tt-cheque", required=True, help="payment method, e.g. IBG")
    p.add_argument(
        "--line",
        action="append",
        required=True,
        metavar="DESC|AMOUNT[|ACCT]",
        help="repeatable; up to six",
    )
    p.add_argument("--prepared-by", required=True)
    p.add_argument("--issued-by", required=True)
    p.add_argument("--approved-by", required=True)
    p.add_argument("--keep-xlsx", action="store_true", help="save the filled workbook too")
    p.add_argument(
        "--live",
        action="store_true",
        help="not available -- live issuance is a separate reviewed change",
    )
    return p


def main() -> int:
    args = build_parser().parse_args()

    if args.live:
        print(
            "ERROR: live issuance is not wired.\n"
            "  This skill runs dry-run only: it peeks the next number without\n"
            "  consuming it and writes a watermarked draft to output/dryrun/.\n"
            "  Wiring counters.reserve()/commit() is a separate change that\n"
            "  needs sign-off. This tool never executes or releases payment.",
            file=sys.stderr,
        )
        return 2

    print("DRY RUN - no counter consumed, no payment executed")
    try:
        destination = generate(args)
    except (GenerationError, ValueError, RuntimeError) as exc:
        print(f"\nFAILED: {exc}", file=sys.stderr)
        return 1

    print(f"\nDraft written: {destination}")
    print("This is a DRAFT awaiting human approval. It is not a payment instruction.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
