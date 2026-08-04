"""Shared generation pipeline for the internal control vouchers (README §8).

The Payment Voucher and the Receiving Voucher differ in four labels and a
reference prefix. Everything else -- validation, the counter, the reference,
the filename, the single Excel session, the post-recalculation verification --
is identical, so it lives here once and each skill is a thin wrapper naming its
document type.

DRY-RUN ONLY. Nothing here can issue a document. The counter is read without
being consumed, output goes to ``output/dryrun/`` with a ``DRAFT_`` prefix, and
``--live`` fails by design. No routine in this module executes, schedules or
releases a payment (§2).
"""

from __future__ import annotations

import argparse
import shutil
import sys
import tempfile
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT / "scripts"))

import counters  # noqa: E402
from amount_in_words import CURRENCIES, words_for_cell  # noqa: E402
from excel_engine import (  # noqa: E402
    excel_app,
    export_worksheet,
    open_workbook,
    pdf_page_size_mm,
)
from fill_template import (  # noqa: E402
    CP_GUARD,
    DEFAULT_CHECKED_BY,
    DEFAULT_ISSUED_BY,
    DEFAULT_PAYMENT_MODE,
    PAYMENT_MODES,
    PRINTED_REFERENCE,
    SHEET_NAMES,
    TOTAL_CELL,
    USABLE_HEIGHT,
    LineItem,
    content_height,
    fill_voucher,
    validate_line_items,
)
from ref_and_filename import (  # noqa: E402
    build_filename,
    build_reference,
    check_path_length,
    counterparty_token,
    load_vendor_index,
)
from registers import (  # noqa: E402
    load_properties,
    load_registers,
    sync_currency,
    sync_mirrors,
)

TEMPLATE = REPO_ROOT / "templates" / "NEK_Document_Templates.xlsx"
REGISTER = REPO_ROOT / "registers" / "NEK_Master_Registers.xlsx"
DRYRUN_DIR = REPO_ROOT / "output" / "dryrun"

# A dry run is marked by its filename and its folder, not on the page. The
# on-page "DRAFT - NOT ISSUED" header was removed on 2026-07-31 because it ate
# the top margin. Worth revisiting when the live counter is wired: at that point
# a draft and an issued voucher become indistinguishable once printed.
DRAFT_PREFIX = "DRAFT_"
PROPERTY_SHEET_NAME = "02 Property & Lease"

SUPPORTED_CURRENCIES = set(CURRENCIES)

# Excel returns a float; the line items are Decimals.
TOTAL_TOLERANCE = Decimal("0.005")

# The voucher prints on the top half of an A4 portrait sheet.
A4_WIDTH_MM, A4_HEIGHT_MM = 210.0, 297.0
PAPER_TOLERANCE_MM = 3.0


class GenerationError(RuntimeError):
    """Raised when a voucher must not be produced."""


@dataclass(frozen=True)
class VoucherType:
    """What distinguishes one voucher from another."""

    doctype: str              # PV / RV
    noun: str                 # "payment voucher"
    counterparty_flag: str    # --pay-to / --received-from
    counterparty_noun: str    # "payee" / "payer"
    direction: str            # "paid to" / "received from"

    @property
    def sheet(self) -> str:
        return SHEET_NAMES[self.doctype]


PAYMENT_VOUCHER = VoucherType(
    doctype="PV",
    noun="payment voucher",
    counterparty_flag="--pay-to",
    counterparty_noun="payee",
    direction="paid to",
)
RECEIVING_VOUCHER = VoucherType(
    doctype="RV",
    noun="receiving voucher",
    counterparty_flag="--received-from",
    counterparty_noun="payer",
    direction="received from",
)


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


def validate(entity_code, bank_code, doc_date, line_items, entities, accounts):
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
        raise GenerationError(
            f"bank code {account.bank_code!r} for {entity_code} is a "
            "placeholder, not a confirmed account. A reference built on it "
            "would be meaningless. Confirm the account in 09 Bank Accounts first."
        )

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


def generate(args: argparse.Namespace, vt: VoucherType) -> Path:
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

    sequence = counters.peek_next(vt.doctype, entity.code, doc_date)
    reference = build_reference(
        vt.doctype, entity.code, account.bank_code, doc_date, sequence
    )

    token = _counterparty_token(args, entity)
    filename = build_filename(
        doc_date, entity.code, vt.doctype, token, reference, "pdf", prefix=DRAFT_PREFIX
    )

    DRYRUN_DIR.mkdir(parents=True, exist_ok=True)
    destination = DRYRUN_DIR / filename
    check_path_length(destination)

    print(f"  entity      {entity.code}  {entity.legal_name}")
    print(f"  account     {account.key}  {account.bank_name} ({account.currency})")
    print(f"  reference   {reference}   [peeked, counter untouched]")
    print(f"  {vt.direction:<11} {args.counterparty}")
    print(f"  total       {total}")
    print(f"  in words    {words}")
    print(f"  counterpty  {token}")
    print(f"  output      {destination.relative_to(REPO_ROOT)}")

    with tempfile.TemporaryDirectory(prefix="nek-voucher-") as tmp:
        working = Path(tmp) / "working.xlsx"
        shutil.copy2(TEMPLATE, working)

        with excel_app() as app, open_workbook(app, working) as wb:
            # Master register wins, every run, before anything is filled.
            sync_mirrors(wb, entities, accounts)
            sync_currency(wb)

            ws = wb.Worksheets(vt.sheet)
            fill_voucher(
                ws,
                entity_code=entity.code,
                bank_code=account.bank_code,
                sequence=sequence,
                doc_date=doc_date,
                counterparty=args.counterparty,
                mode_of_payment=args.mode_of_payment,
                line_items=line_items,
                amount_words=words,
                issued_by=args.issued_by,
                checked_by=args.checked_by,
                approved_by=args.approved_by,
            )

            app.CalculateFullRebuild()
            _verify(ws, reference, total)

            height = content_height(ws)
            if height > USABLE_HEIGHT:
                print(
                    f"  note        content is {height:.0f}pt against a "
                    f"{USABLE_HEIGHT:.0f}pt half-page budget; Excel will scale "
                    "it down to fit. Shorten a description to avoid that."
                )

            export_worksheet(app, wb, vt.sheet, destination, printer=args.printer)

            if args.keep_xlsx:
                wb.SaveCopyAs(str(destination.with_suffix(".xlsx")))

    _check_paper(destination)
    return destination


def _counterparty_token(args: argparse.Namespace, entity) -> str:
    """The filename's counterparty field.

    ``--unit`` files a rental document under the unit rather than the tenant.
    That keeps a private individual's name out of every filename and matches
    how rental income is actually reviewed -- by unit, not by occupant. The
    tenant's name still appears on the document itself.
    """
    if not args.unit:
        return counterparty_token(args.counterparty, load_vendor_index(REGISTER))

    properties = load_properties(REGISTER)
    key = "".join(ch for ch in args.unit.upper() if ch.isalnum())
    prop = properties.get(key)
    if prop is None:
        known = sorted({p.unit or p.code for p in properties.values()})
        raise GenerationError(
            f"unit {args.unit!r} is not in the property register. "
            f"Known: {', '.join(known)}"
        )
    if prop.entity_code and prop.entity_code != entity.code:
        raise GenerationError(
            f"unit {prop.unit or prop.code!r} belongs to {prop.entity_code}, "
            f"not {entity.code}. Filing it under the wrong entity would put the "
            "document beyond reach of that entity's records."
        )

    # Prefer the unit where it identifies the property on its own -- the Penang
    # units are "1G-11-03" and read perfectly in a filename. The Hong Kong ones
    # are bare numbers, so those fall back to the property code ("27-STRP").
    source = prop.unit if any(ch.isalpha() for ch in prop.unit) else prop.code
    token = counterparty_token(source)
    if not any(ch.isalpha() for ch in token):
        raise GenerationError(
            f"unit {prop.unit!r} and code {prop.code!r} both give a filename "
            f"token of {token!r}, which identifies nothing on its own. Give "
            f"this property a distinctive code in {PROPERTY_SHEET_NAME} "
            "(the others use forms like '1G-11-03' and '27-STRP')."
        )
    return token


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


def _check_paper(destination: Path) -> None:
    """Warn when the driver substituted a different paper size."""
    size = pdf_page_size_mm(destination)
    if size is None:
        return
    width, height = size
    if (abs(width - A4_WIDTH_MM) <= PAPER_TOLERANCE_MM
            and abs(height - A4_HEIGHT_MM) <= PAPER_TOLERANCE_MM):
        return
    print(
        f"  warning     exported at {width:.0f} x {height:.0f} mm, not A4 "
        f"({A4_WIDTH_MM:.0f} x {A4_HEIGHT_MM:.0f}).\n"
        "              The Microsoft virtual printers substitute US Letter for\n"
        "              A4. Either set that printer's default paper to A4 in\n"
        "              Windows, or pass --printer with an A4 device."
    )


def build_parser(vt: VoucherType, description: str) -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog=f"generate_{vt.doctype.lower()}",
        description=description,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p.add_argument("--entity", required=True, help="entity code, e.g. NEK")
    p.add_argument("--bank", required=True, help="bank code, e.g. BOC")
    p.add_argument("--date", required=True, help="document date, YYYY-MM-DD")
    p.add_argument(
        vt.counterparty_flag,
        dest="counterparty",
        required=True,
        help=f"{vt.counterparty_noun} name",
    )
    p.add_argument(
        "--mode",
        dest="mode_of_payment",
        default=DEFAULT_PAYMENT_MODE,
        choices=PAYMENT_MODES,
        help=f"mode of payment (default {DEFAULT_PAYMENT_MODE})",
    )
    p.add_argument(
        "--line",
        action="append",
        required=True,
        metavar="DESC|AMOUNT[|ACCT]",
        help="repeatable; up to six",
    )
    p.add_argument(
        "--issued-by",
        default=DEFAULT_ISSUED_BY,
        help=f"initials of the issuer (default {DEFAULT_ISSUED_BY}, Kelvin Ng)",
    )
    p.add_argument(
        "--checked-by",
        default=DEFAULT_CHECKED_BY,
        help=f"initials of the checker (default {DEFAULT_CHECKED_BY}, Ong Hooi Yong)",
    )
    p.add_argument(
        "--approved-by",
        required=True,
        help="initials of the approver -- no default, approval is deliberate",
    )
    p.add_argument(
        "--unit",
        default=None,
        help="property code or unit, e.g. 1G-11-03. When given, the FILENAME is "
        "filed under the unit instead of the counterparty name; the document "
        "itself is unchanged. Use for rental documents.",
    )
    p.add_argument("--keep-xlsx", action="store_true", help="save the filled workbook too")
    p.add_argument(
        "--printer",
        default=None,
        help="printer whose driver lays out the export; use an A4 device if the "
        "default substitutes US Letter",
    )
    p.add_argument(
        "--live",
        action="store_true",
        help="not available -- live issuance is a separate reviewed change",
    )
    return p


def main(vt: VoucherType, description: str) -> int:
    args = build_parser(vt, description).parse_args()

    if args.live:
        print(
            "ERROR: live issuance is not wired.\n"
            "  This skill runs dry-run only: it peeks the next number without\n"
            "  consuming it and writes a draft to output/dryrun/. Wiring\n"
            "  counters.reserve()/commit() is a separate change that needs\n"
            "  sign-off. This tool never executes or releases payment.",
            file=sys.stderr,
        )
        return 2

    print(f"DRY RUN - no counter consumed, no payment executed [{vt.doctype}]")
    try:
        destination = generate(args, vt)
    except (GenerationError, ValueError, RuntimeError) as exc:
        print(f"\nFAILED: {exc}", file=sys.stderr)
        return 1

    print(f"\nDraft written: {destination}")
    print(f"This is a DRAFT {vt.noun} awaiting human approval. "
          "It is not a payment instruction.")
    return 0
