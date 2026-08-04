"""Generate the monthly closing checklist (README §11 item 5).

Reads `04 Recurring Payments` and produces a workbook listing everything that
falls due in a given month, ready to work through and tick off.

This skill issues no document number, touches no counter, and moves no money.
It is a reminder of what to prepare, not an instruction to pay: each line still
goes through the normal voucher and approval route (§2).

Items whose timing the register does not record are listed separately rather
than omitted. An annual audit fee that quietly never appears on any checklist
is the failure this design exists to prevent.
"""

from __future__ import annotations

import argparse
import sys
from collections import defaultdict
from datetime import date
from decimal import Decimal
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(REPO_ROOT / "scripts"))

from excel_engine import excel_app, export_worksheet  # noqa: E402
from recurring import Recurring, load_recurring, split_for_month  # noqa: E402

REGISTER = REPO_ROOT / "registers" / "NEK_Master_Registers.xlsx"
OUTPUT_DIR = REPO_ROOT / "output" / "checklists"

XL_CONTINUOUS, XL_THIN, XL_MEDIUM = 1, 2, -4138
EDGE_LEFT, EDGE_TOP, EDGE_BOTTOM, EDGE_RIGHT = 7, 8, 9, 10
XL_LEFT, XL_CENTER, XL_RIGHT = -4131, -4108, -4152
XL_PAPER_A4, XL_LANDSCAPE = 9, 2
HEADER_FILL = 0x64381F   # BGR for 1F3864, matching the document templates
NOTE_FILL = 0xCCF2FF     # BGR for FFF2CC

COLUMNS = [
    ("Due", 6), ("Ref", 14), ("Entity", 9), ("Counterparty", 30),
    ("Description", 34), ("Category", 18), ("Cur", 5), ("Amount", 12),
    ("Method", 14), ("Bank Account", 26), ("Prepare", 13), ("Approve", 11),
    ("Done", 6),
]


def _row_values(item: Recurring, year: int, month: int) -> list:
    due = item.due_date(year, month)
    return [
        due.day if due else "",
        item.ref, item.entity, item.counterparty, item.description,
        item.category, item.currency,
        float(item.amount) if item.amount is not None else item.amount_text,
        item.method, item.bank_account, item.preparer, item.approver, "",
    ]


def _style_header(ws, row: int, span: int) -> None:
    rng = ws.Range(ws.Cells(row, 1), ws.Cells(row, span))
    rng.Interior.Color = HEADER_FILL
    rng.Font.Color = 0xFFFFFF
    rng.Font.Bold = True
    rng.HorizontalAlignment = XL_CENTER
    rng.BorderAround(LineStyle=XL_CONTINUOUS, Weight=XL_THIN)


def _section(ws, row: int, title: str, note: str = "") -> int:
    ws.Cells(row, 1).Value = title
    ws.Cells(row, 1).Font.Bold = True
    ws.Cells(row, 1).Font.Size = 11
    if note:
        ws.Cells(row + 1, 1).Value = note
        ws.Cells(row + 1, 1).Font.Italic = True
        ws.Cells(row + 1, 1).Font.Size = 8
        return row + 2
    return row + 1


def build(ws, year: int, month: int, due: list[Recurring],
          unknown: list[Recurring], entity: str | None) -> None:
    span = len(COLUMNS)
    period = date(year, month, 1).strftime("%B %Y")
    scope = entity.upper() if entity else "all entities"

    ws.Cells(1, 1).Value = f"MONTHLY CLOSING CHECKLIST - {period.upper()}"
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 14
    ws.Cells(2, 1).Value = (
        f"Scope: {scope}. Generated from 04 Recurring Payments on "
        f"{date.today():%Y-%m-%d}. Each line still requires its own voucher "
        "and approval -- this checklist authorises nothing."
    )
    ws.Cells(2, 1).Font.Italic = True
    ws.Cells(2, 1).Font.Size = 8

    for index, (name, width) in enumerate(COLUMNS, start=1):
        ws.Columns(index).ColumnWidth = width

    row = 4
    for direction in ("Payable", "Receivable"):
        section = [i for i in due if i.direction.lower() == direction.lower()]
        if not section:
            continue
        row = _section(ws, row, f"{direction.upper()}  ({len(section)} items)")
        for index, (name, _) in enumerate(COLUMNS, start=1):
            ws.Cells(row, index).Value = name
        _style_header(ws, row, span)
        row += 1
        first = row
        for item in section:
            for index, value in enumerate(_row_values(item, year, month), start=1):
                ws.Cells(row, index).Value = value
            ws.Cells(row, 8).NumberFormat = "#,##0.00"
            ws.Cells(row, span).Borders(EDGE_BOTTOM).LineStyle = XL_CONTINUOUS
            row += 1
        body = ws.Range(ws.Cells(first, 1), ws.Cells(row - 1, span))
        body.Borders.LineStyle = XL_CONTINUOUS
        body.Borders.Weight = XL_THIN

        by_currency: dict[str, Decimal] = defaultdict(Decimal)
        prose = 0
        for item in section:
            if item.amount is None:
                prose += 1
            else:
                by_currency[item.currency] += item.amount
        totals = "   ".join(f"{cur} {amt:,.2f}" for cur, amt in sorted(by_currency.items()))
        if prose:
            totals += f"   (+{prose} item(s) with no numeric amount recorded)"
        ws.Cells(row, 1).Value = f"Total {direction.lower()}:  {totals}"
        ws.Cells(row, 1).Font.Bold = True
        row += 2

    if unknown:
        row = _section(
            ws, row,
            f"TIMING NOT RECORDED  ({len(unknown)} items)",
            "These recur, but 04 Recurring Payments does not say when. They are "
            "listed every month until a due date is recorded, so that none is "
            "silently missed. Filling in the Due Day column removes them from here.",
        )
        grouped: dict[tuple[str, str, str], list[Recurring]] = defaultdict(list)
        for item in unknown:
            grouped[(item.frequency, item.category, item.counterparty)].append(item)
        headers = ["Frequency", "Category", "Counterparty", "Entities", "Cur", "Amount each"]
        for index, name in enumerate(headers, start=1):
            ws.Cells(row, index).Value = name
        _style_header(ws, row, len(headers))
        row += 1
        first = row
        for (frequency, category, counterparty), group in sorted(grouped.items()):
            entities = ", ".join(sorted({i.entity for i in group}))
            amounts = sorted({i.amount_text for i in group})
            ws.Cells(row, 1).Value = frequency
            ws.Cells(row, 2).Value = category
            ws.Cells(row, 3).Value = counterparty
            ws.Cells(row, 4).Value = entities
            ws.Cells(row, 5).Value = group[0].currency
            ws.Cells(row, 6).Value = " / ".join(amounts)
            row += 1
        block = ws.Range(ws.Cells(first, 1), ws.Cells(row - 1, len(headers)))
        block.Borders.LineStyle = XL_CONTINUOUS
        block.Borders.Weight = XL_THIN
        block.Interior.Color = NOTE_FILL

    ws.Rows(1).RowHeight = 20
    ws.PageSetup.PaperSize = XL_PAPER_A4
    ws.PageSetup.Orientation = XL_LANDSCAPE
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = False
    ws.PageSetup.PrintArea = f"$A$1:${chr(64 + span)}${row}"


def main() -> int:
    p = argparse.ArgumentParser(prog="generate_checklist", description=__doc__,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--month", required=True, help="YYYY-MM")
    p.add_argument("--entity", default=None, help="limit to one entity code")
    p.add_argument("--pdf", action="store_true", help="also export a PDF")
    p.add_argument("--printer", default=None, help="A4 device, if the default substitutes Letter")
    args = p.parse_args()

    try:
        year, month = (int(part) for part in args.month.split("-"))
        date(year, month, 1)
    except (ValueError, TypeError):
        print(f"ERROR: --month {args.month!r} is not YYYY-MM", file=sys.stderr)
        return 1

    items = load_recurring(REGISTER)
    due, unknown = split_for_month(items, year, month, args.entity)

    print(f"Monthly closing checklist - {date(year, month, 1):%B %Y}")
    print(f"  register    {len(items)} active recurring items")
    print(f"  due         {len(due)}")
    print(f"  no timing   {len(unknown)}  (listed separately, not dropped)")

    if not due and not unknown:
        print("\nNothing to list.")
        return 0

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    suffix = f"_{args.entity.upper()}" if args.entity else ""
    destination = OUTPUT_DIR / f"{year}-{month:02d}{suffix}_Monthly-Closing-Checklist.xlsx"

    with excel_app() as app:
        wb = app.Workbooks.Add()
        try:
            ws = wb.Worksheets(1)
            ws.Name = f"{year}-{month:02d}"
            build(ws, year, month, due, unknown, args.entity)
            wb.SaveAs(str(destination))
            if args.pdf:
                export_worksheet(app, wb, ws.Name, destination.with_suffix(".pdf"),
                                 printer=args.printer)
        finally:
            wb.Close(SaveChanges=False)

    print(f"\nWritten: {destination}")
    if args.pdf:
        print(f"         {destination.with_suffix('.pdf')}")
    print("This checklist authorises nothing. Each line still needs its own "
          "voucher and approval.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
