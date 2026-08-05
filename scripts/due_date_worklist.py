"""Produce a worklist of recurring payments that have no due date recorded.

Hand it to whoever knows the answers, get it back, and the dates go into
`04 Recurring Payments`. Every item on this list currently sits in the closing
checklist's "timing not recorded" section, which is where things go to be
forgotten.

Rows are grouped: the annual fees are largely the same handful of services
repeated across eight entities, so the list is ~35 rows but only about six
questions. The Group column makes that visible -- answer one row, copy down.

    python scripts/due_date_worklist.py

Read-only with respect to the register; it writes a workbook to output/.
"""

from __future__ import annotations

import sys
from datetime import date
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT / "scripts"))

from excel_engine import excel_app  # noqa: E402
from recurring import load_recurring  # noqa: E402

REGISTER = REPO_ROOT / "registers" / "NEK_Master_Registers.xlsx"
OUTPUT = REPO_ROOT / "output" / "worklists"

XL_CONTINUOUS, XL_THIN = 1, 2
XL_LEFT, XL_CENTER = -4131, -4108
HEADER_FILL = 0x64381F   # BGR for 1F3864
ANSWER_FILL = 0xCCF2FF   # BGR for FFF2CC
BAND_FILL = 0xF2F2F2

COLUMNS = [
    ("Grp", 5), ("Ref", 15), ("Entity", 10), ("Frequency", 12),
    ("Category", 20), ("Counterparty", 34), ("Description", 32),
    ("Cur", 5), ("Amount", 11), ("Recorded now", 13),
    ("DUE DATE - please complete", 30), ("Notes", 26),
]
ANSWER_COL = 11

INSTRUCTIONS = [
    "Every row below recurs, but 04 Recurring Payments does not record when. "
    "Until a date is entered these never appear in the right month of the "
    "closing checklist.",
    "How to answer -- write the month(s) and the day:",
    "   Annual, one date:        December 31",
    "   Twice a year:            February 28 and August 31",
    "   Quarterly:               March, June, September, December 15",
    "If an item has no fixed date (billed on invoice, or it varies), leave the "
    "cell blank and say so in Notes. It stays in the exceptions list rather "
    "than being given a made-up date.",
]


def build(ws, items) -> int:
    ws.Cells(1, 1).Value = "RECURRING PAYMENTS - DUE DATES TO CONFIRM"
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 14
    ws.Cells(2, 1).Value = (
        f"{len(items)} items | generated {date.today():%Y-%m-%d} from "
        "04 Recurring Payments"
    )
    ws.Cells(2, 1).Font.Italic = True

    row = 4
    for line in INSTRUCTIONS:
        ws.Cells(row, 1).Value = line
        ws.Cells(row, 1).Font.Size = 9
        if line.startswith("   "):
            ws.Cells(row, 1).Font.Bold = True
        row += 1
    row += 1

    for index, (name, width) in enumerate(COLUMNS, start=1):
        ws.Columns(index).ColumnWidth = width
        ws.Cells(row, index).Value = name
    header = ws.Range(ws.Cells(row, 1), ws.Cells(row, len(COLUMNS)))
    header.Interior.Color = HEADER_FILL
    header.Font.Color = 0xFFFFFF
    header.Font.Bold = True
    header.HorizontalAlignment = XL_CENTER
    header.WrapText = True
    header_row = row
    row += 1
    first = row

    group_index = 0
    previous_key = None
    for item in items:
        key = (item.frequency, item.category, item.counterparty)
        if key != previous_key:
            group_index += 1
            previous_key = key
        values = [
            group_index, item.ref, item.entity, item.frequency, item.category,
            item.counterparty, item.description, item.currency,
            float(item.amount) if item.amount is not None else item.amount_text,
            item.due_text or "(blank)", "", "",
        ]
        for index, value in enumerate(values, start=1):
            ws.Cells(row, index).Value = value
        ws.Cells(row, 9).NumberFormat = "#,##0.00"
        if group_index % 2 == 0:
            ws.Range(ws.Cells(row, 1), ws.Cells(row, len(COLUMNS))).Interior.Color = BAND_FILL
        answer = ws.Range(ws.Cells(row, ANSWER_COL), ws.Cells(row, ANSWER_COL + 1))
        answer.Interior.Color = ANSWER_FILL
        row += 1

    table = ws.Range(ws.Cells(header_row, 1), ws.Cells(row - 1, len(COLUMNS)))
    table.Borders.LineStyle = XL_CONTINUOUS
    table.Borders.Weight = XL_THIN
    table.VerticalAlignment = XL_CENTER

    ws.Activate()
    ws.Application.ActiveWindow.FreezePanes = False
    ws.Range(f"A{header_row + 1}").Select()
    ws.Application.ActiveWindow.FreezePanes = True
    return group_index


def main() -> int:
    items = [i for i in load_recurring(REGISTER) if not i.timing_known]
    items.sort(key=lambda i: (i.frequency, i.category, i.counterparty, i.entity))
    if not items:
        print("Every recurring item has a due date recorded. Nothing to ask.")
        return 0

    OUTPUT.mkdir(parents=True, exist_ok=True)
    destination = OUTPUT / f"{date.today():%Y-%m-%d}_Due-Dates-To-Confirm.xlsx"

    with excel_app() as app:
        app.Visible = False
        wb = app.Workbooks.Add()
        try:
            ws = wb.Worksheets(1)
            ws.Name = "Due dates"
            groups = build(ws, items)
            wb.SaveAs(str(destination))
        finally:
            wb.Close(SaveChanges=False)

    print(f"{len(items)} items with no usable due date, in {groups} groups")
    print(f"Written: {destination}")
    print("\nHand this over, then send it back and the dates go into the register.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
