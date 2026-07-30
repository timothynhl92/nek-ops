"""Write inputs into the Payment Voucher's yellow anchor cells.

Cell addresses were read from the template rather than assumed; see the map in
:data:`PV_ANCHORS`. The script writes *only* these cells. Everything else on the
sheet -- letterhead, reference, total -- is a formula and stays untouched (§7).

Deliberately not written: **F23 (RECEIVED BY)**. It is yellow, but it is a
wet-signature field the payee signs on receipt, so it is correctly outside the
input contract in §8. Its blankness is the expected state and is never an error.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from decimal import Decimal

# Excel's day zero. Serial 1 is 1900-01-01; the offset absorbs Excel's
# deliberate 1900 leap-year bug for every date after 1900-03-01.
EXCEL_EPOCH = date(1899, 12, 30)

# Control Panel (columns J:K, outside the print area)
CP_ENTITY = "K3"
CP_BANK = "K4"
CP_RUNNING_NO = "K5"
CP_BANK_NAME = "K6"     # derived
CP_REFERENCE = "K7"     # derived
CP_GUARD = "K8"         # derived

# Document body
PV_ANCHORS = {
    "pay_to": "B7",
    "tt_cheque": "H8",
    "doc_date": "H9",
    "amount_in_words": "B20",
    "prepared_by": "B23",
    "issued_by": "B25",
    "approved_by": "B27",
}
RECEIVED_BY = "F23"  # intentionally never written

# Line-item grid: six rows, description / amount / account code.
LINE_FIRST_ROW = 12
LINE_LAST_ROW = 17
LINE_DESCRIPTION_COL = 1   # A (merged A:D)
LINE_AMOUNT_COL = 5        # E (merged E:F)
LINE_ACCOUNT_COL = 7       # G (merged G:H)
LINE_GRID_LAST_COL = 8     # H -- the grid spans A:H, three merged ranges per row
MAX_LINE_ITEMS = LINE_LAST_ROW - LINE_FIRST_ROW + 1

PRINTED_REFERENCE = "H7"
TOTAL_CELL = "E18"

SHEET_NAME = "Payment Voucher"


class FillError(ValueError):
    """Raised when the inputs cannot be written to the template safely."""


@dataclass(frozen=True)
class LineItem:
    description: str
    amount: Decimal
    account_code: str = ""


def excel_serial(value: date) -> float:
    """Convert a date to Excel's serial number.

    The date must reach the sheet as a real date, because K7 builds the
    reference with ``TEXT(H9,"yyyymm")``. Passing a ``datetime`` through COM is
    not safe: pywin32 converts it to a COM DATE via the local timezone, so
    midnight on the 1st of a month arrives as the last day of the *previous*
    month under a positive UTC offset -- producing a correct-looking voucher
    carrying the wrong reference. The serial number has no timezone and no
    locale, so it cannot drift.
    """
    return float((value - EXCEL_EPOCH).days)


def validate_line_items(items: list[LineItem]) -> None:
    """§8 edge cases: empty, zero, negative, and more rows than the grid holds."""
    if not items:
        raise FillError("a voucher needs at least one line item")
    if len(items) > MAX_LINE_ITEMS:
        raise FillError(
            f"{len(items)} line items, but the Payment Voucher grid holds "
            f"{MAX_LINE_ITEMS} (rows {LINE_FIRST_ROW}-{LINE_LAST_ROW}). "
            "Split this across multiple vouchers, or widen the template first "
            "-- silently dropping rows would understate the total."
        )
    for index, item in enumerate(items, start=1):
        if not item.description.strip():
            raise FillError(f"line {index} has an empty description")
        if item.amount <= 0:
            raise FillError(
                f"line {index} has a non-positive amount ({item.amount}); "
                "a payment voucher must move money"
            )


def fill_payment_voucher(
    ws,
    *,
    entity_code: str,
    bank_code: str,
    sequence: int,
    doc_date: date,
    pay_to: str,
    tt_cheque: str,
    line_items: list[LineItem],
    amount_words: str,
    prepared_by: str,
    issued_by: str,
    approved_by: str,
) -> None:
    """Write every input cell of an open Payment Voucher worksheet."""
    validate_line_items(line_items)

    ws.Range(CP_ENTITY).Value = entity_code.upper()
    ws.Range(CP_BANK).Value = bank_code.upper()
    ws.Range(CP_RUNNING_NO).Value = sequence

    ws.Range(PV_ANCHORS["pay_to"]).Value = pay_to
    ws.Range(PV_ANCHORS["tt_cheque"]).Value = tt_cheque
    ws.Range(PV_ANCHORS["doc_date"]).Value = excel_serial(doc_date)
    ws.Range(PV_ANCHORS["amount_in_words"]).Value = amount_words
    ws.Range(PV_ANCHORS["prepared_by"]).Value = prepared_by
    ws.Range(PV_ANCHORS["issued_by"]).Value = issued_by
    ws.Range(PV_ANCHORS["approved_by"]).Value = approved_by

    # Clear the whole grid first so a shorter run cannot inherit stale rows.
    # Cleared as one block: Excel refuses ClearContents on a single cell inside
    # a merged range, and every row of this grid is three merged ranges.
    ws.Range(
        ws.Cells(LINE_FIRST_ROW, LINE_DESCRIPTION_COL),
        ws.Cells(LINE_LAST_ROW, LINE_GRID_LAST_COL),
    ).ClearContents()

    for offset, item in enumerate(line_items):
        row = LINE_FIRST_ROW + offset
        ws.Cells(row, LINE_DESCRIPTION_COL).Value = item.description
        ws.Cells(row, LINE_AMOUNT_COL).Value = float(item.amount)
        if item.account_code:
            ws.Cells(row, LINE_ACCOUNT_COL).Value = item.account_code
