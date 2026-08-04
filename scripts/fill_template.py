"""Write inputs into the Payment Voucher's yellow anchor cells.

Cell addresses were read from the template rather than assumed; see the map in
:data:`PV_ANCHORS`. The script writes *only* these cells. Everything else on the
sheet -- letterhead, reference, total -- is a formula and stays untouched (§7).

Deliberately not written: the **signature rules** at F23:H23, F25:H25 and
F27:H27. Each sits beside a signatory's initials and is signed by hand once the
voucher is printed, so its blankness is the expected state and never an error.
(The former RECEIVED BY field was removed on 2026-07-31 at the operator's
request; the payee no longer countersigns the voucher.)
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
    "pay_to": "B7",          # merged B7:D7
    "mode_of_payment": "G8",  # merged G8:H8
    "doc_date": "G9",         # merged G9:H9
    # A20:B20 carries the currency label; the words sit in the box beside it.
    "amount_in_words": "C20",
    # Renamed 2026-07-31: the voucher is issued, checked, then approved. The
    # old "prepared by" role is now "issued by", and the old "issued by" is
    # "checked by"; the cells are unchanged.
    "issued_by": "B23",
    "checked_by": "B25",
    "approved_by": "B27",
}
# Each signatory has a blank rule at F:H on their row, signed by hand after
# printing. Never written to. RECEIVED BY was removed on 2026-07-31.
SIGNATURE_RULES = ("F23:H23", "F25:H25", "F27:H27")

# Line-item grid: six rows, description / amount / account code.
LINE_FIRST_ROW = 12
LINE_LAST_ROW = 17
LINE_DESCRIPTION_COL = 1   # A (merged A:D)
LINE_AMOUNT_COL = 5        # E (merged E:F)
LINE_ACCOUNT_COL = 7       # G (merged G:H)
LINE_GRID_LAST_COL = 8     # H -- the grid spans A:H, three merged ranges per row
MAX_LINE_ITEMS = LINE_LAST_ROW - LINE_FIRST_ROW + 1

PRINTED_REFERENCE = "G7"  # merged G7:H7
TOTAL_CELL = "E18"

# Default signatories, used unless a run overrides them.
DEFAULT_ISSUED_BY = "KN"    # Kelvin Ng
DEFAULT_CHECKED_BY = "OHY"  # Ong Hooi Yong

SHEET_NAME = "Payment Voucher"

# The only accepted answers for "Mode of Payment :" (H8). The template carries
# the same three as a dropdown, so a human filling it by hand and a script
# filling it cannot produce different vocabularies.
PAYMENT_MODES = ("IBG", "Cheque", "TT")

# Approximate characters per line in the merged bands, derived from the column
# widths (A13 + B22 + C8 + D13 = 56 for the description band; B..H = 92 for the
# words band). Used only to pick a row height -- see :func:`_fit_row`.
# Calibrated for the 8pt body font on half-A4 stock. Column widths are measured
# in units of the workbook's standard font, so a smaller cell font fits more
# characters per unit -- these are larger than the A4 figures, not smaller.
WIDTH_DESCRIPTION = 70   # A:D
WIDTH_WORDS = 87         # C:H, after the label took A:B
WIDTH_PAYEE = 54         # B:D
LINE_HEIGHT = 10.0
MIN_ROW_HEIGHT = 12.0

# Vertical budget on A5 landscape: 148mm less 0.3in margins, in points. A run
# that would push past this is reported rather than silently scaled down.
USABLE_HEIGHT = 340.0


class FillError(ValueError):
    """Raised when the inputs cannot be written to the template safely."""


@dataclass(frozen=True)
class LineItem:
    description: str
    amount: Decimal
    account_code: str = ""


def content_height(ws, last_row: int = 27) -> float:
    """Total height of the printed rows, in points.

    Worth measuring because the sheet is set to fit one page. If the content
    grows past the sheet, Excel does not spill to a second page -- it silently
    scales everything down, which is exactly what made hairline rules vanish in
    print before. Better to say so than to let it shrink unnoticed.
    """
    return sum(ws.Rows(row).RowHeight for row in range(1, last_row + 1))


def _fit_row(ws, row: int, text: str, width_chars: int) -> None:
    """Grow a row so wrapped text in a merged cell is not clipped.

    Excel auto-fits row height for wrapped text only in *unmerged* cells. Every
    text band on this voucher is merged, so a long payee or description wrapped
    to a second line and then had that line cut off by the fixed row height --
    which is what made long entries look distorted and made the table's ruling
    appear to stop partway down the page.
    """
    lines = max(1, -(-len(str(text)) // width_chars))  # ceiling division
    ws.Rows(row).RowHeight = max(MIN_ROW_HEIGHT, lines * LINE_HEIGHT + 2)


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
    mode_of_payment: str,
    line_items: list[LineItem],
    amount_words: str,
    approved_by: str,
    issued_by: str = DEFAULT_ISSUED_BY,
    checked_by: str = DEFAULT_CHECKED_BY,
) -> None:
    """Write every input cell of an open Payment Voucher worksheet."""
    validate_line_items(line_items)
    if mode_of_payment not in PAYMENT_MODES:
        raise FillError(
            f"mode of payment {mode_of_payment!r} is not one of "
            f"{', '.join(PAYMENT_MODES)}"
        )

    ws.Range(CP_ENTITY).Value = entity_code.upper()
    ws.Range(CP_BANK).Value = bank_code.upper()
    ws.Range(CP_RUNNING_NO).Value = sequence

    ws.Range(PV_ANCHORS["pay_to"]).Value = pay_to
    _fit_row(ws, 7, pay_to, WIDTH_PAYEE)
    ws.Range(PV_ANCHORS["mode_of_payment"]).Value = mode_of_payment
    ws.Range(PV_ANCHORS["doc_date"]).Value = excel_serial(doc_date)
    ws.Range(PV_ANCHORS["amount_in_words"]).Value = amount_words
    _fit_row(ws, 20, amount_words, WIDTH_WORDS)
    ws.Range(PV_ANCHORS["issued_by"]).Value = issued_by
    ws.Range(PV_ANCHORS["checked_by"]).Value = checked_by
    ws.Range(PV_ANCHORS["approved_by"]).Value = approved_by

    # Clear the whole grid first so a shorter run cannot inherit stale rows.
    # Cleared as one block: Excel refuses ClearContents on a single cell inside
    # a merged range, and every row of this grid is three merged ranges.
    ws.Range(
        ws.Cells(LINE_FIRST_ROW, LINE_DESCRIPTION_COL),
        ws.Cells(LINE_LAST_ROW, LINE_GRID_LAST_COL),
    ).ClearContents()

    for row in range(LINE_FIRST_ROW, LINE_LAST_ROW + 1):
        ws.Rows(row).RowHeight = MIN_ROW_HEIGHT

    for offset, item in enumerate(line_items):
        row = LINE_FIRST_ROW + offset
        ws.Cells(row, LINE_DESCRIPTION_COL).Value = item.description
        _fit_row(ws, row, item.description, WIDTH_DESCRIPTION)
        ws.Cells(row, LINE_AMOUNT_COL).Value = float(item.amount)
        if item.account_code:
            ws.Cells(row, LINE_ACCOUNT_COL).Value = item.account_code
