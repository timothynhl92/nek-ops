"""Convert an amount to its written form.

README §7 requires this to be generated, never typed. It is implemented here
rather than pulled from a library so the deterministic layer keeps a stable,
auditable output and adds no third-party dependency (§1: Tier 1 runs locally
with no data egress).

House-style note: the Payment Voucher prints the label "Ringgit Malaysia :" in
A20 and the words in B20. So the cell text must *not* repeat the currency name
-- use :func:`words_for_cell`. :func:`amount_in_words` returns the full phrase
including the currency, for contexts that have no separate label.
"""

from __future__ import annotations

from decimal import ROUND_HALF_UP, Decimal

UNITS = (
    "Zero", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight",
    "Nine", "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen",
    "Sixteen", "Seventeen", "Eighteen", "Nineteen",
)
TENS = (
    "", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy",
    "Eighty", "Ninety",
)
SCALES = ((10**9, "Billion"), (10**6, "Million"), (1000, "Thousand"), (100, "Hundred"))

# currency code -> (major unit name, minor unit name)
CURRENCIES = {
    "MYR": ("Ringgit Malaysia", "Sen"),
    "HKD": ("Hong Kong Dollars", "Cents"),
    "USD": ("US Dollars", "Cents"),
}


class AmountError(ValueError):
    """Raised for amounts that must not reach a document."""


def _under_thousand(n: int) -> str:
    if n < 20:
        return UNITS[n]
    if n < 100:
        tens, rest = divmod(n, 10)
        return TENS[tens] + (f"-{UNITS[rest]}" if rest else "")
    hundreds, rest = divmod(n, 100)
    out = f"{UNITS[hundreds]} Hundred"
    return f"{out} And {_under_thousand(rest)}" if rest else out


def integer_to_words(n: int) -> str:
    """Render a non-negative integer in title-case English."""
    if n < 0:
        raise AmountError("cannot render a negative number in words")
    if n == 0:
        return UNITS[0]

    parts: list[str] = []
    for value, name in SCALES:
        if n >= value:
            count, n = divmod(n, value)
            parts.append(f"{_under_thousand(count) if value < 1000 else integer_to_words(count)} {name}")
    if n:
        # "And" before a trailing sub-hundred remainder reads as it is spoken.
        parts.append(("And " if parts and n < 100 else "") + _under_thousand(n))
    return " ".join(parts)


def _split(amount: Decimal | float | int | str) -> tuple[int, int]:
    """Return ``(major, minor)`` after rounding to two decimals."""
    value = Decimal(str(amount)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    if value < 0:
        raise AmountError(f"amount must not be negative: {amount}")
    major = int(value)
    minor = int((value - major) * 100)
    return major, minor


def words_for_cell(amount: Decimal | float | int | str, currency: str) -> str:
    """Words *without* the major-unit name, for a template that labels it."""
    _require_known(currency)
    major, minor = _split(amount)
    _, minor_name = CURRENCIES[currency.upper()]
    words = integer_to_words(major)
    if minor:
        words = f"{words} And {minor_name} {integer_to_words(minor)}"
    return f"{words} Only"


def amount_in_words(amount: Decimal | float | int | str, currency: str) -> str:
    """The full phrase, e.g. ``Ringgit Malaysia Two Thousand ... Only``."""
    _require_known(currency)
    major_name, _ = CURRENCIES[currency.upper()]
    return f"{major_name} {words_for_cell(amount, currency)}"


def _require_known(currency: str) -> None:
    if currency.upper() not in CURRENCIES:
        raise AmountError(
            f"no wording defined for currency {currency!r}; "
            f"known: {', '.join(sorted(CURRENCIES))}"
        )


if __name__ == "__main__":
    for sample in ("2793.00", "2793.50", "1000000", "0.05", "115.00", "21.00"):
        print(f"{sample:>12}  MYR  {words_for_cell(sample, 'MYR')}")
