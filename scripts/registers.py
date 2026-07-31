"""Read the master register and push it into a working copy's mirror sheets.

README §6 makes ``registers/NEK_Master_Registers.xlsx`` the single source of
truth. The templates carry ``_EntityData`` / ``_BankAccounts`` sheets that are
labelled "interim mirror" of it, and nothing keeps the two in step.

Rather than detect that drift and fail, every run *overwrites* the working
copy's mirrors from the master register before filling anything. Master always
wins, so the two cannot disagree at the moment a document is produced.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook

ENTITY_SHEET = "01 Entity"
BANK_SHEET = "09 Bank Accounts"
HEADER_ROW = 4  # data starts on the row below

MIRROR_ENTITY = "_EntityData"
MIRROR_BANK = "_BankAccounts"
MIRROR_FIRST_ROW = 3
# Lookup formulas span rows 3:100; anything past that is invisible to them.
MIRROR_LAST_ROW = 100

# Values the register uses to mean "not applicable". Written through to the
# mirror they would print as a literal "Co. Reg. No. N/A" in the letterhead.
NULL_TOKENS = {"n/a", "na", "-", "--", "tbc", "[to confirm]"}


class RegisterError(RuntimeError):
    """Raised when the register is missing, malformed, or over-full."""


@dataclass(frozen=True)
class Entity:
    code: str
    legal_name: str
    reg_no: str
    address: str
    currency: str
    tel: str
    email: str


@dataclass(frozen=True)
class BankAccount:
    entity_code: str
    bank_code: str
    bank_name: str
    currency: str
    account_name: str
    account_no: str

    @property
    def key(self) -> str:
        """The composite key the templates match on."""
        return f"{self.entity_code}|{self.bank_code}"


def _key(value: object) -> str:
    """Normalise a *key* cell: trim only, never blank.

    Key fields must not go through :func:`_clean`. WT's bank code is literally
    ``TBC``, which is also a null token -- blanking it silently deleted the
    whole account from the register as the generation layer saw it. A row that
    quietly does not exist is far worse than a row that fails a later check.
    """
    return "" if value is None else str(value).strip()


def _clean(value: object) -> str:
    """Normalise a *descriptive* cell, blanking placeholder tokens.

    Only for fields that are printed or ignored -- never for keys.
    """
    if value is None:
        return ""
    text = str(value).strip()
    return "" if text.lower() in NULL_TOKENS else text


def load_registers(register_path: str | Path) -> tuple[dict[str, Entity], dict[str, BankAccount]]:
    """Return ``({entity_code: Entity}, {key: BankAccount})`` from the master file."""
    path = Path(register_path).resolve()
    if not path.is_file():
        raise RegisterError(f"master register not found: {path}")

    wb = load_workbook(path, data_only=True, read_only=True)
    try:
        for required in (ENTITY_SHEET, BANK_SHEET):
            if required not in wb.sheetnames:
                raise RegisterError(f"{path.name} has no {required!r} sheet")

        entities: dict[str, Entity] = {}
        ws = wb[ENTITY_SHEET]
        for row in ws.iter_rows(min_row=HEADER_ROW + 1, values_only=True):
            code = _key(row[0] if row else None)
            if not code:
                continue
            entities[code] = Entity(
                code=code,
                legal_name=_clean(row[1]),
                reg_no=_clean(row[3]),      # D: Registration No.
                address=_clean(row[9]),     # J: Registered Address
                currency=_clean(row[6]),    # G: Functional Currency
                tel=_clean(row[14]),        # O: Tel
                email=_clean(row[15]),      # P: Email
            )

        accounts: dict[str, BankAccount] = {}
        ws = wb[BANK_SHEET]
        for row in ws.iter_rows(min_row=HEADER_ROW + 1, values_only=True):
            bank_code = _key(row[0] if row else None)
            entity_code = _key(row[1]) if row and len(row) > 1 else ""
            if not bank_code or not entity_code:
                continue
            account = BankAccount(
                entity_code=entity_code,
                bank_code=bank_code,
                bank_name=_clean(row[2]),
                currency=_clean(row[6]),
                account_name=_clean(row[4]),
                account_no=_clean(row[5]),
            )
            if account.key in accounts:
                raise RegisterError(
                    f"duplicate bank key {account.key!r} in {BANK_SHEET}; "
                    "the (entity, bank) pair must be unique"
                )
            accounts[account.key] = account
    finally:
        wb.close()

    if not entities:
        raise RegisterError(f"no entities found in {ENTITY_SHEET}")
    if not accounts:
        raise RegisterError(f"no bank accounts found in {BANK_SHEET}")

    _assert_fits(len(entities), "entities")
    _assert_fits(len(accounts), "bank accounts")
    return entities, accounts


def _assert_fits(count: int, what: str) -> None:
    capacity = MIRROR_LAST_ROW - MIRROR_FIRST_ROW + 1
    if count > capacity:
        raise RegisterError(
            f"{count} {what} exceed the {capacity}-row mirror capacity "
            f"(rows {MIRROR_FIRST_ROW}-{MIRROR_LAST_ROW}). Widen the lookup "
            "ranges in every document sheet before adding more."
        )


def sync_mirrors(wb, entities: dict[str, Entity], accounts: dict[str, BankAccount]) -> None:
    """Overwrite an open workbook's mirror sheets from the register.

    ``wb`` is a live Excel COM workbook -- the per-run working copy, never the
    template on disk.
    """
    ws = wb.Worksheets(MIRROR_ENTITY)
    _clear_mirror(ws, columns=7)
    for offset, entity in enumerate(sorted(entities.values(), key=lambda e: e.code)):
        row = MIRROR_FIRST_ROW + offset
        ws.Cells(row, 1).Value = entity.code
        ws.Cells(row, 2).Value = entity.legal_name
        ws.Cells(row, 3).Value = entity.reg_no
        ws.Cells(row, 4).Value = entity.address
        ws.Cells(row, 5).Value = entity.currency
        ws.Cells(row, 6).Value = entity.tel
        ws.Cells(row, 7).Value = entity.email

    ws = wb.Worksheets(MIRROR_BANK)
    _clear_mirror(ws, columns=7)
    for offset, account in enumerate(sorted(accounts.values(), key=lambda a: a.key)):
        row = MIRROR_FIRST_ROW + offset
        ws.Cells(row, 1).Value = account.key
        ws.Cells(row, 2).Value = account.bank_code
        ws.Cells(row, 3).Value = account.entity_code
        ws.Cells(row, 4).Value = account.bank_name
        ws.Cells(row, 5).Value = account.currency
        ws.Cells(row, 6).Value = account.account_name
        ws.Cells(row, 7).Value = account.account_no


def _clear_mirror(ws, columns: int) -> None:
    """Blank the whole mirror band so removed rows cannot linger."""
    ws.Range(
        ws.Cells(MIRROR_FIRST_ROW, 1),
        ws.Cells(MIRROR_LAST_ROW, columns),
    ).ClearContents()
