# Decisions log

Append-only. Newest last. Each entry: date, decision, why, and what it
supersedes in the brief (if anything).

---

## 2026-07-27 — Recalc/PDF engine is Excel COM, not LibreOffice

**Decision.** `scripts/excel_engine.py` drives Microsoft Excel over COM for
formula recalculation and PDF export.

**Why.** README §7 and §3 assume LibreOffice (`soffice`). LibreOffice is not
installed on the build machine; Microsoft 365 is. Excel recalculates with the
engine the templates were authored against and exports PDF with exact fidelity
to the template print setup (§4: A4 portrait, fit to 1 page wide, print area =
document body). For documents issued to payees and employees, that fidelity is
worth more than LibreOffice's headless portability.

**Supersedes.** The `scripts/recalc.py` + `scripts/export_pdf.py` split in §3,
and the LibreOffice references in §7. **Not superseded:** the §7 rule to stick
to Excel-2007-era functions. Excel would evaluate `XLOOKUP`/`FILTER` happily,
but keeping the restriction preserves the option of moving back to LibreOffice
later, and costs nothing today.

**Known hazard, handled in code.** `ExportAsFixedFormat` lays out pages through
the *active printer's* driver. The build machine's default printer is on a WSD
port and blocks indefinitely when the device sleeps, while Windows still
reports its status as "Normal" — the export never returns. The engine selects a
local printer for the duration of the export and restores the previous setting
afterwards.

---

## 2026-07-31 — Rental documents are filed by unit, not by tenant

`--unit` sets the filename's counterparty field to the unit or property code.
The tenant's name still appears on the document under `RECEIVED FROM :`.

**Why.** Rental income is reviewed by unit rather than by occupant, so the
filing matches the use. It also keeps a private individual's name out of every
filename, which matters more as the archive grows — and it sidesteps
romanisation: the Sha Tin tenant's name is Chinese, and 吴 romanises as "Wu" in
Mandarin but "Ng" in Cantonese. Guessing wrong would put a wrong name on a
legal document.

The unit is validated against `02 Property & Lease`. It fails on an unknown
unit, on a unit belonging to a different entity (which would put the document
beyond reach of that entity's records), and on one whose token identifies
nothing.

**Hong Kong property codes regularised** the same day: `Sha Tin` → `27-STRP`,
`Victoria` → `28-VC`, following the Penang `<unit>-<property>` pattern. They
were descriptions, not codes. The token is the **unit** where it stands alone
(`1G-11-03`) and the **property code** otherwise (`27-STRP`), because the HK
units are bare numbers.

---

## 2026-07-31 — Receiving Voucher wording

`RECEIVED FM :` → `RECEIVED FROM :`, and the table header `IN PAYMENT FOR` →
`BEING PAYMENT FOR`. The Payment Voucher keeps `IN PAYMENT FOR`, which is
correct for money going out.

---

## 2026-07-31 — Currency labels are formula-driven; MYR restriction lifted

`E11` (currency code) and `A20` (amount-in-words label) on the Payment Voucher,
Receiving Voucher, Invoice and Official Receipt now derive from the **bank
account's** currency via a new `K9` control cell.

**The account governs, not the entity.** A voucher pays out of a specific
account. HHIL settles the question: functional currency HKD, BOC account MYR —
testing the entity would have blocked a perfectly ordinary MYR payment.

Wording lives in `CURRENCIES` in `scripts/amount_in_words.py` and is written
into a `_Currency` sheet on every run, the same way the register mirrors are.
Keeping it in two places would eventually print "Hong Kong Dollars : Two
Thousand Ringgit …". Adding a currency now needs no template edit.

**Salary Slip deliberately untouched.** It hard-codes MYR in 17 cells, and
Malaysian payroll with EPF/SOCSO/PCB is MYR by law. Revisit only if an entity
ever runs a foreign payroll.

Verified with a real HKD case (NCL|BOCOM → Cavatina, HK property management):
the document prints `HKD` and `Hong Kong Dollars :`, and the MYR path is
unchanged.

---

## 2026-07-31 — `INV` means the invoice we issue; supplier invoices become `PINV`

`08 Code Lists` had `INV` = "Supplier invoice (payable)" and `SIV` = "Sales
invoice issued (receivable)". README §4 and §5 say the opposite: `INV` belongs
to the outward-facing family alongside `OR`.

**Settled by the template, not by preference.** The `Invoice` sheet reads
`Billed To :` and builds `"INV/"&…` — the outward meaning is already
implemented in the artefact the whole generation layer depends on.

So `INV` = invoice issued by us (counterparty: customer / tenant). The supplier
invoice we *receive* keeps its concept under the new code `PINV`. `SIV` is
retired, having become a duplicate of `INV`.

This matters for filing, not for PV. It had to be settled before
`generate-invoice` is built, since the wrong choice misfiles a year of
documents.

---

## 2026-07-31 — Vendor register seeded with mnemonic codes

17 vendors, derived from the distinct Payable counterparties in
`04 Recurring Payments`.

Codes are **mnemonic** (`KWSP`, `TNB`, `KCK-AUD`), not sequential. The vendor
code becomes the counterparty field of every filename, so `V-001` would make
the document store unsearchable.

Specific decisions:

- **KCK is two vendors.** `KCK-AUD` (KCK & Associates PLT, auditor) and
  `KCK-TAX` (KCK Consultancy Services Sdn Bhd, tax agent) are different legal
  entities providing different services.
- **`TMUNIFI`, not `TM`.** WT's masked counterparty renders as `TM*`; two
  unrelated companies sharing a filename token would be unrecoverable.
- **Payroll excluded.** Salary runs through `SAL`, not a vendor PV.
- **Receivable counterparties excluded.** Tenants are not vendors. They need
  their own home before `generate-official-receipt` is built.
- **Nothing invented.** Registration numbers, contacts and payment terms are
  blank rather than guessed; the bank-detail verification columns stay blank
  for a human, per §2.

**Left unmapped:** WT's two payments go to a counterparty masked as `TM*`, as
its bank is masked as `BN*`. Belongs with the WT open items in §9.

---

## 2026-07-31 — CSK|BOCOM verification recorded

`09 Bank Accounts` row 16: verified 2026-07-31 by Timothy Ng, on his explicit
attestation that the Sha Tin rental is transacted through NCL's Bank of
Communications account.

Note the §2 control was only lightly engaged here: no new account number was
introduced (the row is a re-keyed copy of `NCL|BOCOM`), and the instruction
came directly from the operator rather than from an inbound request that could
have been spoofed.

---

## 2026-07-31 — Excel can hand back a read-only copy and swallow the save

Two runs of the vendor migration reported success and wrote nothing. The
register was open in another Excel window, so `Workbooks.Open` returned a
read-only copy rather than refusing, and with `DisplayAlerts = False` the
subsequent `Save()` was a silent no-op.

`excel_engine.open_workbook()` now raises when it is handed a read-only
workbook it did not ask for. Verified by holding the file open in a second
Excel instance and confirming the guard fires.

**Operational note:** close the register and template before running anything
that writes to them.

---

## 2026-07-30 — Counter starts at 001 from the 2026-09 cutover

Clean break; the pre-cutover manual sequence is not continued. Document dates
before 2026-09 are refused outright, since numbers issued for those months
belong to the old sequence. `counters.initialise()` (`--init-from`) remains
available for any doctype/entity that later needs a different start.

`counters/counters.json` does not exist yet and will not until live issuance is
wired — dry-run reads the next number without creating the file.

---

## 2026-07-30 — Counterparty token rule for filenames

Fixed rule, implemented in `scripts/ref_and_filename.py`:

1. If the payee matches a vendor **code or name** in `05 Vendor`, use the
   vendor code.
2. Otherwise: ASCII-fold, `&` → `AND`, uppercase, non-alphanumerics → spaces,
   collapse to single hyphens, truncate to 40 characters on a word boundary.
3. An empty result is an error, never a blank field.

Rows flagged as example data are excluded from the vendor match. The shipped
register contains exactly one vendor row, marked "EXAMPLE ROW - delete before
use"; matching it would have stamped a placeholder code on a real document.
Until the register is populated, rule 2 applies to every payee.

---

## 2026-07-30 — Payment Voucher restricted to MYR

The template hard-codes `MYR` (E11) and `Ringgit Malaysia :` (A20). Non-MYR
accounts are refused rather than mis-labelled. Per the operator's decision the
next change after PV works is to make both cells formula-driven from the entity
currency; `amount_in_words` already covers HKD and USD.

Note the currency test uses the **bank account's** currency, not the entity's
functional currency. A PV pays out of a specific account, so the account
governs. This matters for HHIL, whose functional currency is HKD while its BOC
account is MYR — testing the entity would have wrongly blocked it.

---

## 2026-07-30 — Dates are written to Excel as serial numbers, not datetimes

Passing a `datetime` through COM let pywin32 convert it via the local timezone,
so midnight on the 1st of a month arrived in the sheet as the last day of the
*previous* month. A voucher dated 2026-09-01 rendered the reference
`PV/NEK/BOC/202608/001` — a correct-looking document carrying the wrong number.

`fill_template.excel_serial()` now writes the raw serial, which carries no
timezone and no locale. The defect was caught by the step-6 verification, not
by inspection, which is the argument for keeping that step.

---

## 2026-07-30 — Printer must be switched before any PageSetup write

Already known for `ExportAsFixedFormat`; it also applies to *setting* a page
header, which round-trips to the active printer's driver and stalls on the WSD
default just as the export does. `export_worksheet()` selects the local printer
before touching `PageSetup`.

---

## 2026-07-27 — Repository folder case aligned to §3

`Templates/` renamed to `templates/` via a two-step `git mv`, since git is
case-sensitive and Windows' filesystem is not. No content change.
