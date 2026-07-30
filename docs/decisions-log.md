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
