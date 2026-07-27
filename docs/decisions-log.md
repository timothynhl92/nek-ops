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

## 2026-07-27 — Repository folder case aligned to §3

`Templates/` renamed to `templates/` via a two-step `git mv`, since git is
case-sensitive and Windows' filesystem is not. No content change.
