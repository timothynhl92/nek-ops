# NEK Group — Back-Office Automation: Project Brief & Handoff

**Purpose of this document.** This is the "start here" for the Claude Code project that will build and maintain NEK Group's finance/admin document automation. It captures every decision made during the design phase so that whoever opens Claude Code — you or a contractor — starts fully oriented instead of re-deriving the context. Read it top to bottom once before writing any code.

**Status at handoff.** Design phase complete. Two artifacts exist and are the foundation for everything below:

- `NEK_Master_Registers.xlsx` — the reference-data layer (single source of truth).
- `NEK_Document_Templates.xlsx` — five standardized, entity-driven document templates.

**Status now (2026-08-05).** Three skills are built:

| Skill | State |
|---|---|
| `generate-payment-voucher` | Works. **Dry-run only** — cannot issue a document. |
| `generate-receiving-voucher` | Works. **Dry-run only.** Shares the PV pipeline. |
| `monthly-closing-checklist` | Works and is **usable now** — consumes no document number, so it is not gated on the counter. |

The two vouchers produce correctly-named, correctly-numbered PDFs from
structured inputs, but read the next sequence number **without consuming it**;
`--live` fails by design. Wiring the counter is a separate, reviewed change
waiting on the operator's confirmed starting numbers.

The recurring register was completed on 2026-08-05: 64 of 65 items now carry a
usable due date, so the closing checklist places every obligation in its proper
month.

Sections below have been corrected where the build diverged from the design.
Every such change is dated and explained in `docs/decisions-log.md` — read that
alongside this document, not instead of it. **Where this brief and the code
disagree, the code is what runs; treat the disagreement as a bug in one of
them and resolve it deliberately.**

---

## 1. Guiding philosophy (do not drift from these)

1. **Registers are the product; AI is just the interface.** The value lives in the structured data (entities, leases, recurring payments, deadlines, bank accounts). Automation reads from it. If the registers aren't maintained, nothing downstream works.
2. **The accounting software remains the system of record.** This project produces drafts, documents, checks and reminders. It never becomes the ledger.
3. **Three tiers, by mechanism — not by department.**
   - **Tier 1 — Deterministic (no AI at runtime).** A template + structured data + a script. AI writes the script *once*; the script then runs forever with no model call. Identical output every time, auditable, free per run, and **no data leaves the premises when it runs.** Target ~60–70% of all automation. *The generation layer is Tier 1.*
   - **Tier 2 — AI-in-the-loop (human verification gate, always).** Invoice/bank-statement extraction, letter drafting, variance detection. Needs a language model, is non-deterministic, will occasionally be confidently wrong. Every output gets human eyes before it has any consequence.
   - **Tier 3 — Agentic orchestration (last, gated hard).** Multi-step chains. Only once the individual Tier 1/Tier 2 pieces are proven. **Never touches payment execution.**
4. **Build once, run cheaply.** Build and maintain the code in Claude Code (git repo). Run deterministic work locally; run AI-in-the-loop work via Cowork. Humans keep the approval gates.
5. **One named owner** maintains the registers, with a quarterly review. Registers with diffuse ownership decay within a quarter.

---

## 2. What must NEVER be automated (permanent guardrails)

- **Payment execution.** The pipeline stops at "draft/voucher awaiting approval." A human reviews and releases payment in the bank portal. Always.
- **Any change to a vendor's or entity's bank details.** Verified only by telephone to a number already held — never a number supplied in the request. Logged with a date.
- **Tax and statutory submission.** Preparation yes; submission no.
- **Final approvals.** Dual approval above a threshold the business sets, unchanged from today.
- **Investment research and decisions.** Out of scope entirely for now.

---

## 3. Proposed repository structure

```
nek-ops/
├── README.md                     # this brief
├── registers/
│   └── NEK_Master_Registers.xlsx # SINGLE SOURCE OF TRUTH — do not fork
├── templates/
│   └── NEK_Document_Templates.xlsx  # 8 sheets: 5 documents + 3 lookup sheets
├── .claude/
│   └── skills/                   # Skills, checked into the repo
│       ├── generate-payment-voucher/
│       │   ├── SKILL.md
│       │   └── generate_pv.py    # thin wrapper over scripts/voucher.py
│       ├── generate-receiving-voucher/
│       │   ├── SKILL.md
│       │   └── generate_rv.py    # ditto
│       └── monthly-closing-checklist/
│           ├── SKILL.md
│           └── generate_checklist.py
├── scripts/                      # shared, reusable Python
│   ├── voucher.py                # the shared PV/RV pipeline
│   ├── recurring.py              # read 04 Recurring Payments; work out timing
│   ├── due_date_worklist.py      # ask the operator for missing due dates
│   ├── excel_engine.py           # Excel COM: recalculation + PDF export (see §7)
│   ├── registers.py              # read the register; refresh the template mirrors
│   ├── counters.py               # the sequence authority (see §5)
│   ├── fill_template.py          # write inputs into a template's yellow anchor cells
│   ├── ref_and_filename.py       # build reference string + safe filename
│   ├── amount_in_words.py        # number -> words; also the currency wording table
│   └── audit_registers.py        # standing consistency check over the register
├── counters/
│   └── counters.json             # running sequence per doctype/entity/year
├── output/                       # generated PDFs  (git-ignored)
│   ├── dryrun/                   # watermark-free drafts, DRAFT_ prefixed
│   ├── checklists/               # monthly closing checklists
│   └── worklists/                # register-gap forms for the operator
└── docs/
    ├── house-style.md
    ├── naming-convention.md
    ├── decisions-log.md          # dated record of every divergence from this brief
    └── archive/                  # superseded artefacts, kept not deleted
```

`templates/` also holds five single-sheet stubs (`Invoice Template.xlsx` and
friends). They are **empty shells with no formulas** — never generate from
them. The only live template is `NEK_Document_Templates.xlsx`.

**Run `python scripts/audit_registers.py` after any hand-edit to the register.**
It checks referential integrity, duplicate keys, filename safety of every
vendor code, and lists which fields are still placeholders.

Keep skills in `.claude/skills/` and **check them into the repo** so every future Claude Code session and every contributor inherits the conventions automatically. The repo, not any individual's laptop, is the home of the automation — this is the defence against the single-person "black box" risk.

---

## 4. House style (the fixed standard — already implemented in the templates)

| Item | Standard |
|---|---|
| Font | Arial throughout |
| Currency label | `MYR` (never "RM"/"Ringgit Malaysia" as a label; "Ringgit Malaysia" only in the amount-in-words line). Both the code and the words label are **derived from the bank account's currency**, not typed |
| Minor unit | `Cents` for every currency, MYR included — the documents are English throughout |
| Amount format | Accounting, two decimals: `#,##0.00` |
| Date format | `YYYY-MM-DD` everywhere |
| Letterhead | Entity-driven: full legal name, `Co. Reg. No.`, office address, `Tel: 012-4820853`, `Email: nekgroup84@gmail.com` (no fax) |
| Page setup | A4 portrait, fit to 1 page **wide only** (never fit-to-height — it scales the sheet down until hairline rules drop out in print), print area = document body only, margins 0.35" sides / 0.3" top-bottom |
| Page usage | Vouchers occupy the **top half** of the A4 sheet, which is then guillotined in two. `scripts/fill_template.py` reports when content would overrun that half |
| Footer | "Computer-generated" line on outward-facing + salary docs. Internal vouchers carry **no footer line** |
| Signatories | Internal vouchers: **Issued / Checked / Approved**, each with an initials field and a printed rule signed by hand after printing. Issuer defaults to `KN`, checker to `OHY`; **the approver never defaults** |

### Document families
- **Outward-facing:** Invoice (`INV`), Official Receipt (`OR`). Branded letterhead, no signatory block.
- **Internal control:** Payment Voucher (`PV`), Receiving Voucher (`RV`). Signatory block.
- **Employee-facing:** Salary Slip (`SAL`). No signatory block.

"Receiving Voucher" is the internal record of money received. "Official Receipt" is the tenant-facing document. The two are never both called "receipt".

The two vouchers carry the same body wording except where direction demands
otherwise: the PV reads `PAY TO :` / `IN PAYMENT FOR`, the RV reads
`RECEIVED FROM :` / `BEING PAYMENT FOR`. Their sheets are otherwise
cell-for-cell identical, which is why one module drives both.

**`INV` is the invoice we issue** (counterparty: customer / tenant). A supplier
invoice we *receive* is `PINV`. `08 Code Lists` originally had these the other
way round; the template settles it — the `Invoice` sheet reads `Billed To :`.

---

## 5. Reference number & filename rules (important, easy to get wrong)

**On-document reference (printed on the document):**

```
DOCTYPE / ENTITY / BANKCODE / YYYYMM / NNN
Example:  PV/NEK/BOC/202607/014
```

- `BANKCODE` = the account the money is paid **from** (PV) or **into** (OR, INV). Codes live in the Bank Account register.
- `NNN` = sequential **per document type, per entity**, reset to `001` at the start of each calendar year. `YYYYMM` records the issue month; the counter itself is annual, not monthly.
- Doc type codes: `PV`, `RV`, `OR`, `INV`, `SAL`.

**Filename (when saved to disk) — MUST convert slashes:** a filename cannot contain `/` (it is a path separator and breaks file systems and cloud sync). So the generation script converts slashes to hyphens and wraps the reference inside the naming convention:

```
YYYY-MM-DD_ENTITY_DOCTYPE_COUNTERPARTY_REFERENCE.ext
Example:  2026-07-25_NEK_PV_KWSP_PV-NEK-BOC-202607-014.pdf
```

Filename rules: underscores separate fields, hyphens join words *within* a field, no spaces, no `/ \ & % # $ ( ) , ' "`, ISO date is the **document** date, total path < 200 characters.

**The COUNTERPARTY token** (fixed rule, in `scripts/ref_and_filename.py`):

1. If the payee matches a **vendor code or vendor name** in `05 Vendor`, use
   that vendor code. Codes are **mnemonic** (`KWSP`, `TNB`, `KCK-AUD`), never
   sequential — `V-001` in a filename makes the document store unsearchable.
2. Otherwise: ASCII-fold, `&` → `AND`, uppercase, non-alphanumerics → spaces,
   collapse to single hyphens, truncate to 40 characters on a word boundary.
3. An empty result is an **error**, never a blank field.

Vendor rows flagged as example data are excluded from the match.

**Rental documents are filed by unit, not by tenant** (decided 2026-07-31).
Pass `--unit`; the counterparty field becomes the unit or property code, while
the tenant's name still appears on the document itself. Rental income is
reviewed by unit rather than by occupant, and it keeps a private individual's
name out of every filename.

```
2026-09-01_HHIL_RV_1G-11-03_RV-HHIL-BOC-202609-001.pdf
```

The unit is validated against `02 Property & Lease` — an unknown unit, a unit
belonging to a different entity, or one whose token would identify nothing all
fail the run. The token is the **unit** where the unit stands alone
(`1G-11-03`), otherwise the **property code** (`27-STRP`), because the Hong
Kong units are bare numbers.

**The sequence counter (`counters/counters.json`)** must track the last-used `NNN` per `(doctype, entity, year)` and increment on each successful generation. Design the counter so it is the single authority — never derive the number by guessing, and guard against gaps and duplicates (auditors notice both).

> **Superseded 2026-07-31.** Clean break rather than continuing the old
> numbering: every `(doctype, entity)` starts at **001 from the 2026-09
> cutover**, and document dates before that month are **refused** — those
> numbers belong to the previous manual sequence. `--init-from` remains for any
> doctype/entity that later needs a different start.

`counters/counters.json` **does not exist yet** and will not until live issuance
is wired. Dry-run reads the next number without creating the file.

---

## 6. The registers (single source of truth)

`NEK_Master_Registers.xlsx` sheets:

- `00 README` — rules for maintaining the file.
- `01 Entity` — every entity, keyed by **Entity Code** (the master key). Now includes Tel/Email columns.
- `02 Property & Lease` — 6 units (4 Penang "Quayside" under HHIL; 2 Hong Kong under NCL/CSK). Drives renewal alerts.
- `03 Employee` — reference data only. **Deliberately excludes** IC/passport numbers, bank accounts, medical data.
- `04 Recurring Payments` — every repeating payment in/out; the basis for the monthly closing checklist.
- `05 Vendor` — to be populated later (fraud-control sheet; bank-detail verification columns).
- `06 Compliance Calendar` — statutory/audit/tax deadlines per entity.
- `07 Naming Convention` — the filename standard and folder structure.
- `08 Code Lists` — controlled vocabularies (doc type codes, statuses).
- `09 Bank Accounts` — one row per account; the source for bank codes. Seeded from recurring-payments data; account names/numbers are placeholders to confirm.

### Entities (11)

| Code | Legal name | Jurisdiction | FYE | Notes |
|---|---|---|---|---|
| NEK | Ng Eng Kee & Sons Sdn Bhd | Malaysia | 31 Dec | Parent; banks at BOC **and** Maybank |
| NEKGV | NEK Global Venture Sdn Bhd | Malaysia | 31 Dec | |
| MDL | Marque De Luxe Sdn Bhd | Malaysia | 31 Dec | |
| HHIL | Honour Harvest International Limited | Hong Kong | 31 Dec | HK company, SSM-registered as a foreign company; owns the Penang rental units and files Malaysian tax |
| NGCS | NGCS Capital Sdn Bhd | Malaysia | 31 Dec | |
| NCLCAP | NCL Capital Sdn Bhd | Malaysia | 31 Dec | |
| SBOXCAP | Sandbox Capital Sdn Bhd | Malaysia | **31 Jul** | Non-Dec FYE — see §9 |
| SBOXAI | Sandbox AI Ventures Sdn Bhd | Malaysia | **31 Jul** | Non-Dec FYE — see §9 |
| NCL | Ng Chong Lam | Malaysia (individual) | — | Holds HK property |
| CSK | Chiang Sau Kuen | Hong Kong (individual) | — | Holds HK property |
| WT | Wintrace Investment | **BVI** | — | Offshore; annual return + economic-substance filing |

---

## 7. Technical notes for the generation layer

**The templates are two-layer by design.** A small set of yellow **input cells** plus a **Control Panel** (columns J:K, outside the print area) hold everything variable. Everything else — letterhead, reference number, totals — is a formula. The script writes only to the input cells; it never touches the formatted layout.

- **Control Panel cells** (per template): Entity Code (`K3`), Bank Code (`K4`), Running No. (`K5`), plus a Payroll Date on the Salary Slip. Derived below them: bank name (`K6`), reference (`K7`), mismatch guard (`K8`), currency (`K9`). The letterhead and reference derive from these via `INDEX/MATCH` against the `_EntityData` / `_BankAccounts` lookup sheets.
- **Lookup ranges span rows 3:100.** They were originally `3:13` — exactly the number of rows present — so a 12th entity would have fallen outside every `INDEX/MATCH` and been swallowed by `IFERROR`, blanking the letterhead silently. Keep the headroom.
- **Bank–entity mismatch guard:** each template shows a warning if the chosen bank code is not linked to the entity. The template's guard is display-only; the script performs its own check and **fails** the run, *before* the counter is touched.

**Recalculation and PDF export use Microsoft Excel over COM, not LibreOffice.**
LibreOffice is not installed on the build machine and Microsoft 365 is, so
`scripts/excel_engine.py` drives Excel directly. It recalculates with the
engine the templates were authored against and exports PDF with exact fidelity
to the print setup. openpyxl is still used for *reading*, since it writes
formulas with no cached values — until recalculated, formula cells read back as
`None`. A green recalc proves formulas *evaluate*, not that they are *right*:
the generator re-checks the printed reference, the mismatch guard and the total
against independently computed values before it exports anything. That check
has caught two real defects that inspection did not.

**Four hazards, all handled in code — do not undo them:**

1. **Never pass a `datetime` to a cell through COM.** pywin32 converts via the
   local timezone, so midnight on the 1st arrives as the previous month under
   a positive UTC offset — a correct-looking voucher carrying the wrong
   reference. Use `fill_template.excel_serial()`.
2. **Switch the printer before any `PageSetup` write or export.** Both
   round-trip to the active printer's driver, and a WSD-port printer that is
   asleep blocks indefinitely while Windows still reports it "Normal".
3. **Excel silently substitutes US Letter for A4** on the Microsoft virtual
   printers, while still reporting A4 to the application. Every run measures
   the finished PDF and warns. Set "Microsoft Print to PDF" to A4 in Windows,
   or pass `--printer` with a real A4 device.
4. **`Borders(edge)` is unreliable on a *merged* range.** Use
   `Range.BorderAround()`, or boxes silently lose edges.

**Formula compatibility.** Stick to Excel-2007-era functions (`INDEX`, `MATCH`,
`SUMIFS`, `IFERROR`, `TEXT`). Excel would evaluate `XLOOKUP`/`FILTER` happily,
but keeping the restriction preserves the option of moving back to LibreOffice
and costs nothing today.

**Amount-in-words** is generated by `scripts/amount_in_words.py`, never typed.
That module is also the **single source for currency wording**: it is written
into the template's `_Currency` sheet on every run, so the label the template
prints and the words the script writes cannot drift apart. Adding a currency is
a one-line change there and needs no template edit.

**Single-source enforcement.** Every run *overwrites* the template's
`_EntityData`, `_BankAccounts` and `_Currency` sheets from the master register
before filling anything. Master always wins, so the two cannot disagree at the
moment a document is produced. **Never hand-edit those sheets** — edit the
register.

**Python dependencies:** `openpyxl`, `pywin32`. Amount-in-words is implemented
in-repo rather than pulled from a library, so the deterministic layer adds no
third-party dependency and no data leaves the machine (§1, Tier 1).

---

## 8. First Skill to build — `generate-payment-voucher`

Build this first; it establishes the pattern the other four templates reuse.

**Built and working in dry-run.** See `.claude/skills/generate-payment-voucher/SKILL.md`
for the live contract; this section records the design intent.

**Inputs**
- `entity_code` (e.g. `NEK`) — must exist in the Entity register.
- `bank_code` (e.g. `BOC`) — must be linked to `entity_code` in the Bank Account register; else fail.
- `date` (document date, `YYYY-MM-DD`) — must be on or after the 2026-09 cutover.
- `pay_to` (payee name).
- `mode_of_payment` — `IBG`, `Cheque` or `TT`. **Defaults to `IBG`.**
- `line_items`: list of `{description, amount, account_code}`, one to six.
- `signatories`: `{issued_by, checked_by, approved_by}` (initials). Issuer
  defaults to `KN`, checker to `OHY`. **The approver has no default and is
  required** — an approval nobody chose is not an approval (§2).
- Running number handled by the counter (see §5), not passed in by hand.

**Process**
1. Validate entity, then the (entity, bank) pair, then the currency, the line
   items and the date. **Everything that can fail runs before the counter is
   touched and before Excel is launched.**
2. Read the counter for `(PV, entity, year)`; build the reference
   `PV/ENTITY/BANK/YYYYMM/NNN`. *Dry-run reads without consuming.*
3. Copy the template to a temp working file. In **one** Excel session: refresh
   the mirror sheets from the register, write the yellow anchor cells, set the
   Control Panel. (One session, not five — launching Excel costs ~50s.)
4. Compute + write amount-in-words.
5. Recalculate, then verify: guard cell empty, printed reference matches, sheet
   total equals the independently computed sum. Abort on any disagreement.
6. Export the Payment Voucher sheet to a single-page PDF, top half of A4.
7. Build the safe filename (slashes → hyphens; naming convention) and save.
8. Optionally keep the filled `.xlsx` alongside the PDF for the audit trail.

**Outputs:** a correctly-named PDF (and optional xlsx). In dry-run these land in
`output/dryrun/` with a `DRAFT_` prefix and the counter is untouched.

**Edge cases handled:** bank not linked to entity; placeholder bank code (`TBC`);
a currency with no wording defined; empty/zero/negative amount; more than six
line items (hard error — silently dropping rows would understate the total);
long payee names; a document date before the cutover; a counter file that does
not yet exist.

**`SKILL.md` structure:** YAML frontmatter (`name`, `description`, `allowed-tools`) + a clear "when to use", the input contract above, the step-by-step process, and the guardrail that this skill produces a *draft for approval* and never executes payment.

---

## 9. Open items to confirm before/while building (do not guess these)

- **Bank Account register:** confirm account names/numbers, WT's bank, and which account each recurring payment actually uses (NEK has both BOC and Maybank). Decide whether full account numbers live in this register at all, or only in the accounting system, per your data-protection review.
- **Reference sequence starting numbers** per doc type per entity, as at cutover date.
- **Compliance deadlines — verify every row with KCK (tax agent) and Exceliz (secretary).** These were entered during design and some may be wrong or outdated; my knowledge of Malaysian/HK statutory dates has a January 2026 cutoff and may be stale. Specific things to check: Form C for the two 31-July-FYE entities (SBOXCAP, SBOXAI) is ~28 February, **not** 31 July; the Annual Return in Malaysia is anniversary-of-incorporation based, not a fixed date; and any current-year (FY2025) filings clustered at end-July may be imminent. **Do not let automation populate statutory dates.**
- **HHIL identifiers:** capture both the Hong Kong Business Registration number (70597712, seen on the existing invoice) and the Malaysian SSM foreign-company number, clearly labelled.
- ~~**HK property ownership:** reconcile whether the Sha Tin unit belongs to CSK or NCL.~~ **Resolved 2026-07-31:** standardised to **NCL** throughout. CSK remains in the entity register as a director of NCLCAP, marked `Dormant`, with no bank account of its own.
- **Chart of accounts.** The `ACCOUNTS CODE` column is free text pending the accountant. `08 Code Lists` holds *document-type* codes and controlled vocabularies — it has no chart of accounts, so one needs a home (a new `10 Chart of Accounts` sheet is the natural fit).
- **Sha Tin tenant name.** `02 Property & Lease` and `04 Recurring Payments` both record the tenant as "Chinese national" — a description, not a name. The real name is a Chinese one, which renders correctly in the document body (verified) but cannot form a filename token. **The filename half is resolved** — rental documents file by unit (§5) — so what remains is simply recording the tenant's actual name in the register.
- **Hong Kong units are stored as numbers.** `02 Property & Lease` holds units `27` and `28` as numeric cells, so Excel returns them as `27.0`. Handled in `registers._key()`, but the register would be tidier with them as text.
- **Vendor detail.** 17 vendors are seeded from the recurring payments, but registration numbers, contacts and payment terms are blank rather than guessed. `scripts/audit_registers.py` lists exactly what is outstanding.
- ~~**Recurring payment due dates — 36 of 65 rows record `N/A`.**~~ **Resolved 2026-08-05.** 64 of 65 items now carry a usable due date. The remaining one is HHIL's secretarial fee, recorded at **RM 0** — as a foreign company it has no fee to pay, so it needs no due date and is reported under "recorded, no payment due" rather than as a gap.
- **The completeness report over-reports.** It flags `N/A` in `01 Entity` for NCL, CSK and WT — natural persons and a BVI entity, for which no registration number or financial year end is *correct*, not missing. Same for the Malaysian quit-rent and assessment columns on the Hong Kong property. Roughly 19 of the 21 entity "gaps" and all 3 property ones are legitimate. Worth teaching the audit to tell "not applicable" from "not yet filled", or it trains the reader to ignore it.
- **Salary slip statutory figures** (EPF/SOCSO/EIS/PCB) must come from the actual payroll computation each month, never be re-keyed or defaulted.

---

## 10. Data protection & privacy

- Entities span **Malaysia (PDPA)** and **Hong Kong (PDPO)**. Confirm obligations with legal counsel before routing any personal data through third-party/AI tools.
- The registers deliberately exclude IC/passport numbers, bank accounts and medical data — keep it that way so the file carries lighter obligations.
- **Deterministic generation runs locally with no data egress** — a real privacy advantage of building this as Tier 1 code.
- Any Tier 2 (AI-in-the-loop) step sends data to Anthropic; review current data-handling terms first. Use a dedicated `AI-Staging` folder as the access-control boundary — point AI tools only at that folder, never the whole drive; clear it on a schedule.
- e-Invoicing (LHDN MyInvois) and SST: confirmed **not currently applicable** to any entity. Re-check if turnover or rules change.

---

## 11. Recommended build order (advance on proof, not calendar)

1. ~~`generate-payment-voucher`~~ — **built, dry-run only.** The pattern-setter (§8).
2. ~~`generate-receiving-voucher`~~ — **built, dry-run only.** Money received; "RECEIVED FM"; `RV` code. Shares `scripts/voucher.py` with the Payment Voucher — the two sheets are cell-for-cell identical apart from four labels and the reference prefix, so each skill is a thin wrapper naming its document type.
3. `generate-salary-slips` — batch: read the month's payroll figures → one slip per employee.
4. `generate-official-receipt` / `generate-invoice` — pull the tenant + unit address from the Property register.
5. ~~`monthly-closing-checklist`~~ — **built and usable now.** Generated from the Recurring Payments register. Consumes no document number, so unlike the vouchers it is not gated on the counter.
6. `renewal-and-compliance-alerts` — from the Property + Compliance registers.

Gate each: three consecutive clean monthly runs with no manual correction before you build on top of it. Only move to Tier 2 (extraction, drafting, variance detection) once the Tier 1 set is stable.

**What the gate does and does not block.** It blocks building *on top of* a
skill — above all Tier 2, where an unproven foundation compounds. It does not
block building a *sibling* Tier 1 skill that merely replicates a proven pattern.
Receiving Voucher can be built while Payment Voucher accumulates its three
clean runs; extraction and drafting cannot.

---

## 12. Where to pick up

Start every session with:

> Read `README.md` and `docs/decisions-log.md`, then run
> `python scripts/audit_registers.py` to see the register's current state.

### Working commands

```bash
# Payment voucher (dry run). Only the approver must be stated.
python .claude/skills/generate-payment-voucher/generate_pv.py \
  --entity NEK --bank BOC --date 2026-09-01 --pay-to "KWSP (EPF)" \
  --line "EPF Payable - Aug 2026|2949.00|5100-01" --approved-by NCL

# Receiving voucher. Rental receipts file under the unit.
python .claude/skills/generate-receiving-voucher/generate_rv.py \
  --entity HHIL --bank BOC --date 2026-09-01 --received-from "Yan Zhou" \
  --unit 1G-11-03 --line "Rental 1G-11-03 - Sep 2026|10000.00|4100-01" \
  --approved-by NCL

# Monthly closing checklist. Usable now; no counter involved.
python .claude/skills/monthly-closing-checklist/generate_checklist.py \
  --month 2026-09 --pdf
```

**Add `--printer "Brother DCP-L2550DW series"` to anything producing a PDF**
until "Microsoft Print to PDF" is set to A4 in Windows — the Microsoft virtual
printers substitute US Letter while still reporting A4 (§7).

### Next: `generate-official-receipt` and `generate-invoice` (§11 item 4)

Everything they were blocked on is settled — the counterparty rule (§5, file by
unit), the tenant's name, and the Hong Kong property codes. Six rental receipts
a month flow through them, making them the highest-volume outward documents.

What remains is a **layout review**. The `Invoice` and `Official Receipt`
sheets have had the letterhead, paper-size and currency passes, but **not** the
border, signatory or spacing work the vouchers went through. Expect one or two
rounds of printed review; the Payment Voucher took four before the pattern was
established, and those lessons are in §7 and the decisions log.

Note both are outward-facing (§4): branded letterhead, **no signatory block**,
and a "Computer-generated" footer — unlike the vouchers.

### Blocked, and on what

- **Live issuance** — needs the operator's confirmed starting numbers from the
  accountant, and a deliberate decision to begin consuming real document
  numbers. **All three voucher-type skills are gated on this.** Note that with
  the on-page draft watermark removed, a dry run and an issued voucher are
  indistinguishable once printed; decide how drafts are marked before wiring.
- **Chart of accounts** — with the accountant. `ACCOUNTS CODE` is free text
  until it exists, and `08 Code Lists` has no home for it (§9).
- **Salary slips** — need the payroll figures, which §9 says must never be
  defaulted or re-keyed. The `Salary Slip` sheet also still hard-codes MYR in
  seventeen cells (deliberately — Malaysian payroll is MYR by law).

**Do not** implement anything that executes, schedules or releases a payment.
Show the plan before writing code.

---

*End of brief. Keep this document in the repo root and update the decisions log in `docs/` as choices change.*
