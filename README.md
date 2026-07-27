# NEK Group — Back-Office Automation: Project Brief & Handoff

**Purpose of this document.** This is the "start here" for the Claude Code project that will build and maintain NEK Group's finance/admin document automation. It captures every decision made during the design phase so that whoever opens Claude Code — you or a contractor — starts fully oriented instead of re-deriving the context. Read it top to bottom once before writing any code.

**Status at handoff.** Design phase complete. Two artifacts exist and are the foundation for everything below:

- `NEK_Master_Registers.xlsx` — the reference-data layer (single source of truth).
- `NEK_Document_Templates.xlsx` — five standardized, entity-driven document templates.

Nothing is automated yet. The next phase builds the **generation layer**: code that takes structured inputs, fills a template, and produces a correctly-named PDF.

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
│   └── NEK_Document_Templates.xlsx
├── .claude/
│   └── skills/                   # Skills, checked into the repo
│       ├── generate-payment-voucher/
│       │   ├── SKILL.md
│       │   └── generate_pv.py
│       └── ...
├── scripts/                      # shared, reusable Python
│   ├── fill_template.py          # write inputs into a template's yellow anchor cells
│   ├── recalc.py                 # LibreOffice recalculation (see §7)
│   ├── export_pdf.py             # xlsx sheet -> single-page A4 PDF
│   ├── ref_and_filename.py       # build reference string + safe filename
│   └── amount_in_words.py        # number -> "Ringgit Malaysia ... only" / HKD variant
├── counters/
│   └── counters.json             # running sequence per doctype/entity/year
├── output/                       # generated PDFs  (git-ignored)
└── docs/
    ├── house-style.md
    ├── naming-convention.md
    └── decisions-log.md
```

Keep skills in `.claude/skills/` and **check them into the repo** so every future Claude Code session and every contributor inherits the conventions automatically. The repo, not any individual's laptop, is the home of the automation — this is the defence against the single-person "black box" risk.

---

## 4. House style (the fixed standard — already implemented in the templates)

| Item | Standard |
|---|---|
| Font | Arial throughout |
| Currency label | `MYR` (never "RM"/"Ringgit Malaysia" as a label; "Ringgit Malaysia" only in the amount-in-words line) |
| Amount format | Accounting, two decimals: `#,##0.00` |
| Date format | `YYYY-MM-DD` everywhere |
| Letterhead | Entity-driven: full legal name, `Co. Reg. No.`, office address, `Tel: 012-4820853`, `Email: nekgroup84@gmail.com` (no fax) |
| Page setup | A4 portrait, fit to 1 page wide, print area = document body only, margins 0.5" sides / 0.6" top-bottom |
| Footer | "Computer-generated" line on outward-facing + salary docs; Prepared/Issued/Approved/Received signatory block on internal vouchers |

### Document families
- **Outward-facing:** Invoice (`INV`), Official Receipt (`OR`). Branded letterhead, no signatory block.
- **Internal control:** Payment Voucher (`PV`), Receiving Voucher (`RV`). Signatory block.
- **Employee-facing:** Salary Slip (`SAL`). No signatory block.

"Receiving Voucher" is the internal record of money received. "Official Receipt" is the tenant-facing document. The two are never both called "receipt".

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

**The sequence counter (`counters/counters.json`)** must track the last-used `NNN` per `(doctype, entity, year)` and increment on each successful generation. The operator will supply the *current* sequence position for each doc type at cutover (we are mid-year, so counters do not start at 001). Design the counter so it is the single authority — never derive the number by guessing, and guard against gaps and duplicates (auditors notice both).

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

- **Control Panel cells** (per template): Entity Code, Bank Code, Running No., plus a Payroll Date on the Salary Slip. The letterhead and reference derive from these via `INDEX/MATCH` against the `_EntityData` / `_BankAccounts` lookup sheets.
- **Bank–entity mismatch guard:** each template shows a warning if the chosen bank code is not linked to the entity. Replicate this check in the script and **fail** the run on mismatch.

**openpyxl + LibreOffice recalculation.** openpyxl writes formulas as strings with no cached values — until recalculated, formula cells read back as `None`. After filling a template, run it through LibreOffice (`soffice`) to recalculate, then export to PDF. A green recalc proves formulas *evaluate*, not that they are *right*: spot-check a couple of values.

**Formula compatibility.** Stick to Excel-2007-era functions (`INDEX`, `MATCH`, `SUMIFS`, `IFERROR`, `TEXT`). Avoid `XLOOKUP`, `FILTER`, `UNIQUE`, `SORT`, `SEQUENCE` — LibreOffice cannot evaluate them reliably in this pipeline.

**Amount-in-words** must be generated by the script (e.g. a `num2words`-style routine), not typed. Handle currency: MYR → "Ringgit Malaysia … only"; HKD → "Hong Kong Dollars … only" (relevant for NCL/CSK HK documents).

**Python dependencies:** `openpyxl`, `pandas`, a number-to-words library, and LibreOffice for recalc + PDF export.

---

## 8. First Skill to build — `generate-payment-voucher`

Build this first; it establishes the pattern the other four templates reuse.

**Inputs**
- `entity_code` (e.g. `NEK`) — must exist in the Entity register.
- `bank_code` (e.g. `BOC`) — must be linked to `entity_code` in the Bank Account register; else fail.
- `date` (document date, `YYYY-MM-DD`).
- `pay_to` (payee name).
- `tt_cheque` (payment method, e.g. `IBG`).
- `line_items`: list of `{description, amount, account_code}`.
- `signatories`: `{prepared_by, issued_by, approved_by}` (initials).
- Running number handled by the counter (see §5), not passed in by hand.

**Process**
1. Validate entity and (entity, bank) pair. Fail clearly on mismatch.
2. Read + increment the counter for `(PV, entity, year)`; build the reference `PV/ENTITY/BANK/YYYYMM/NNN`.
3. Open the Payment Voucher template; write inputs to the yellow anchor cells and set the Control Panel (entity, bank, running no.).
4. Compute + write amount-in-words.
5. Recalculate via LibreOffice; sanity-check the total against the sum of line items.
6. Export the Payment Voucher sheet to a single-page A4 PDF.
7. Build the safe filename (slashes → hyphens; naming convention) and save to `output/`.
8. Optionally keep the filled `.xlsx` alongside the PDF for the audit trail.

**Outputs:** a correctly-named PDF (and optional xlsx) in `output/`, plus the updated counter.

**Edge cases to handle:** bank not linked to entity (fail); empty/zero/negative amount; multi-line vouchers; long payee names; non-MYR currency for HK entities; a counter file that doesn't yet exist (initialise from the operator-supplied starting number).

**`SKILL.md` structure:** YAML frontmatter (`name`, `description`, `allowed-tools`) + a clear "when to use", the input contract above, the step-by-step process, and the guardrail that this skill produces a *draft for approval* and never executes payment.

---

## 9. Open items to confirm before/while building (do not guess these)

- **Bank Account register:** confirm account names/numbers, WT's bank, and which account each recurring payment actually uses (NEK has both BOC and Maybank). Decide whether full account numbers live in this register at all, or only in the accounting system, per your data-protection review.
- **Reference sequence starting numbers** per doc type per entity, as at cutover date.
- **Compliance deadlines — verify every row with KCK (tax agent) and Exceliz (secretary).** These were entered during design and some may be wrong or outdated; my knowledge of Malaysian/HK statutory dates has a January 2026 cutoff and may be stale. Specific things to check: Form C for the two 31-July-FYE entities (SBOXCAP, SBOXAI) is ~28 February, **not** 31 July; the Annual Return in Malaysia is anniversary-of-incorporation based, not a fixed date; and any current-year (FY2025) filings clustered at end-July may be imminent. **Do not let automation populate statutory dates.**
- **HHIL identifiers:** capture both the Hong Kong Business Registration number (70597712, seen on the existing invoice) and the Malaysian SSM foreign-company number, clearly labelled.
- **HK property ownership:** reconcile whether the Sha Tin unit belongs to CSK or NCL (the property and recurring-payment sheets currently disagree).
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

1. `generate-payment-voucher` — the pattern-setter (this brief, §8).
2. `generate-receiving-voucher` — near-identical (money received; "RECEIVED FM"; RV code).
3. `generate-salary-slips` — batch: read the month's payroll figures → one slip per employee.
4. `generate-official-receipt` / `generate-invoice` — pull the tenant + unit address from the Property register.
5. `monthly-closing-checklist` — generated from the Recurring Payments register.
6. `renewal-and-compliance-alerts` — from the Property + Compliance registers.

Gate each: three consecutive clean monthly runs with no manual correction before you build on top of it. Only move to Tier 2 (extraction, drafting, variance detection) once the Tier 1 set is stable.

---

## 12. First prompt to give Claude Code

> "Read `README.md`, then scaffold the repo structure in §3. Inspect `templates/NEK_Document_Templates.xlsx` — the `Payment Voucher` sheet and the `_EntityData` / `_BankAccounts` lookup sheets — and confirm the input (yellow) cell addresses and Control Panel cells you'll write to. Then draft (a) `scripts/ref_and_filename.py`, (b) `scripts/amount_in_words.py`, and (c) the `generate-payment-voucher` skill per §8, with a dry-run mode that fills the template and exports a PDF without touching the counter. Do not implement anything that executes a payment. Show me the plan before writing code."

---

*End of brief. Keep this document in the repo root and update the decisions log in `docs/` as choices change.*
