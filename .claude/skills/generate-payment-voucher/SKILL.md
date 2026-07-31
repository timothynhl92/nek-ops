---
name: generate-payment-voucher
description: >
  Produce a Payment Voucher (PV) draft from the master register and the
  standard template, as a watermarked PDF for human approval. Use when someone
  asks to raise, prepare, or draft a payment voucher for a group entity. Runs
  dry-run only; it does not issue document numbers and never executes payment.
allowed-tools: Read, Bash
---

# Generate a Payment Voucher

## This skill never executes payment

It produces **a draft awaiting approval** and nothing else (README §2). It does
not, and must not be extended to, release funds, schedule a transfer, log into a
bank portal, or alter anyone's bank details. The pipeline stops at "voucher
awaiting approval"; a human reviews and releases payment in the bank portal.

It is also **dry-run only** at present: it reads the next sequence number
without consuming it, and writes to `output/dryrun/` with a `DRAFT - NOT ISSUED`
page header. `--live` deliberately fails. Wiring live issuance is a separate
change requiring sign-off.

## When to use

Someone asks to raise, prepare or draft a payment voucher — typically to pay a
statutory body, supplier or contractor from a group entity's bank account.

Do **not** use it for money *received* (that is a Receiving Voucher, `RV`), for
a tenant-facing receipt (`OR`), or for an invoice (`INV`).

## Inputs

| Input | Required | Notes |
|---|---|---|
| `--entity` | yes | Entity code, e.g. `NEK`. Must exist in `01 Entity`. |
| `--bank` | yes | Bank code, e.g. `BOC`. Must be linked to the entity in `09 Bank Accounts`, else the run fails. |
| `--date` | yes | Document date, `YYYY-MM-DD`. Must be on or after the 2026-09 cutover. |
| `--pay-to` | yes | Payee name. |
| `--tt-cheque` | yes | Payment method, e.g. `IBG`. |
| `--line` | yes | Repeatable, `description\|amount[\|account_code]`. One to six. |
| `--prepared-by` / `--issued-by` / `--approved-by` | yes | Signatory initials. |
| `--keep-xlsx` | no | Also save the filled workbook for the audit trail. |

The running number is **never** passed by hand — it comes from the counter (§5).

`RECEIVED BY` (F23) is intentionally left blank: the payee signs it on receipt.

## Process

1. Load `01 Entity` and `09 Bank Accounts` from the master register.
2. Validate the entity, then the `(entity, bank)` pair — the §7 mismatch guard.
   **This happens before the counter is touched.**
3. Reject non-MYR accounts (see restriction below), bad line items, and dates
   before cutover.
4. Read the next sequence number *without consuming it*; build the reference
   `PV/ENTITY/BANK/YYYYMM/NNN`.
5. Copy the template to a temp working file. In one Excel session: refresh the
   `_EntityData` / `_BankAccounts` mirrors **from the master register**, write
   the yellow input cells, and rebuild all formulas.
6. Verify the rebuilt sheet: the guard cell is empty, the printed reference
   matches the expected one, and the sheet total equals the independently
   computed sum. Abort on any disagreement.
7. Export the `Payment Voucher` sheet to a single-page A4 PDF in
   `output/dryrun/`, named per §5.

Step 6 exists because a clean recalculation proves formulas *evaluate*, not
that they are *right* (§7). It has already caught one real defect.

## Currency

The **bank account's** currency governs, not the entity's — a voucher pays out
of a specific account. HHIL is the case that proves it: functional currency
HKD, but its BOC account is MYR.

`MYR`, `HKD` and `USD` are supported. The currency code (E11) and the
amount-in-words label (A20) both derive from the account, via a `_Currency`
sheet that is rewritten from `scripts/amount_in_words.py` on every run — so the
label the template prints and the words the script writes cannot disagree.

To add a currency, add it to `CURRENCIES` in `amount_in_words.py`. The template
picks it up automatically; no template edit is needed.

## Current restrictions

- **Six line items.** The grid is rows 12–17. A seventh is a hard error, never
  a silent drop.
- **On or after 2026-09.** Earlier dates belong to the previous manual
  sequence.
- **No placeholder bank codes.** An account recorded as `TBC` and similar is
  refused: the reference would be meaningless.

## Example

```bash
python .claude/skills/generate-payment-voucher/generate_pv.py \
  --entity NEK --bank BOC --date 2026-09-01 \
  --pay-to "Kumpulan Wang Simpanan Pekerja" --tt-cheque IBG \
  --line "EPF Payable - Aug 2026|2793.00|5100-01" \
  --prepared-by TN --issued-by KY --approved-by NCL
```
