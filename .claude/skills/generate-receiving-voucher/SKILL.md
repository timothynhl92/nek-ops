---
name: generate-receiving-voucher
description: >
  Produce a Receiving Voucher (RV) draft — the internal record of money
  received — from the master register and the standard template, as a PDF for
  human approval. Use when someone asks to raise or record a receipt of funds
  such as rent, a refund or a reimbursement. Runs dry-run only; it does not
  issue document numbers.
allowed-tools: Read, Bash
---

# Generate a Receiving Voucher

## What this is, and what it is not

A Receiving Voucher is the **internal control record of money received**. It is
*not* the tenant-facing document — that is an **Official Receipt** (`OR`).
README §4 is explicit: the two are never both called "receipt". If someone
wants a document to give a tenant, they want an `OR`, not this.

## This skill never moves money

It records a receipt that has already happened and produces **a draft awaiting
approval** (README §2). It does not reconcile against a bank, confirm that
funds cleared, or touch a bank portal.

It is also **dry-run only** at present: it reads the next sequence number
without consuming it, and writes to `output/dryrun/` with a `DRAFT_` prefix.
`--live` deliberately fails.

## When to use

Someone asks to raise, record or draft a receiving voucher — typically rent
received from a tenant, a refund from a supplier, or a reimbursement.

Do **not** use it for money *paid out* (that is a Payment Voucher, `PV`), for a
tenant-facing receipt (`OR`), or for an invoice (`INV`).

## Inputs

| Input | Required | Notes |
|---|---|---|
| `--entity` | yes | Entity code, e.g. `HHIL`. Must exist in `01 Entity`. |
| `--bank` | yes | Bank code, e.g. `BOC`. The account the money was received **into**. Must be linked to the entity, else the run fails. |
| `--date` | yes | Document date, `YYYY-MM-DD`. On or after the 2026-09 cutover. |
| `--received-from` | yes | Payer's name. |
| `--mode` | no | How the money arrived — `IBG`, `Cheque` or `TT`. Defaults to `IBG`. |
| `--line` | yes | Repeatable, `description\|amount[\|account_code]`. One to six. |
| `--issued-by` | no | Defaults to `KN` (Kelvin Ng). |
| `--checked-by` | no | Defaults to `OHY` (Ong Hooi Yong). |
| `--approved-by` | **yes** | **No default** — approval is a deliberate act. |
| `--keep-xlsx` | no | Also save the filled workbook for the audit trail. |
| `--printer` | no | Use an A4 device if the default substitutes US Letter. |

The running number comes from the counter (§5), never passed by hand. `BANKCODE`
in the reference is the account the money was received **into**.

## Process

Identical to the Payment Voucher, and implemented in the same module
(`scripts/voucher.py`) — the two sheets are cell-for-cell identical apart from
four labels and the reference prefix. See
`.claude/skills/generate-payment-voucher/SKILL.md` for the step-by-step, the
currency handling and the printer note.

The reference is `RV/ENTITY/BANK/YYYYMM/NNN`, and `RV` has its own counter
sequence, independent of `PV`.

## Example

```bash
python .claude/skills/generate-receiving-voucher/generate_rv.py \
  --entity HHIL --bank BOC --date 2026-09-01 \
  --received-from "Yan Zhou" \
  --line "Rental 1G-11-03 - Sep 2026|10000.00|4100-01" \
  --approved-by NCL
```

## Note on the counterparty token

Payers are usually **tenants**, who are not in `05 Vendor`. The filename token
therefore falls back to the sanitised name (`Yan Zhou` → `YAN-ZHOU`). That is
correct today, but §9 records an open decision on whether rental documents
should be filed under the **property code** instead — which would also keep
tenants' names out of filenames. Resolve that before `OR`/`INV` are built, and
this skill should follow whatever is decided.
