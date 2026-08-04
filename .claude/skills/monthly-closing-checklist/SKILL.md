---
name: monthly-closing-checklist
description: >
  Produce the month-end closing checklist from the recurring payment register —
  everything falling due in a given month, ready to work through and tick off.
  Use when preparing for month-end close, or when someone asks what is due this
  month. Issues no document numbers and authorises no payment.
allowed-tools: Read, Bash
---

# Monthly closing checklist

## What this is

A worklist generated from `04 Recurring Payments`: every repeating payment and
receipt falling due in a chosen month, with who prepares it, who approves it,
and which account it moves through.

**It authorises nothing.** Each line still needs its own voucher and its own
approval (README §2). This is a reminder of what to prepare, not an instruction
to pay — and unlike the voucher skills it consumes no document number, so it
can be regenerated as often as you like.

## When to use

Preparing for month-end close, or answering "what's due this month?". Also
useful mid-month to see what is still outstanding.

## Inputs

| Input | Required | Notes |
|---|---|---|
| `--month` | yes | `YYYY-MM`, e.g. `2026-09`. |
| `--entity` | no | Limit to one entity code. Omit for the whole group. |
| `--pdf` | no | Also export a PDF alongside the workbook. |
| `--printer` | no | Use an A4 device if the default substitutes US Letter. |

Output lands in `output/checklists/`, which is git-ignored like all generated
material.

## What it produces

Two sections in one sheet:

1. **Due this month** — split into Payable and Receivable, sorted by day of the
   month, with a `Done` column to tick. Totals are given per currency; an item
   whose amount is recorded as prose rather than a number (there is one,
   "Around 400") is counted separately rather than guessed at.

2. **Timing not recorded** — items that recur but whose due date the register
   does not hold. Grouped by frequency, category and counterparty so the list
   stays short.

## Why section 2 exists

Over half the register's recurring items record their Due Day as `N/A` — 36 of
65 at the time of writing, including every entity's annual audit and tax fees.
They are not unscheduled; the timing was simply never captured.

A checklist that silently omitted them would be worse than useless: it would
look complete while missing every annual obligation the group has. So they are
listed every month, marked clearly, until someone records a due date. **Filling
in the Due Day column in `04 Recurring Payments` moves an item out of that
section and into the proper month** — the checklist gets shorter as the
register gets better.

## What it will not do

- It will not guess a due date. A quarterly item with no anchor month appears
  under "timing not recorded", not in an arbitrary month.
- It will not populate statutory dates (§9 forbids it). Compliance deadlines
  live in `06 Compliance Calendar` and are verified with KCK and Exceliz.

## Example

```bash
python .claude/skills/monthly-closing-checklist/generate_checklist.py \
  --month 2026-09 --pdf
```

For one entity:

```bash
python .claude/skills/monthly-closing-checklist/generate_checklist.py \
  --month 2026-09 --entity NEK
```
