# O/S Report Maker

## What This Does
Generates the weekly **MYP&SRD outstanding report** from:
1. A Tally **debtors export** (customer ↔ outstanding amount)
2. **Per-customer ledger Excel exports** from Tally (paired-row format)

Replaces the manual workflow of:
- Exporting each ledger
- Running an Excel formula to extract `INV` numbers from narration
- FIFO-picking debit entries (newest backward) until cumulative ≥ outstanding
- Adjusting the last entry by the difference
- Pasting into the O/S report template

…for ~200 customers per week.

## How the FIFO works
For each customer, walk debit entries newest-first (filtered to `date ≤ as-on date`), accumulate amounts until cumulative ≥ outstanding, then **adjust the final (oldest in picked set) row** by `cumulative − outstanding`. The picked rows become the outstanding-invoice list. Each row goes into the matching ageing bucket (`0-30 / 31-60 / 61-90 / 91-120 / 121-Above`) by `as-on − invoice-date`.

This matches the user's manual method exactly. Verified end-to-end against a real DEVA JCB ledger: produces 9 rows summing to ₹16,466,589.48 with the last row adjusted from ₹171,872.90 → ₹156,783.31.

## Outputs
A single Excel workbook with:
- **MYP&SRD** — ready-to-share report in the template format
- **Reconciliation** — per-customer audit (Outstanding · Picked rows · FIFO sum · Variance · Status)
- **Unmatched** — debtors where no ledger file was provided

## Build on Windows

### Requirements
- Windows 10/11 x64
- Node.js 18+ LTS from https://nodejs.org

### Steps
1. Extract this folder to e.g. `C:\os-report`
2. Open Command Prompt:
   ```
   cd C:\os-report
   npm install
   npm run build
   ```
3. Find installer in `dist\` folder.

### Quick run (no build):
```
npm install
npm start
```

## Tally Export Steps

### Debtors export
`Gateway of Tally → Display More Reports → Statements of Accounts → Outstandings → Group → Sundry Debtors → Alt+E → Excel`. Two columns: `Particulars`, `Debit`.

### Per-customer ledgers
`Gateway of Tally → Display More Reports → Account Books → Ledger → <pick customer> → set period → Alt+E → Excel`. Standard paired-row Tally export — no reformatting needed. Either pick files individually in the app or drop a whole folder of them.

## INV extraction
Wider than the original Excel formula. Handles:
- `ref inv.251420019` (your existing pattern)
- `ref inv. 252240227` (stray leading space)
- `ref inv.251423218\r\r\n` (CRLF tails)
- `REF 251610287` (uppercase, no period)
- `Invoice:253400899` (alternative prefix)
- Bare 8+ digit narrations
- `PARTS SALE` / `OILS SALE` rows (no narration ref) → falls back to **Vch No.** column, which IS the invoice number for those rows

## Tests
```
npm test
```
22 tests cover INV extraction, date parsing, ageing buckets, FIFO selection (including edge cases), debtors parsing, ledger parsing, and an end-to-end smoke test against the real DEVA JCB sample.
