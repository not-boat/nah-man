# HO–Branch Reconciliation Engine v2.0

## What This Does
Inter-company ledger reconciliation between Branch books and HO books.

Upload 3 Tally exports:
1. Branch ledger: "Debit to HO"
2. HO ledger: "Credit to Branch"
3. Suspense ledger (optional but recommended)

Outputs 5 categories:
- ✅ Reconciled (matched, net = 0)
- 🟡 Suspense → Branch (found in suspense, rebook needed)
- 🟠 HO → Suspense (in HO but no branch debit, park in suspense)
- 🔴 Wrong Ledger/Branch (not anywhere, investigate)
- 🟣 Manual Review (low confidence, round amounts)

Exports: Excel report + Tally XML import file for journal adjustments.

---

## Build on Windows

### Requirements
- Windows 10/11 x64
- Node.js 18+ LTS from https://nodejs.org

### Steps
1. Extract this folder to e.g. `C:\ho-recon`
2. Open Command Prompt in that folder:
```
cd C:\ho-recon
npm install
npm run build
```
3. Find installer in `dist\` folder

### Quick run (no build):
```
npm install
npm start
```

---

## Tally Export Steps
1. Gateway of Tally → Display More Reports → Account Books → Ledger
2. Select ledger → set date range → Alt+E → Export → Excel
3. Do this for: Branch "Debit to HO", HO "Credit to Branch", Suspense
4. Upload as-is — no reformatting needed

## Tally XML Import (after export from this tool)
1. Gateway of Tally → Import → XML Data
2. Select the exported Tally_Adjustments_YYYY-MM-DD.xml file
3. Tally will create Journal vouchers for each adjustment
4. Review and accept vouchers

## Important
- Set exact Tally ledger names in the "Tally Ledger Names" config section before exporting XML
- Names must match exactly as they appear in Tally
