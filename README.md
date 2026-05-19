# Inventory Reconciliation App

Month-end reconciliation tool: QBO Bills vs TicketVault Purchase Details.

## What it does

**Check 1 — Daily $ Reconciliation**
Compares daily QBO Bill totals per company against TicketVault daily Total Cost,
using the built-in company mapping. Flags any date+company combos where amounts differ.

**Check 2 — Duplicate Bills**
Finds any QBO Bills where the same Company + Bill # (Num) appears more than once
across all uploaded files.

**Check 3 — Description Company Mismatch**
For both Bills and Expenses, extracts the company name in parentheses at the end
of the Description field and verifies it matches the QBO Company column.

## Running locally

```bash
pip install -r requirements.txt
python app.py
# Visit http://localhost:5000
```

## Deploying to Railway

1. Push this folder to a GitHub repo
2. Go to [railway.app](https://railway.app) → New Project → Deploy from GitHub
3. Select the repo — Railway auto-detects the Procfile and deploys
4. Your app will be live at the Railway-provided URL

## Inputs

- **QBO Export(s):** One consolidated report and/or one or more single-company reports (.xlsx)
- **TicketVault Report:** Purchase Details export (.xlsx)

## Output

A single Excel file with:
- **Summary** tab — pass/fail overview of all 3 checks
- **1. Daily Reconciliation** — full date × company comparison with variance
- **2. Duplicate Bills** — all duplicate bill rows highlighted
- **3. Description Mismatches** — all Bills/Expenses where description company ≠ QBO company

## Company Mapping

Defined in `reconciler.py` → `COMPANY_MAPPING`. Update this dict if companies are
added or renamed in QBO or TicketVault.

```python
COMPANY_MAPPING = {
    "Y&S Tickets":        ["YS Tickets", "YS-SeatGeek2"],
    "Damona & Crew":      ["Damon and Crew"],
    "The Ticket Guy LLC": ["The Ticket Guy"],
    # ...
}
```
