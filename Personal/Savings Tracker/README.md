# Savings Tracker

Personal finance Excel workbooks for tracking spending, budgeting, savings balances, and interest growth — all formula-driven with no macros required.

## Workbooks

| File | Purpose |
| --- | --- |
| `Overall Money Tracker.xlsx` | Account-level balance tracking with compound interest estimation |
| `Personal Finance Tracker.xlsx` | Transaction-level budgeting, cash flow analysis, and savings health |

Both files are standard `.xlsx` format — no macros, no VBA, no add-ins. Every calculation is powered by native Excel formulas that update automatically.

## Overall Money Tracker

### Worksheet structure

| Sheet | Purpose |
| --- | --- |
| `Accounts` | Master list of bank accounts, e-wallets, and cash with interest calculations |
| `Balance History` | Manual snapshots tracking total money over time with growth analysis |
| `README` | In-workbook user guide |

### Accounts — fields and formulas

| Column | Type | Description |
| --- | --- | --- |
| Account | Input | Account name (e.g. BDO Savings, GCash) |
| Type | Input | `Bank`, `E-Wallet`, or `Cash / Other` |
| Bank / Provider | Input | Institution name |
| Recorded Balance | Input | Last known balance |
| Interest Rate | Input | Rate as a percentage (e.g. 3.50%) |
| Rate Basis | Input | `Annual` or `Monthly` |
| Est. Monthly Interest | Formula | Converts rate to monthly equivalent |
| Est. Annual Interest | Formula | Converts rate to annual equivalent |
| Accrued Interest | Formula | Compound growth from Last Updated → TODAY() |
| Calculated Balance Today | Formula | Recorded Balance + Accrued Interest |
| Last Updated | Input | Date the recorded balance was accurate |
| Include in Total | Input | `Yes` or `No` — controls summary aggregation |
| Status | Formula | `✅ Current` or `⚠️ Stale` (>30 days since update) |

### Interest calculation logic

- **Annual basis**: Rate is an effective annual rate. Monthly estimate = `Balance × ((1 + Rate)^(1/12) − 1)`. Accrued = `Balance × ((1 + Rate)^(days/365) − 1)`.
- **Monthly basis**: Rate is an effective monthly rate. Annual estimate = `Balance × ((1 + Rate)^12 − 1)`. Accrued = `Balance × ((1 + Rate)^(days/(365/12)) − 1)`.
- Set `0%` for non-interest-bearing accounts.
- Interest is a pre-tax estimate and does not account for balance tiers or promotional conditions.

### Summary dashboard (Accounts sheet)

| Metric | Description |
| --- | --- |
| Recorded Balance Included | Sum of balances for included accounts |
| Accrued Interest | Total accrued across included accounts |
| Calculated Total | Balance + accrued for included accounts |
| Est. Monthly Interest | Combined monthly interest |
| Est. Annual Interest | Combined annual interest |
| Bank / E-Wallet / Cash Totals | Breakdown by account type |
| Active Accounts | Count of included accounts |
| Top Earning Account | Account with highest annual interest |
| Stale Accounts | Count of accounts not updated in >30 days |

### Balance History

Manual snapshot table with automatic calculations:

| Column | Type | Description |
| --- | --- | --- |
| Snapshot Date | Input | Date of the checkpoint |
| Bank / E-Wallet / Cash Totals | Input | Category balances at snapshot time |
| Total Money | Formula | Sum of the three categories |
| Change | Formula | Difference from previous snapshot |
| Change (%) | Formula | Percentage change from previous snapshot |
| Ann. Growth (%) | Formula | Annualized return from first snapshot |

**History Insights** (below the table): Average Monthly Change, Total Growth, Total Growth %, Snapshots Recorded, Best Month, Worst Month.

---

## Personal Finance Tracker

### Worksheet structure

| Sheet | Purpose |
| --- | --- |
| `Dashboard` | KPIs, budget vs actual, cash flow, emergency fund, savings goal, top spending |
| `Transactions` | Raw transaction log (date, description, category, type, amount, account) |
| `Settings` | Categories & budgets, accounts list, period engine, interest settings, savings goal |
| `README` | In-workbook user guide |

### Dashboard KPIs

| KPI | Description |
| --- | --- |
| 💵 Income | Total income for the selected period |
| 💸 Expenses | Total expenses for the selected period |
| 🏦 Net Savings | Income − Expenses |
| 📈 Savings Rate | Net Savings / Income |
| Savings Balance | Total calculated savings balance with accrued interest |

### Budget vs Actual

Per-category breakdown showing Monthly Budget, Period Budget, Actual spend, Remaining, % Used, and a traffic-light Status (🟢 On track / 🟡 Getting close / 🔴 Over budget).

### Cash Flow — Last 12 Months

Rolling 12-month view showing Income, Expenses, and Net per month.

### Emergency Fund & Savings Health

| Metric | Description |
| --- | --- |
| Savings Balance (w/ Interest) | Pulled from Settings interest calculations |
| Avg Monthly Expenses | Period expenses ÷ months in period |
| Emergency Fund Coverage | Savings ÷ Avg Monthly Expenses (in months) |
| Coverage Status | 🟢 6+ months / 🟡 3–6 months / 🔴 <3 months |
| Est. Monthly Interest Earned | Sum of monthly interest from all savings accounts |
| Est. Annual Interest Earned | Sum of annual interest from all savings accounts |

### Savings Goal Progress

| Metric | Description |
| --- | --- |
| Target Amount | User-set savings goal |
| Current Balance | Calculated savings balance |
| Progress | Percentage of goal reached |
| Target Date | User-set deadline |
| Months Remaining | Months until target date |
| Monthly Savings Needed | Gap ÷ months remaining |
| Status | 🔴 → 🟠 → 🟡 → 🟢 → 🏆 milestone indicators |

### Top 5 Spending Categories

Ranked table of the five highest-spend expense categories for the selected period, showing amount and percentage of total expenses.

### Interest & Savings Settings (Settings sheet)

Replicates the compound interest engine from the Overall Money Tracker:

| Field | Type | Description |
| --- | --- | --- |
| Account | Input | Savings account name |
| Recorded Balance | Input | Current balance |
| Interest Rate | Input | Rate as percentage |
| Rate Basis | Input | `Annual` or `Monthly` |
| Last Updated | Input | Date balance was last verified |
| Est. Monthly Int. | Formula | Monthly interest estimate |
| Est. Annual Int. | Formula | Annual interest estimate |
| Accrued Interest | Formula | Compound growth to today |
| Calc. Balance Today | Formula | Balance + accrued |

---

## Quick start

1. **Overall Money Tracker**: Open `Accounts`, replace sample accounts with your real balances and interest rates. Set `Include in Total` to `Yes`. Add snapshots to `Balance History` periodically.
2. **Personal Finance Tracker**: Log transactions in the `Transactions` sheet. Set your budgets in `Settings`. Configure your savings accounts in the Interest & Savings section. Set a savings goal. The Dashboard does the rest.

## Color legend

| Style | Meaning |
| --- | --- |
| Yellow cells with blue text | User-editable inputs |
| Light green cells | Computed formulas — do not overwrite |
| Blue header cells | Section/column headers |
| Status emojis (🟢 🟡 🔴) | Traffic-light indicators |

## Maintenance

- This README and root [AI Documentation Notes.md](../../AI%20Documentation%20Notes.md) are the technical references.
- Environment setup is in [Tech Stack Setup Guide.md](../../Tech%20Stack%20Setup%20Guide.md).
- Both workbooks include in-file `README` sheets for end-user guidance.
- No macros — works in Excel desktop, Excel for the web, and Google Sheets.
