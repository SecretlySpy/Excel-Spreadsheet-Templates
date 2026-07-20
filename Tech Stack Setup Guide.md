# Tech Stack Setup Guide

## Core Tech Stack

| Component | Personal Savings Tracker | Adorama Campaign Tracker |
| --- | --- | --- |
| **Language / Platform** | Microsoft Excel (Formulas only) | Microsoft Excel (Formulas + VBA) |
| **File Format** | `.xlsx` (no macros) | `.xlsm` (macro-enabled) |
| **Runtime** | Excel Calculation Engine | Excel Calculation Engine + VBA Runtime |
| **Key Functions** | `SUMIFS`, `INDEX`, `MATCH`, `IFERROR`, `EOMONTH`, `DATEDIF`, `LARGE`, `LOOKUP`, `COUNTIF`, `AVERAGEIFS` | Same + VBA `modEmailProductionTracker` |
| **Version** | Excel 2016+ / Microsoft 365 | Excel 2016+ / Microsoft 365 (desktop for VBA) |

## Setup Instructions

### Windows
1. **Install Microsoft Excel** via Microsoft 365 (recommended) or standalone Office 2016+.
2. **Clone or download** the repository to your local machine.
3. **Open** any `.xlsx` file by double-clicking. For `.xlsm` files, click **Enable Content** when prompted.
4. **Enable Editing** if prompted with a "Protected View" banner.
5. The Personal Savings workbooks require **no macros** — formulas calculate automatically.

### macOS
1. **Install Microsoft Excel** for Mac via Microsoft 365.
2. **Clone or download** the repository.
3. **Open** `.xlsx` files directly. For `.xlsm` files, enable macros when prompted.
4. VBA compatibility may vary — core formulas work on all platforms.

### Linux
1. Use **LibreOffice Calc** for `.xlsx` files (full formula compatibility).
2. Alternatively, use **Excel for the Web** via browser (free with Microsoft account).
3. VBA macros (`.xlsm` files) are **not supported** on Linux — use web or Windows/Mac.

## Data Flow Architecture

### Personal Finance Tracker

```mermaid
graph TD
    T[Transactions Sheet<br/>Raw Data Input] -->|SUMIFS by Category & Date| D[Dashboard]
    S[Settings Sheet<br/>Categories, Budgets, Periods] -->|INDEX/MATCH| D
    I[Interest Settings<br/>Accounts, Rates, Dates] -->|Compound Interest Formulas| D
    G[Savings Goal<br/>Target Amount & Date] -->|Progress Calculation| D
    D --> K[KPIs: Income, Expenses,<br/>Net Savings, Savings Rate]
    D --> B[Budget vs Actual Table]
    D --> CF[Cash Flow — 12 Months]
    D --> EF[Emergency Fund Tracker]
    D --> SG[Savings Goal Progress]
    D --> T5[Top 5 Spending Categories]
```

### Overall Money Tracker

```mermaid
graph TD
    A[Accounts Sheet<br/>Balances & Interest Rates] -->|Compound Interest Engine| C[Calculated Balances]
    C --> SM[Summary Metrics<br/>Totals by Type, Interest, Staleness]
    A -->|Manual Snapshots| BH[Balance History]
    BH --> GR[Growth Analytics<br/>Ann. Growth %, Insights]
```

### Adorama Campaign Tracker

```mermaid
graph TD
    E[Email Campaigns Table] --> DB[Dashboard<br/>Sunday–Saturday View]
    S[SMS Campaigns Table] --> DB
    DB --> KPI[Summary KPIs<br/>Active, Sending, Pending]
    VBA[VBA Module] -->|Audit Stamps, Links| E
    VBA -->|Audit Stamps, Links| S
    VBA -->|Refresh, Format| DB
    CAL[SharePoint Calendars] -->|External Links| MC[Monthly Calendar Sheets]
```

## Common Troubleshooting Tips

| Issue | Solution |
| --- | --- |
| **Formulas Not Updating** | Go to `Formulas > Calculation Options` and set to `Automatic` |
| **`#NAME?` Errors** | Ensure you're using Excel 2016+ (older versions may not support `IFERROR`, `EOMONTH`, `DATEDIF`) |
| **`#DIV/0!` Errors** | Should not occur due to `IFERROR` wrappers — verify Settings has non-zero values |
| **Interest shows ₱0** | Check that `Last Updated` date is in the past and `Rate Basis` is exactly `Annual` or `Monthly` |
| **Macros not running** | `.xlsm` files require Excel desktop with macros enabled — not supported on web/Linux |
| **Protected View** | Click `Enable Editing` in the yellow banner when first opening downloaded files |
| **Stale account warnings** | Update the `Last Updated` date in Accounts/Settings after verifying the balance |
