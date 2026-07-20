# Excel-Spreadsheet-Templates

Production-ready **Microsoft Excel** templates for marketing operations and personal finance management.

This repository is workbook-centric: runtime behavior is embedded in the files themselves (tables, formulas, conditional formatting, and VBA where applicable). There is no application server and no committed `tools/` package.

## Repository layout

```text
Excel-Spreadsheet-Templates/
├── Personal/
│   └── Savings Tracker/
│       ├── Overall Money Tracker.xlsx       # Account balances & compound interest
│       ├── Personal Finance Tracker.xlsx    # Budgeting, cash flow & savings health
│       └── README.md                        # Savings Tracker documentation
├── Work Related/
│   └── Adorama/
│       ├── Production Tracker/              # Email & SMS Campaign Tracker (.xlsm)
│       ├── Project Tracker/                 # Loyalty & PLCC plan template (.xlsx)
│       ├── Reporting Analysis/              # Placeholder (empty)
│       └── README.md                        # Adorama product documentation
├── AI Documentation Notes.md                # Machine-parseable technical reference
├── Tech Stack Setup Guide.md                # Setup for Windows / macOS / Linux
├── AGENTS.md / CLAUDE.md / GEMINI.md
└── LICENSE
```

## What's included

| Area | Path | Format | Description |
| --- | --- | --- | --- |
| **Overall Money Tracker** | `Personal/Savings Tracker/` | `.xlsx` | Account-level balance tracking with compound interest, staleness warnings, and growth analytics |
| **Personal Finance Tracker** | `Personal/Savings Tracker/` | `.xlsx` | Transaction-based budgeting with dashboard KPIs, emergency fund tracker, savings goal, and interest engine |
| **Email & SMS Campaign Tracker** | `Work Related/Adorama/Production Tracker/` | `.xlsm` | Dual-channel production tracker with Dashboard, workflow checklists, timed links, SharePoint calendars, and desktop VBA |
| **Loyalty & PLCC Email Plan** | `Work Related/Adorama/Project Tracker/` | `.xlsx` | Macro-free project grid + 2027 monthly calendars driven by formulas |

### Personal Savings workbooks

| File | Features |
| --- | --- |
| `Overall Money Tracker.xlsx` | Multi-account balances, compound interest (Annual/Monthly basis), accrued interest, staleness warnings, balance history with annualized growth |
| `Personal Finance Tracker.xlsx` | Transaction logging, period-based budgeting, 12-month cash flow, emergency fund coverage, savings goal tracking, top 5 spending, interest engine |

### Campaign tracker workbooks

| File | Role |
| --- | --- |
| `Email & SMS Campaign Tracker.xlsm` | Active production copy |
| `Email & SMS Campaign Tracker Template.xlsm` | Clean template |
| `Email & SMS Campaign Tracker_backup.xlsm` | Backup snapshot |

## Quick start

### Personal finance
1. **Overall Money Tracker**: Open `Accounts`, enter your real bank/e-wallet balances and interest rates. Snapshots go in `Balance History`.
2. **Personal Finance Tracker**: Log transactions in `Transactions`. Set budgets and savings accounts in `Settings`. The Dashboard auto-updates.

### Campaign tracker
1. Open the workbook from **SharePoint** (preferred) or from a local clone.
2. For full automation (audit stamps, `RefreshDashboard`, timed-link scheduling), use **Excel desktop** with macros enabled.
3. Read the in-workbook guide on **`Notes - Instructions`**.

For environment setup and troubleshooting, see [Tech Stack Setup Guide.md](Tech%20Stack%20Setup%20Guide.md).

## Documentation map

| Audience | Document |
| --- | --- |
| End users | `README` sheet inside each workbook |
| Savings Tracker details | [Personal/Savings Tracker/README.md](Personal/Savings%20Tracker/README.md) |
| Adorama details | [Work Related/Adorama/README.md](Work%20Related/Adorama/README.md) |
| Maintainers / AI agents | [AI Documentation Notes.md](AI%20Documentation%20Notes.md) |
| New machine setup | [Tech Stack Setup Guide.md](Tech%20Stack%20Setup%20Guide.md) |
| AI coding assistants | [AGENTS.md](AGENTS.md) |

## License

See [LICENSE](LICENSE).
