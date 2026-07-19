# Excel-Spreadsheet-Templates

Production-ready **Microsoft Excel** templates for Adorama marketing operations.

This repository is workbook-centric: runtime behavior for the campaign tracker is embedded in `.xlsm` files (tables, formulas, conditional formatting, and VBA). There is no application server and no committed `tools/` package.

## Repository layout

```text
Excel-Spreadsheet-Templates/
├── Adorama/
│   ├── Production Tracker/          # Email & SMS Campaign Tracker (.xlsm)
│   ├── Project Tracker/             # Loyalty & PLCC plan template (.xlsx)
│   ├── Reporting Analysis/          # Placeholder (empty)
│   └── README.md                    # Adorama product documentation
├── AI Documentation Notes.md        # Machine-parseable technical reference
├── Tech Stack Setup Guide.md        # Setup for Windows / macOS / Linux
├── AGENTS.md / CLAUDE.md / GEMINI.md
└── LICENSE
```

## What’s included

| Area | Path | Format | Description |
| --- | --- | --- | --- |
| **Email & SMS Campaign Tracker** | `Adorama/Production Tracker/` | `.xlsm` | Dual-channel production tracker with Dashboard, workflow checklists, timed links, SharePoint calendars, and desktop VBA |
| **Loyalty & PLCC Email Plan** | `Adorama/Project Tracker/` | `.xlsx` | Macro-free project grid + 2027 monthly calendars driven by formulas |

### Campaign tracker workbooks

| File | Role |
| --- | --- |
| `Email & SMS Campaign Tracker.xlsm` | Active production copy |
| `Email & SMS Campaign Tracker Template.xlsm` | Clean template |
| `Email & SMS Campaign Tracker_backup.xlsm` | Backup snapshot |

## Quick start

1. Open the workbook from **SharePoint** (preferred) or from a local clone.
2. For full automation (audit stamps, `RefreshDashboard`, timed-link scheduling), use **Excel desktop** with macros enabled.
3. Read the in-workbook guide on **`Notes - Instructions`**.
4. For environment setup and troubleshooting, see [Tech Stack Setup Guide.md](Tech%20Stack%20Setup%20Guide.md).
5. For maintainer-level architecture and every VBA entry point, see [AI Documentation Notes.md](AI%20Documentation%20Notes.md).

Adorama-specific product details (fields, KPIs, calendar workflow) live in [Adorama/README.md](Adorama/README.md).

## Documentation map

| Audience | Document |
| --- | --- |
| End users | `Notes - Instructions` sheet inside each campaign tracker |
| Operators / team leads | [Adorama/README.md](Adorama/README.md) |
| Maintainers / AI agents | [AI Documentation Notes.md](AI%20Documentation%20Notes.md) |
| New machine setup | [Tech Stack Setup Guide.md](Tech%20Stack%20Setup%20Guide.md) |
| AI coding assistants | [AGENTS.md](AGENTS.md) |

## License

See [LICENSE](LICENSE).
