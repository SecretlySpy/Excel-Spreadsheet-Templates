# System Production Adorama

Adorama marketing Excel applications: the self-contained **Email & SMS Campaign Tracker** and the **Loyalty and PLCC Email Plan** template.

## Workbooks

### Production Tracker (`Production Tracker/`)

| File | Role |
| --- | --- |
| `Email & SMS Campaign Tracker.xlsm` | Active tracker (production data) |
| `Email & SMS Campaign Tracker Template.xlsm` | Clean template |
| `Email & SMS Campaign Tracker_backup.xlsm` | Backup copy |

All formulas, formatting, validation, and VBA required at runtime are embedded in each `.xlsm` file.

### Project Tracker (`Project Tracker/`)

| File | Role |
| --- | --- |
| `Loyalty and PLCC Email Plan Template.xlsx` | Macro-free Loyalty/PLCC project plan with 2027 monthly calendars |

### Reporting Analysis

`Reporting Analysis/` is reserved for future reporting assets and is currently empty.

## Email & SMS Campaign Tracker

### Worksheet structure

| Sheet | Purpose |
| --- | --- |
| `Dashboard` | Current Sunday through next Saturday campaign view and summary KPIs |
| `Email Campaigns` | Email campaign source table and workflow checkboxes |
| `SMS Campaigns` | SMS campaign source table and workflow checkboxes |
| `Notes - Instructions` | Plain-language end-user guide (protected) |
| `June 2026 Calendar` | Monthly calendar mirrored from the SharePoint planning file |
| `May 2026 Calendar` | Hidden monthly calendar (active tracker only) |
| `Template for Duplicate` | Pre-formatted calendar used to create a new month |
| `Dropdowns` | Hidden Campaign Type validation source |
| `Automation Log` | Hidden desktop automation and error log |

Monthly calendar sheets mirror the team’s SharePoint Email/SMS planning workbooks through external links. To add a month, copy `Template for Duplicate`, rename it `<Month> YYYY Calendar`, and re-point its source via **Data → Edit Links**. Last Week versus Current Week delivery comparisons remain retired.

### Campaign fields

#### Canonical column order (Template & Backup)

Both campaign tables include a leading **`Week Number`** column (formula-driven `Week N` label from `Send Date`) so campaigns can be grouped and filtered by week.

Then both tables contain:

- `Send Date`
- `Send Time`
- `Campaign Name`
- `Campaign Type`
- `Current Stage`
- `Owner`
- channel-specific workflow checkboxes (plus `Scheduled`)
- links
- `Est. Audience`
- `Delivered`
- `Last Updated`
- `Last Updated By`
- `Notes`

#### Active tracker variance

The **active** workbook (`Email & SMS Campaign Tracker.xlsm`) currently starts at `Send Date` and does **not** include the leading `Week Number` column on the Email/SMS tables. The Dashboard still shows a Week Number KPI tile and a per-row Week Number derived from the campaign name’s leading `MMDDYY` code (fallback: Send Date). When aligning structure, prefer the Template column set.

Email workflow checkboxes:

1. `Campaign Name and UTM Parameter (Source Code)`
2. `Creative Brief, SL & PH`
3. `SKUs`
4. `In-Design`
5. `Build, QA`
6. `Route`
7. `Approval`
8. `Segments`
9. `Scheduled` (booking flag used by schedule-gap highlighting)

SMS workflow checkboxes:

1. `Send SMS Options`
2. `Send Test`
3. `Approval`
4. `Segments`
5. `Scheduled`

`Current Stage` lists all checked workflow fields rather than selecting only one stage.

### Date and time input

- `Send Date` displays as `dddd, mmmm d, yyyy`, for example `Wednesday, June 10, 2026`.
- Numeric `Send Time` values display in 12-hour format.
- `Send Time` also accepts text such as `STO` and `Local Timezone`.

### Dashboard

The Dashboard combines Email and SMS campaigns scheduled from the current Sunday through the following Saturday.

Approval displays `Done` or `Not Yet`. Segments displays `Provided` or `Pending`. A campaign is excluded from both the Dashboard feed and summary KPIs when either `Current Stage` or `Notes` is exactly `Cancelled` or `Canceled`, ignoring capitalization and surrounding spaces.

Summary KPIs use expanding structured table references:

- Active Work
- Sending Today
- Email Active
- SMS Active
- Approval Pending
- Sent

A **Week Number** tile beside the KPIs shows the current two-week window as a span (for example `25-26`), matching the Sunday-through-next-Saturday feed. Beneath it, a per-campaign **Week Number** column (next to `Bluecore/Attentive`) shows a `Week N` label for each row, derived from the leading `MMDDYY` code in the campaign string (e.g. `061726-STO-...` → `Week 25`) and falling back to the Send Date when there is no code.

Do not edit Dashboard helper columns **`AA:AL`**.

### Timed link labels

The `Jira Link`, `ClickUp Link`, `Bluecore/Attentive Link`, and `Proof of Schedule` link columns use native `HYPERLINK` formulas. They display the full URL until exactly seven days after `Send Date` and numeric `Send Time`, then display the clean platform name while preserving the same clickable URL.

For `STO`, `Local Timezone`, or a blank Send Time, the seven-day period starts at midnight on the Send Date. Excel for the web updates the label when the workbook recalculates; desktop Excel also schedules the next due refresh while the workbook remains open.

Very long links (over 255 characters, such as some Bluecore/Attentive `compose/design` URLs) are stored as a real clickable hyperlink showing the platform name instead of the timed formula — Excel cannot hold a URL that long inside a formula without returning `#VALUE!`.

### Schedule-gap highlighting

On the `Email Campaigns` and `SMS Campaigns` sheets, a row is highlighted when a deployment is due in the next working window but its `Scheduled` checkbox is still unchecked. Email rows fill orange and SMS rows fill yellow. The window is the next day, and on Fridays it extends through Saturday, Sunday, and Monday so weekend and Monday sends are flagged before everyone leaves. Cancelled rows are never highlighted, and checking the `Scheduled` box clears the highlight. The rule is native conditional formatting, so it updates with the date and works in desktop Excel and Excel for the web.

The same highlighting also appears on the **Dashboard** feed. Because the Dashboard has no `Scheduled` column, a feed row is treated as unscheduled when its `Stage` is not yet `Scheduled` or `Sent`; Email rows fill orange and SMS rows fill yellow over the same next-day/Friday-weekend window.

### Excel and SharePoint compatibility

Native formulas, tables, filters, formats, and saved checkbox values work in desktop Excel and Excel for the web.

VBA does **not** execute in Excel for the web. Open the workbook in desktop Excel with macros enabled for:

- automatic audit timestamps and editor names
- event-driven formatting repair
- checkbox double-click fallback
- Dashboard refresh commands
- automation logging
- scheduled timed-link label refresh (`Application.OnTime`)

Use SharePoint version history as the authoritative editor record for web changes.

### Maintenance

The `Notes - Instructions` sheet inside each workbook is the **end-user guide**: plain, non-technical wording covering features still in use.

This README and root [AI Documentation Notes.md](../AI%20Documentation%20Notes.md) are the technical/maintainer references. Environment setup is in [Tech Stack Setup Guide.md](../Tech%20Stack%20Setup%20Guide.md).

The `Notes - Instructions` worksheet password is `Adorama@042026_` for the active tracker and template; the legacy backup copy may still open with `adorama2024`. Protection is an accidental-edit safeguard, not encryption.

Do not rename workbook tables or headers, expose or edit Dashboard helper columns `AA:AL`, add blank lines after VBA continuation characters, or convert the files to `.xlsx`.

Useful desktop entry points: `RefreshDashboard`, `RefreshProductionStatus`, `ValidateWorkbookConfiguration`, `ApplyTimedCampaignLinks`.

## Loyalty and PLCC Email Plan Template

Macro-free `.xlsx` planner:

| Sheet | Purpose |
| --- | --- |
| `Dashboard` | Summary view |
| `Project Tracker` | Project grid (`ProjectTrackerTable`) |
| `January 2027 Calendar` … `December 2027 Calendar` | Monthly calendars fed by tracker formulas |
| `Lists` | Status list and formula helpers |

**Project Tracker columns:** `Jira Ticket Link`, `Campaign`, `Status`, `Send Date`, `Notes`, `Calendar Key`, `Calendar Display`, `Future Send Helper`.

Calendar day cells look up `TEXT(date,"yyyymmdd")&"|n"` keys against `Project Tracker!$F$2:$G$201` so multiple campaigns can appear on one day.

**Status values** (on `Lists`): Upcoming, In Progress, Completed, Cancelled.
