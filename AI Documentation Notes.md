# AI Documentation Notes

<!-- Machine-parseable technical documentation for Excel-Spreadsheet-Templates / Adorama trackers. -->
<!-- Last regenerated: 2026-07-19 from static scan of .xlsm/.xlsx + olevba extraction. -->

# Systemic Operational Mechanics

## High-Level Architecture
- Repository `Excel-Spreadsheet-Templates` hosts Adorama marketing production Excel templates.
- Primary application: self-contained **Email & SMS Campaign Tracker** (`.xlsm`) under `Adorama/Production Tracker/`.
- Secondary template: formula-only **Loyalty and PLCC Email Plan Template** (`.xlsx`) under `Adorama/Project Tracker/`.
- Runtime logic for the campaign tracker lives inside VBA module `modEmailProductionTracker` (~124 procedures).
- Sheet event modules (`ThisWorkbook`, Email `Sheet1`, SMS `Sheet17`) are thin delegators.
- End-user guidance is embedded on protected sheet `Notes - Instructions`.
- There is **no** committed `tools/` Python package in this repository; automation is workbook-embedded.

## Data Flow
1. Users enter/edit rows on `Email Campaigns` (`EmailCampaignsTable`) and `SMS Campaigns` (`SMSCampaignsTable`).
2. Checklist columns (native Excel checkboxes / TRUE-FALSE) and fields such as Send Date, Owner, links, Est. Audience, Delivered, Notes are source-of-truth data.
3. `Current Stage` is a calculated structured formula listing checked workflow steps.
4. `Dashboard` consumes both tables via native spill formulas in helper columns `AA:AL` and displays the Sunday-through-next-Saturday window.
5. Summary KPIs (Active Work, Sending Today, Email Active, SMS Active, Approval Pending, Sent, Week Number) use expanding structured-table formulas.
6. Monthly calendar sheets mirror SharePoint planning workbooks via external links (`Data > Edit Links`).
7. Desktop VBA stamps `Last Updated` / `Last Updated By`, repairs formats, manages timed link labels, and can log to `Automation Log`.

## Control Flow
```
Workbook_Open
  -> UnfreezeWorkbookViews
  -> ApplyCampaignEntryFormats
  -> ApplyTimedCampaignLinks
  -> RefreshDashboard

Worksheet_Change (Email or SMS table body)
  -> HandleCampaignChange
       -> audit stamp (UpdateRowTimestampAndUser)
       -> FormatHyperlinksInChangedCells (link columns)
       -> RefreshNativeOutputs (light Dashboard calc)

Worksheet_BeforeDoubleClick (checklist cell)
  -> ToggleInventoryChecklist

Workbook_BeforeClose
  -> CancelCampaignLinkRefresh
```

## Key Dependencies
- Microsoft Excel desktop (VBA macros) for full automation; Excel for the web for native tables/formulas/checkboxes only.
- Excel tables: `EmailCampaignsTable`, `SMSCampaignsTable`, `DashboardWorkTable`, `AutomationLogTable`, `NotesTable`.
- Hidden sheets: `Dropdowns` (campaign type lists), `Automation Log`.
- SharePoint external links for monthly Email/SMS planning calendars.
- Optional Windows identity via `Environ("Username")` for audit fields.
- `Application.OnTime` for deferred timed-link label refresh while workbook remains open on desktop.

## Workbook Inventory (scanned 2026-07-19)

| File | Role | Sheets | VBA |
| --- | --- | ---: | --- |
| `Adorama/Production Tracker/Email & SMS Campaign Tracker.xlsm` | Active tracker with production data | 9 | Yes (`modEmailProductionTracker`) |
| `Adorama/Production Tracker/Email & SMS Campaign Tracker Template.xlsm` | Clean template | 8 | Yes (same module) |
| `Adorama/Production Tracker/Email & SMS Campaign Tracker_backup.xlsm` | Backup snapshot | 8 | Yes (same module) |
| `Adorama/Project Tracker/Loyalty and PLCC Email Plan Template.xlsx` | Loyalty/PLCC project planner | 15 | No |

### Active tracker sheet map
| Sheet | Visibility | Purpose |
| --- | --- | --- |
| Dashboard | visible | KPI tiles + Sunday–next-Saturday campaign feed (`DashboardWorkTable`); helper cols AA:AL |
| Email Campaigns | visible | Email source table + 9 checklist fields including Scheduled |
| SMS Campaigns | visible | SMS source table + 5 checklist fields including Scheduled |
| Notes - Instructions | visible (protected) | End-user feature guide table |
| May 2026 Calendar | hidden (active only) | SharePoint-mirrored May plan |
| June 2026 Calendar | visible | SharePoint-mirrored June plan |
| Template for Duplicate | visible | Starter calendar layout for new months |
| Dropdowns | veryHidden | Campaign Type / stage / owner lists |
| Automation Log | veryHidden | Desktop automation log table |

### Campaign table columns (canonical / Template & Backup)
Email (`EmailCampaignsTable`): `Week Number`, `Send Date`, `Send Time`, `Campaign Name`, `Campaign Type`, `Current Stage`, `Owner`, checklist (`Campaign Name and UTM Parameter (Source Code)`, `Creative Brief, SL & PH`, `SKUs`, `In-Design`, `Build, QA`, `Route`, `Approval`, `Segments`, `Scheduled`), `Jira Link`, `ClickUp Link`, `Bluecore/Attentive Link`, `Est. Audience`, `Delivered`, `Last Updated`, `Last Updated By`, `Notes`.

SMS (`SMSCampaignsTable`): `Week Number`, `Send Date`, `Send Time`, `Campaign Name`, `Campaign Type`, `Current Stage`, `Owner`, checklist (`Send SMS Options`, `Send Test`, `Approval`, `Segments`, `Scheduled`), `Proof of Schedule`, `Bluecore/Attentive Link`, `Est. Audience`, `Delivered`, `Last Updated`, `Last Updated By`, `Notes`.

### Known structural variance (Active vs Template)
- **Active** `Email & SMS Campaign Tracker.xlsm` currently omits the leading `Week Number` column on both campaign tables (starts at `Send Date`).
- **Template** and **Backup** include `Week Number` as column A (formula-driven week label).
- Dashboard still exposes a Week Number KPI tile and a per-row Week Number column derived from campaign name `MMDDYY` prefix / send date.
- Active workbook includes hidden `May 2026 Calendar`; Template/Backup do not.

### Loyalty and PLCC Email Plan Template.xlsx
- Sheets: `Dashboard`, `Project Tracker` (`ProjectTrackerTable`), monthly `January 2027 Calendar` … `December 2027 Calendar`, `Lists`.
- Project Tracker columns: `Jira Ticket Link`, `Campaign`, `Status`, `Send Date`, `Notes`, `Calendar Key`, `Calendar Display`, `Future Send Helper`.
- Monthly calendars VLOOKUP from `Project Tracker!$F$2:$G$201` using `yyyymmdd|n` calendar keys.
- Status list values on `Lists`: Upcoming, In Progress, Completed, Cancelled.
- No VBA; pure worksheet formulas.

### Features / Capabilities
- Dual-channel campaign production tracking (Email + SMS).
- Workflow checklists with native checkbox preference and double-click fallback.
- Calculated multi-stage `Current Stage` (all checked steps, not single-state).
- Dashboard window: current Sunday through next Saturday; cancel exclusion via Notes/Stage = Cancelled/Canceled.
- Timed link labels (full URL for 7 days post-send, then platform name).
- Long-URL hyperlink fallback (>255 chars).
- Schedule-gap conditional formatting (email orange / SMS yellow).
- Campaign Type dropdown + optional Others custom prompt (desktop).
- Audit timestamps and editor names (desktop VBA).
- SharePoint-linked monthly planning calendars.
- Protected end-user instruction sheet.
- Validation helpers: `ValidateWorkbookConfiguration`.
- Migration helper: `MigrateProductionInventoryStructure` (creates backup).

### Module constants (`modEmailProductionTracker`)
| Constant | Value |
| --- | --- |
| `SH_EMAIL` / `SH_INVENTORY` | Email Campaigns |
| `SH_SMS` | SMS Campaigns |
| `SH_DASHBOARD` | Dashboard |
| `SH_LOG` | Automation Log |
| `SH_INSTRUCTIONS` | Notes - Instructions |
| `TBL_INVENTORY` | EmailCampaignsTable |
| `TBL_SMS` | SMSCampaignsTable |
| `TBL_DASHBOARD` | DashboardWorkTable |
| `FIRST_DASHBOARD_ROW` | 11 |
| `TIMESTAMP_FORMAT` | MM/DD/YYYY h:mm AM/PM |

# Module / File: ThisWorkbook

## Function: Workbook_Open
- **Purpose**: On open: unfreezes views, applies date/time formats, installs timed links, and refreshes the Dashboard.
- **Inputs**:
  - None
- **Outputs**: None (Sub)
- **Dependencies**: `modEmailProductionTracker`; Excel Application event model
- **Behavior**: Event entry point `Workbook_Open` on `ThisWorkbook` delegates to the production tracker module while preserving prior `Application.EnableEvents` state.
- **Side Effects**: May modify worksheet cells, formats, hyperlinks, Application.OnTime schedule, or Dashboard formulas

# Module / File: ThisWorkbook

## Function: Workbook_BeforeClose
- **Purpose**: On close: cancels any scheduled Application.OnTime timed-link refresh callback.
- **Inputs**:
  - `Cancel` (`Boolean`): Standard BeforeClose cancel flag (unused for cancel logic)
- **Outputs**: None (Sub)
- **Dependencies**: `modEmailProductionTracker`; Excel Application event model
- **Behavior**: Event entry point `Workbook_BeforeClose` on `ThisWorkbook` delegates to the production tracker module while preserving prior `Application.EnableEvents` state.
- **Side Effects**: May modify worksheet cells, formats, hyperlinks, Application.OnTime schedule, or Dashboard formulas

# Module / File: Sheet1 (Email Campaigns)

## Function: Worksheet_BeforeDoubleClick
- **Purpose**: Delegates double-clicks on campaign table body cells to ToggleInventoryChecklist.
- **Inputs**:
  - `Target` (`Range`): Double-clicked cell
  - `Cancel` (`Boolean`): Set True when checklist toggle handles the click
- **Outputs**: None (Sub)
- **Dependencies**: `modEmailProductionTracker`; Excel Application event model
- **Behavior**: Event entry point `Worksheet_BeforeDoubleClick` on `Sheet1 (Email Campaigns)` delegates to the production tracker module while preserving prior `Application.EnableEvents` state.
- **Side Effects**: May modify worksheet cells, formats, hyperlinks, Application.OnTime schedule, or Dashboard formulas

# Module / File: Sheet1 (Email Campaigns)

## Function: Worksheet_Change
- **Purpose**: Delegates campaign table body changes to HandleCampaignChange with events temporarily disabled.
- **Inputs**:
  - `Target` (`Range`): Changed cell range
- **Outputs**: None (Sub)
- **Dependencies**: `modEmailProductionTracker`; Excel Application event model
- **Behavior**: Event entry point `Worksheet_Change` on `Sheet1 (Email Campaigns)` delegates to the production tracker module while preserving prior `Application.EnableEvents` state.
- **Side Effects**: May modify worksheet cells, formats, hyperlinks, Application.OnTime schedule, or Dashboard formulas

# Module / File: Sheet17 (SMS Campaigns)

## Function: Worksheet_BeforeDoubleClick
- **Purpose**: Delegates double-clicks on SMS table body cells to ToggleInventoryChecklist.
- **Inputs**:
  - `Target` (`Range`): Double-clicked cell
  - `Cancel` (`Boolean`): Set True when checklist toggle handles the click
- **Outputs**: None (Sub)
- **Dependencies**: `modEmailProductionTracker`; Excel Application event model
- **Behavior**: Event entry point `Worksheet_BeforeDoubleClick` on `Sheet17 (SMS Campaigns)` delegates to the production tracker module while preserving prior `Application.EnableEvents` state.
- **Side Effects**: May modify worksheet cells, formats, hyperlinks, Application.OnTime schedule, or Dashboard formulas

# Module / File: Sheet17 (SMS Campaigns)

## Function: Worksheet_Change
- **Purpose**: Delegates SMS table body changes to HandleCampaignChange with events temporarily disabled.
- **Inputs**:
  - `Target` (`Range`): Changed cell range
- **Outputs**: None (Sub)
- **Dependencies**: `modEmailProductionTracker`; Excel Application event model
- **Behavior**: Event entry point `Worksheet_Change` on `Sheet17 (SMS Campaigns)` delegates to the production tracker module while preserving prior `Application.EnableEvents` state.
- **Side Effects**: May modify worksheet cells, formats, hyperlinks, Application.OnTime schedule, or Dashboard formulas

# Module / File: Sheet2

## Function: (none)
- **Purpose**: Code-behind placeholder for a worksheet with no custom event handlers.
- **Inputs**:
  - None
- **Outputs**: None
- **Dependencies**: None
- **Behavior**: Module exists in the VBA project with no Sub/Function procedures.
- **Side Effects**: none

# Module / File: Sheet3

## Function: (none)
- **Purpose**: Code-behind placeholder for a worksheet with no custom event handlers.
- **Inputs**:
  - None
- **Outputs**: None
- **Dependencies**: None
- **Behavior**: Module exists in the VBA project with no Sub/Function procedures.
- **Side Effects**: none

# Module / File: Sheet4

## Function: (none)
- **Purpose**: Code-behind placeholder for a worksheet with no custom event handlers.
- **Inputs**:
  - None
- **Outputs**: None
- **Dependencies**: None
- **Behavior**: Module exists in the VBA project with no Sub/Function procedures.
- **Side Effects**: none

# Module / File: Sheet5

## Function: (none)
- **Purpose**: Code-behind placeholder for a worksheet with no custom event handlers.
- **Inputs**:
  - None
- **Outputs**: None
- **Dependencies**: None
- **Behavior**: Module exists in the VBA project with no Sub/Function procedures.
- **Side Effects**: none

# Module / File: Sheet6

## Function: (none)
- **Purpose**: Code-behind placeholder for a worksheet with no custom event handlers.
- **Inputs**:
  - None
- **Outputs**: None
- **Dependencies**: None
- **Behavior**: Module exists in the VBA project with no Sub/Function procedures.
- **Side Effects**: none

# Module / File: Sheet7

## Function: (none)
- **Purpose**: Code-behind placeholder for a worksheet with no custom event handlers.
- **Inputs**:
  - None
- **Outputs**: None
- **Dependencies**: None
- **Behavior**: Module exists in the VBA project with no Sub/Function procedures.
- **Side Effects**: none

# Module / File: Sheet8

## Function: (none)
- **Purpose**: Code-behind placeholder for a worksheet with no custom event handlers.
- **Inputs**:
  - None
- **Outputs**: None
- **Dependencies**: None
- **Behavior**: Module exists in the VBA project with no Sub/Function procedures.
- **Side Effects**: none

# Module / File: modEmailProductionTracker

Core standard module (~132k characters of VBA). Public members are safe desktop entry points; Private members support migration, formatting, dashboard formulas, calendars, and validation.

## Function: FormatHyperlinksInChangedCells
- **Purpose**: Reformats hyperlink display text in changed cells within a campaign table using platform display names.
- **Inputs**:
  - `lo` (`ListObject`): Excel ListObject (table)
  - `changedCells` (`Range`): Subset of changed cells inside the table
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Sub `FormatHyperlinksInChangedCells`. Signature: `Private Sub FormatHyperlinksInChangedCells( ByVal lo As ListObject, ByVal changedCells As Range)`. Reformats hyperlink display text in changed cells within a campaign table using platform display names.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: MigrateProductionInventoryStructure
- **Purpose**: One-time/migration routine that reshapes campaign table columns, checklists, and structure; creates a backup first.
- **Inputs**:
  - None
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Public Sub `MigrateProductionInventoryStructure`. Signature: `Public Sub MigrateProductionInventoryStructure()`. One-time/migration routine that reshapes campaign table columns, checklists, and structure; creates a backup first.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: RefreshProductionStatus
- **Purpose**: Full production refresh: recalculates stages, formats, timed links, and dashboard-related outputs for email and SMS tables.
- **Inputs**:
  - None
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Public Sub `RefreshProductionStatus`. Signature: `Public Sub RefreshProductionStatus()`. Full production refresh: recalculates stages, formats, timed links, and dashboard-related outputs for email and SMS tables.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: ApplyAllConfigurations
- **Purpose**: Compatibility no-op retained for older callers; use RefreshDashboard instead.
- **Inputs**:
  - None
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Public Sub `ApplyAllConfigurations`. Signature: `Public Sub ApplyAllConfigurations()`. Compatibility no-op retained for older callers; use RefreshDashboard instead.
- **Side Effects**: none

## Function: GenerateCampaignID
- **Purpose**: Builds a normalized campaign identifier string from send date and campaign name.
- **Inputs**:
  - `sendDate` (`Variant`): Campaign send date
  - `campaignName` (`String`): Campaign display name
- **Outputs**: `String` — return value of `GenerateCampaignID`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Public Function `GenerateCampaignID`. Signature: `Public Function GenerateCampaignID( ByVal sendDate As Variant, ByVal campaignName As String) As String`. Builds a normalized campaign identifier string from send date and campaign name.
- **Side Effects**: none

## Function: CalculateCurrentStage
- **Purpose**: Returns a human-readable stage string from checked workflow checklist items on a campaign row.
- **Inputs**:
  - `ws` (`Worksheet`): Worksheet being processed
  - `r` (`Long`): Worksheet row number for the campaign
- **Outputs**: `String` — return value of `CalculateCurrentStage`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Public Function `CalculateCurrentStage`. Signature: `Public Function CalculateCurrentStage( ByVal ws As Worksheet, ByVal r As Long) As String`. Returns a human-readable stage string from checked workflow checklist items on a campaign row.
- **Side Effects**: none

## Function: CheckedChecklistItemsForRow
- **Purpose**: Collects checked checklist header labels for one email or SMS row into a comma-separated string.
- **Inputs**:
  - `ws` (`Worksheet`): Worksheet being processed
  - `r` (`Long`): Worksheet row number for the campaign
- **Outputs**: `String` — return value of `CheckedChecklistItemsForRow`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Function `CheckedChecklistItemsForRow`. Signature: `Private Function CheckedChecklistItemsForRow( ByVal ws As Worksheet, ByVal r As Long) As String`. Collects checked checklist header labels for one email or SMS row into a comma-separated string.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: CalculateRiskLevel
- **Purpose**: Legacy risk classification for a campaign row based on delivered count, checklist progress, and days until send.
- **Inputs**:
  - `ws` (`Worksheet`): Worksheet being processed
  - `r` (`Long`): Worksheet row number for the campaign
- **Outputs**: `String` — return value of `CalculateRiskLevel`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Public Function `CalculateRiskLevel`. Signature: `Public Function CalculateRiskLevel( ByVal ws As Worksheet, ByVal r As Long) As String`. Legacy risk classification for a campaign row based on delivered count, checklist progress, and days until send.
- **Side Effects**: none

## Function: ApplyDashboardKpiFormulas
- **Purpose**: Writes structured-table SUMPRODUCT KPI formulas onto the Dashboard summary tiles.
- **Inputs**:
  - None
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Public Sub `ApplyDashboardKpiFormulas`. Signature: `Public Sub ApplyDashboardKpiFormulas()`. Writes structured-table SUMPRODUCT KPI formulas onto the Dashboard summary tiles.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: ApplyTimedCampaignLinks
- **Purpose**: Installs timed HYPERLINK formulas (or long-link fallbacks) on email and SMS link columns.
- **Inputs**:
  - None
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Public Sub `ApplyTimedCampaignLinks`. Signature: `Public Sub ApplyTimedCampaignLinks()`. Installs timed HYPERLINK formulas (or long-link fallbacks) on email and SMS link columns.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: ApplyTimedLinksToTable
- **Purpose**: Applies timed link installation across a specified list of link column headers on one ListObject.
- **Inputs**:
  - `lo` (`ListObject`): Excel ListObject (table)
  - `linkHeaders` (`Variant`): Array of link column header names
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Sub `ApplyTimedLinksToTable`. Signature: `Private Sub ApplyTimedLinksToTable( ByVal lo As ListObject, ByVal linkHeaders As Variant)`. Applies timed link installation across a specified list of link column headers on one ListObject.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: CampaignLinkDisplayName
- **Purpose**: Maps a link column header to the clean platform label (Jira, ClickUp, Bluecore/Attentive, Proof of Schedule).
- **Inputs**:
  - `headerName` (`String`): Table column header text
- **Outputs**: `String` — return value of `CampaignLinkDisplayName`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Function `CampaignLinkDisplayName`. Signature: `Private Function CampaignLinkDisplayName( ByVal headerName As String) As String`. Maps a link column header to the clean platform label (Jira, ClickUp, Bluecore/Attentive, Proof of Schedule).
- **Side Effects**: none

## Function: CampaignLinkAddress
- **Purpose**: Extracts the URL from a cell hyperlink, formula, or plain text address.
- **Inputs**:
  - `cell` (`Range`): Single cell range
- **Outputs**: `String` — return value of `CampaignLinkAddress`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Function `CampaignLinkAddress`. Signature: `Private Function CampaignLinkAddress(ByVal cell As Range) As String`. Extracts the URL from a cell hyperlink, formula, or plain text address.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: InstallTimedCampaignLink
- **Purpose**: Writes a timed HYPERLINK formula into a link cell that switches label after seven days post-send.
- **Inputs**:
  - `lo` (`ListObject`): Excel ListObject (table)
  - `cell` (`Range`): Single cell range
  - `displayName` (`String`): Platform display label for the hyperlink
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Sub `InstallTimedCampaignLink`. Signature: `Private Sub InstallTimedCampaignLink( ByVal lo As ListObject, ByVal cell As Range, ByVal displayName As String)`. Writes a timed HYPERLINK formula into a link cell that switches label after seven days post-send.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: InstallLongCampaignLink
- **Purpose**: Installs a real hyperlink with platform label for URLs longer than 255 characters (formula-safe limit).
- **Inputs**:
  - `cell` (`Range`): Single cell range
  - `linkAddress` (`String`): URL string for the campaign link
  - `displayName` (`String`): Platform display label for the hyperlink
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Sub `InstallLongCampaignLink`. Signature: `Private Sub InstallLongCampaignLink( ByVal cell As Range, ByVal linkAddress As String, ByVal displayName As String)`. Installs a real hyperlink with platform label for URLs longer than 255 characters (formula-safe limit).
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: TimedCampaignLinkFormula
- **Purpose**: Builds the Excel HYPERLINK formula string for timed platform label switching.
- **Inputs**:
  - `linkAddress` (`String`): URL string for the campaign link
  - `displayName` (`String`): Platform display label for the hyperlink
- **Outputs**: `String` — return value of `TimedCampaignLinkFormula`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Function `TimedCampaignLinkFormula`. Signature: `Private Function TimedCampaignLinkFormula( ByVal linkAddress As String, ByVal displayName As String) As String`. Builds the Excel HYPERLINK formula string for timed platform label switching.
- **Side Effects**: none

## Function: RefreshTimedCampaignLinks
- **Purpose**: Recalculates timed link columns and reschedules the next Application.OnTime refresh.
- **Inputs**:
  - None
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Public Sub `RefreshTimedCampaignLinks`. Signature: `Public Sub RefreshTimedCampaignLinks()`. Recalculates timed link columns and reschedules the next Application.OnTime refresh.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: CalculateTimedCampaignLinkColumns
- **Purpose**: Forces recalculation of timed link formulas on email and SMS tables.
- **Inputs**:
  - None
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Sub `CalculateTimedCampaignLinkColumns`. Signature: `Private Sub CalculateTimedCampaignLinkColumns()`. Forces recalculation of timed link formulas on email and SMS tables.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: ScheduleNextCampaignLinkRefresh
- **Purpose**: Schedules Application.OnTime for the next link-label maturity using RefreshTimedCampaignLinks.
- **Inputs**:
  - None
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Public Sub `ScheduleNextCampaignLinkRefresh`. Signature: `Public Sub ScheduleNextCampaignLinkRefresh()`. Schedules Application.OnTime for the next link-label maturity using RefreshTimedCampaignLinks.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: CancelCampaignLinkRefresh
- **Purpose**: Cancels any pending Application.OnTime callback for timed campaign link refresh.
- **Inputs**:
  - None
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Public Sub `CancelCampaignLinkRefresh`. Signature: `Public Sub CancelCampaignLinkRefresh()`. Cancels any pending Application.OnTime callback for timed campaign link refresh.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: NextCampaignLinkMaturity
- **Purpose**: Returns the earliest future datetime when any campaign link label should flip.
- **Inputs**:
  - None
- **Outputs**: `Date` — return value of `NextCampaignLinkMaturity`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Function `NextCampaignLinkMaturity`. Signature: `Private Function NextCampaignLinkMaturity() As Date`. Returns the earliest future datetime when any campaign link label should flip.
- **Side Effects**: none

## Function: NextTableLinkMaturity
- **Purpose**: Scans one campaign table for the next timed-link maturity datetime.
- **Inputs**:
  - `lo` (`ListObject`): Excel ListObject (table)
  - `linkHeaders` (`Variant`): Array of link column header names
- **Outputs**: `Date` — return value of `NextTableLinkMaturity`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Function `NextTableLinkMaturity`. Signature: `Private Function NextTableLinkMaturity( ByVal lo As ListObject, ByVal linkHeaders As Variant) As Date`. Scans one campaign table for the next timed-link maturity datetime.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: CampaignLinkMaturity
- **Purpose**: Computes the maturity datetime for one campaign row based on Send Date and Send Time.
- **Inputs**:
  - `lo` (`ListObject`): Excel ListObject (table)
  - `rowIndex` (`Long`): Parameter `rowIndex`
- **Outputs**: `Date` — return value of `CampaignLinkMaturity`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Function `CampaignLinkMaturity`. Signature: `Private Function CampaignLinkMaturity( ByVal lo As ListObject, ByVal rowIndex As Long) As Date`. Computes the maturity datetime for one campaign row based on Send Date and Send Time.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: RefreshDashboard
- **Purpose**: Primary desktop refresh: KPI formulas, timed links, and native spill formulas into Dashboard helper columns AA:AL.
- **Inputs**:
  - None
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Public Sub `RefreshDashboard`. Signature: `Public Sub RefreshDashboard()`. Primary desktop refresh: KPI formulas, timed links, and native spill formulas into Dashboard helper columns AA:AL.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: RefreshNativeOutputs
- **Purpose**: Lightweight edit-time recalculation of Dashboard audit header and KPI ranges only.
- **Inputs**:
  - None
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Public Sub `RefreshNativeOutputs`. Signature: `Public Sub RefreshNativeOutputs()`. Lightweight edit-time recalculation of Dashboard audit header and KPI ranges only.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: ApplyDashboardAuditHeader
- **Purpose**: Writes Last Refresh / Last Edited By labels and formulas on the Dashboard header area.
- **Inputs**:
  - `ws` (`Worksheet`): Worksheet being processed
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Sub `ApplyDashboardAuditHeader`. Signature: `Private Sub ApplyDashboardAuditHeader(ByVal ws As Worksheet)`. Writes Last Refresh / Last Edited By labels and formulas on the Dashboard header area.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: DashboardLastEditedByFormula
- **Purpose**: Returns the LET formula that picks the most recently updated email or SMS editor name.
- **Inputs**:
  - None
- **Outputs**: `String` — return value of `DashboardLastEditedByFormula`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Function `DashboardLastEditedByFormula`. Signature: `Private Function DashboardLastEditedByFormula() As String`. Returns the LET formula that picks the most recently updated email or SMS editor name.
- **Side Effects**: none

## Function: ApplyDashboardNativeFormulas
- **Purpose**: Applies spill and display formulas that feed the Dashboard work table from helper columns.
- **Inputs**:
  - `wsD` (`Worksheet`): Dashboard worksheet
  - `dashboardTable` (`ListObject`): DashboardWorkTable ListObject
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Sub `ApplyDashboardNativeFormulas`. Signature: `Private Sub ApplyDashboardNativeFormulas( ByVal wsD As Worksheet, ByVal dashboardTable As ListObject)`. Applies spill and display formulas that feed the Dashboard work table from helper columns.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: ApplyFormulaCompat
- **Purpose**: Writes a formula using Formula2 when available, otherwise Formula, for Excel version compatibility.
- **Inputs**:
  - `targetRange` (`Range`): Range that receives a formula
  - `formulaText` (`String`): Excel formula string to write
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Sub `ApplyFormulaCompat`. Signature: `Private Sub ApplyFormulaCompat( ByVal targetRange As Range, ByVal formulaText As String)`. Writes a formula using Formula2 when available, otherwise Formula, for Excel version compatibility.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: DashboardNativeLinkFormula
- **Purpose**: Builds a formula that surfaces Jira/ClickUp/Bluecore links from the spilled helper array.
- **Inputs**:
  - `sourceColumn` (`Long`): Source column index in the spilled helper array
  - `displayText` (`String`): Parameter `displayText`
- **Outputs**: `String` — return value of `DashboardNativeLinkFormula`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Function `DashboardNativeLinkFormula`. Signature: `Private Function DashboardNativeLinkFormula( ByVal sourceColumn As Long, ByVal displayText As String) As String`. Builds a formula that surfaces Jira/ClickUp/Bluecore links from the spilled helper array.
- **Side Effects**: none

## Function: DashboardNativeSpillFormula
- **Purpose**: Builds the large FILTER/HSTACK spill formula that unions email and SMS rows for the Sunday–next-Saturday window.
- **Inputs**:
  - None
- **Outputs**: `String` — return value of `DashboardNativeSpillFormula`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Function `DashboardNativeSpillFormula`. Signature: `Private Function DashboardNativeSpillFormula() As String`. Builds the large FILTER/HSTACK spill formula that unions email and SMS rows for the Sunday–next-Saturday window.
- **Side Effects**: none

## Function: DashboardNativeBlankRowFormula
- **Purpose**: Builds a blank-row placeholder formula used when the Dashboard spill has no matching campaigns.
- **Inputs**:
  - None
- **Outputs**: `String` — return value of `DashboardNativeBlankRowFormula`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Function `DashboardNativeBlankRowFormula`. Signature: `Private Function DashboardNativeBlankRowFormula() As String`. Builds a blank-row placeholder formula used when the Dashboard spill has no matching campaigns.
- **Side Effects**: none

## Function: CountDashboardMatches
- **Purpose**: Counts how many table rows fall into the Dashboard date window and are not cancelled.
- **Inputs**:
  - `lo` (`ListObject`): Excel ListObject (table)
  - `weekStart` (`Date`): Parameter `weekStart`
  - `weekEnd` (`Date`): Parameter `weekEnd`
- **Outputs**: `Long` — return value of `CountDashboardMatches`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Function `CountDashboardMatches`. Signature: `Private Function CountDashboardMatches( ByVal lo As ListObject, ByVal weekStart As Date, ByVal weekEnd As Date) As Long`. Counts how many table rows fall into the Dashboard date window and are not cancelled.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: AppendDashboardRows
- **Purpose**: Legacy path that materializes Dashboard rows into the work table (superseded by native spill where used).
- **Inputs**:
  - `lo` (`ListObject`): Excel ListObject (table)
  - `channelName` (`String`): Email or SMS channel label
  - `weekStart` (`Date`): Parameter `weekStart`
  - `weekEnd` (`Date`): Parameter `weekEnd`
  - `dashboardValues` (`Variant`): Parameter `dashboardValues`
  - `jiraLinks` (`String`): Parameter `jiraLinks`
  - `clickUpLinks` (`String`): Parameter `clickUpLinks`
  - `bluecoreLinks` (`String`): Parameter `bluecoreLinks`
  - `displayIndex` (`Long`): Parameter `displayIndex`
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Sub `AppendDashboardRows`. Signature: `Private Sub AppendDashboardRows( ByVal lo As ListObject, ByVal channelName As String, ByVal weekStart As Date, ByVal weekEnd As Date, ByRef dashboardValues() As Variant, ByRef jiraLinks() As String, ByRef clickUpLinks() As String, ByRef bluecoreLinks() As String, ByRef displayIndex As Long)`. Legacy path that materializes Dashboard rows into the work table (superseded by native spill where used).
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: ApprovalDashboardStatus
- **Purpose**: Maps a checklist Approval value to Done or Not Yet for Dashboard display.
- **Inputs**:
  - `statusValue` (`Variant`): Approval or Segments checkbox/status value
- **Outputs**: `String` — return value of `ApprovalDashboardStatus`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Function `ApprovalDashboardStatus`. Signature: `Private Function ApprovalDashboardStatus(ByVal statusValue As Variant) As String`. Maps a checklist Approval value to Done or Not Yet for Dashboard display.
- **Side Effects**: none

## Function: SegmentsDashboardStatus
- **Purpose**: Maps a checklist Segments value to Provided or Pending for Dashboard display.
- **Inputs**:
  - `statusValue` (`Variant`): Approval or Segments checkbox/status value
- **Outputs**: `String` — return value of `SegmentsDashboardStatus`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Function `SegmentsDashboardStatus`. Signature: `Private Function SegmentsDashboardStatus(ByVal statusValue As Variant) As String`. Maps a checklist Segments value to Provided or Pending for Dashboard display.
- **Side Effects**: none

## Function: RemoveLegacyCalendarAndComparisonArtifacts
- **Purpose**: Deletes retired calendar sheet names and clears legacy Dashboard comparison artifacts from a workbook.
- **Inputs**:
  - `wb` (`Workbook`): Workbook being processed
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Public Sub `RemoveLegacyCalendarAndComparisonArtifacts`. Signature: `Public Sub RemoveLegacyCalendarAndComparisonArtifacts(ByVal wb As Workbook)`. Deletes retired calendar sheet names and clears legacy Dashboard comparison artifacts from a workbook.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: RemoveLegacyDashboardComparisons
- **Purpose**: Removes retired last-week/current-week comparison tables/charts from the Dashboard sheet.
- **Inputs**:
  - `ws` (`Worksheet`): Worksheet being processed
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Sub `RemoveLegacyDashboardComparisons`. Signature: `Private Sub RemoveLegacyDashboardComparisons(ByVal ws As Worksheet)`. Removes retired last-week/current-week comparison tables/charts from the Dashboard sheet.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: IsLegacyCalendarSheetName
- **Purpose**: Returns True when a worksheet name matches retired calendar or Calendar Import naming patterns.
- **Inputs**:
  - `sheetName` (`String`): Worksheet name to evaluate
- **Outputs**: `Boolean` — return value of `IsLegacyCalendarSheetName`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Function `IsLegacyCalendarSheetName`. Signature: `Private Function IsLegacyCalendarSheetName(ByVal sheetName As String) As Boolean`. Returns True when a worksheet name matches retired calendar or Calendar Import naming patterns.
- **Side Effects**: none

## Function: CreateDeliveredComparison
- **Purpose**: Legacy builder for delivered-volume comparison tables (retired feature path).
- **Inputs**:
  - `ws` (`Worksheet`): Worksheet being processed
  - `emailTable` (`ListObject`): Parameter `emailTable`
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Sub `CreateDeliveredComparison`. Signature: `Private Sub CreateDeliveredComparison( ByVal ws As Worksheet, ByVal emailTable As ListObject)`. Legacy builder for delivered-volume comparison tables (retired feature path).
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: CreateDeliveredComparisonChart
- **Purpose**: Legacy builder for delivered-volume comparison charts (retired feature path).
- **Inputs**:
  - `ws` (`Worksheet`): Worksheet being processed
  - `comparisonTable` (`ListObject`): Parameter `comparisonTable`
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Sub `CreateDeliveredComparisonChart`. Signature: `Private Sub CreateDeliveredComparisonChart( ByVal ws As Worksheet, ByVal comparisonTable As ListObject)`. Legacy builder for delivered-volume comparison charts (retired feature path).
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: StyleDashboard
- **Purpose**: Applies fonts, column widths, number formats, and visual styling to the Dashboard sheet.
- **Inputs**:
  - `ws` (`Worksheet`): Worksheet being processed
  - `dashboardTable` (`ListObject`): DashboardWorkTable ListObject
  - `summaryRow` (`Long`): Parameter `summaryRow`
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Sub `StyleDashboard`. Signature: `Private Sub StyleDashboard( ByVal ws As Worksheet, ByVal dashboardTable As ListObject, ByVal summaryRow As Long)`. Applies fonts, column widths, number formats, and visual styling to the Dashboard sheet.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: ApplyDashboardStatusFormatting
- **Purpose**: Applies conditional/status formatting to Approval and Segments columns on the Dashboard table.
- **Inputs**:
  - `dashboardTable` (`ListObject`): DashboardWorkTable ListObject
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Sub `ApplyDashboardStatusFormatting`. Signature: `Private Sub ApplyDashboardStatusFormatting(ByVal dashboardTable As ListObject)`. Applies conditional/status formatting to Approval and Segments columns on the Dashboard table.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: AddDashboardLink
- **Purpose**: Creates or updates an internal hyperlink cell on the Dashboard (e.g., calendar navigation).
- **Inputs**:
  - `targetCell` (`Range`): Checklist cell to set
  - `linkAddress` (`String`): URL string for the campaign link
  - `displayText` (`String`): Parameter `displayText`
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Sub `AddDashboardLink`. Signature: `Private Sub AddDashboardLink( ByVal targetCell As Range, ByVal linkAddress As String, ByVal displayText As String)`. Creates or updates an internal hyperlink cell on the Dashboard (e.g., calendar navigation).
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: CreateDailyDigest
- **Purpose**: Builds a text digest of outstanding email/SMS campaigns with stage and risk context for logging or messaging.
- **Inputs**:
  - None
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Public Sub `CreateDailyDigest`. Signature: `Public Sub CreateDailyDigest()`. Builds a text digest of outstanding email/SMS campaigns with stage and risk context for logging or messaging.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: LogAction
- **Purpose**: Appends a timestamped action row to AutomationLogTable on the Automation Log sheet.
- **Inputs**:
  - `actionName` (`String`): Short log action label
  - `details` (`String`): Log detail text
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Public Sub `LogAction`. Signature: `Public Sub LogAction(ByVal actionName As String, ByVal details As String)`. Appends a timestamped action row to AutomationLogTable on the Automation Log sheet.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: InventoryColumnNumber
- **Purpose**: Returns the absolute worksheet column number for a named header on EmailCampaignsTable.
- **Inputs**:
  - `headerName` (`String`): Table column header text
- **Outputs**: `Long` — return value of `InventoryColumnNumber`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Public Function `InventoryColumnNumber`. Signature: `Public Function InventoryColumnNumber(ByVal headerName As String) As Long`. Returns the absolute worksheet column number for a named header on EmailCampaignsTable.
- **Side Effects**: none

## Function: AddChecklistColumns
- **Purpose**: Ensures required checklist ListColumns exist on a campaign table and configures them.
- **Inputs**:
  - `lo` (`ListObject`): Excel ListObject (table)
  - `checklistHeaders` (`Variant`): Array of checklist column headers
  - `insertAfterHeader` (`String`): Parameter `insertAfterHeader`
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Sub `AddChecklistColumns`. Signature: `Private Sub AddChecklistColumns( ByVal lo As ListObject, ByVal checklistHeaders As Variant, ByVal insertAfterHeader As String)`. Ensures required checklist ListColumns exist on a campaign table and configures them.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: ConfigureChecklistColumn
- **Purpose**: Configures one checklist column with native or legacy checkbox display and defaults.
- **Inputs**:
  - `lo` (`ListObject`): Excel ListObject (table)
  - `headerName` (`String`): Table column header text
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Sub `ConfigureChecklistColumn`. Signature: `Private Sub ConfigureChecklistColumn( ByVal lo As ListObject, ByVal headerName As String)`. Configures one checklist column with native or legacy checkbox display and defaults.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: ApplyNativeCheckboxControl
- **Purpose**: Attempts to enable Excel native cell checkbox control on a range; returns success Boolean.
- **Inputs**:
  - `rng` (`Range`): Target range for formatting/validation
- **Outputs**: `Boolean` — return value of `ApplyNativeCheckboxControl`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Function `ApplyNativeCheckboxControl`. Signature: `Private Function ApplyNativeCheckboxControl(ByVal rng As Range) As Boolean`. Attempts to enable Excel native cell checkbox control on a range; returns success Boolean.
- **Side Effects**: none

## Function: CheckboxControlType
- **Purpose**: Returns the CellControl type code for a range (2 indicates native checkbox when available).
- **Inputs**:
  - `rng` (`Range`): Target range for formatting/validation
- **Outputs**: `Long` — return value of `CheckboxControlType`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Function `CheckboxControlType`. Signature: `Private Function CheckboxControlType(ByVal rng As Range) As Long`. Returns the CellControl type code for a range (2 indicates native checkbox when available).
- **Side Effects**: none

## Function: ApplyLegacyCheckboxDisplay
- **Purpose**: Applies checkbox-like number formats/symbols for environments without native checkboxes.
- **Inputs**:
  - `rng` (`Range`): Target range for formatting/validation
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Sub `ApplyLegacyCheckboxDisplay`. Signature: `Private Sub ApplyLegacyCheckboxDisplay(ByVal rng As Range)`. Applies checkbox-like number formats/symbols for environments without native checkboxes.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: SetChecklistValue
- **Purpose**: Sets a checklist cell to complete/incomplete using native control or legacy boolean/symbol values.
- **Inputs**:
  - `targetCell` (`Range`): Checklist cell to set
  - `isComplete` (`Boolean`): True when checklist item should be marked complete
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Sub `SetChecklistValue`. Signature: `Private Sub SetChecklistValue(ByVal targetCell As Range, ByVal isComplete As Boolean)`. Sets a checklist cell to complete/incomplete using native control or legacy boolean/symbol values.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: LegacyCheckboxNumberFormat
- **Purpose**: Returns the custom number format string used for legacy checkbox display.
- **Inputs**:
  - None
- **Outputs**: `String` — return value of `LegacyCheckboxNumberFormat`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Function `LegacyCheckboxNumberFormat`. Signature: `Private Function LegacyCheckboxNumberFormat() As String`. Returns the custom number format string used for legacy checkbox display.
- **Side Effects**: none

## Function: EnsureNotesColumn
- **Purpose**: Ensures a Notes column exists at the end of a campaign table with wrap text formatting.
- **Inputs**:
  - `lo` (`ListObject`): Excel ListObject (table)
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Sub `EnsureNotesColumn`. Signature: `Private Sub EnsureNotesColumn(ByVal lo As ListObject)`. Ensures a Notes column exists at the end of a campaign table with wrap text formatting.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: ConfigureOwnerColumn
- **Purpose**: Configures Owner column validation/list behavior on a campaign table.
- **Inputs**:
  - `lo` (`ListObject`): Excel ListObject (table)
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Sub `ConfigureOwnerColumn`. Signature: `Private Sub ConfigureOwnerColumn(ByVal lo As ListObject)`. Configures Owner column validation/list behavior on a campaign table.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: ConfigureCampaignTypeColumn
- **Purpose**: Applies Campaign Type data validation from Dropdowns sheet or hardcoded options, with Others prompt support.
- **Inputs**:
  - `lo` (`ListObject`): Excel ListObject (table)
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Sub `ConfigureCampaignTypeColumn`. Signature: `Private Sub ConfigureCampaignTypeColumn(ByVal lo As ListObject)`. Applies Campaign Type data validation from Dropdowns sheet or hardcoded options, with Others prompt support.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: ConfigureSmsCampaignTypeColumn
- **Purpose**: Delegates SMS Campaign Type configuration to ConfigureCampaignTypeColumn.
- **Inputs**:
  - `lo` (`ListObject`): Excel ListObject (table)
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Sub `ConfigureSmsCampaignTypeColumn`. Signature: `Private Sub ConfigureSmsCampaignTypeColumn(ByVal lo As ListObject)`. Delegates SMS Campaign Type configuration to ConfigureCampaignTypeColumn.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: CampaignTypeOptions
- **Purpose**: Returns the fixed array of campaign type strings (Promo, Services, Loyalty, PLCC, Newsletters, Events, NPA, Others).
- **Inputs**:
  - None
- **Outputs**: `Variant` — return value of `CampaignTypeOptions`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Function `CampaignTypeOptions`. Signature: `Private Function CampaignTypeOptions() As Variant`. Returns the fixed array of campaign type strings (Promo, Services, Loyalty, PLCC, Newsletters, Events, NPA, Others).
- **Side Effects**: none

## Function: CampaignTypeOptionsRange
- **Purpose**: Returns the Dropdowns sheet range used as Campaign Type list source.
- **Inputs**:
  - None
- **Outputs**: `Range` — return value of `CampaignTypeOptionsRange`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Function `CampaignTypeOptionsRange`. Signature: `Private Function CampaignTypeOptionsRange() As Range`. Returns the Dropdowns sheet range used as Campaign Type list source.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: EnsureUpdatedByColumn
- **Purpose**: Ensures Last Updated and Last Updated By columns exist and are formatted.
- **Inputs**:
  - `lo` (`ListObject`): Excel ListObject (table)
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Sub `EnsureUpdatedByColumn`. Signature: `Private Sub EnsureUpdatedByColumn(ByVal lo As ListObject)`. Ensures Last Updated and Last Updated By columns exist and are formatted.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: HandleInventorySelection
- **Purpose**: Compatibility no-op; native checkboxes handle selection clicks.
- **Inputs**:
  - `Target` (`Range`): Changed or double-clicked cell range
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Public Sub `HandleInventorySelection`. Signature: `Public Sub HandleInventorySelection(ByVal Target As Range)`. Compatibility no-op; native checkboxes handle selection clicks.
- **Side Effects**: none

## Function: ToggleInventoryChecklist
- **Purpose**: Double-click fallback that toggles a checklist cell and returns whether the click was handled (Cancel).
- **Inputs**:
  - `Target` (`Range`): Changed or double-clicked cell range
- **Outputs**: `Boolean` — return value of `ToggleInventoryChecklist`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Public Function `ToggleInventoryChecklist`. Signature: `Public Function ToggleInventoryChecklist(ByVal Target As Range) As Boolean`. Double-click fallback that toggles a checklist cell and returns whether the click was handled (Cancel).
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: HandleInventoryChange
- **Purpose**: Compatibility wrapper that forwards to HandleCampaignChange for the target worksheet.
- **Inputs**:
  - `Target` (`Range`): Changed or double-clicked cell range
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Public Sub `HandleInventoryChange`. Signature: `Public Sub HandleInventoryChange(ByVal Target As Range)`. Compatibility wrapper that forwards to HandleCampaignChange for the target worksheet.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: HandleCampaignChange
- **Purpose**: Central change handler: stamps audit fields, formats hyperlinks, and refreshes native dashboard outputs for edited campaign rows.
- **Inputs**:
  - `ws` (`Worksheet`): Worksheet being processed
  - `Target` (`Range`): Changed or double-clicked cell range
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Public Sub `HandleCampaignChange`. Signature: `Public Sub HandleCampaignChange( ByVal ws As Worksheet, ByVal Target As Range)`. Central change handler: stamps audit fields, formats hyperlinks, and refreshes native dashboard outputs for edited campaign rows.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: CalculateStageRow
- **Purpose**: Forces recalculation/update of the Current Stage cell for one table data row.
- **Inputs**:
  - `rowNumber` (`Long`): Absolute worksheet row number
  - `lo` (`ListObject`): Excel ListObject (table)
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Sub `CalculateStageRow`. Signature: `Private Sub CalculateStageRow(ByVal rowNumber As Long, ByVal lo As ListObject)`. Forces recalculation/update of the Current Stage cell for one table data row.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: UpdateRowTimestampAndUser
- **Purpose**: Writes Now and CurrentUserName into Last Updated / Last Updated By for a table row.
- **Inputs**:
  - `rowNumber` (`Long`): Absolute worksheet row number
  - `lo` (`ListObject`): Excel ListObject (table)
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Public Sub `UpdateRowTimestampAndUser`. Signature: `Public Sub UpdateRowTimestampAndUser(ByVal rowNumber As Long, ByVal lo As ListObject)`. Writes Now and CurrentUserName into Last Updated / Last Updated By for a table row.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: IsUserEditableColumn
- **Purpose**: Returns True when a changed column should trigger audit stamping (excludes calculated/audit columns).
- **Inputs**:
  - `lo` (`ListObject`): Excel ListObject (table)
  - `absoluteColumn` (`Long`): Parameter `absoluteColumn`
- **Outputs**: `Boolean` — return value of `IsUserEditableColumn`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Function `IsUserEditableColumn`. Signature: `Private Function IsUserEditableColumn( ByVal lo As ListObject, ByVal absoluteColumn As Long) As Boolean`. Returns True when a changed column should trigger audit stamping (excludes calculated/audit columns).
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: CurrentUserName
- **Purpose**: Returns the Windows username from Environ, with Application.UserName fallback.
- **Inputs**:
  - None
- **Outputs**: `String` — return value of `CurrentUserName`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Function `CurrentUserName`. Signature: `Private Function CurrentUserName() As String`. Returns the Windows username from Environ, with Application.UserName fallback.
- **Side Effects**: none

## Function: UpdateCalendarTabs
- **Purpose**: Compatibility no-op; programmatic monthly calendar rebuild is intentionally retired.
- **Inputs**:
  - None
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Public Sub `UpdateCalendarTabs`. Signature: `Public Sub UpdateCalendarTabs()`. Compatibility no-op; programmatic monthly calendar rebuild is intentionally retired.
- **Side Effects**: none

## Function: ApplyCampaignEntryFormats
- **Purpose**: Applies Send Date/Time formats and calculated column formulas to email, SMS, and dashboard tables.
- **Inputs**:
  - None
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Public Sub `ApplyCampaignEntryFormats`. Signature: `Public Sub ApplyCampaignEntryFormats()`. Applies Send Date/Time formats and calculated column formulas to email, SMS, and dashboard tables.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: FormatSendDateColumn
- **Purpose**: Sets Send Date number format to long weekday form (dddd, mmmm d, yyyy).
- **Inputs**:
  - `lo` (`ListObject`): Excel ListObject (table)
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Sub `FormatSendDateColumn`. Signature: `Private Sub FormatSendDateColumn(ByVal lo As ListObject)`. Sets Send Date number format to long weekday form (dddd, mmmm d, yyyy).
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: FormatSendTimeColumn
- **Purpose**: Sets Send Time number format for 12-hour times while allowing text labels like STO.
- **Inputs**:
  - `lo` (`ListObject`): Excel ListObject (table)
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Sub `FormatSendTimeColumn`. Signature: `Private Sub FormatSendTimeColumn(ByVal lo As ListObject)`. Sets Send Time number format for 12-hour times while allowing text labels like STO.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: ApplyCalculatedColumns
- **Purpose**: Installs structured formulas for Current Stage (and related calculated fields) on a campaign table.
- **Inputs**:
  - `lo` (`ListObject`): Excel ListObject (table)
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Sub `ApplyCalculatedColumns`. Signature: `Private Sub ApplyCalculatedColumns(ByVal lo As ListObject)`. Installs structured formulas for Current Stage (and related calculated fields) on a campaign table.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: CurrentStageFormulaText
- **Purpose**: Returns the structured formula text that lists all checked workflow stages for email or SMS.
- **Inputs**:
  - `isSmsTable` (`Boolean`): True for SMS checklist formulas; False for email
- **Outputs**: `String` — return value of `CurrentStageFormulaText`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Function `CurrentStageFormulaText`. Signature: `Private Function CurrentStageFormulaText(ByVal isSmsTable As Boolean) As String`. Returns the structured formula text that lists all checked workflow stages for email or SMS.
- **Side Effects**: none

## Function: ChecklistStatusFormulaText
- **Purpose**: Returns the TEXTJOIN/IF checklist aggregation fragment for email or SMS checklist columns.
- **Inputs**:
  - `isSmsTable` (`Boolean`): True for SMS checklist formulas; False for email
- **Outputs**: `String` — return value of `ChecklistStatusFormulaText`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Function `ChecklistStatusFormulaText`. Signature: `Private Function ChecklistStatusFormulaText(ByVal isSmsTable As Boolean) As String`. Returns the TEXTJOIN/IF checklist aggregation fragment for email or SMS checklist columns.
- **Side Effects**: none

## Function: ChecklistFormulaItem
- **Purpose**: Builds one IF(structuredRef) contribution used inside checklist aggregation formulas.
- **Inputs**:
  - `structuredReference` (`String`): Parameter `structuredReference`
  - `labelText` (`String`): Parameter `labelText`
- **Outputs**: `String` — return value of `ChecklistFormulaItem`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Function `ChecklistFormulaItem`. Signature: `Private Function ChecklistFormulaItem( ByVal structuredReference As String, ByVal labelText As String) As String`. Builds one IF(structuredRef) contribution used inside checklist aggregation formulas.
- **Side Effects**: none

## Function: FormulaString
- **Purpose**: Wraps a string in Excel double quotes for safe formula concatenation.
- **Inputs**:
  - `TextValue` (`String`): Literal text to quote for formula use
- **Outputs**: `String` — return value of `FormulaString`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Function `FormulaString`. Signature: `Private Function FormulaString(ByVal TextValue As String) As String`. Wraps a string in Excel double quotes for safe formula concatenation.
- **Side Effects**: none

## Function: FormatLastUpdatedColumn
- **Purpose**: Applies TIMESTAMP_FORMAT to a Last Updated ListColumn body range.
- **Inputs**:
  - `updatedColumn` (`ListColumn`): Last Updated ListColumn
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Sub `FormatLastUpdatedColumn`. Signature: `Private Sub FormatLastUpdatedColumn(ByVal updatedColumn As ListColumn)`. Applies TIMESTAMP_FORMAT to a Last Updated ListColumn body range.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: DeleteTableColumnIfPresent
- **Purpose**: Deletes a ListColumn by header name when present; no error if missing.
- **Inputs**:
  - `lo` (`ListObject`): Excel ListObject (table)
  - `headerName` (`String`): Table column header text
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Sub `DeleteTableColumnIfPresent`. Signature: `Private Sub DeleteTableColumnIfPresent( ByVal lo As ListObject, ByVal headerName As String)`. Deletes a ListColumn by header name when present; no error if missing.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: RequireTableColumn
- **Purpose**: Raises a controlled failure when a required table header is missing.
- **Inputs**:
  - `lo` (`ListObject`): Excel ListObject (table)
  - `headerName` (`String`): Table column header text
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Sub `RequireTableColumn`. Signature: `Private Sub RequireTableColumn( ByVal lo As ListObject, ByVal headerName As String)`. Raises a controlled failure when a required table header is missing.
- **Side Effects**: none

## Function: EnsureCampaignSheets
- **Purpose**: Creates or repairs core sheets (Email, SMS, Dashboard, Notes, Dropdowns, Log) and base tables if missing.
- **Inputs**:
  - None
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Sub `EnsureCampaignSheets`. Signature: `Private Sub EnsureCampaignSheets()`. Creates or repairs core sheets (Email, SMS, Dashboard, Notes, Dropdowns, Log) and base tables if missing.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: GetInventoryTable
- **Purpose**: Returns EmailCampaignsTable ListObject from the Email Campaigns sheet.
- **Inputs**:
  - None
- **Outputs**: `ListObject` — return value of `GetInventoryTable`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Function `GetInventoryTable`. Signature: `Private Function GetInventoryTable() As ListObject`. Returns EmailCampaignsTable ListObject from the Email Campaigns sheet.
- **Side Effects**: none

## Function: GetSmsTable
- **Purpose**: Returns SMSCampaignsTable ListObject from the SMS Campaigns sheet.
- **Inputs**:
  - None
- **Outputs**: `ListObject` — return value of `GetSmsTable`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Function `GetSmsTable`. Signature: `Private Function GetSmsTable() As ListObject`. Returns SMSCampaignsTable ListObject from the SMS Campaigns sheet.
- **Side Effects**: none

## Function: FindTableColumn
- **Purpose**: Finds a ListColumn by normalized header name; returns Nothing when not found.
- **Inputs**:
  - `lo` (`ListObject`): Excel ListObject (table)
  - `headerName` (`String`): Table column header text
- **Outputs**: `ListColumn` — return value of `FindTableColumn`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Function `FindTableColumn`. Signature: `Private Function FindTableColumn( ByVal lo As ListObject, ByVal headerName As String) As ListColumn`. Finds a ListColumn by normalized header name; returns Nothing when not found.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: EnsureRenamedColumn
- **Purpose**: Renames a column from an old header to a new header when the old name is present.
- **Inputs**:
  - `lo` (`ListObject`): Excel ListObject (table)
  - `oldHeader` (`String`): Previous column header name
  - `newHeader` (`String`): New column header name
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Sub `EnsureRenamedColumn`. Signature: `Private Sub EnsureRenamedColumn( ByVal lo As ListObject, ByVal oldHeader As String, ByVal newHeader As String)`. Renames a column from an old header to a new header when the old name is present.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: NormalizeHeaderKey
- **Purpose**: Trims and lowercases a header string for case-insensitive matching.
- **Inputs**:
  - `headerName` (`String`): Table column header text
- **Outputs**: `String` — return value of `NormalizeHeaderKey`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Function `NormalizeHeaderKey`. Signature: `Private Function NormalizeHeaderKey(ByVal headerName As String) As String`. Trims and lowercases a header string for case-insensitive matching.
- **Side Effects**: none

## Function: ValueByHeader
- **Purpose**: Reads the cell value at a worksheet row for a named table header.
- **Inputs**:
  - `ws` (`Worksheet`): Worksheet being processed
  - `rowNumber` (`Long`): Absolute worksheet row number
  - `headerName` (`String`): Table column header text
- **Outputs**: `Variant` — return value of `ValueByHeader`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Function `ValueByHeader`. Signature: `Private Function ValueByHeader( ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal headerName As String) As Variant`. Reads the cell value at a worksheet row for a named table header.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: CampaignTableForSheet
- **Purpose**: Returns the campaign ListObject for Email Campaigns or SMS Campaigns worksheet names.
- **Inputs**:
  - `ws` (`Worksheet`): Worksheet being processed
- **Outputs**: `ListObject` — return value of `CampaignTableForSheet`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Function `CampaignTableForSheet`. Signature: `Private Function CampaignTableForSheet(ByVal ws As Worksheet) As ListObject`. Returns the campaign ListObject for Email Campaigns or SMS Campaigns worksheet names.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: IsChecked
- **Purpose**: Interprets Boolean, checkbox, TRUE/FALSE text, and legacy symbols as a checked state.
- **Inputs**:
  - `inputValue` (`Variant`): Value to interpret
- **Outputs**: `Boolean` — return value of `IsChecked`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Function `IsChecked`. Signature: `Private Function IsChecked(ByVal inputValue As Variant) As Boolean`. Interprets Boolean, checkbox, TRUE/FALSE text, and legacy symbols as a checked state.
- **Side Effects**: none

## Function: HasText
- **Purpose**: Returns True when a Variant holds non-blank, non-error text/content.
- **Inputs**:
  - `inputValue` (`Variant`): Value to interpret
- **Outputs**: `Boolean` — return value of `HasText`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Function `HasText`. Signature: `Private Function HasText(ByVal inputValue As Variant) As Boolean`. Returns True when a Variant holds non-blank, non-error text/content.
- **Side Effects**: none

## Function: TextValue
- **Purpose**: Coerces a Variant to trimmed string, returning empty string for errors/nulls.
- **Inputs**:
  - `inputValue` (`Variant`): Value to interpret
- **Outputs**: `String` — return value of `TextValue`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Function `TextValue`. Signature: `Private Function TextValue(ByVal inputValue As Variant) As String`. Coerces a Variant to trimmed string, returning empty string for errors/nulls.
- **Side Effects**: none

## Function: CheckedSymbol
- **Purpose**: Returns the Unicode ballot-box-with-check character used in legacy checkbox display.
- **Inputs**:
  - None
- **Outputs**: `String` — return value of `CheckedSymbol`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Function `CheckedSymbol`. Signature: `Private Function CheckedSymbol() As String`. Returns the Unicode ballot-box-with-check character used in legacy checkbox display.
- **Side Effects**: none

## Function: UncheckedSymbol
- **Purpose**: Returns the Unicode ballot box character used in legacy checkbox display.
- **Inputs**:
  - None
- **Outputs**: `String` — return value of `UncheckedSymbol`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Function `UncheckedSymbol`. Signature: `Private Function UncheckedSymbol() As String`. Returns the Unicode ballot box character used in legacy checkbox display.
- **Side Effects**: none

## Function: FullColumnReference
- **Purpose**: Builds a full structured column reference string for a ListObject column.
- **Inputs**:
  - `lo` (`ListObject`): Excel ListObject (table)
  - `headerName` (`String`): Table column header text
- **Outputs**: `String` — return value of `FullColumnReference`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Function `FullColumnReference`. Signature: `Private Function FullColumnReference( ByVal lo As ListObject, ByVal headerName As String) As String`. Builds a full structured column reference string for a ListObject column.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: PendingCountFormula
- **Purpose**: Builds a formula fragment used by KPI/pending counts for a campaign reference.
- **Inputs**:
  - `campaignRef` (`String`): Parameter `campaignRef`
  - `checklistRef` (`String`): Parameter `checklistRef`
- **Outputs**: `String` — return value of `PendingCountFormula`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Function `PendingCountFormula`. Signature: `Private Function PendingCountFormula( ByVal campaignRef As String, ByVal checklistRef As String) As String`. Builds a formula fragment used by KPI/pending counts for a campaign reference.
- **Side Effects**: none

## Function: CreateBackupCopy
- **Purpose**: Saves a timestamped backup copy of the workbook next to the original and returns its path.
- **Inputs**:
  - `wb` (`Workbook`): Workbook being processed
- **Outputs**: `String` — return value of `CreateBackupCopy`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Function `CreateBackupCopy`. Signature: `Private Function CreateBackupCopy(ByVal wb As Workbook) As String`. Saves a timestamped backup copy of the workbook next to the original and returns its path.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: CountBrokenReferences
- **Purpose**: Counts cells containing broken reference error markers (#REF!) via Find.
- **Inputs**:
  - `wb` (`Workbook`): Workbook being processed
- **Outputs**: `Long` — return value of `CountBrokenReferences`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Function `CountBrokenReferences`. Signature: `Private Function CountBrokenReferences(ByVal wb As Workbook) As Long`. Counts cells containing broken reference error markers (#REF!) via Find.
- **Side Effects**: none

## Function: ValidateWorkbookConfiguration
- **Purpose**: Validates core sheets, tables, dropdown configuration, and returns a status string.
- **Inputs**:
  - None
- **Outputs**: `String` — return value of `ValidateWorkbookConfiguration`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Public Function `ValidateWorkbookConfiguration`. Signature: `Public Function ValidateWorkbookConfiguration() As String`. Validates core sheets, tables, dropdown configuration, and returns a status string.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: DashboardChartExists
- **Purpose**: Returns True when a chart object with the given name exists on the Dashboard.
- **Inputs**:
  - `chartName` (`String`): Chart object name on Dashboard
- **Outputs**: `Boolean` — return value of `DashboardChartExists`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Function `DashboardChartExists`. Signature: `Private Function DashboardChartExists(ByVal chartName As String) As Boolean`. Returns True when a chart object with the given name exists on the Dashboard.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: ValidateCampaignTypeOptions
- **Purpose**: Ensures Campaign Type dropdown source options are present and consistent.
- **Inputs**:
  - None
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Sub `ValidateCampaignTypeOptions`. Signature: `Private Sub ValidateCampaignTypeOptions()`. Ensures Campaign Type dropdown source options are present and consistent.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: CampaignTypeDropdownIsConfigured
- **Purpose**: Returns True when a range has list validation pointing at campaign types.
- **Inputs**:
  - `rng` (`Range`): Target range for formatting/validation
- **Outputs**: `Boolean` — return value of `CampaignTypeDropdownIsConfigured`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Function `CampaignTypeDropdownIsConfigured`. Signature: `Private Function CampaignTypeDropdownIsConfigured(ByVal rng As Range) As Boolean`. Returns True when a range has list validation pointing at campaign types.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: ValidateCampaignTable
- **Purpose**: Validates required columns and checklist configuration on one campaign ListObject.
- **Inputs**:
  - `lo` (`ListObject`): Excel ListObject (table)
  - `requiredHeaders` (`Variant`): Array of required table headers
  - `checklistHeaders` (`Variant`): Array of checklist column headers
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Sub `ValidateCampaignTable`. Signature: `Private Sub ValidateCampaignTable( ByVal lo As ListObject, ByVal requiredHeaders As Variant, ByVal checklistHeaders As Variant)`. Validates required columns and checklist configuration on one campaign ListObject.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: RangeHasValidation
- **Purpose**: Returns True when the range has any data validation rule.
- **Inputs**:
  - `rng` (`Range`): Target range for formatting/validation
- **Outputs**: `Boolean` — return value of `RangeHasValidation`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Function `RangeHasValidation`. Signature: `Private Function RangeHasValidation(ByVal rng As Range) As Boolean`. Returns True when the range has any data validation rule.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: RebuildMonthlyCalendars
- **Purpose**: Compatibility no-op; monthly calendar generation is intentionally retired in favor of SharePoint-linked sheets.
- **Inputs**:
  - None
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Public Sub `RebuildMonthlyCalendars`. Signature: `Public Sub RebuildMonthlyCalendars()`. Compatibility no-op; monthly calendar generation is intentionally retired in favor of SharePoint-linked sheets.
- **Side Effects**: none

## Function: GetOrCreateCalendarSheet
- **Purpose**: Legacy helper that gets or creates a monthly calendar worksheet by name.
- **Inputs**:
  - `wb` (`Workbook`): Workbook being processed
  - `monthNumber` (`Long`): Parameter `monthNumber`
- **Outputs**: `Worksheet` — return value of `GetOrCreateCalendarSheet`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Function `GetOrCreateCalendarSheet`. Signature: `Private Function GetOrCreateCalendarSheet( ByVal wb As Workbook, ByVal monthNumber As Long) As Worksheet`. Legacy helper that gets or creates a monthly calendar worksheet by name.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: BuildCalendarSheet
- **Purpose**: Legacy builder that lays out a monthly calendar grid and formulas.
- **Inputs**:
  - `ws` (`Worksheet`): Worksheet being processed
  - `monthNumber` (`Long`): Parameter `monthNumber`
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Sub `BuildCalendarSheet`. Signature: `Private Sub BuildCalendarSheet( ByVal ws As Worksheet, ByVal monthNumber As Long)`. Legacy builder that lays out a monthly calendar grid and formulas.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: SetDynamicFormula
- **Purpose**: Writes a formula to a range with error-safe Formula2/Formula assignment.
- **Inputs**:
  - `targetRange` (`Range`): Range that receives a formula
  - `formulaText` (`String`): Excel formula string to write
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Sub `SetDynamicFormula`. Signature: `Private Sub SetDynamicFormula( ByVal targetRange As Range, ByVal formulaText As String)`. Writes a formula to a range with error-safe Formula2/Formula assignment.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: ClearCalendarSheet
- **Purpose**: Clears contents and formatting from a calendar worksheet prior to rebuild.
- **Inputs**:
  - `ws` (`Worksheet`): Worksheet being processed
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Sub `ClearCalendarSheet`. Signature: `Private Sub ClearCalendarSheet(ByVal ws As Worksheet)`. Clears contents and formatting from a calendar worksheet prior to rebuild.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: StyleCalendarSheet
- **Purpose**: Applies layout, colors, and typography to a monthly calendar sheet.
- **Inputs**:
  - `ws` (`Worksheet`): Worksheet being processed
  - `monthNumber` (`Long`): Parameter `monthNumber`
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Sub `StyleCalendarSheet`. Signature: `Private Sub StyleCalendarSheet( ByVal ws As Worksheet, ByVal monthNumber As Long)`. Applies layout, colors, and typography to a monthly calendar sheet.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: AddDashboardCalendarLinks
- **Purpose**: Adds Dashboard hyperlinks that navigate to monthly calendar sheets.
- **Inputs**:
  - `wb` (`Workbook`): Workbook being processed
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Sub `AddDashboardCalendarLinks`. Signature: `Private Sub AddDashboardCalendarLinks(ByVal wb As Workbook)`. Adds Dashboard hyperlinks that navigate to monthly calendar sheets.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: GetOrCreateNotesInstructionSheet
- **Purpose**: Returns the Notes - Instructions worksheet, creating it if missing.
- **Inputs**:
  - `wb` (`Workbook`): Workbook being processed
- **Outputs**: `Worksheet` — return value of `GetOrCreateNotesInstructionSheet`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Function `GetOrCreateNotesInstructionSheet`. Signature: `Private Function GetOrCreateNotesInstructionSheet(ByVal wb As Workbook) As Worksheet`. Returns the Notes - Instructions worksheet, creating it if missing.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: ApplyWorkbookWrapText
- **Purpose**: Enables wrap text on key used ranges across core workbook sheets.
- **Inputs**:
  - `wb` (`Workbook`): Workbook being processed
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Sub `ApplyWorkbookWrapText`. Signature: `Private Sub ApplyWorkbookWrapText(ByVal wb As Workbook)`. Enables wrap text on key used ranges across core workbook sheets.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: InstructionSheetIsReady
- **Purpose**: Returns True when Notes - Instructions already contains the expected instruction table content.
- **Inputs**:
  - None
- **Outputs**: `Boolean` — return value of `InstructionSheetIsReady`
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Function `InstructionSheetIsReady`. Signature: `Private Function InstructionSheetIsReady() As Boolean`. Returns True when Notes - Instructions already contains the expected instruction table content.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: StyleCoreWorkbookSheets
- **Purpose**: Applies shared styling to Dashboard, Email Campaigns, SMS Campaigns, and related core sheets.
- **Inputs**:
  - `wb` (`Workbook`): Workbook being processed
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Sub `StyleCoreWorkbookSheets`. Signature: `Private Sub StyleCoreWorkbookSheets(ByVal wb As Workbook)`. Applies shared styling to Dashboard, Email Campaigns, SMS Campaigns, and related core sheets.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: StyleCampaignSheet
- **Purpose**: Applies header, filter, column width, and checkbox styling to a campaign sheet table.
- **Inputs**:
  - `ws` (`Worksheet`): Worksheet being processed
  - `lo` (`ListObject`): Excel ListObject (table)
  - `tabColor` (`Long`): Parameter `tabColor`
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Sub `StyleCampaignSheet`. Signature: `Private Sub StyleCampaignSheet( ByVal ws As Worksheet, ByVal lo As ListObject, ByVal tabColor As Long)`. Applies header, filter, column width, and checkbox styling to a campaign sheet table.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: OrderWorkbookSheets
- **Purpose**: Reorders worksheets into the canonical tab sequence.
- **Inputs**:
  - `wb` (`Workbook`): Workbook being processed
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Sub `OrderWorkbookSheets`. Signature: `Private Sub OrderWorkbookSheets(ByVal wb As Workbook)`. Reorders worksheets into the canonical tab sequence.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: UnfreezeWorkbookViews
- **Purpose**: Clears freeze panes and split settings on every worksheet in the workbook.
- **Inputs**:
  - `wb` (`Workbook`): Workbook being processed
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Public Sub `UnfreezeWorkbookViews`. Signature: `Public Sub UnfreezeWorkbookViews(ByVal wb As Workbook)`. Clears freeze panes and split settings on every worksheet in the workbook.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: ConfigureWorkbookViews
- **Purpose**: Sets preferred zoom/view settings for core sheets after structural updates.
- **Inputs**:
  - `wb` (`Workbook`): Workbook being processed
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Sub `ConfigureWorkbookViews`. Signature: `Private Sub ConfigureWorkbookViews(ByVal wb As Workbook)`. Sets preferred zoom/view settings for core sheets after structural updates.
- **Side Effects**: none

## Function: SetInternalLink
- **Purpose**: Creates an internal worksheet hyperlink with display text on a target cell.
- **Inputs**:
  - `ws` (`Worksheet`): Worksheet being processed
  - `targetCell` (`Range`): Checklist cell to set
  - `displayText` (`String`): Parameter `displayText`
  - `subAddress` (`String`): Parameter `subAddress`
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Sub `SetInternalLink`. Signature: `Private Sub SetInternalLink( ByVal ws As Worksheet, ByVal targetCell As Range, ByVal displayText As String, ByVal subAddress As String)`. Creates an internal worksheet hyperlink with display text on a target cell.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: DeleteSheetIfPresent
- **Purpose**: Deletes a worksheet by name when present, suppressing alerts.
- **Inputs**:
  - `wb` (`Workbook`): Workbook being processed
  - `sheetName` (`String`): Worksheet name to evaluate
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Sub `DeleteSheetIfPresent`. Signature: `Private Sub DeleteSheetIfPresent( ByVal wb As Workbook, ByVal sheetName As String)`. Deletes a worksheet by name when present, suppressing alerts.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: BuildNotesInstructionSheet
- **Purpose**: Rebuilds the end-user Notes - Instructions feature table content.
- **Inputs**:
  - `wb` (`Workbook`): Workbook being processed
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Sub `BuildNotesInstructionSheet`. Signature: `Private Sub BuildNotesInstructionSheet(ByRef wb As Workbook)`. Rebuilds the end-user Notes - Instructions feature table content.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: AddInstructionRow
- **Purpose**: Writes one feature-guide row (feature, action, description, do/avoid, limits) onto Notes - Instructions.
- **Inputs**:
  - `ws` (`Worksheet`): Worksheet being processed
  - `rowNum` (`Long`): Destination row on Notes - Instructions
  - `comp` (`String`): Feature/component name for instruction row
  - `act` (`String`): User action description
  - `desc` (`String`): Plain-language how-it-works text
  - `dep` (`String`): Please do / please avoid guidance
  - `lim` (`String`): Where it works / limits text
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Sub `AddInstructionRow`. Signature: `Private Sub AddInstructionRow(ByRef ws As Worksheet, ByRef rowNum As Long, ByVal comp As String, ByVal act As String, ByVal desc As String, ByVal dep As String, ByVal lim As String)`. Writes one feature-guide row (feature, action, description, do/avoid, limits) onto Notes - Instructions.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

## Function: PromptCustomCampaignType
- **Purpose**: Prompts the user for a custom Campaign Type value when Others is selected.
- **Inputs**:
  - `lo` (`ListObject`): Excel ListObject (table)
  - `changedCells` (`Range`): Subset of changed cells inside the table
- **Outputs**: None (Sub)
- **Dependencies**: Excel object model; campaign tables/sheets; sibling procedures in `modEmailProductionTracker`
- **Behavior**: Private Sub `PromptCustomCampaignType`. Signature: `Private Sub PromptCustomCampaignType( ByVal lo As ListObject, ByVal changedCells As Range)`. Prompts the user for a custom Campaign Type value when Others is selected.
- **Side Effects**: Modifies workbook state (cells, formulas, formats, tables, sheets, Application settings, and/or Automation Log)

# Module / File: Loyalty and PLCC Email Plan Template.xlsx

## Function: (worksheet formulas — no VBA)
- **Purpose**: Plan Loyalty/PLCC email projects on a tracker grid and mirror them onto monthly calendars.
- **Inputs**:
  - User-entered rows on `Project Tracker` (`Jira Ticket Link`, `Campaign`, `Status`, `Send Date`, `Notes`)
  - Helper columns `Calendar Key`, `Calendar Display`, `Future Send Helper` (formula-driven)
- **Outputs**: Monthly calendar day cells show campaign display text via `VLOOKUP` on calendar keys
- **Dependencies**: Excel formulas only; sheets `Dashboard`, `Project Tracker`, twelve `* 2027 Calendar` sheets, `Lists`
- **Behavior**: Each calendar day formula looks up `TEXT(date,"yyyymmdd")&"|n"` against `Project Tracker!$F$2:$G$201` and concatenates matching display values for multi-campaign days.
- **Side Effects**: none (recalculation only)

# QA Snapshot

## Result: `QA_PASSED`

Static QA of repository content (2026-07-19):
- VBA project present and extractable from all three `.xlsm` workbooks; core module signatures consistent across Active/Template/Backup.
- Required tables and core sheets present on Active/Template/Backup.
- Native Dashboard spill + KPI formulas present; cancelled-campaign exclusion encoded in KPI formulas.
- Loyalty template opens as formula-only `.xlsx` with 12 monthly calendars + Project Tracker table.

## Non-blocking observations
- Active tracker missing leading `Week Number` column present on Template/Backup (document structural variance above).
- `Adorama/Reporting Analysis/` directory is empty (placeholder).
- No `tools/*.py` automation package is present despite older setup notes referring to it.
- Notes - Instructions unlock password is operational (accidental-edit protection, not encryption): `Adorama@042026_` (active/template); legacy backup may still use `adorama2024`.


# Module / File: Others/Personal Finance Tracker.xlsx

## Function: Interest Rate Savings Projection
- **Purpose**: Calculates projected net savings including user-specified annual interest rate over a dynamic period.
- **Inputs**:
  - `Income` (`currency`): Total income over the selected period.
  - `Expense` (`currency`): Total expense over the selected period.
  - `Annual Interest Rate` (`percentage`): User input cell at Dashboard `L4`.
  - `Months` (`integer`): Number of months in the selected period (calculated at Dashboard `N6`).
- **Outputs**: Projected Net Savings including simple interest (Dashboard `H7`).
- **Dependencies**: `Settings` configuration sheet, `Transactions` sheet data.
- **Behavior**: Computes base net savings (`Income - Expense`) and multiplies it by `(1 + (Annual Rate / 12) * Months)`.
- **Side Effects**: None.
