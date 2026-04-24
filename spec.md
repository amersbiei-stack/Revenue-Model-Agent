# Revenue Model — Monthly Rollover Agent

## Context

Automates the monthly rollover of the Revenue Model Excel file. The file is rolled over approximately 4 days after month-end close. The agent handles archiving, renaming, formula rolling, Prophix data refresh, validation checks, and executive summary email delivery — fully end to end.

| Detail | Value |
|---|---|
| File Location | `C:\Users\amers\Downloads\Revenue Model\` |
| Archive Folder | `C:\Users\amers\Downloads\Revenue Model\Archive\` |
| File Name Format | `Revenue Model - MMDDYYYY (Internal).xlsm` |
| Current File Name Example | `Revenue Model - 04062026 (Internal).xlsm` |
| Rollover Timing | Approximately 4 days after month-end close |
| Email Recipient | asbiei@prophix.com |
| File Link in Email | Local file path (SharePoint URL to be added after testing) |

---

## Step 1: Archive and Rename the File

Implemented in pure Python (`agent/step1_archive_rename.py`) — no Excel/VBA involvement. Chosen for stability: no Trust-Center prompts, no macro-enable dialogs, no dependence on the hard-coded SharePoint path inside `Module1`, and no risk of the workbook being held open read-only by Excel.

#### 1a — Archive
- Source: `C:\Users\amers\Downloads\Revenue Model\Revenue Model - MMDDYYYY (Internal).xlsm`
- **Copy** (`shutil.copy2`, preserving timestamps) to: `C:\Users\amers\Downloads\Revenue Model\Archive\`
- Archive filename inserts `" Backup"` before the extension: `Revenue Model - MMDDYYYY (Internal) Backup.xlsm`.
- Collision-safe: if `" Backup.xlsm"` already exists, try `" Backup (2).xlsm"`, `" Backup (3).xlsm"`, ... until a free name is found.
- Post-check: the new backup file must exist and be non-empty before proceeding to 1b. If it is missing the step raises `RuntimeError` and halts.

#### 1b — Rename the Working File
- Target name uses today's actual system date (not the close-period date): `Revenue Model - MMDDYYYY (Internal).xlsm`.
- **Idempotent:** if the source path is already today's dated name, the rename is a no-op (the archive copy was still made in 1a, so the safety net is intact).
- **Same-day re-run:** if a different file already occupies today's dated path, it is deleted and overwritten by the fresh rename.
- After renaming, the main folder must contain the newly renamed file; the source path must no longer exist (unless source == target per the idempotent case).
- All subsequent steps reference the returned renamed path, not the archived copy.

#### Failure modes
- **File locked** (Excel has the workbook open): `shutil.copy2` or `Path.rename` raises `PermissionError`. Close Excel and retry. The agent does not attempt to force a close.
- **Archive folder missing**: created automatically via `mkdir(parents=True, exist_ok=True)`.
- **Source missing**: `FileNotFoundError` is raised immediately; no partial state.

---

## Step 2: Open the File
- Open the renamed file from: `C:\Users\amers\Downloads\Revenue Model\`
- If the file is already open by another user, proceed anyway. The archive backup already exists and it is safe to continue.

---

## Step 3: Roll Formulas Month-over-Month

### Option 1 — Run VBA Macro (Preferred)
- Run the macro: `Step2_RollFormulas_TwoMonthsBack_SourceToTarget_Reliable` (in `Module2`).
- Wait for the macro to finish. Completion is confirmed by `TimedPopup`: `"Step 2 complete: Rolled formulas [srcLabel] -> [dstLabel]"`.
- Do NOT proceed to Step 4 until the pop-up confirmation appears.

**Date-math guard (pre-flight check — recommended):** The macro derives `srcDate = fileDate - 2 months` and `dstDate = fileDate - 1 month` from the filename. If the agent runs too early (before month-end close), `dstDate` will be two or more months behind the intended close month and the wrong months will be rolled. Before running, verify:

```
expected_dstMonth = month to populate (the closing month)
actual_dstMonth   = Month(fileDate) - 1
ASSERT expected_dstMonth == actual_dstMonth
```

If the assertion fails, stop and alert — the filename date is out of sync with the intended close.

### Option 2 — Manual Formula Drag (Fallback if Macro Fails)
- If the macro fails, manually drag formulas column by column using the `FormulaR1C1` method.
- Apply to ALL of the following 15 tabs:
  - `NA`, `NA_USD`, `NA_CAD`, `UK_GBP`, `Europe & MEA`, `Beneleux`, `DACH`, `Nordics`, `Western Europe`, `Eastern Europe`, `ME&A`, `LATAM PX`, `LATAM AS`, `Asia`, `Pacific (FG)`
- Use the following row blocks on each tab:

| Block | Row Range |
|---|---|
| 1 | 19:19 |
| 2 | 22:37 |
| 3 | 41:56 |
| 4 | 60:75 |
| 5 | 79:94 |
| 6 | 102:104 |
| 7 | 111:111 |
| 8 | 114:130 |
| 9 | 133:148 |
| 10 | 152:167 |
| 11 | 171:187 |
| 12 | 195:195 |

### VBA Macro Code Reference

```vba
Public Sub Step2_RollFormulas_TwoMonthsBack_SourceToTarget_Reliable()
  Dim wb As Workbook : Set wb = ActiveWorkbook
  If wb Is Nothing Then Exit Sub
  Dim fileDate As Date
  On Error GoTo DateFail
  fileDate = GetFileDateFromWorkbookName_MMDDYYYY(wb.Name)
  On Error GoTo 0
  Dim srcDate As Date, dstDate As Date
  srcDate = DateAdd("m", -2, DateSerial(Year(fileDate), Month(fileDate), 1))
  dstDate = DateAdd("m", -1, DateSerial(Year(fileDate), Month(fileDate), 1))
  Dim srcLabel As String, dstLabel As String
  srcLabel = Year(srcDate) & "M" & Format(Month(srcDate), "00")
  dstLabel = Year(dstDate) & "M" & Format(Month(dstDate), "00")
  Dim tabs As Variant
  tabs = Array("NA","NA_USD","NA_CAD","UK_GBP","Europe & MEA","Beneleux",
    "DACH","Nordics","Western Europe","Eastern Europe","ME&A",
    "LATAM PX","LATAM AS","Asia","Pacific (FG)")
  Dim rowBlocks As Variant : rowBlocks = GetRowBlocks()
  Application.ScreenUpdating = False
  Application.EnableEvents = False
  Dim originalSheet As Worksheet : Set originalSheet = wb.ActiveSheet
  Dim i As Long
  For i = LBound(tabs) To UBound(tabs)
    Dim ws As Worksheet : Set ws = Nothing
    On Error Resume Next
    Set ws = wb.Worksheets(CStr(tabs(i)))
    On Error GoTo 0
    If ws Is Nothing Then
      RestoreAppState originalSheet
      TimedPopup "ERROR: Sheet not found -> " & CStr(tabs(i)), 10, "Revenue Model - ERROR"
      Exit Sub
    End If
    ws.Activate : ws.Calculate
    Dim srcCol As Long, dstCol As Long
    srcCol = FindMonthColumnInRow4_Robust(ws, srcLabel)
    dstCol = FindMonthColumnInRow4_Robust(ws, dstLabel)
    If srcCol = 0 Or dstCol = 0 Then
      RestoreAppState originalSheet
      TimedPopup "ERROR: Month column not found on: " & ws.Name, 10, "Revenue Model - ERROR"
      Exit Sub
    End If
    RollFormulasByBlocks_DragFix ws, srcCol, dstCol, rowBlocks
  Next i
  RestoreAppState originalSheet
  TimedPopup "Step 2 complete: Rolled " & srcLabel & " to " & dstLabel, 5, "Revenue Model"
  Exit Sub
DateFail:
  RestoreAppState Nothing
  TimedPopup "ERROR: Could not parse date from filename: " & wb.Name, 10, "Revenue Model - ERROR"
End Sub
```

---

## Step 4: Refresh Prophix Analyzer

1. Navigate to the `Units Bookings DV` tab.
2. Press `Alt + N + Y1` to open the Prophix Analyzer CV2 add-in.
3. Locate the refresh button in the bottom section of Prophix Analyzer, first button on the left.
4. Click the arrow next to the refresh button. A menu appears with two options: **This Sheet** and **All Sheets**.
5. Select **All Sheets** and click it.
6. **Wait 10 minutes, then proceed to Step 5.**
   - VBA: `Application.Wait Now + TimeValue("00:10:00")`
   - Python: `time.sleep(600)`
   - This is a fixed sleep, not a poll. The Prophix Analyzer CV2 add-in does not expose a programmatic completion signal the agent can observe.
   - Step 5's zero-check is the primary safety net: if the refresh fails entirely, any tab with zero in the new month column will trigger the alert before Step 6 runs.
   - **Known risk:** if Prophix runs longer than 10 minutes (rare under normal load, but possible under heavy load or network latency), tabs may land partial data. Step 5's zero-check will not catch a partial tab — only a fully empty one. If historical refresh times suggest >10 minutes is plausible, consider adding a post-sleep re-check loop (see Error Handling).

---

## Step 5: Verify Data Populated Successfully

After the Prophix refresh completes, verify that data has populated in the new month column across all required tabs. Data is successfully populated if `SUM(data_range, new_month_col) > 0` on each tab. If any tab fails, stop immediately and send an alert email to asbiei@prophix.com before proceeding.

### 5a — Derive the new month label

The new month label is computed from the filename date, identical to the `dstLabel` inside the Step 2 macro:

```
fileDate      = date parsed from workbook filename (MMDDYYYY)
newMonthDate  = DateAdd("m", -1, DateSerial(Year(fileDate), Month(fileDate), 1))
newMonthLabel = Year(newMonthDate) & "M" & Format(Month(newMonthDate), "00")
```

Rule: if the filename date falls in month `N`, the new month label is month `N-1` (previous month). Year wraps are handled by `DateSerial` (January filename → previous December).

| Filename month | New month label |
|---|---|
| May 2026 | `2026M04` |
| June 2026 | `2026M05` |
| July 2026 | `2026M06` |
| January 2027 | `2026M12` |

### 5b — Locate the new month column and verify per tab

Every one of the 5 tabs below stores month headers in **row 2**. Find the new month column by matching `newMonthLabel` against row 2 of the tab (label lookup, not hard-coded column numbers — column positions differ across tabs).

| Tab | Header Row | Data Range | Check |
|---|---|---|---|
| `Units Bookings DV` | 2 | 3:152 | `SUM(row 3 : row 152, new_month_col) > 0` |
| `$$ Bookings DV` | 2 | 3:602 | `SUM(row 3 : row 602, new_month_col) > 0` |
| `Sub Europe Units` | 2 | 3:182 | `SUM(row 3 : row 182, new_month_col) > 0` |
| `Migration Units` | 2 | 3:17 | `SUM(row 3 : row 17, new_month_col) > 0` |
| `AS` | 2 | 3:122 | `SUM(row 3 : row 122, new_month_col) > 0` (covers Dollars block 3:62 + Units block 63:122) |

Notes on the tabs:
- **`Migration Units`** uses a different column base than the other four tabs (2026M04 at col BN=66 rather than BP=68). Label lookup in row 2 handles this transparently.
- **`AS`** contains two stacked blocks (`Dollars` rows 3:62 and `Units` rows 63:122). Summing the combined range catches zero in either.
- **All 5 tabs have `*~pR0Ph1x~*` in cell `B1`** — this is Prophix Analyzer's marker and does not affect the data check. Do not include row 1 in the SUM.
- **None of these tabs have pre-computed total cells.** The check is always a live `SUM` over the data range.

### 5c — Failure handling

If any tab's check fails (sum ≤ 0), stop immediately. Do **not** proceed to Step 6. Send alert email to asbiei@prophix.com with body:

```
Step 5 failure — Prophix data did not populate for [newMonthLabel].

Failed tabs:
  [tab name]: SUM = [value]
  [tab name]: SUM = [value]
```

Do not resume until the Prophix refresh has been re-run and all 5 tabs pass.

**Example:** For April 2026 close run in May 2026:
- `newMonthLabel = "2026M04"`
- On each of the 5 tabs, find the column in row 2 whose value equals `2026M04`.
- Run the SUM over the tab-specific data range.
- All 5 sums must be `> 0` to proceed.

---

## Step 6: Run Q&A Validation Checks

Run the following tie-out checks between the `Summary By Region` tab and the `$$ Bookings DV` / `Units Bookings DV` source tabs. If any check fails, flag the error, continue to Step 7, and list all failed checks at the top of the executive summary email.

> **NOTE:** `Summary By Region` does not reference `$$ Bookings DV` directly. It sums FX-converted values from the per-country tabs (`NA_USD`, `NA_CAD`, `UK_GBP`, `Europe & MEA`, `Asia`, `Pacific (FG)`, `LATAM PX`, `LATAM AS`). The tie-out against `$$ Bookings DV` therefore doubles as an FX-consistency check: if Prophix-side FX rates diverge from the `FX` tab rates, these checks can fail even when the rollover itself is clean.

### 6a — Dollar Bookings Checks (`$$ Bookings DV` vs `Summary By Region`)

Row ranges are validated against the live workbook structure (Apr 2026). Region boundaries in `$$ Bookings DV` are defined by column B labels.

**Top-level regional ties:**

| Region | `Summary By Region` Row | `$$ Bookings DV` Row Range | Notes |
|---|---|---|---|
| NA | 22 | 3:182 | US + Mexico & Caribbean + Canada |
| UK | 29 | 183:242 | UK & Ireland |
| Europe & MEA | 36 | 243:362 | Europe + ME&A |
| **APAC (Asia + Pacific)** | 64 | **363:482** | Asia (363:422) + Pacific (423:482) |
| **LATAM (PX + AS)** | 85 | **483:602** | LATAM PX (483:542) + LATAM AS (543:602) |
| **Total Bookings** | 92 | **3:602** | Excludes rows 603+ (Beneleux, DACH, etc.) — those are sub-Europe breakdowns already rolled into Europe & MEA |

**Subregion ties (recommended — pinpoint where a break occurs):**

| Subregion | `Summary By Region` Row | `$$ Bookings DV` Row Range |
|---|---|---|
| Asia | 50 | 363:422 |
| Pacific | 57 | 423:482 |
| LATAM PX | 71 | 483:542 |
| LATAM AS | 78 | 543:602 |

### 6b — Units Bookings Check

| Check | `Summary By Region` Row | `Units Bookings DV` Row Range |
|---|---|---|
| Total Units | **169** | 3:152 |

> **Note:** `Summary By Region` row 164 carries the (mis-spelled) header `Total Regions Untis`. Rows 165–168 are the four sub-categories (`New`, `Migrations`, `Cross-sell`, `Upgrades`). Row 169 is the `Total` line — that is the row to read for the Total Units tie-out. (Earlier drafts of this spec said row 165, which is the `New` sub-row, not the Total.)

### 6c — Tolerance bands and severity tiers

Each check is classified into one of three severity tiers based on `ABS(variance)`:

| Tier | Subregional checks | Total Bookings | Behavior |
|---|---|---|---|
| **PASS** | `≤ $0.01` | `≤ $1.00` | Silent. Check passes, nothing logged. |
| **WARN** | `> $0.01` and `≤ $1,000` | `> $1.00` and `≤ $1,000` | Record as warning. **Executive summary email still sends**, with a banner at the top listing all warning-tier failures. |
| **BLOCK** | `> $1,000` | `> $1,000` | **Do NOT send executive summary email.** Send a variance-alert email instead (see 6e) and halt the workflow. Manual review required before retry. |

**Rationale for tier boundaries:**

- **PASS threshold** separates floating-point/FX noise (`< $0.01`) from any real discrepancy.
- **WARN/BLOCK threshold** is a uniform **$1,000** for both subregional and Total Bookings checks. Variances under $1,000 are likely FX drift, rounding residuals, or small data-entry differences — worth surfacing but not worth halting. Variances over $1,000 are almost always indicative of a partial Prophix refresh, a broken formula, or a wrong-month roll, and must be investigated before sending the exec email.

**Thresholds are adjustable.** Tune after the first few real runs if clean-run variances routinely land near the $1,000 boundary.

**Comparison logic (pseudocode):**

```
SUBREGIONAL_TOLERANCE = 0.01
TOTAL_TOLERANCE       = 1.00
BLOCK_LIMIT           = 1000.00     # uniform across subregional and total

warnings = []
blockers = []

for each (region, summary_row, bookings_range, is_total) in checks:
    summary_value  = SummaryByRegion[summary_row, new_month_col]
    bookings_value = SUM($$ Bookings DV[bookings_range, new_month_col])
    variance       = summary_value - bookings_value        # signed

    tolerance = TOTAL_TOLERANCE if is_total else SUBREGIONAL_TOLERANCE

    if ABS(variance) <= tolerance:
        continue  # PASS
    elif ABS(variance) <= BLOCK_LIMIT:
        warnings.append((region, variance, summary_row, bookings_range))
    else:
        blockers.append((region, variance, summary_row, bookings_range))

if blockers:
    send_variance_alert_email(blockers + warnings)   # Section 6e
    halt_workflow()
else:
    proceed_to_step_7_with_warnings(warnings)        # warnings banner in exec email
```

**VBA implementation sketch:**

```vba
Const BLOCK_LIMIT As Double = 1000#

Dim variance As Double
variance = summaryValue - bookingsSum

Dim tol As Double
If isTotal Then tol = 1# Else tol = 0.01

If Abs(variance) <= tol Then
    ' PASS — silent
ElseIf Abs(variance) <= BLOCK_LIMIT Then
    warnings.Add region & "|" & Format(variance, "#,##0.00") & "|" & summaryRow & "|" & bookingsRange
Else
    blockers.Add region & "|" & Format(variance, "#,##0.00") & "|" & summaryRow & "|" & bookingsRange
End If
```

### 6d — Warning reporting format (exec email banner)

When the workflow proceeds but one or more checks were WARN-tier, prepend the following banner to the executive summary email body (above the normal bullet points). Use **signed** variance (positive = Summary higher, negative = Summary lower):

```
Check the below warnings (variance within tolerated band, please review):
  [Region]: Variance of [+/-$ amount] between Summary By Region Row [X] and $$ Bookings DV rows [Y:Z]
```

Example:
```
Check the below warnings (variance within tolerated band, please review):
  APAC: Variance of -$3,247.42 between Summary By Region Row 64 and $$ Bookings DV rows 363:482
```

### 6e — Blocking variance — alert email and halt

When one or more checks hit the BLOCK tier, the workflow halts. The executive summary email is **not** sent. Instead, send a variance-alert email:

| Field | Value |
|---|---|
| To | asbiei@prophix.com |
| Subject | `Revenue Model — [Month] [Year] Close \| VARIANCE ALERT — Exec Summary Blocked` |
| Body | See template below |

```
Executive summary for [Month] [Year] close was NOT sent due to material variance(s) in validation.

BLOCKING variances (exceeded $1,000 — manual review required):
  [Region]: Variance of [+/-$ amount] between Summary By Region Row [X] and $$ Bookings DV rows [Y:Z]
  [Region]: Variance of [+/-$ amount] ...

Warning-tier variances (informational):
  [Region]: Variance of [+/-$ amount] between Summary By Region Row [X] and $$ Bookings DV rows [Y:Z]

File: C:\Users\amers\Downloads\Revenue Model\[filename]

Next steps:
  1. Review the Summary By Region tab and $$ Bookings DV tab for the listed rows in the new month column.
  2. If the issue is a failed/partial Prophix refresh, re-run Step 4 (Prophix refresh) and Step 5 (data verification).
  3. If the issue is a formula/FX drift, correct the source and re-run Step 6 only.
  4. Once checks pass (or are below the blocking threshold), the executive summary email can be sent.

The workflow is halted. Do not re-run from Step 1 unless the underlying file is corrupted — the archive backup exists if needed.
```

After sending the alert, the agent must exit cleanly. Do **not** attempt Step 7 or Step 8.

---

## Step 7: Create and Send Executive Summary Email

### 7a — Email Details

| Field | Value |
|---|---|
| To | asbiei@prophix.com |
| Subject | Revenue Model — [Month] [Year] Close \| Executive Summary |
| Delivery | **Saved to Outlook Drafts for manual review** via Outlook COM `MailItem.Save()` (NOT `.Send()`). Reviewer opens Drafts, inspects, sends manually. |
| Body format | HTML with inline CID-embedded PNG charts |
| File Link | Paste local file path: `C:\Users\amers\Downloads\Revenue Model\` |

### 7b — Executive Summary Bullet Points
- If any validation checks failed in Step 6, list them at the very top of the email before the summary bullets.
- Write 8 to 10 bullet points covering actual performance vs FY26 Plan (**Internal plan only**).
- Cover MTD, QTD, and YTD performance using the calculation rules below.

#### Fiscal year convention

**Calendar fiscal year** (verified from workbook `FY26 Plan!C14='Q1-2026'`, `F14='Q2-2026'`).

| Quarter | Months |
|---|---|
| Q1 | Jan, Feb, Mar |
| Q2 | Apr, May, Jun |
| Q3 | Jul, Aug, Sep |
| Q4 | Oct, Nov, Dec |

Given a close month `m`, QTD = months in the containing quarter up to and including `m`.

#### Source sheet column rules

The two source sheets have **different column layouts**. The agent MUST use the correct sheet-specific column, not the same column number across both sheets.

**`Summary By Region` (Actuals)** — month headers in row 17, continuous multi-year layout starting at `2021M01` = col C (3).

| Month | Column | |  | Month | Column |
|---|---|---|---|---|---|
| 2026M01 | BK (63) | | | 2026M07 | BQ (69) |
| 2026M02 | BL (64) | | | 2026M08 | BR (70) |
| 2026M03 | BM (65) | | | 2026M09 | BS (71) |
| 2026M04 | BN (66) | | | 2026M10 | BT (72) |
| 2026M05 | BO (67) | | | 2026M11 | BU (73) |
| 2026M06 | BP (68) | | | 2026M12 | BV (74) |

Formula: `col = 63 + (month - 1)` for calendar year 2026. Safer: match `YYYYMmm` header in row 17.

**`FY26 Plan` — Internal block only (cols B–U)**. The External block (cols W–AO) is **not** used for the executive summary.

| Month | Column | |  | Month | Column |
|---|---|---|---|---|---|
| 2026M01 | C (3) | | | 2026M07 | I (9) |
| 2026M02 | D (4) | | | 2026M08 | J (10) |
| 2026M03 | E (5) | | | 2026M09 | K (11) |
| 2026M04 | F (6) | | | 2026M10 | L (12) |
| 2026M05 | G (7) | | | 2026M11 | M (13) |
| 2026M06 | H (8) | | | 2026M12 | N (14) |

Precomputed Internal quarterly subtotals at cols P (Q1), Q (Q2), R (Q3), S (Q4). Internal annual total at col U.

Formula: `col = 3 + (month - 1)`. Safer: match `YYYYMmm` header in row 17 within cols C–N.

#### Metric calculation table

For a close in month `m` of calendar year `yyyy`:

| Metric | `Summary By Region` (Actual) | `FY26 Plan` Internal (Plan) |
|---|---|---|
| MTD | Single column for month `m` | Single column for month `m` |
| QTD | Sum of columns for the quarter months up to and including `m` | Sum of columns for the quarter months up to and including `m` — **or** read the precomputed quarterly subtotal column (P/Q/R/S) if the close is the last month of the quarter (Mar, Jun, Sep, Dec) |
| YTD | Sum of columns from Jan (col BK) through month `m` | Sum of columns from Jan (col C) through month `m` |

#### Worked example — April 2026 close

| Metric | Summary By Region (Actual) | FY26 Plan Internal (Plan) |
|---|---|---|
| MTD | col BN (66) | col F (6) |
| QTD | col BN (66) (April is first month of Q2, so QTD = April only) | col F (6) |
| YTD | cols BK:BN (63:66) | cols C:F (3:6) |

#### Worked example — June 2026 close (end of Q2)

| Metric | Summary By Region (Actual) | FY26 Plan Internal (Plan) |
|---|---|---|
| MTD | col BP (68) | col H (8) |
| QTD | cols BN:BP (66:68) | cols F:H (6:8) — or read precomputed col Q (Q2-2026 Internal subtotal) |
| YTD | cols BK:BP (63:68) | cols C:H (3:8) |

#### Bullet point template and formatting rules

**Tone:** Professional, direct, data-driven. Big-4 consultant style. No fluff.
**Audience:** CFO, CRO, CEO.
**Length:** 8 to 10 bullets maximum.
**Structure:** Follow this exact order.

**Bullet 1 — Overall MTD Headline**
> Total ACV Bookings for [Month] came in at **[$ Actual]** CAD, [above/below] Plan by [$ variance] ([% variance]%) and [above/below] Prior Month by [$ variance] ([% variance]%).

**Bullet 2 — Overall YTD Headline**
> Year-to-date ACV Bookings through [Month] stand at **[$ Actual]** CAD, tracking [above/below] the FY26 Plan by [$ variance] ([% variance]%), reflecting [ahead of/behind] pace for full-year target.

**Bullet 3 — Overall QTD Headline**
> Q[X] quarter-to-date Bookings of **[$ Actual]** CAD are [above/below] the quarterly Plan by [$ variance] ([% variance]%), with [X] months remaining in the quarter.

**Bullet 4 — Strongest Regional Performer (MTD)**
> [Region] was the strongest contributor in [Month], delivering [$ Actual] CAD against a Plan of [$ Plan], [above/below] by [$ variance] ([% variance]%).

**Bullet 5 — Weakest Regional Performer (MTD)**
> [Region] was the largest miss in [Month] at [$ Actual] CAD vs a Plan of [$ Plan], [below] by [$ variance] ([% variance]%). [One sentence on what drove it if identifiable from the data.]

**Bullet 6 — NA Performance (MTD and YTD)**
> North America posted MTD Bookings of [$ Actual] CAD ([above/below] Plan by [% variance]%) and YTD of [$ Actual] CAD ([above/below] Plan by [% variance]%).

**Bullet 7 — EMEA Performance (MTD and YTD)**
> EMEA delivered MTD Bookings of [$ Actual] CAD ([above/below] Plan by [% variance]%) and YTD of [$ Actual] CAD ([above/below] Plan by [% variance]%).

**Bullet 8 — APAC and LATAM Performance (MTD)**
> APAC contributed [$ Actual] CAD MTD ([above/below] Plan by [% variance]%) while LATAM delivered [$ Actual] CAD ([above/below] Plan by [% variance]%).

**Bullet 9 — Prior Month Comparison (MoM trend)**
> Compared to [Prior Month], total Bookings [increased/decreased] by [$ variance] ([% variance]%), driven primarily by [strongest/weakest region MoM].

**Bullet 10 — Forward-Looking Closing Statement**
> Based on current QTD trajectory, the business is [on track/at risk] to achieve the Q[X] Plan of [$ Plan] CAD, requiring [$ remaining] in [remaining months].

**Formatting rules (strict — enforce in code):**
- All dollar amounts in CAD, formatted as `$X,XXX,XXX`.
- All variances shown as both `$` and `%`.
- Favorable variance = "above Plan". Never use "beat" or "exceeded".
- Unfavorable variance = "below Plan". Never use "missed" or "fell short".
- Prior month label = actual month name (e.g., "March", not "prior month").
- Never use em dashes (—). Use commas or parentheses.
- Never use the word "significant" or "strong" without a number to back it up.
- Bold the dollar figures for MTD, YTD, and QTD actuals in each bullet (HTML `<b>...</b>`).

**Step 6 WARN-tier override:** If any validation check was WARN-tier in Step 6, insert this as **Bullet 1** and push all other bullets down (becoming 2–11 max):
> [WARNING] Validation checks detected variances in [Region]. Numbers should be reviewed before distribution.

This is *additional to*, not a replacement for, the Section 6d banner (which lists the per-region variance line items above the bullets). The banner gives the numbers; this bullet gives the executive call-out.

### 7c — Bar Chart Specifications

Three separate **clustered horizontal bar charts** are generated in **Python via matplotlib** as PNG bytes, attached to the Outlook `MailItem`, and referenced inline in the HTML body via **Content-ID (CID)**. One chart per time window (MTD / QTD / YTD), stacked vertically in the email body.

#### Chart source rows (per-region CAD totals)

Both `Summary By Region` (Actual) and `FY26 Plan` Internal (Plan) use the **same row numbers** for pre-summed regional totals — one cell read per region per sheet. Do not sum sub-regional rows manually; the workbook already aggregates.

| Region | Actual row (`Summary By Region`) | Plan row (`FY26 Plan` Internal) | Col B label to verify |
|---|---|---|---|
| NA | 22 | 22 | `NA $$` block → `Total` |
| EMEA | 43 | 43 | `EMEA $$` block → `Total` (pre-sums UK + Europe & MEA) |
| APAC | 64 | 64 | `APAC $$` block → `Total` (pre-sums Asia + Pacific) |
| LATAM | 85 | 85 | `LATAM $$` block → `Total` (pre-sums LATAM PX + LATAM AS) |

**Label-check guard (required before every read):** Assert that column B at the target row reads `Total`, and the section header 5 rows above (e.g., `B38`) matches the expected region label (`EMEA $$` for EMEA, etc.). If either assertion fails, stop and alert — rows have shifted and the mapping is stale.

**Month column lookup:** Match `YYYYMmm` label in `Summary By Region` row 17 and `FY26 Plan` row 2. Never hard-code column numbers — both sheets advance by one column per month and layouts drift across fiscal years.

**MTD / QTD / YTD per region:** sum the column range defined in Step 7b for the selected time window, for each region's Actual row and Plan row.

#### Chart styling

| Chart Element | Specification |
|---|---|
| Chart type | Clustered horizontal bar (matplotlib `ax.barh` with grouped series, offset by half bar width) |
| Regions (Y axis, top to bottom) | NA, EMEA, APAC, LATAM |
| X axis | Total Bookings in CAD |
| Series | Budget (from `FY26 Plan` Internal), Actual (from `Summary By Region`) — side by side |
| Budget bar color | `#E9EAEF` |
| Actual bar color | `#095A87` |
| Data labels | `$#.#M` at end of each bar (e.g., `$12.3M`) using matplotlib `ax.bar_label` |
| Chart titles | `MTD: [Month] [Year] vs Plan` / `QTD: [Q#] [Year] vs Plan` / `YTD: FY[YY] vs Plan` |
| Output format | PNG bytes via `savefig(BytesIO(), format='png', dpi=150, bbox_inches='tight')` |
| Figure size | Recommended 8in × 4in (wide enough for data labels without clipping) |

#### Embedding in the Outlook draft (CID)

Each PNG is attached to the `MailItem`, assigned a Content-ID, and referenced from HTML via `<img src="cid:...">`. Outlook's attachment API requires a file path, so PNGs are written to a temp directory first and deleted after `.Save()` returns.

```python
# Python sketch — Outlook COM via pywin32
import win32com.client as win32
from tempfile import NamedTemporaryFile

outlook = win32.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)  # 0 = olMailItem
mail.To = "asbiei@prophix.com"
mail.Subject = f"Revenue Model — {month_name} {year} Close | Executive Summary"

PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
cid_map = {"mtd_chart": mtd_png, "qtd_chart": qtd_png, "ytd_chart": ytd_png}
temp_paths = []

for cid, png_bytes in cid_map.items():
    tmp = NamedTemporaryFile(delete=False, suffix=".png")
    tmp.write(png_bytes); tmp.close()
    temp_paths.append(tmp.name)
    att = mail.Attachments.Add(tmp.name)
    att.PropertyAccessor.SetProperty(PR_ATTACH_CONTENT_ID, cid)

mail.HTMLBody = build_html_body(
    warn_banner=warn_banner,   # Section 6d banner, if any
    bullets=exec_bullets,      # Section 7b template output
    chart_cids=list(cid_map.keys()),
)
mail.Save()                    # Drafts — NOT .Send()

# Clean up temp files after Outlook has copied them into the draft
for p in temp_paths: os.unlink(p)
```

HTML body references the images inline:

```html
<img src="cid:mtd_chart" style="display:block; margin:12px 0;">
<img src="cid:qtd_chart" style="display:block; margin:12px 0;">
<img src="cid:ytd_chart" style="display:block; margin:12px 0;">
```

> **Important:** A stacked bar chart adds series values into a single combined bar (e.g., Budget $10M + Actual $8M = $18M bar). That pattern is for showing composition, not comparison. For Budget vs Actual per region, a clustered bar chart places the two values side-by-side so the variance is visible at a glance. The original spec said "stacked" but the described layout (two distinct series colors, regions on Y axis) only makes visual sense as clustered.

### 7c.1 — Optional alternative interpretations (confirm before coding)

If the "stacked" wording was intentional and a different visualization is wanted, there are two plausible alternatives — confirm with the recipient before locking in:

1. **Stacked by booking type within region (one bar per region, one chart per period per scenario).** Each stacked bar composed of New + Migration + Cross-sell + Upgrades summing to total. Would require 2× the charts (Budget and Actual shown separately) or 8 charts total if split by period as well.
2. **Overlaid bars with transparency (target-line style).** Actual rendered on top of Budget at the same origin, with Budget showing through as a faded backdrop. Not a standard Excel chart type but common in dashboards.

Default (clustered) is the recommended interpretation unless overridden.

---

## Step 8: Save and Close
- Save the file after the email is confirmed sent.
- Close the file.

---

## Error Handling Rules

| Scenario | Action |
|---|---|
| Workbook is locked (Excel has it open) during Step 1 | `shutil.copy2` or `Path.rename` raises `PermissionError`. Close Excel and retry. The agent does not force-close the workbook. |
| Archive folder missing at Step 1 start | Created automatically (`mkdir(parents=True, exist_ok=True)`). No failure. |
| VBA Macro fails (Step 2) | Switch to Step 3 Option 2: manually drag formulas using the row blocks and tab list defined in Step 3. |
| Step 2 date-math guard fails | The filename date is out of sync with the intended close month. Stop and alert. Do not run the rollover — it will populate the wrong month. |
| Prophix Analyzer fails to open after Alt+N+Y1 | Retry up to 3 times. If still failing, stop and send alert to asbiei@prophix.com: *Prophix Analyzer failed to open after 3 attempts.* |
| Prophix refresh hangs beyond 30 minutes | With the fixed 10-minute sleep in Step 4, this rule only fires if the optional post-sleep re-check loop is enabled. Without the re-check, the agent always proceeds at 10 minutes regardless of Prophix state. If 30-minute enforcement is required, implement the re-check loop below. |
| Prophix partial-refresh suspected (data looks stale) | Optional post-sleep re-check: after Step 4's 10-minute sleep, re-run Step 5's zero-check. If any required tab shows zero, sleep an additional 5 minutes and re-check once. If still zero after the second check (total 15 minutes elapsed), send alert to asbiei@prophix.com and do not proceed to Step 6. |
| Outlook is closed when sending email | Open Outlook automatically and proceed with sending the email. |
| Archive folder does not exist | Not applicable. The folder always exists at: `C:\Users\amers\Downloads\Revenue Model\Archive\` |
| Validation check fails in Step 6 (WARN tier) | Flag the error. Continue to Step 7. List all WARN-tier checks in the banner at the top of the executive summary email per Section 6d. |
| Validation check fails in Step 6 (BLOCK tier) | **Do NOT send the executive summary email.** Send the variance-alert email per Section 6e to asbiei@prophix.com and halt the workflow. Manual review required before retry. The archive backup exists if a rollback is needed. |
| Any data tab shows zero after Prophix refresh | Stop immediately. Send alert to asbiei@prophix.com listing which tab(s) showed zero. Do not proceed until resolved. |

---

## Implementation Architecture

The agent is a **Python orchestrator** that drives Excel and Outlook via COM automation. Step 1 (archive + rename) is pure Python file ops. Step 3 invokes one existing VBA macro for formula rolling. All validation, tolerance math, chart generation, email drafting, and logging lives in Python.

### Division of responsibility

| Layer | Owns |
|---|---|
| Python (orchestrator) | Filename parsing, date math, workflow control, Step 5 verification, Step 6 tolerance math, Step 7 MTD/QTD/YTD computation, chart PNG generation, exec summary bullet composition, Outlook draft creation, logging |
| Excel workbook VBA (existing) | `Step2_RollFormulas_TwoMonthsBack_SourceToTarget_Reliable` (Step 3) — invoked by Python via Excel COM `.Run()`. Step 1's VBA macro is no longer used; file ops are in Python. |
| Excel COM (`pywin32`) | Open/close workbook, invoke macros, read cell ranges for Steps 5/6/7 |
| Outlook COM (`pywin32`) | Build HTML `MailItem`, embed PNG attachments via CID, save to Drafts for manual review |

### Dependencies

| Package | Purpose |
|---|---|
| `pywin32` | Excel + Outlook COM |
| `openpyxl` | Offline inspection / unit tests (non-COM) |
| `matplotlib` | Chart PNG generation |
| `pillow` | PNG encoding support (matplotlib dependency) |
| `python-dateutil` | Month-offset date math |
| `python-docx` | (already used) spec-to-docx conversion |

All pip-installable. Windows-only (acceptable — this is a Windows workflow using Excel, VBA, and Outlook).

### Repo layout

```
Revenue Model/
  Revenue Model - MMDDYYYY (Internal).xlsm
  Archive/
  agent/
    main.py                  # top-level orchestrator
    config.py                # paths, recipient, thresholds, tab lists
    logging_setup.py         # rotating per-run log files
    step1_archive_rename.py  # calls Step1 VBA macro via Excel COM
    step3_roll_formulas.py   # date-math guard + calls Step2 VBA macro
    step4_prophix.py         # SendKeys Alt+N+Y1, 10-min wait
    step5_verify.py          # per-tab SUM > 0 checks, failure email
    step6_validate.py        # tie-out checks, WARN/BLOCK tolerance tier
    step7_email.py           # exec summary compose + Outlook draft
    charts.py                # matplotlib clustered horizontal bar -> PNG
    outlook_draft.py         # Outlook COM helper (CID embed, Drafts)
    excel_com.py             # Excel COM helper (open, Run macro, read ranges)
    logs/                    # rotating per-run log files
  tests/
    test_date_math.py
    test_step5_checks.py
    test_step6_tolerance.py
    test_bullet_template.py
  spec.md
  spec.docx
```

### Run mode

- **Now (local testing):** manual invocation `py agent/main.py`.
- **Later (production):** Windows Task Scheduler triggered approximately 4 days after month-end close.
- **Python never sends email** — it always drafts. User reviews in Outlook Drafts and sends manually. Matches Step 7 "draft for review" decision and provides a natural checkpoint for WARN-tier review.

### Logging

Every run writes a timestamped log file to `agent/logs/run_YYYYMMDD_HHMMSS.log` capturing: filename date, derived month labels, step-by-step outcome, Step 5 SUM values per tab, Step 6 variance per tie-out check with tier (PASS/WARN/BLOCK), chart computation values per region per time window, Outlook draft EntryID. This log is the primary audit trail and the first place to look when a run needs debugging.

---

## Testing Note

This agent is currently configured for local testing. Once local testing is confirmed working, provide the SharePoint URL to replace the local file path in all file access and email link steps.

- Current local path: `C:\Users\amers\Downloads\Revenue Model\`
- SharePoint URL: *To be provided after successful local test.*

---

## Change Log

- **2026-04-21** — Corrected Step 6a row ranges. APAC (row 64) now spans `$$ Bookings DV` 363:482 (was 363:422; previously missed Pacific). LATAM (row 85) now spans 483:602 (was 483:542; previously missed LATAM AS). Total Bookings (row 92) now spans 3:602 (was 3:603; row 603 starts Beneleux, a sub-Europe breakdown already rolled into Europe & MEA). Added subregion tie-out table and a note that `Summary By Region` does its own FX conversion so these ties also validate FX consistency.
- **2026-04-21** — Replaced "100% exact match, no rounding tolerance" with tolerance bands in Step 6c: $0.01 for subregional checks, $1.00 for Total Bookings. Added pseudocode, VBA, and cell-formula implementations. Step 6d failure format now uses signed variance so direction of the miss is visible.
- **2026-04-21** — Verified all 7 helper subs/functions the Step 2 macro depends on (`GetFileDateFromWorkbookName_MMDDYYYY`, `GetRowBlocks`, `FindMonthColumnInRow4_Robust`, `RollFormulasByBlocks_DragFix`, `RestoreAppState`, `TimedPopup`, plus the main `Step2_...`) exist in the workbook's VBA project. Discovered `Step1_BackupLiveThenRename_OverwriteLiveDatedName` already automates Step 1 — restructured Step 1 into Option 1 (macro, preferred) and Option 2 (manual fallback). Flagged three constraints: hard-coded SharePoint backup path (blocks local testing without a one-line VBA edit), macro exits on read-only workbook (contradicts original spec's "proceed anyway" rule — fallback required), no guard against running before month-end (macro would roll wrong months). Added date-math pre-flight assertion to Step 3 Option 1. Minor cleanup note: `TimedPopup` and `RemoveExtension` are duplicated across Module1/Module2 — not a compile blocker but should be deduplicated.
- **2026-04-21** — Step 7c chart type corrected from "stacked bar" to "clustered horizontal bar". Stacked combines Budget + Actual into one bar (composition), which doesn't visually represent a Budget vs Actual comparison. Clustered places them side-by-side so the gap is readable. Added data labels, chart titles, and a note on two alternate interpretations (Step 7c.1) in case "stacked" was intentional and a different visualization is preferred.
- **2026-04-21** — Step 4 Prophix wait replaced "wait for natural completion signal" (not programmatically observable) with fixed 10-minute sleep. Rationale: Prophix Analyzer CV2 does not expose a completion event, and Step 5's existing zero-check serves as the safety net for total refresh failures. Known limitation: a partial-refresh (some tabs land new data, others stay stale) can slip past Step 5's zero-check. Documented optional post-sleep re-check in Error Handling for cases where Prophix timing is variable. 30-minute timeout rule now only applies if the re-check is enabled.
- **2026-04-21** — Step 7b rewritten with explicit column-lookup tables for both source sheets and worked examples. Calendar fiscal year convention (Q1=Jan-Mar, Q2=Apr-Jun, Q3=Jul-Sep, Q4=Oct-Dec) verified from workbook `FY26 Plan!C14` and locked in. Corrected spec bug where "MTD Plan = same closing month column as actuals" was wrong: `Summary By Region` April column is BN(66) while `FY26 Plan` Internal April column is F(6) — different sheets use different column layouts. FY26 Plan structure clarified: workbook has two parallel blocks (Internal cols B–U, External cols W–AO); **per business owner decision, executive summary uses Internal plan only — External block is ignored entirely**. Noted precomputed quarterly subtotals at FY26 Plan cols P–S for quarter-end optimization.
- **2026-04-21** — Step 5 rewritten with concrete verification rules. Added Section 5a (new month label derivation from filename date — same logic as Step 2 macro's `dstLabel`). Added Section 5b with verified per-tab header row, data row range, and check formula for all five tabs (`Units Bookings DV` 3:152, `$$ Bookings DV` 3:602, `Sub Europe Units` 3:182, `Migration Units` 3:17, `AS` 3:122). All five tabs use **row 2** for month headers (differs from the regional tabs which use row 4). Label-lookup approach handles the `Migration Units` column-base discrepancy transparently. Added Section 5c concrete failure-handling with alert email template.
- **2026-04-21** — Step 6 tolerance replaced with tiered model: PASS (`≤ $0.01` subregional / `≤ $1.00` total), WARN (exceeds tolerance and `≤ $1,000` → exec email still sends with banner), BLOCK (`> $1,000` → exec email NOT sent, variance-alert sent to asbiei@prophix.com, workflow halts). Uniform $1,000 block threshold across subregional and Total Bookings checks. Added Section 6e variance-alert email template. Error Handling updated to split WARN-tier from BLOCK-tier validation failures. Threshold flagged as adjustable pending first real runs.
- **2026-04-21** — Locked in architecture and email/chart mechanics. Step 7a now specifies Outlook Drafts (via `MailItem.Save()`, not `.Send()`) for manual reviewer approval. Step 7b gained the full bullet-point template (10-bullet exec summary structure, CFO/CRO/CEO audience, strict formatting rules: no em dashes, no "beat"/"missed" language, bolded MTD/YTD/QTD actuals, WARN-tier Bullet 1 override). Step 7c rewritten: charts generated in Python via matplotlib (not Excel native), embedded as CID-referenced PNGs in the HTML body. Added chart source row table (NA=22, EMEA=43, APAC=64, LATAM=85 on both `Summary By Region` and `FY26 Plan` Internal — pre-summed regional totals, same row on both sheets) with required label-check guard. Added new `Implementation Architecture` section: Python orchestrator + existing VBA macros + `pywin32` COM; dependencies, repo layout, logging spec, and production run mode (Windows Task Scheduler).
- **2026-04-21** — Pre-build audit fix. Step 6b Total Units row corrected from 165 to **169**. Row 165 in `Summary By Region` is the `New` sub-row of the `Total Regions Untis` block (note: workbook contains a typo — header reads "Untis" not "Units"); the `Total` line is at row 169. All other row claims in Steps 5, 6a, and 7c verified against the live workbook (Apr 2026 file) and confirmed correct.
- **2026-04-22** — Step 1 rewritten as pure Python file ops per owner preference for stability. The `Step1_BackupLiveThenRename_OverwriteLiveDatedName` VBA macro is no longer used. Trade-offs: file copy + rename happen as two operations (previously atomic inside VBA), but the order is safe (archive first, rename second) so a failure between them still leaves the backup in place. Eliminates the Trust Center / macro-enable / hard-coded SharePoint path prerequisites for Step 1. Same collision semantics kept: `" Backup"` suffix with numeric `(n)` fallback; same-day re-run overwrites today's dated target; idempotent when source already matches today's name. Error Handling table and Implementation Architecture table updated; Step 1's "Option 1" section removed.
