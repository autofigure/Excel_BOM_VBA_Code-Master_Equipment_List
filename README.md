# Master Equipment List System - README

## Table of Contents
1. [System Overview](#system-overview)
2. [Architecture](#architecture)
3. [Setup Instructions](#setup-instructions)
4. [User Workflow](#user-workflow)
5. [Module Reference](#module-reference)
6. [Data Flow](#data-flow)
7. [Protection & Locking](#protection--locking)
8. [Troubleshooting](#troubleshooting)
9. [Advanced Configuration](#advanced-configuration)

---

## System Overview

### Purpose
The Master Equipment List system is a comprehensive VBA-based data management tool that consolidates Bills of Materials (BOMs) from four engineering disciplines (Electrical, Hydraulic, Pneumatic, Mechanical) into a unified Master Equipment List, then distributes relevant data to specialized destination sheets.

### Key Features
- **Automated BOM Integration**: Power Query combines data from separate discipline BOM files
- **Intelligent Part Ownership**: ELEC discipline owns shared parts (sensors appear in multiple BOMs)
- **Multi-Destination Distribution**: Automatically populates IO Lists, P&ID Tag Lists, Utility Load, and Heat/Noise tables
- **Smart Data Protection**: Row-level and column-level locking based on data source
- **Manual Entry Support**: Users can add custom equipment entries alongside BOM data
- **Removed Part Tracking**: Parts deleted from BOMs are flagged rather than immediately deleted
- **Dropbox Integration**: Supports both local and Dropbox-synced project folders

---

## Architecture

### System Components

```
┌─────────────────────────────────────────────────────────────┐
│                      SOURCE BOM FILES                        │
│  (Separate Excel files maintained by each discipline)       │
├──────────────┬──────────────┬──────────────┬────────────────┤
│ ELEC BOM     │ HYD BOM      │ PNU BOM      │ MECH BOM       │
│ (INTERNAL)   │ (INTERNAL)   │ (INTERNAL)   │ (INTERNAL)     │
└──────┬───────┴──────┬───────┴──────┬───────┴────────┬───────┘
       │              │              │                │
       └──────────────┴──────────────┴────────────────┘
                              │
                    ┌─────────▼──────────┐
                    │   POWER QUERY      │
                    │  (Resilient Load)  │
                    └─────────┬──────────┘
                              │
                    ┌─────────▼──────────┐
                    │  All_BOM_Parts     │
                    │  (Staging Table)   │
                    └─────────┬──────────┘
                              │
                    ┌─────────▼──────────┐
                    │ Master Equipment   │
                    │      List          │
                    │  (Central Hub)     │
                    └─────────┬──────────┘
                              │
       ┌──────────────────────┼──────────────────────┐
       │                      │                      │
┌──────▼───────┐    ┌────────▼────────┐   ┌────────▼────────┐
│   IO List    │    │  P&ID Tag List  │   │ Utility Load &  │
│              │    │  & Loops        │   │  Heat/Noise     │
└──────────────┘    └─────────────────┘   └─────────────────┘
```

### Data Hierarchy (Part Ownership)

When a part appears in multiple BOMs, **ELEC has highest priority**:

**Priority Order**: ELEC > HYD > PNU > MECH

**Example**: Pressure sensor appears in all BOMs
- **Owner**: ELEC (provides Part Number, Assy QTY, QTY, Description)
- **Contributors**: HYD, PNU, MECH (provide Need QTY and Tags)
- **Result**: Master list shows ELEC data with combined Need QTY

---

## Setup Instructions

### Prerequisites
1. Excel 2016 or later (Power Query support required)
2. VBA macros enabled
3. Project folder structure:
   ```
   Project Root/
   ├── Job Documentation/
   │   ├── Electrical/
   │   │   └── [JobNum]-ELEC BOM (INTERNAL).xlsx
   │   ├── Hydraulic/
   │   │   └── [JobNum]-HYD BOM (INTERNAL).xlsx
   │   ├── Pneumatic/
   │   │   └── [JobNum]-PNU BOM (INTERNAL).xlsx
   │   └── Mechanical/
   │       └── [JobNum]-MECH BOM (INTERNAL).xlsx
   ```

### Initial Setup

#### Step 1: Configure Project Information
On the **Master Equipment List** sheet, fill in:
- **Project Number** (Cell C2): e.g., "20570"
- **Customer** (Cell C3)
- **Project Name** (Cell C4)
- **Disconnect Rating** (Cell C5)
- **Project Folder** (Cell C6): Full path or Dropbox subfolder
- **Dropbox Filepath?** (Cell E6): "Y" or "N"

#### Step 2: Configure Config Sheet
On the **Config** sheet:
- **ThisWorkbookFolder**: Full path to where this Excel file is saved
- **DropboxStatus**: "Y" if using Dropbox, "N" if using direct paths

#### Step 3: Create Power Query
1. Run macro: **`Create_BOM_Query`**
2. This creates the query that loads from all BOM files
3. Handles missing BOMs gracefully (project may not have all 4 disciplines)

#### Step 4: First Sync
1. Manually load Power Query to All_BOM_Parts table:
   - Data → Queries & Connections → All_BOM_Parts → Right-click → Load To
   - Choose: Table, Existing worksheet, Cell A3 in All_BOM_Parts sheet
2. Run macro: **`Run_Master_Equipment_Sync`**
3. Parts will populate Master Equipment List

---

## User Workflow

### Normal Operation

#### Daily Workflow
1. Engineers update their discipline BOM files (ELEC, HYD, PNU, MECH)
2. User opens Master Equipment List workbook
3. Click **"Sync from BOMs"** button → Runs `Run_Master_Equipment_Sync`
   - Refreshes Power Query from BOM files
   - Updates Master Equipment List
   - Syncs destination tables
   - Takes 10-30 seconds depending on BOM sizes

#### Configuring Destination Tables
1. After sync, new parts appear with all "Include in..." columns set to "N"
2. User sets "Include" flags to "Y" for relevant destinations:
   - **Include in I/O List?** → "Y" to add to IO List
   - **Include in Utility Load Table?** → "Y" to add to Utility Load
   - **Include in Heat Load & Noise Table?** → "Y" to add to Heat/Noise
3. Click **"Apply Changes"** button → Runs `Apply_Changes`
   - Updates destination tables based on Include flags
   - No BOM refresh (fast, ~2-5 seconds)

#### Adding Manual Equipment
1. Add new row to Master Equipment List table
2. Leave **Source** column blank → Auto-fills to "MAN"
3. Fill in equipment details
4. Set "Include in..." flags as needed
5. Click **"Apply Changes"**

#### Managing Removed Parts
When a part is deleted from a BOM:
- Master list marks it **"Removed from BOM" = "Y"**
- Row remains visible but fully unlocked
- User can review and decide whether to delete

To clean up removed parts:
1. Click **"Delete Removed Parts"** button → Runs `Delete_Removed_Parts`
2. Choose deletion mode:
   - **Yes**: Bulk delete all removed parts
   - **No**: Review each one-by-one (Keep/Delete/Cancel)
   - **Cancel**: Exit without changes

---

## Module Reference

### Module_Config.txt
**Purpose**: Centralized configuration and constants

**Key Functions**:
- `GetProjectPath()`: Builds project path considering Dropbox status
- Worksheet accessor functions: `wsMaster()`, `wsIOList()`, etc.

**Constants**:
- Sheet names: `SHEET_MASTER`, `SHEET_IO_LIST`, etc.
- Table names: `TABLE_MASTER`, `TABLE_IO_LIST`, etc.
- BOM column names: `BOM_COL_SOURCE`, `BOM_COL_PART_NUMBER`, etc.

---

### Module_Global_Functions.txt
**Purpose**: Core infrastructure functions used across all modules

**Key Functions**:

#### Protection Management
- **`Global_Unprotect()`**: Unprotects all sheets for editing
- **`Global_Protect()`**: Applies protection with intelligent cell-level locking
  - Master Equipment List: Row-level locking (BOM vs MAN vs Removed)
  - IO List: Column locking (master-sourced columns locked)
  - P&ID Tag List: Column locking with manual entry support
  - Utility Load: Column locking
  - Heat & Noise: Column locking
  - Unlocks project info cells (merged cells C2:C7, E6)

#### Table Helpers
- **`GetTableColIndex(lo, colName)`**: Returns column index in ListObject
- **`TableHasRows(lo)`**: Checks if table has data
- **`IsInArray(val, arr)`**: Array membership check

#### Data Helpers
- **`GetCellValue(rowRange, colName, lo)`**: Safe cell value extraction
- **`SetCellValue(lo, rowRange, colName, val)`**: Safe cell value setting

#### String Formatting
- **`PadRight(txt, width)`**: Right-pad text with spaces
- **`PadLeft(txt, width)`**: Left-pad text with spaces
- **`PadBoth(txt, width)`**: Center text in fixed width

#### BOM Management
- **`DetectMissingBOMs(loSrc, expected)`**: Identifies which BOMs are missing
- **`Create_BOM_Query()`**: Dynamically generates Power Query M code
  - Handles missing BOM files gracefully
  - Uses try...otherwise null for resilience
  - Combines only non-empty tables

#### Logging
- **`WriteSyncLog(statusText, detailText)`**: Writes to SYNC_Log table (optional)

---

### Module_Sync_Master_Equipment.txt
**Purpose**: Syncs All_BOM_Parts → Master Equipment List

**Main Function**: `Run_Master_Equipment_Sync()`

**Workflow**:
1. Refreshes Power Query connection
2. Calls `Sync_Master_Equipment_List()`
3. Calls `Post_Sync_Format()`
4. Calls `Global_Protect()`

**Key Algorithm**: `Sync_Master_Equipment_List()`
1. **Build BOM Data Dictionary**:
   - Groups All_BOM_Parts by Part Number
   - For each part, stores all BOM sources that contain it
   - Determines "owner" based on hierarchy: ELEC > HYD > PNU > MECH
   
2. **Pass 1 - Update Existing Rows**:
   - Loop through Master Equipment List
   - If part exists in BOM data:
     - Update with owner's data (Part#, Assy QTY, QTY, Description)
     - Sum Need QTY across all BOMs
     - Concatenate Source (e.g., "ELEC, HYD, PNU")
     - Collect tags by BOM (ELEC Tags, HYD Tags, etc.)
     - Mark "Removed from BOM" = "N"
   - If part NOT in BOM data:
     - Mark "Removed from BOM" = "Y"
     - Keep existing data (for user review)
   
3. **Pass 2 - Add New Parts**:
   - Add rows for parts in BOM but not in Master
   - Set all "Include in..." columns to "N" (user decides later)
   - Assign sequential Item numbers

**Part Ownership Logic**:
```vba
Priority:
  ELEC = 4 (highest)
  HYD  = 3
  PNU  = 2
  MECH = 1 (lowest)

Owner provides: Part Number, Assy QTY, QTY, Description
All sources provide: Need QTY, Tags
```

**Item Numbering**:
- Permanent sequential numbers
- Survives part removal and re-addition
- Numbers never decrease, only increase

---

### Module_Sync_Destinations.txt
**Purpose**: Syncs Master Equipment List → Destination tables

**Main Function**: `Sync_All_Destinations()`

Calls individual sync functions:
- `Sync_IO_List()`
- `Sync_PID_Tag_List()`
- `Sync_Utility_Load()`
- `Sync_Heat_Noise()`

#### Sync_IO_List()
**Purpose**: Populates IO List with electrical I/O points

**Logic**:
1. Build dictionary of parts where "Include in I/O List?" = "Y"
2. For each ELEC Tag, create separate row (one part → multiple rows if multiple tags)
3. Pass 1: Update/delete existing rows
4. Pass 2: Add new rows
5. Sort by Master Equipment List Item#

**Key**: `ItemNum|Tag` (e.g., "42|PIT-100")

#### Sync_PID_Tag_List()
**Purpose**: Populates P&ID Tag List with instrumentation

**Logic**:
1. Parse P&ID Tags column (comma-separated)
2. Create row for each tag
3. Supports manual entries (blank Item#)
4. Sort by Item#, then Instrument/Equipment, then Loop/Equipment Number

**Key**: `ItemNum|Tag`

#### Sync_Utility_Load()
**Purpose**: Tracks utility consumption (electric, air, hydraulic)

**Logic**:
1. Include parts where "Include in Utility Load Table?" = "Y"
2. One row per part (no tag splitting)
3. Copies all tag columns (ELEC, HYD, PNU)
4. Sort by Item#

**Key**: `ItemNum`

#### Sync_Heat_Noise()
**Purpose**: Tracks heat generation and noise levels

**Logic**:
1. Include parts where "Include in Heat Load & Noise Table?" = "Y"
2. One row per part
3. Sort by Item#

**Key**: `ItemNum`

**Column Locking**:
All destination syncs lock columns that come from Master Equipment List:
- Master Equipment List Item (always locked)
- Manufacturer (always locked)
- Part Number (always locked)
- Tags (always locked)
- User columns (Description, Notes, calculations) remain editable

---

### Module_Post_Sync_Format.txt
**Purpose**: Data normalization and visual formatting (NO protection logic)

**Main Function**: `Post_Sync_Format()`

**Key Operations**:

#### ManualEntryDef_Master()
1. Sets blank "Source" → "MAN"
2. Sets defaults for new manual entries:
   - P&ID Tags: ""
   - Removed from BOM: "N"
   - Notes: ""
3. **Auto-fills Include columns**: Sets blank "Include in..." → "N"
4. Clears inherited shading
5. Unlocks manual rows

#### MarkLockedItems_Master()
Visual indicator only - shades locked BOM rows gray (RGB 230, 230, 230)

**Important**: This module does NOT handle protection. All locking logic is in `Global_Protect()`.

---

### Module_Delete_Removed_Parts.txt
**Purpose**: Cleanup utility for removed parts

**Main Function**: `Delete_Removed_Parts()`

**User Options**:
1. **Yes**: Bulk delete ALL rows where "Removed from BOM" = "Y"
2. **No**: Review each removed row individually:
   - Shows: Source, Manufacturer, Part Number, Description
   - User chooses: Yes (delete), No (keep), Cancel (stop)
3. **Cancel**: Exit without changes

**Workflow**:
1. Unprotect sheets
2. Find all rows with "Removed from BOM" = "Y"
3. Delete based on user choice
4. Run `Post_Sync_Format()` (defaults, shading)
5. Run `Global_Protect()` (locking)
6. Show deletion count

---

### Sheet_Master_Equipment_List.txt
**Purpose**: Worksheet event handler for on-the-fly defaults

**Event**: `Worksheet_Change()`

**Behavior**:
When user edits a cell in Master Equipment List table:
1. Check if row has blank "Source"
2. If yes:
   - Set Source = "MAN"
   - Set defaults for Include columns (if blank) → "N"
   - Unlock row
3. Re-apply protection

**Smart Defaulting**:
- Only sets values if currently blank
- Doesn't overwrite user entries
- Runs only on changed rows (efficient)

---

## Data Flow

### Full Sync Flow (Run_Master_Equipment_Sync)

```
1. User clicks "Sync from BOMs" button
   ↓
2. Global_Unprotect() - unlock all sheets
   ↓
3. Refresh Power Query connection
   ↓
4. All_BOM_Parts table populates from source BOMs
   ↓
5. Sync_Master_Equipment_List()
   ├─ Build dictionary from All_BOM_Parts
   ├─ Update existing master rows
   ├─ Mark removed parts (Removed from BOM = Y)
   └─ Add new parts
   ↓
6. Sync_All_Destinations()
   ├─ Sync_IO_List()
   ├─ Sync_PID_Tag_List()
   ├─ Sync_Utility_Load()
   └─ Sync_Heat_Noise()
   ↓
7. Post_Sync_Format()
   ├─ Set defaults (blank Source → MAN, Include → N)
   └─ Apply visual shading
   ↓
8. Global_Protect()
   ├─ Apply row-level locking (Master)
   ├─ Apply column-level locking (Destinations)
   └─ Protect all sheets
   ↓
9. Show completion message
```

### Quick Update Flow (Apply_Changes)

```
1. User sets Include flags, clicks "Apply Changes"
   ↓
2. Global_Unprotect()
   ↓
3. Sync_All_Destinations() - NO BOM refresh
   ↓
4. Post_Sync_Format()
   ↓
5. Global_Protect()
   ↓
6. Show completion message
```

---

## Protection & Locking

### Locking Strategy

The system uses **three-tier protection**:

#### Tier 1: Sheet Protection
All sheets protected with:
- Password: "" (blank - easy to manually unprotect if needed)
- AllowFiltering: True
- AllowSorting: True
- AllowInsertingRows: True (for user entries)
- AllowFormattingColumns: True
- AllowFormattingRows: True

#### Tier 2: Cell-Level Locking
Individual cells have `.Locked` property set based on data source.

#### Tier 3: Visual Indicators
Gray shading (RGB 230, 230, 230) shows locked BOM data.

### Master Equipment List Locking Rules

| Source | Locking Behavior |
|--------|------------------|
| **ELEC, HYD, PNU, MECH** | Most columns locked<br>P&ID Tags, Include columns, User Entries, Notes: unlocked |
| **MAN** (Manual Entry) | Fully unlocked (user's data) |
| **Removed = "Y"** | Fully unlocked (for user review/edit) |

**Editable Columns (even for BOM rows)**:
- P&ID Tags
- Include in I/O List?
- Include in Utility Load Table?
- Include in Heat Load & Noise Table?
- User Entries
- Notes

**Always Unlocked**:
- Project info cells (C2:C7, E6) - merged cells

### Destination Table Locking Rules

#### IO List, Utility Load, Heat & Noise
**Locked columns** (from Master):
- Master Equipment List Item
- Manufacturer
- Part Number
- All tag columns (ELEC Tags, HYD Tags, PNU Tags)
- QTY (Utility/Heat only)

**Unlocked columns** (user editable):
- Description
- Notes
- Calculations
- User-specific columns

#### P&ID Tag List
**Manual entries** (blank Item#): Fully unlocked
**Synced entries** (has Item#):
- Locked: Item, Manufacturer, Part Number, P&ID Tag
- Unlocked: All other columns

---

## Troubleshooting

### Power Query Issues

#### Problem: "All_BOM_Parts returned 0 rows"
**Cause**: BOM files not found or empty
**Solution**:
1. Verify project path in Config sheet
2. Check BOM files exist in expected folders:
   - `[ProjPath]\Job Documentation\Electrical\[JobNum]-ELEC BOM (INTERNAL).xlsx`
3. Verify each BOM file has a table named "BOM"
4. Run `Create_BOM_Query` to regenerate

#### Problem: Query shows error "Column 'X' not found"
**Cause**: BOM table column names don't match All_BOM_Parts
**Solution**:
1. Check All_BOM_Parts table column names
2. Verify source BOM tables have matching columns
3. Column names are case-sensitive
4. Run `Create_BOM_Query` to sync column list

#### Problem: Power Query doesn't load to table
**Cause**: Connection created but not loaded to worksheet
**Solution**:
1. Data → Queries & Connections
2. Right-click "All_BOM_Parts" → Load To...
3. Choose: Table, Existing worksheet, Cell A3 (All_BOM_Parts sheet)
4. Subsequent syncs will work automatically

### Sync Issues

#### Problem: Parts not appearing in Master Equipment List
**Check**:
1. Is All_BOM_Parts table populated? (View the table)
2. Are part numbers spelled identically across BOMs?
3. Run sync with VBA Immediate Window open (Ctrl+G) to see errors

#### Problem: Destination tables not updating
**Check**:
1. Are "Include in..." columns set to "Y"?
2. Is Master Equipment List Item# present?
3. Did you run `Apply_Changes` or full sync?
4. Check for protection errors in Immediate Window

#### Problem: Deleted parts keep reappearing
**Cause**: Parts still exist in source BOMs
**Solution**:
1. Remove from source BOM files first
2. Run sync - marks "Removed from BOM" = "Y"
3. Run `Delete_Removed_Parts` to permanently remove

### Protection Issues

#### Problem: Can't edit cells that should be editable
**Cause**: Protection applied incorrectly
**Solution**:
1. Manually unprotect sheet (Review → Unprotect, password is blank)
2. Run `Apply_Changes` to reapply correct locking
3. Check if cell is in editable column list

#### Problem: Entire row is locked when it shouldn't be
**Check**:
1. Is Source = "MAN"? Should be fully unlocked
2. Is "Removed from BOM" = "Y"? Should be fully unlocked
3. Run `Apply_Changes` to refresh locking

### Data Issues

#### Problem: Need QTY doesn't match expected total
**Cause**: Part appears in multiple BOMs with different quantities
**Verify**:
1. Check All_BOM_Parts table - see all instances of part
2. Need QTY = sum of all BOMs' Need QTY columns
3. Owner's QTY (ELEC) may differ from total Need QTY

#### Problem: Wrong BOM owns a shared part
**Cause**: Part appears in multiple BOMs, hierarchy determines owner
**Current Hierarchy**: ELEC > HYD > PNU > MECH
**Solution**: This is by design. ELEC always owns shared parts (sensors).

---

## Advanced Configuration

### Modifying BOM Hierarchy

To change which discipline owns shared parts:

1. Open **Module_Sync_Master_Equipment.txt**
2. Find `UpdateMasterRow()` function (around line 312)
3. Modify priority values:
   ```vba
   Dim currentPriority As Long
   currentPriority = 0
   If bomSource = "MECH" Then currentPriority = 1  ' Lowest
   If bomSource = "PNU" Then currentPriority = 2
   If bomSource = "HYD" Then currentPriority = 3
   If bomSource = "ELEC" Then currentPriority = 4  ' Highest (Owner)
   ```
4. Higher number = higher priority = owner

### Adding New Destination Tables

To create additional destination tables:

1. **Create Table**: Add new sheet with Excel table
2. **Update Module_Config.txt**:
   ```vba
   Public Const SHEET_NEW_TABLE As String = "New Table"
   Public Const TABLE_NEW_TABLE As String = "New_Table"
   ```
3. **Create Sync Function** in Module_Sync_Destinations.txt:
   ```vba
   Private Sub Sync_New_Table()
       ' Build dictionary from Master where Include column = "Y"
       ' Update/delete existing rows
       ' Add new rows
       ' Sort by Item#
   End Sub
   ```
4. **Call from Sync_All_Destinations()**:
   ```vba
   Public Sub Sync_All_Destinations()
       Sync_IO_List
       Sync_PID_Tag_List
       Sync_Utility_Load
       Sync_Heat_Noise
       Sync_New_Table  ' Add here
   End Sub
   ```
5. **Add Protection** in Module_Global_Functions.txt → Global_Protect():
   ```vba
   ' Lock columns from master
   ' Leave user columns unlocked
   ```

### Modifying Column Names

If BOM files use different column names:

1. **Update Module_Config.txt** constants:
   ```vba
   Public Const BOM_COL_PART_NUMBER As String = "Part No"  ' If BOMs use "Part No"
   ```
2. **Run Create_BOM_Query** to regenerate Power Query with new column names

### Customizing Editable Columns

To change which columns stay editable for BOM rows:

1. Open **Module_Global_Functions.txt**
2. Find `Global_Protect()` → Master Equipment List section
3. Modify `editableCols` array:
   ```vba
   editableCols = Array("P&ID Tags", "Include in I/O List?", _
                        "Include in Utility Load Table?", _
                        "Include in Heat Load & Noise Table?", _
                        "User Entries", "Notes", _
                        "Your New Column")  ' Add here
   ```

### Performance Optimization

For large projects (1000+ parts):

1. **Disable Screen Updating** during operations (already done)
2. **Reduce Logging**: Comment out `WriteSyncLog()` calls
3. **Limit Destination Syncs**: Only sync tables you actually use
4. **Optimize Sorting**: Remove sort operations if tables stay filtered

---

## File Structure

### Required Sheets
- **Master Equipment List**: Central data hub
- **IO List**: Electrical I/O points
- **P&ID Tag List**: Instrumentation tags
- **P&ID Loops**: Control loops (optional, not synced)
- **Utility Load Table**: Utility consumption tracking
- **Heat Load & Noise**: Thermal/acoustic data
- **All_BOM_Parts**: Power Query staging table (hidden from users)
- **Config**: System configuration
- **SYNC_Log**: Optional logging (commented out in code)

### VBA Modules
- **Module_Config**: Constants and configuration
- **Module_Global_Functions**: Core infrastructure
- **Module_Sync_Master_Equipment**: BOM → Master sync
- **Module_Sync_Destinations**: Master → Destinations sync
- **Module_Post_Sync_Format**: Data defaults and formatting
- **Module_Delete_Removed_Parts**: Cleanup utility
- **Sheet_Master_Equipment_List**: Worksheet event handler

### External Dependencies
- Source BOM files (4 separate Excel workbooks)
- Power Query (built into Excel 2016+)

---

## Best Practices

### Daily Use
1. **Morning**: Run "Sync from BOMs" to get latest from engineers
2. **Configure**: Set Include flags for new parts
3. **Distribute**: Run "Apply Changes" to update destinations
4. **Clean Up**: Weekly run "Delete Removed Parts" to remove obsolete items

### Data Entry
- **Let BOMs Drive**: Don't manually edit BOM-sourced data in Master
- **Use Manual Entries**: Add custom equipment with Source = "MAN"
- **Tag Organization**: Use consistent tag naming (PIT-100, FCV-200, etc.)

### Protection
- **Don't Override**: Never manually unlock BOM-sourced cells to edit
- **Request Changes**: If BOM data is wrong, fix it in source BOM file
- **Respect Locks**: Gray shaded cells are locked for data integrity

### Performance
- **Batch Changes**: Set multiple Include flags, then run Apply Changes once
- **Minimize Full Syncs**: Use "Sync from BOMs" only when BOMs actually changed
- **Keep BOMs Clean**: Remove test/dummy data from source BOMs

---

## Version History

**Version 1.0** (Current)
- Initial release with full BOM integration
- ELEC-priority hierarchy for shared parts
- Multi-destination sync with intelligent locking
- Manual entry support
- Removed part tracking
- Dropbox integration
- Resilient Power Query (handles missing BOMs)
- Auto-fill Include columns
- Worksheet event handler for on-the-fly defaults

---

## Support & Modifications

### Common Modifications
See **Advanced Configuration** section for:
- Changing BOM hierarchy
- Adding destination tables
- Customizing editable columns
- Modifying column names

### Debugging
1. Enable **Immediate Window** (Ctrl+G in VBA editor)
2. Add `Debug.Print` statements to trace execution
3. Set breakpoints on error-prone lines
4. Check `LAST_SYNC_DATE` named cell for last successful sync

### Error Messages
All user-facing errors use `MsgBox` with clear instructions. Check the message text for troubleshooting steps.

---

## Glossary

**BOM** (Bill of Materials): List of parts from a discipline (ELEC, HYD, PNU, MECH)

**Owner**: The discipline whose data is used when a part appears in multiple BOMs

**Source**: Which BOM(s) contain a part (e.g., "ELEC, HYD")

**Item#**: Sequential number assigned to each part in Master Equipment List

**Need QTY**: Total quantity needed across all BOMs (sum)

**Assy QTY**: Quantity from owner BOM (not summed)

**Include Flags**: User settings determining which destination tables receive each part

**Manual Entry**: User-added equipment (Source = "MAN")

**Removed from BOM**: Parts deleted from source BOMs but kept in Master for review

**Tag**: Instrument/equipment identifier (e.g., PIT-100, FCV-200)

**Power Query**: Excel's data transformation engine that loads BOMs

**Protection**: Excel sheet/cell locking to prevent accidental edits

**Destination Tables**: Specialized sheets that receive filtered data from Master

---

## Quick Reference

### User Buttons
| Button | Macro | Use When |
|--------|-------|----------|
| **Sync from BOMs** | `Run_Master_Equipment_Sync` | BOMs changed, need full refresh |
| **Apply Changes** | `Apply_Changes` | Include flags changed, quick update |
| **Delete Removed Parts** | `Delete_Removed_Parts` | Clean up obsolete parts |

### Key Concepts
- **ELEC owns shared parts** (sensors in multiple BOMs)
- **Include flags control destinations** (user decides where each part goes)
- **Gray = Locked** (BOM data, don't edit)
- **White = Editable** (user data, go ahead)
- **MAN = Manual Entry** (user's equipment, fully editable)

### Keyboard Shortcuts
- **Ctrl+G**: Immediate Window (debugging)
- **Alt+F11**: VBA Editor
- **F5**: Run macro (in VBA editor)
- **Ctrl+Break**: Stop running macro

---

*End of README*
