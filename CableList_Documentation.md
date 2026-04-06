# Cable List — Program Documentation

**Version:** Cable 4.05  
**Platform:** Microsoft Access (MDB format)  
**Purpose:** Substation control cable design — from wire connections to full cable schedules and procurement reports

---

---

# SECTION 1 — USER GUIDE

## What This Program Does

Cable List takes your electrical connection data (which terminal connects to which terminal, across which locations) and automatically:

- Groups connections into cables
- Assigns cable tags and core numbers
- Accounts for shield and armour cores
- Calculates cable lengths from route distances
- Produces all the reports needed for construction and procurement

---

## Getting Started — Opening a Project

When the program starts, it opens with a blank (new) project loaded automatically.

To work on an existing project, use **File → Open** from the menu. The program looks for files with the `.cbl` extension. These are your project files. When you open one, all the project data is loaded into the program ready to work with.

To start a fresh project, use **File → New**.

Always save your work using **File → Save** or **File → Save As** before closing.

---

## Step-by-Step Workflow

The correct order to use the program is:

### Step 1 — Set Up Project Details
Go to **Project** in the menu and fill in the project information form. This includes:
- Project name, client name, consultant name
- Document number, revision, project code
- Designer, checker, approver, preparer names
- Cable tag starting number and step (e.g. start at 1000, step by 10)
- Cable tag prefix (e.g. "W-" so cables appear as W-1000, W-1010...)

### Step 2 — Enter Connection Data
Go to **Connection** in the menu. This opens the `tblConnection` table directly. Enter one row per core connection with:

| Field | Description |
|---|---|
| Location1 / Location2 | The two panel/cubicle locations the cable runs between |
| Device1 / Device2 | The device (relay, switch, terminal block name) at each end |
| Terminal1 / Terminal2 | The terminal number at each end |
| Document1 / Document2 | The drawing/document reference at each end |
| CableClass | The class of cable required (links to your cable class table) |

The program automatically enforces that Location1 is alphabetically before Location2 (it swaps sides if needed), so you don't have to worry about the order you enter them.

### Step 3 — Set Up Reference Tables
Before allocating cables, the following tables must be filled in correctly. Use the menus to access each one:

**Cable Class** — defines each cable class with its required cross-section, number of cores, type, and whether it is half-cross-section.

**Cable Type** — defines each cable type and whether it has a shield (screen) and/or armour.

**Cable Stock** — lists the cable types you have available, with core/cross-section/type details, the colour coding pattern, gland code, and total stock length.

**Split Connection** — defines how connections should be split across cables when more than one cable is needed between two locations for the same class. For example, if you have 24 cores needed, you may split them as 12+12 or 19+5.

**Distance** — the cable route distance in metres between each pair of locations. Use **Distance** in the menu to open the distance entry. The program fills in location pairs automatically from your connection data.

**Grounding Priority** — a priority number for each location. When a cable has a shield or armour, the earth connection goes to the location with the lower priority number.

**Gland Information** — gland types and their metric/PG sizes, linked to gland codes in the stock table.

**Section Information** — cable tray/duct section data for route-based reports.

**Determine Levels** — used to define tray routing levels for the section report.

### Step 4 — Check and Repair Data
Before allocating cables, always run **Check & Repair** from the menu. This does several important things automatically:

- Fixes any null device or terminal entries (sets them to "-")
- Corrects location ordering (swaps sides where Location1 > Location2)
- Checks that the number of cores between each location pair does not exceed the cable class capacity
- Fills in any missing distance entries with zero (and flags them as errors)
- Fills in any missing grounding priority entries with zero
- Adds any missing split connection rules
- Warns about terminals that have more than one core connected

Errors are shown in red, warnings in yellow in the repair table that opens. **Errors must be fixed before you can allocate.** Warnings are advisory.

### Step 5 — Allocate Cables
Once all data is correct with no errors, run **Allocate** from the menu. The program:

1. Groups connections by location pair and cable class
2. Determines how many cables are needed (using the split connection rules)
3. Assigns cable tags (starting from the number in Project settings, incrementing by the step)
4. Assigns core numbers to each connection
5. Adds spare cores to each cable
6. Adds shield and armour cores where the cable type requires them
7. Earths shield/armour at the location with lower grounding priority
8. Checks the total is consistent and shows a success message

After allocation, the **Balance** form opens automatically showing cable stock versus usage.

### Step 6 — Manual Adjustments (Optional)

After automatic allocation you can make manual changes:

**Edit Cores** — opens `frmEditCore`. Select a cable from the list, then select any core. You can change the device and terminal at either end. You can also mark a used core as spare (freeing it up) or fill in a spare core with connection details.

**Add Cable** — opens `frmAddCable`. Lets you manually add a completely new cable that was not in the original connection data. You enter the connection details core by core into a temporary table, then confirm the cable type and tag. The program validates against stock and checks for duplicate tags.

**Edit Cable** — opens `frmEditCable`. Lets you:
- Change the tag, core count, cross-section, or type of an existing allocated cable
- **Merge** two or more cables between the same locations into one larger cable
- **Split** one cable into two or more smaller cables, redistributing the cores

Any manual change flags the project as "edited" which prevents re-running Allocate (to protect your manual work). If you need to re-allocate, you would need to clear the editing flag in `tblProject`.

### Step 7 — Generate Reports

All reports are accessed from the menu. Each report collects and prepares data then opens a print preview. The available reports are:

| Menu Item | Report Description |
|---|---|
| **Two Connection List** | Full cable schedule showing both ends of every core |
| **One Connection List** | Single-side cable schedule (each location's view of its cables) |
| **Cutting List** | Cable cutting lengths — cable tag, route, distance |
| **Used Cable** | Summary of total cable length used per type and size |
| **Allocated Cable Length** | Query showing allocated cable details |
| **Gland Punching** | List of cable glands needed per location and device |
| **Gland Order** | Total gland quantities by gland code for procurement |
| **Cable Lugs** | Cable lug quantities by cross-section (calculated with 20% excess) |
| **Core Tag Order** | Ferrule/core tag labels for ordering |
| **Cable Tags** | Cable nameplate tag labels |
| **Cable Path / Section** | Cable route and tray section report |
| **Accessories Cover** | Cover sheet for the accessories reports |
| **Balance** | Stock versus usage balance sheet |

---

## File Menu Summary

| Menu Item | What It Does |
|---|---|
| **New** | Clears current data and starts a blank project |
| **Open** | Opens an existing `.cbl` project file |
| **Save** | Saves current data back to the open project file |
| **Save As** | Saves current data to a new file location |

When you close the program, it will ask if you want to save changes. The program also automatically keeps a backup copy called `back.cbl` in the program folder.

---

---

# SECTION 2 — DEVELOPER / TECHNICAL GUIDE

## Architecture Overview

The program is a Microsoft Access MDB split into two parts:

**Main Program (this database)** — contains all forms, queries, reports, VBA modules, and a working copy of all data tables. This is what the user runs.

**Project Files (.cbl)** — renamed MDB files that contain only data tables (no forms, queries, or code). They are opened and saved by the main program using VBA. The `.cbl` extension was chosen deliberately to disguise them from casual users.

All data manipulation happens inside the main program's tables. Loading and saving copies the table contents to/from the project file.

---

## Module Structure

### modDeclearation
Global constants and the `Distancece()` and `color()` subroutines.

**Global Constants:**
- `c_ProgramName` = "Cable List" — used in message box titles
- `c_ShieldCore` = 254 — the special CoreNo value used to represent a shield core
- `c_ArmourCore` = 255 — the special CoreNo value used to represent an armour core

**`Distancece()`** — Called before opening the Distance table. It first re-runs the location swap repair (same logic as in `RepairCheckCableData`) to ensure Location1 < Location2, then synchronises `tblDistance` with the current connection data, preserving any distances already entered and filling new location pairs with nulls.

**`color()`** — Assigns colour codes to `CoreTag` in `tblConnection` based on the colour string stored in `tblCableStock`. Calls `updatingStock` first. It reads the colour character sequence from the stock table and maps each character (formatted to lowercase) to the corresponding core's `CoreTag`. Shield cores get tag "C", armour cores get tag "R", and numeric cores get zero-padded numbers by default.

---

### modDeclearation (also contains)
The `color()` sub uses `Format(Mid(color, i, 1), "<")` to convert colour characters to lowercase. This is the Access VBA way of applying a lowercase format mask.

---

### modFile
Handles all file operations. Uses two module-level variables:
- `m_State` — Byte flag: `1` = a project file is currently open, `0` = new/unsaved project
- `m_dbName` — String holding the full path of the current project file

**`C_CABLE_DIR_PATH`** — Must point to the folder where the program support files live (`new`, `start.cbl`, `final.cbl`, `back.cbl`). **This constant must be updated if the program is moved to a new machine.**

**`C_CABLE_FILE_EXT`** = `.cbl` — the project file extension filter used in file dialogs.

**`Program_Begin()`** — Called from `Form_Load` of `frmMain`. Initialises the program by loading the blank `new` template file. Should check that `tblConnection` exists in the main program before proceeding (see Bug 1 fix). Sets the title bar to show the version and current file name.

**`program_end(Cancel)`** — Called from `Form_Unload` of `frmMain`. Asks the user whether to save. If the user clicks Cancel, sets `Cancel = -1` to abort the close. Otherwise saves to `final.cbl` and allows the close.

**`mnuOpen()`** — File open dialog filtered to `.cbl` files. Validates the selected file contains `tblConnection` before proceeding (to reject non-project files). If a file is already open, backs up current data to `final.cbl` first. Then calls `read_tables_from()` and `copy_file()` to back up the opened file to `start.cbl`.

**`mnuNew()`** — Backs up current data if open, then loads the blank `new` template.

**`mnuSave()` / `mnuSaveAs()`** — Save current table data to the project file. `SaveAs` first copies the blank `new` template to the chosen location to create an empty container, then writes data into it.

**`copy_file(src, dst)`** — Uses `Scripting.FileSystemObject` to copy a file. Used to create backups and to copy the blank template when saving a new file.

**`copy_tables_to(dbName)`** — Saves all data tables from the main program into the external project file. Uses `DELETE ... FROM table IN 'file'` followed by `INSERT INTO table IN 'file' SELECT ... FROM table` for each table. Tables saved: tblConnection, tblAllocatedCable, tblCableClass, tblCableStock, tblCableType, tblDescription, tblDetermineLevels, tblDistance, tblGlandInformation, tblGroundingPriority, tblProject, tblSectionInformation, tblSplitConnection.

**`read_tables_from(dbName)`** — Loads all data tables from the external project file into the main program. For `tblConnection` specifically it uses `DeleteObject` + `TransferDatabase acImport` rather than DELETE/INSERT — this is intentional to preserve the AutoNumber sequence in the imported data. All other tables use DELETE then INSERT from the external file.

> **⚠ Critical Risk:** The delete-then-import of `tblConnection` is the root cause of `tblConnection` disappearing from the main program. If the import fails after the delete, the table is permanently gone until restored from `start.cbl`. Error handling must wrap these two lines together.

---

### modCable
Contains the core algorithmic logic of the program.

**`Allocate()`** — The main cable allocation engine. Flow:

1. Calls `RepairCheckCableData()` — aborts if any errors found
2. Deletes spare/shield/armour rows from `tblConnection` to start clean
3. Creates `tblSumOfConnection` — groups connections by Location1, Location2, CableClass and counts cores needed
4. Reads `CableTagStart` and `CableTagStep` from `tblProject`
5. Joins `tblSumOfConnection` with `tblCableClass`, `tblCableType`, and `tblSplitConnection` to determine which cable(s) satisfy each group
6. For groups where `core <> 0` (single cable satisfies): adds one record to `tblAllocatedCable` directly
7. For groups where `core = 0` (multiple cables needed): reads `core1` through `core10` fields from `tblSplitConnection` to determine the split. Calculates spare core distribution proportionally across the split cables using integer rounding with remainder correction
8. Assigns sequential cable tags to all allocated cables (TagStart, TagStart+TagStep, ...)
9. Assigns core numbers to each connection record in `tblConnection`, matched by location/class/device/terminal ordering
10. Adds spare core rows to `tblConnection` for each cable
11. Adds shield (CoreNo=254) and armour (CoreNo=255) rows where required, earthed at the lower-priority location
12. Validates that sum of `UsedCore` in `tblAllocatedCable` equals count of real connections in `tblConnection`
13. Calls `mnuBalance()` to show the stock balance

**`RepairCheckCableData(hasError)`** — Pre-allocation data validation and repair. Flow:

1. Clears `tblRepairAndCheck`
2. Replaces null Device/Terminal values with "-"
3. Swaps Location1/Location2 (and all side1/side2 fields) where Location1 > Location2 — ensures consistent ordering
4. Checks used cores per location/class does not exceed cable class capacity — writes errors to `tblRepairAndCheck`
5. Checks `tblSplitConnection` for any missing entries and auto-adds them with `Core1 = SumOfConnection` (single cable, no split)
6. Checks for terminals used by more than one core — writes warnings
7. Fills missing grounding priority entries with 0
8. Fills missing distance entries with 0 and writes errors for zero distances
9. Sets `hasError = True` if any "Error" level entries exist in `tblRepairAndCheck`
10. Opens `tblRepairAndCheck` for the user if any errors or warnings exist

**`updatingStock()`** — Ensures `tblCableStock` contains a row for every cable type used in `tblAllocatedCable`. Any cable in allocated but missing from stock gets added with `Stock = 0` and a warning written to `tblRepairAndCheck`.

---

### ModMenu
Contains one sub per menu item. Each sub prepares data (usually by running DELETE/INSERT INTO queries to build report-specific tables) and then opens the relevant form or report.

Notable items:

**`mnuAllocate()`** — Checks `EditingData` flag in `tblProject` before calling `Allocate()`. If the project has been manually edited, allocation is blocked to protect manual work.

**`mnuBalance()`** — Creates `tblUsedCable` (cable usage by type/size) and `tblBalance` (stock minus usage), then opens `frmBalance`.

**`mnuGlandPunching()`, `mnuGlandOrder()`** — Both call the `Gland()` function (not in the uploaded modules — likely in a form or missing module) before building gland report tables.

**`mnuTwoConnection()`, `mnuOneConnection()`** — Both call `color()` first (to assign core colour tags), then build the report tables and open the respective report.

**`mnuSectionInformation()`** — Loops 1 to 20 inserting section/level data from tblDistance and tblAllocatedCable into `tempSection`, then builds `tblDetailOfSections` for the report.

**`getfile()`** — Incomplete function, dead code. Contains an unfinished `If Err = 0 Then` with no body and no `End If`. Should be removed.

---

### modLock
Copy-protection mechanism — all code commented out. When active it checked:
1. That the current date was not before a base date (`C_DEF_DATE` = 10/12/2006)
2. That a key file existed at a network path (`O:\KAMANKESH-HAMID\BSAF.DLL`)
3. That the key file contained a number of months, and that the current date had not passed the base date plus that number of months

The same checks (now commented out) also appear in `Form_Load` and `Form_Unload` of `frmMain`. All checks are disabled and the `checkLock` sub is now effectively empty.

---

### other
Contains two utility items:

**`copyrec()`** — Uses `DoMenuItem` to perform Select All / Copy / Paste Append. This is an old Access 97/2000 style record duplication helper, likely assigned to a button on a form.

**`HandingEdit()`** — Sets `EditingData = True` in `tblProject`. Called after every manual change (add cable, edit cable, edit core, merge, split). This flags the project as having been manually modified, which blocks the Allocate function from overwriting manual work.

---

## Form Code Summary

### frmMain (Form_frmMain)
The main switchboard — the startup form of the application.

- **`Form_Load`** — Calls `Program_Begin()`. Lock check code is present but commented out.
- **`Form_Unload`** — Calls `program_end(Cancel)` which handles the save-on-close prompt and can cancel the close if the user clicks Cancel.
- **`Form_Timer`** — Timer event present but commented out. Was intended to call `backup()` every 15 minutes automatically.

**To restore the menu:** In Access Startup settings (File → Options → Current Database), set Display Form to `frmMain`.

---

### frmAddCable (Form_frmAddCable)
Manual cable addition form. Uses a temporary table `tblNewConnection` where the user types in core connections before confirming.

**`Add_Click()`** — Validation checks performed in order:
1. Required fields (Location1, Location2, Core, CrossSec, Type, CableTag) must not be null
2. Cable tag must not already exist in `tblAllocatedCable`
3. Number of cores entered must not exceed the new cable's core count
4. The cable type/size must exist in `tblCableStock`
5. At least one core must have been entered in `tblNewConnection`

If all checks pass:
- Applies location swap if Location1 > Location2
- Assigns CoreNo sequentially to each row in `tblNewConnection`
- Adds spare cores to make up the full cable count
- Adds shield and/or armour cores based on cable type, earthed at the lower-priority location
- Inserts the new cable into `tblAllocatedCable`
- Moves all rows from `tblNewConnection` into `tblConnection`
- Calls `HandingEdit()` to flag manual edit
- Calls `color()` to update core colour tags
- Opens `frmDeterminLevels` after completion

---

### frmEditCable (Form_frmEditCable)
Shows all allocated cables in a list (`lstAllocate`). Supports three operations:

**`lstAllocate_Click()`** — Populates module-level variables `m_Tag`, `m_Core`, `m_cross`, `m_type`, `m_loc1`, `m_loc2` from the selected row. Also fills the edit text boxes.

**`cmbEditCable_Click()` (Edit button)** — Validates and updates the cable's tag, core count, cross-section, and type. If cores are increased, adds spare rows to `tblConnection`. If cores are decreased, deletes spare rows from the bottom (by CoreNo descending).

**`cmbMerge_Click()` (Merge button)** — Merges all selected cables in `lstAllocate` into the first selected cable's tag. All selected cables must have the same Location1, Location2, and CrossSec. Deletes shield/armour/spare cores from the merged cables, reassigns all connection rows to the surviving cable tag, renumbers cores sequentially, adds spares, and re-adds shield/armour cores.

**`cmbSplit_Click()` (Split button)** — Splits the selected cable into up to 5 new cables. The user enters a new tag and core count for each split cable (txtTag1..txtTag5, txtCore1..txtCore5). Validates that total cores across splits >= used cores. Distributes spare cores proportionally. Re-assigns connection rows to the new cable tags, adds spare rows, adds shield/armour cores for each new cable.

---

### frmEditCore (Form_frmEditCore)
Shows all cores of all allocated cables in a list (`lstConnection`). Allows editing individual core assignments.

**`lstConnection_Click()`** — Populates the form fields from the selected core. Saves previous device/terminal values in `prvDevice1/2`, `prvTerminal1/2` for change detection. Disables the Spare button for shield/armour cores or already-spare cores. Disables terminal fields for shield/armour cores (only "earth" is valid for those).

**`cmdEdit_Click()` (Edit button)** — For shield/armour cores: only allows "earth" or blank as device values. For normal cores: checks that the new device/terminal combination does not already exist at the same location in another connection (duplicate terminal check). Prevents setting side1 and side2 to identical values. Calls `updatAndrefresh()`. If the previous device was "spare", increments `UsedCore` in `tblAllocatedCable`.

**`cmdSpare_Click()` (Mark Spare button)** — Sets the selected core's device1/2 to "spare", clears terminal and document fields. Decrements `UsedCore` in `tblAllocatedCable` by 1.

**`updatAndrefresh()`** — Runs the UPDATE SQL on `tblConnection` using `ItemNo` as the key, then clears the text boxes and refreshes the list.

---

## Key Tables Reference

| Table | Purpose |
|---|---|
| `tblConnection` | One row per core — the central data table |
| `tblAllocatedCable` | One row per allocated cable |
| `tblProject` | Single-row project settings and title block info |
| `tblCableClass` | Maps cable class codes to core/cross-section/type specs |
| `tblCableType` | Cable types with Shield and Armour (ar) boolean flags |
| `tblCableStock` | Available cable inventory with colour codes and gland codes |
| `tblSplitConnection` | Rules for splitting N connections across multiple cables |
| `tblDistance` | Route distances between location pairs |
| `tblGroundingPriority` | Earthing priority per location (lower = earth side) |
| `tblGlandInformation` | Gland sizes (metric and PG) per gland code |
| `tblSectionInformation` | Cable tray/duct section data |
| `tblDetermineLevels` | Tray level configuration |
| `tblRepairAndCheck` | Temporary — errors and warnings from Check & Repair |
| `tblDescription` | User-defined descriptions (used in reports) |
| `tblNewConnection` | Temporary — staging table for frmAddCable |

---

## Special CoreNo Values

| Value | Meaning |
|---|---|
| 01, 02, 03... | Normal cable cores |
| 254 (`c_ShieldCore`) | Shield (screen) core |
| 255 (`c_ArmourCore`) | Armour core |

---

## Backup File Strategy

| File | Created When | Purpose |
|---|---|---|
| `start.cbl` | On Open | Snapshot of file as it was when opened — recovery point |
| `final.cbl` | On Close / New / Open (if file open) | Last saved state before switching files |
| `back.cbl` | `backup()` sub (timer — currently disabled) | Rolling backup |
| `new` | Shipped with program | Blank template used for New File and SaveAs |

---

## Known Issues and Recommendations

| Issue | Location | Recommendation |
|---|---|---|
| `tblConnection` can be permanently deleted if import fails after delete | `read_tables_from()` in modFile | Wrap delete+import in error handler; restore from `start.cbl` on failure |
| `Program_Begin()` references undefined variable `dbName` | modFile line 93 | Remove the IN clause; check main program's own tblConnection instead |
| `getfile()` function is incomplete — no `End If`, no return value | ModMenu line 183 | Delete the entire function if unused |
| Loose `mnuAllocatedCableLength` statement outside any sub | ModMenu line 191 | Delete the stray line (now fixed) |
| `m_State` written as `State` | Program_Begin in modFile | Fixed — use `m_State` consistently |
| `C_CABLE_DIR_PATH` hardcoded to original developer PC path | modFile line 5 | Update to current machine path (now fixed) |
| `EditingData` flag has no reset mechanism visible in code | tblProject / frmMain | Add a menu option or button to reset `EditingData = False` if re-allocation is needed after manual edits |
| Timer-based backup is disabled | frmMain Form_Timer | Consider re-enabling with `Me.TimerInterval = 900000` in Form_Load |
