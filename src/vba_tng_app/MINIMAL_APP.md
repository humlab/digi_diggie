# Goal
Generate **VBA code** (MS Access) that **creates a minimal, robust, court-case–oriented set of forms** for linked PostgreSQL tables (ODBC). This is a **pilot data registration tool**: simple UX, stable forms, and **NO combo boxes that load all placenames** (50k+).

## Architecture

The VBA code is split into two modules for clarity:

1. **`minimal_app_generator.vba`** - Form generation code
   - Run `BuildAllForms()` to create all forms, queries, and controls
   - Contains all `Create_*` functions
   - Only needed during initial setup or form regeneration

2. **`minimal_app_runtime.vba`** - Runtime event handlers
   - Contains all event handlers called by form controls
   - Helper functions for search, navigation, etc.
   - Must remain in the database for forms to function

Both modules must be imported into your Access database:
- Import `minimal_app_generator.vba` → Run `BuildAllForms()` → Forms created
- Import `minimal_app_runtime.vba` → Event handlers available to forms

---

# Assumptions / Constraints
- Tables are **linked** in Access with the same names as in PostgreSQL schema:  
  `court_case`, `court_case_entry`, `ruling`, `person`, `person_entry`, `person_outcome`, plus lookup tables.
- Primary keys are `*_id` integer, already present in linked tables.
- Use **DAO** (`CurrentDb`, `CreateQueryDef`, `CreateForm`, `CreateControl`, `CreateProperty`) and standard form objects.
- Avoid heavy row sources in datasheet combos. Specifically:
  - `placename` has **50,000+** rows → use **popup search**, not combo.
  - `person` has a couple thousand → also prefer popup search for performance + usability.
- Everything is “good enough”, but must be robust (basic error handling, avoid invalid selections, etc.).

---

# Deliverables (Forms)
Create these saved forms with the given names:

## 1) frmCourtCase (Main workspace)
- RecordSource: `court_case`
- Default View: Single Form
- Controls (bound):
  - cboSource (source_id) -> combo, small list OK
  - txtReferenceNumber (reference_number)
  - txtCaseYear (case_year)
  - txtDistrictCourtName (district_court_name)

- Subforms on frmCourtCase:
  1) sfrmCourtCaseEntries (datasheet/continuous)
     - SourceObject: Form.sfrmCourtCaseEntries
     - LinkMasterFields: `court_case_id`
     - LinkChildFields: `court_case_id`
  2) sfrmRuling (single)
     - SourceObject: Form.sfrmRuling
     - LinkMasterFields: `court_case_id`
     - LinkChildFields: `court_case_id`
  3) sfrmPersonOutcomes (datasheet)
     - SourceObject: Form.sfrmPersonOutcomes
     - IMPORTANT: This links to `ruling_id`, so host it in a container that can reference sfrmRuling.
       Easiest pilot approach: place outcomes subform on frmCourtCase and set its RecordSource via code when ruling changes
       OR place outcomes as a subform inside sfrmRuling (recommended; see below).

- Buttons:
  - cmdNewCase: DoCmd.GoToRecord , , acNewRec
  - cmdOpenEntryDetail: opens frmCourtCaseEntryDetail for currently selected entry in sfrmCourtCaseEntries

## 2) sfrmCourtCaseEntries (Subform grid)
- RecordSource: `court_case_entry`
- Default View: Datasheet (or Continuous)
- Bound fields shown:
  - entry_year
  - season_id (combo OK)
  - land_use_id (combo OK)
  - placename_id (NOT a full combo; show as numeric + a read-only calculated display if possible)
  - original_placename
  - curated_text (can be shown or edited in popup later)
- Include a button per row (or in footer/header) to open:
  - frmCourtCaseEntryDetail filtered by court_case_entry_id
  - Name: cmdEntryDetail

## 3) frmCourtCaseEntryDetail (Popup detail for one entry)
- RecordSource: `court_case_entry`
- Default View: Single Form
- PopUp = True, Modal = True
- Fields:
  - entry_year, season_id (combo), land_use_id (combo), original_placename, curated_text
  - Show placename as:
    - txtPlacenameId (bound to placename_id) + txtPlacenameDisplay (unbound)
- Button:
  - cmdPickPlacename → opens frmPlacenameSearch, passes calling form + target control in OpenArgs

- Embed subform:
  - sfrmPersonEntryByEntry
  - LinkMasterFields: `court_case_entry_id`
  - LinkChildFields: `court_case_entry_id`

## 4) sfrmPersonEntryByEntry (Grid: people in entry)
- RecordSource: `person_entry`
- Default View: Datasheet (or Continuous)
- Fields:
  - person_id (do NOT load all persons combo in grid; use popup picker)
  - community_id (combo OK)
  - land_rights_status_id (combo OK)
  - role_id (combo OK)
  - curated_text
- Add button:
  - cmdPickPerson → opens frmPersonSearch, writes selected person_id to current row

## 5) sfrmRuling (Single: 0/1 ruling per case)
- RecordSource: `ruling`
- Default View: Single Form
- Bound fields:
  - ruling_year
  - ruling_type_id (combo)
  - legal_source_id (combo)
  - description
- Button:
  - cmdCreateRuling: if no ruling exists for current case, INSERT new row with current court_case_id and required ruling_type_id (pick default or prompt minimal)
- Embed subform inside sfrmRuling (recommended):
  - sfrmPersonOutcomes
  - LinkMasterFields: `ruling_id`
  - LinkChildFields: `ruling_id`

## 6) sfrmPersonOutcomes (Grid outcomes for ruling)
- RecordSource: `person_outcome`
- Default View: Datasheet
- Fields:
  - person_id (popup picker preferred; optional “limit to people in case” can be phase 2)
  - outcome_type_id (combo OK)
  - description
- Button:
  - cmdPickPersonOutcomePerson → opens frmPersonSearch, writes person_id

## 7) frmPlacenameSearch (Popup picker)
- Purpose: search placename without loading 50k rows.
- PopUp = True, Modal = True
- Controls:
  - txtSearch (unbound)
  - cmdSearch
  - lstResults (listbox) showing columns:
    - placename_id, placename, parish_name, serial_number
  - cmdSelect
  - cmdCancel
- Behavior:
  - cmdSearch builds/requeries results with LIKE on placename and/or parish_name
  - Use TOP 200 and ORDER BY placename
  - cmdSelect writes selected placename_id back to caller:
    - caller/target specified in OpenArgs (e.g., JSON-like or delimited string)
    - Example OpenArgs: `caller=frmCourtCaseEntryDetail;target=placename_id`
  - After writeback, close search form

## 8) frmPersonSearch (Popup picker)
- PopUp = True, Modal = True
- Controls:
  - txtSearch (unbound)
  - cmdSearch
  - lstResults: person_id, full_name, birth_year, death_year
  - cmdSelect, cmdNewPerson, cmdCancel
- cmdNewPerson opens frmPerson with DataEntry=True; after closing, requery results
- cmdSelect writes person_id back to caller/target in OpenArgs

## 9) frmPerson (Basic person editor)
- RecordSource: `person`
- Fields:
  - given_name, patronymic, surname, birth_year, death_year, community_name, note
- (No relationship subform for pilot)

## 10) frmLookups (Tabbed maintenance; optional if time)
- Tab control with datasheet subforms for small lookup tables:
  - source, season, land_use, land_rights_status, role_type, role, ruling_type, legal_source, outcome_type, parish, community
- Each subform RecordSource = table, DefaultView = Datasheet

---

# Lookup combos (small lists only)
Create combo RowSources for:
- source: `SELECT source_id, source_abbreviation, source_name FROM source ORDER BY source_abbreviation, source_name;`
- season: `SELECT season_id, season_name FROM season ORDER BY season_name;`
- land_use: `SELECT land_use_id, description FROM land_use ORDER BY description;`
- land_rights_status: `SELECT land_rights_status_id, land_rights_status FROM land_rights_status ORDER BY land_rights_status;`
- role: `SELECT role_id, role_name FROM role ORDER BY role_name;`
- ruling_type: `SELECT ruling_type_id, ruling_type FROM ruling_type ORDER BY ruling_type;`
- legal_source: `SELECT legal_source_id, legal_source_name FROM legal_source ORDER BY legal_source_name;`
- outcome_type: `SELECT outcome_type_id, outcome_type_name FROM outcome_type ORDER BY outcome_type_name;`
- community: `SELECT community_id, community_name FROM community ORDER BY community_name;`

Combo properties:
- BoundColumn = 1
- ColumnCount = 2 or 3
- ColumnWidths = "0cm;5cm;..." (hide id)
- LimitToList = True

---

# Queries to Create (Saved QueryDefs)
Create these saved queries (Access select queries) to support search listboxes:

## qPlacenameSearch
A parameterized query (DAO QueryDef) with `PARAMETERS [pSearch] Text ( 255 );`
SQL example (Access SQL):
- `SELECT TOP 200 placename_id, placename, parish_name, serial_number
   FROM placename
   WHERE placename LIKE [pSearch] OR parish_name LIKE [pSearch]
   ORDER BY placename;`
At runtime set pSearch to `"*" & Nz(txtSearch,"") & "*"`.

## qPersonSearch
`PARAMETERS [pSearch] Text ( 255 );`
- `SELECT TOP 200 person_id, full_name, birth_year, death_year
   FROM person
   WHERE full_name LIKE [pSearch]
   ORDER BY full_name;`

(If Access SQL/ODBC has issues with TOP on linked tables, fallback: pull results via pass-through query or add WHERE + ORDER and let Access limit in listbox RowSource; but try TOP first.)

---

# Event Logic Requirements (VBA)
Implement these common helpers:

## BuildAllForms()
- Calls:
  - CreateQueries()
  - Create_frmCourtCase()
  - Create_sfrmCourtCaseEntries()
  - Create_sfrmRuling()
  - Create_sfrmPersonOutcomes()
  - Create_frmCourtCaseEntryDetail()
  - Create_sfrmPersonEntryByEntry()
  - Create_frmPlacenameSearch()
  - Create_frmPersonSearch()
  - Create_frmPerson()
  - (optional) Create_frmLookups()
- Shows a MsgBox when done.

## OpenArgs convention
Use a simple delimiter format:
`caller=<FormName>;target=<ControlName>`
Example:
`caller=frmCourtCaseEntryDetail;target=placename_id`
Write parse function `ParseOpenArgs(ByVal s As String) As Scripting.Dictionary` OR a simple Split-based parser (avoid external refs if possible).

## Picker write-back
In frmPlacenameSearch / frmPersonSearch:
- On Select:
  - Read caller and target from OpenArgs
  - `Forms(caller).Controls(target).Value = selectedId`
  - `Forms(caller).Dirty = False` (or leave dirty)
  - Optionally update a display textbox (e.g., txtPlacenameDisplay) by DLookup or a small query
  - Close picker

## Create ruling button
In sfrmRuling:
- cmdCreateRuling_Click:
  - If the subform has no record for current `court_case_id`:
    - Insert into `ruling` with `court_case_id` and required `ruling_type_id`
    - Choose a default ruling_type_id: the first row from ruling_type table (DMin / DLookup)
  - Requery the ruling form and outcomes subform

## Open Entry Detail button
In sfrmCourtCaseEntries:
- cmdEntryDetail_Click:
  - `DoCmd.OpenForm "frmCourtCaseEntryDetail", , , "court_case_entry_id=" & Me.court_case_entry_id, , acDialog`

---

# Implementation Notes (Must Follow)
- Use `DoCmd.SetWarnings False/True` sparingly; prefer DAO execute with dbFailOnError.
- Include `Option Compare Database` and `Option Explicit` in all modules.
- Add basic error handling in each public Sub/Function:
  - `On Error GoTo ErrHandler`
  - show MsgBox with procedure name and Err.Description
- Set reasonable form properties:
  - NavigationButtons = True on main forms; False on subforms
  - RecordSelectors = True for datasheet subforms
  - AllowAdditions/AllowEdits = True
  - DataEntry = False except when explicitly creating a new person
- Layout: simple vertical stacking; do not over-design.

---

# Usage Instructions

## Setup (First Time)

1. **Link PostgreSQL Tables to Access**
   - Follow [LINKED-DATABASE.md](../../docs/LINKED-DATABASE.md) guide
   - Create ODBC DSN to your PostgreSQL database
   - Link all tables from `digidiggie_tng` schema

2. **Import VBA Modules**
   - Open VBA Editor (`Alt + F11`)
   - Import both modules:
     - File → Import File... → `minimal_app_generator.vba`
     - File → Import File... → `minimal_app_runtime.vba`

3. **Generate Forms**
   - In VBA Editor, open Immediate Window (`Ctrl + G`)
   - Type: `BuildAllForms` and press Enter
   - Wait 30-60 seconds for form generation
   - Message box will confirm completion

4. **Start Using**
   - Close VBA Editor
   - Open form `frmCourtCase` to begin data entry

## Regenerating Forms

If you need to recreate forms (e.g., after database changes):

1. **Delete Existing Forms** (optional but recommended)
   - In Access Navigation Pane, delete old forms

2. **Re-run Generator**
   - VBA Editor → Immediate Window
   - Type: `BuildAllForms` and press Enter

**Note**: Custom changes to forms will be lost. Modify the generator VBA instead for persistent changes.

## Module Responsibilities

### minimal_app_generator.vba
- **Purpose**: Creates all forms and queries
- **Main Function**: `BuildAllForms()` - Entry point
- **When Needed**: Only during initial setup or regeneration
- **Can Remove**: After forms are created (but keep for future regeneration)

### minimal_app_runtime.vba
- **Purpose**: Provides event handlers for form controls
- **Functions**: All `frm*` and `sfrm*` event handlers
- **When Needed**: Always - must remain in database
- **Cannot Remove**: Forms depend on these functions

## Troubleshooting

**Forms don't work after generation:**
- Ensure `minimal_app_runtime.vba` is imported
- Check button OnClick properties reference correct function names

**"Can't find table" error:**
- Verify all tables are linked in Access
- Check table names match exactly (case-sensitive in some cases)

**"Compile error" when opening forms:**
- Tools → References in VBA Editor
- Ensure Microsoft DAO 3.6 Object Library is checked
- Remove any missing references

---

# Files

- **`minimal_app_generator.vba`** - Form generation code (~870 lines)
- **`minimal_app_runtime.vba`** - Runtime event handlers (~390 lines)
- **`minimal_app.vba`** - DEPRECATED - Combined version (kept for reference)
- **`MINIMAL_APP.md`** - This documentation