# Goal
Generate **VBA code** (MS Access) that **creates a minimal, robust, court-case–oriented set of forms** for linked PostgreSQL tables (ODBC). This is a **pilot data registration tool**: simple UX, stable forms, and **NO combo boxes that load all placenames** (50k+).

## Architecture

The core app is split into two modules for clarity (plus `automated_setup.vba` and `materialize_linked_tables.vba` for one-shot setup — see Files):

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
  - txtSourceText (source_text) -> multiline court case text

- Tab control (tabMain) with two pages:
  1) "Court Case Entries" page hosts sfrmCourtCaseEntries (datasheet)
     - SourceObject: Form.sfrmCourtCaseEntries
     - LinkMasterFields: `court_case_id`
     - LinkChildFields: `court_case_id`
  2) "Ruling" page hosts sfrmRuling (single)
     - SourceObject: Form.sfrmRuling
     - LinkMasterFields: `court_case_id`
     - LinkChildFields: `court_case_id`

  sfrmPersonOutcomes is embedded inside sfrmRuling (it links to `ruling_id`), see §5.

- Buttons:
  - cmdPrevious / cmdNext: DoCmd.GoToRecord , , acPrevious / acNext
  - cmdSaveCase: saves current record, enabled while the form is dirty
  - cmdNewCase: DoCmd.GoToRecord , , acNewRec
  - cmdOpenEntryDetail: opens frmCourtCaseEntryDetail for currently selected entry in sfrmCourtCaseEntries (disabled until a saved entry is selected)
  - cmdAddEntry: inserts a blank court_case_entry for the case and opens frmCourtCaseEntryDetail on it
  - cmdCreateRuling: creates the ruling for the current case (visible/enabled only while no ruling exists)

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
- Buttons:
  - cmdPickPlacename → opens frmPlacenameSearch, passes calling form + target control in OpenArgs
  - cmdAddPersonEntry → opens frmPersonSearch with `target=_return`; the selected person is INSERTed as a `person_entry` row for this entry
  - cmdDeletePersonEntry → deletes the selected row in sfrmPersonEntryByEntry (with confirmation)

- Embed subform:
  - sfrmPersonEntryByEntry
  - LinkMasterFields: `court_case_entry_id`
  - LinkChildFields: `court_case_entry_id`

## 4) sfrmPersonEntryByEntry (Grid: people in entry)
- RecordSource: `person_entry` LEFT JOIN `person` (for read-only name display)
- Default View: Datasheet (or Continuous)
- Fields:
  - person_id (do NOT load all persons combo in grid; use popup picker)
  - person_name (read-only, from the join)
  - community_id (combo OK)
  - land_rights_status_id (combo OK)
  - role_id (combo OK)
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
- Buttons:
  - cmdAddOutcome / cmdDeleteOutcome: add/delete person_outcome rows (AddOutcome inserts with a default outcome_type_id)
  - Note: cmdCreateRuling lives on `frmCourtCase`, not here (visible only while no ruling exists)
- Embeds subform sfrmPersonOutcomes:
  - LinkMasterFields: `ruling_id`
  - LinkChildFields: `ruling_id`

## 6) sfrmPersonOutcomes (Grid outcomes for ruling)
- RecordSource: `person_outcome`
- Default View: Datasheet
- Fields:
  - person_id (combo, limited to persons already linked to the current court case)
  - outcome_type_id (combo OK)
  - description

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
  - On load: shows top 50 placenames (ORDER BY placename)
  - cmdSearch requeries with LIKE on placename and/or parish_name (no row cap)
  - cmdSelect writes selected placename_id back to caller:
    - caller/target specified in OpenArgs (delimited string)
    - Example OpenArgs: `caller=frmCourtCaseEntryDetail;target=txtPlacenameId`
    - Also refreshes the caller's txtPlacenameDisplay via DLookup
  - After writeback, close search form

## 8) frmPersonSearch (Popup picker)
- PopUp = True, Modal = True
- Controls:
  - txtSearch (unbound)
  - cmdSearch
  - lstResults: person_id, full_name, birth_year, community_name
  - cmdSelect, cmdNewPerson, cmdCancel
- On load: shows top 50 persons (ORDER BY full_name)
- cmdNewPerson opens frmPerson with DataEntry=True and the same OpenArgs; on close the new person_id is written back to the caller and the search form closes (or, without OpenArgs, results are just requered)
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

## 11) frmCourtCaseList (Continuous list of all cases)
- RecordSource: `court_case` (ordered), includes a read-only "Has ruling" indicator
- cmdOpen: opens frmCourtCase on the selected case

## 12) frmCourtCaseEntriesList (Continuous list of all entries)
- RecordSource: `court_case_entry` LEFT JOIN `placename` (for name display)
- cmdOpenDetail: opens frmCourtCaseEntryDetail for the current row
- cmdPickPlacename: same placename picker as in the detail dialog

## 13) frmPersonEntryList (Continuous list of person entries)
- RecordSource: `person_entry` JOIN `court_case_entry`, with lookup combos (person, community, land rights, role) and curated text

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
These saved queries are created by `BuildAllForms()` for completeness. The picker
forms themselves use inline SQL (see above) and do not reference these QueryDefs:

## qPlacenameSearch
A parameterized query (DAO QueryDef) with `PARAMETERS [pSearch] Text ( 255 );`
Actual SQL (Access SQL):
- `SELECT TOP 200 p1.placename_id, p1.placename,
   IIf(Nz(pr.parish, '') <> '', pr.parish, p1.parish_name) AS parish_display,
   p1.serial_number
   FROM placename AS p1
   LEFT JOIN parish AS pr ON Val(p1.parish_code) = pr.parish_id
   WHERE (p1.placename LIKE [pSearch]) OR (p1.parish_name LIKE [pSearch]) OR (pr.parish LIKE [pSearch])
   ORDER BY p1.placename;`

## qPersonSearch
`PARAMETERS [pSearch] Text ( 255 );`
- `SELECT TOP 200 person_id, full_name, birth_year, community_name
   FROM person
   WHERE full_name LIKE [pSearch]
   ORDER BY full_name;`

(If Access SQL/ODBC has issues with TOP on linked tables, fallback: pull results via pass-through query or add WHERE + ORDER and let Access limit in listbox RowSource; but try TOP first.)

---

# Event Logic Requirements (VBA)
Implement these common helpers:

## BuildAllForms()
- Calls (subforms first, then parents, then list forms):
  - CreateQueries()
  - Create_sfrmCourtCaseEntries()
  - Create_sfrmPersonEntryByEntry()
  - Create_sfrmPersonOutcomes()
  - Create_sfrmRuling()
  - Create_frmCourtCaseEntryDetail()
  - Create_frmCourtCase()
  - Create_frmCourtCaseList()
  - Create_frmCourtCaseEntriesList()
  - Create_frmPlacenameSearch()
  - Create_frmPersonEntryList()
  - Create_frmPersonSearch()
  - Create_frmPerson()
  - (optional) Create_frmLookups()
- Also closes open forms/tables, deletes previously generated forms, applies datasheet captions.

## OpenArgs convention
Use a simple delimiter format:
`caller=<FormName>;target=<ControlName>`
Example:
`caller=frmCourtCaseEntryDetail;target=txtPlacenameId`
Special target `_return` (used by cmdAddPersonEntry): the picker INSERTs a `person_entry` row for the current entry instead of writing to a control.
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
On frmCourtCase:
- cmdCreateRuling_Click:
  - If no ruling exists for current `court_case_id`:
    - Insert into `ruling` with `court_case_id` and a default `ruling_type_id` (first row, DMin)
  - Requery the ruling subform, switch to the Ruling tab, hide the button

## Open Entry Detail button
In sfrmCourtCaseEntries:
- cmdEntryDetail_Click:
  - `DoCmd.OpenForm "frmCourtCaseEntryDetail", , , "court_case_entry_id=" & Me.court_case_entry_id, , acDialog`

## Add Court Case Entry button
On frmCourtCase (action row, between `cmdOpenEntryDetail` and `cmdCreateRuling`):
- cmdAddEntry: initially disabled + hidden (until first OnCurrent)
  - Visible + enabled for any saved case (has court_case_id) — a case may have MANY entries, so no existence check
  - Visible but disabled while the case is unsaved (mirrors cmdCreateRuling)
  - Click: INSERT a blank `court_case_entry` row for the current case, Requery the entries datasheet, then open `frmCourtCaseEntryDetail` on the new entry
  - State synced by `UpdateAddEntryButtonState()` from `frmCourtCase_OnCurrent` and `sfrmCourtCaseEntries_OnCurrent`

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
   - Alternative: import `automated_setup.vba` and run `AutomatedFullSetup()` — it does the linking, form build, and local materialization in one go

2. **Import VBA Modules**
   - Open VBA Editor (`Alt + F11`)
   - Import both modules:
     - File → Import File... → `minimal_app_generator.vba`
     - File → Import File... → `minimal_app_runtime.vba`

3. **Generate Forms**
   - In VBA Editor, open Immediate Window (`Ctrl + G`)
   - Type: `BuildAllForms` and press Enter
   - Wait 30-60 seconds for form generation
   - Immediate Window prints “BuildAllForms completed successfully.”

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

- **`minimal_app_generator.vba`** - Form generation code (~1600 lines)
- **`minimal_app_runtime.vba`** - Runtime event handlers (~1050 lines)
- **`automated_setup.vba`** - One-shot setup: `AutomatedFullSetup()` links tables, builds forms, runs the materialization pipeline
- **`materialize_linked_tables.vba`** - Copies linked PostgreSQL tables to local tables with indexes and FK relations
- **`MINIMAL_APP.md`** - This documentation
- **`ARCHITECTURE_OVERVIEW.md`** - Form interconnections and runtime behavior overview