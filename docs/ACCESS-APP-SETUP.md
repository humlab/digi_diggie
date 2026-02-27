# MS Access App Setup Guide

Complete step-by-step guide to create and configure the DigiDiggie TNG MS Access application from scratch.

## Table of Contents

1. [Prerequisites](#prerequisites)
2. [Quick Start](#quick-start)
3. [Detailed Setup Steps](#detailed-setup-steps)
4. [Running the VBA Form Generator](#running-the-vba-form-generator)
5. [Post-Setup Configuration](#post-setup-configuration)
6. [Troubleshooting](#troubleshooting)

---

## Prerequisites

### Required Software

- **Windows 11** (or Windows 10)
- **Microsoft Access 2016 or later** (part of Office Professional or Microsoft 365)
- **PostgreSQL ODBC Driver** (32-bit or 64-bit matching your Access installation)
- **PostgreSQL Database** with `digidiggie_tng` schema (local or remote)

### Required Knowledge

- Basic familiarity with MS Access
- Basic understanding of VBA (for running macros)
- Access to PostgreSQL connection details (host, port, database, username, password)

### Database Schema

Your PostgreSQL database must have the `digidiggie_tng` schema created. The schema includes tables such as:
- `court_case`, `court_case_entry`, `person`, `person_entry`
- `community`, `parish`, `source`, `ruling`, `role`
- `placename` (with PostGIS geometry), and lookup tables

**Schema file location**: [`src/schema/digidiggie_en_tng.sql`](../src/schema/digidiggie_en_tng.sql)

---

## Quick Start

For experienced users, here's the condensed version:

```bash
# 1. Install PostgreSQL ODBC driver
# 2. Create System DSN named "DigiDiggie_TNG" pointing to your PostgreSQL database
# 3. Create blank Access database (.accdb)
# 4. Link all tables from digidiggie_tng schema via External Data → ODBC
# 5. Import VBA module from src/vba_tng_app/generate_tng_app.vba
# 6. Run BuildForms_DigiDiggie_TNG() macro
# 7. Done! Your forms are generated.
```

---

## Detailed Setup Steps

### Step 1: Install PostgreSQL ODBC Driver

Follow the comprehensive guide in [`LINKED-DATABASE.md`](./LINKED-DATABASE.md#step-1-install-postgresql-odbc-driver).

**Key points:**
- Download from https://www.postgresql.org/ftp/odbc/releases/
- Choose 32-bit or 64-bit to match your Access installation
  - Check Access version: File → Account → About Access
- Install with default settings
- Verify in ODBC Data Source Administrator (`odbcad32.exe`)

### Step 2: Create System DSN

Follow the guide in [`LINKED-DATABASE.md`](./LINKED-DATABASE.md#step-2-create-a-system-dsn-data-source-name).

**Recommended DSN configuration:**

| Field        | Value                                      |
| ------------ | ------------------------------------------ |
| Data Source  | `DigiDiggie_TNG`                           |
| Database     | `digidiggie`                               |
| Server       | `localhost` (or your PostgreSQL hostname)  |
| Port         | `5432` (default) or your custom port       |
| User Name    | Your PostgreSQL username                   |
| Password     | Your PostgreSQL password                   |
| SSL Mode     | `prefer` (or `disable` for local dev)      |

**Test the connection** before saving!

### Step 3: Create a New Access Database

1. **Launch Microsoft Access**
   - Start → Microsoft Access

2. **Create Blank Database**
   - Click "Blank Database"
   - Choose location (e.g., `C:\Projects\DigiDiggie\`)
   - Name: `digidiggie_tng_app.accdb`
   - Click "Create"

### Step 4: Link PostgreSQL Tables to Access

See [`LINKED-DATABASE.md`](./LINKED-DATABASE.md#step-3-link-tables-in-ms-access) for the full walkthrough.

#### Using External Data Wizard (Recommended)

1. **Start Link Process**
   - Ribbon: **External Data** tab
   - Click **New Data Source** → From Other Sources → **ODBC Database**
   - Select **Link to the data source by creating a linked table**
   - Click OK

2. **Select DSN**
   - Go to **Machine Data Source** tab
   - Select `DigiDiggie_TNG` (the DSN you created)
   - Click OK
   - Enter password if prompted

3. **Select Tables**
   - Select all tables from the `digidiggie_tng` schema:
     - ✅ `audit_log`
     - ✅ `community`
     - ✅ `court_case`
     - ✅ `court_case_entry`
     - ✅ `land_rights_status`
     - ✅ `legal_source`
     - ✅ `outcome_type`
     - ✅ `parish`
     - ✅ `person`
     - ✅ `person_entry`
     - ✅ `person_outcome`
     - ✅ `person_relationship`
     - ✅ `placename`
     - ✅ `relationship_type`
     - ✅ `role`
     - ✅ `role_type`
     - ✅ `ruling`
     - ✅ `ruling_type`
     - ✅ `source`
   - **Exclude** system tables like `spatial_ref_sys`, `geography_columns`, `geometry_columns`
   - Click OK

4. **Select Primary Keys**
   - Access will auto-detect most primary keys
   - For tables where prompted, select the `*_id` column as unique record identifier
     - Example: `court_case_id` for `court_case` table

5. **Verify Linked Tables**
   - Linked tables appear with 🌐 (globe) icon in Navigation Pane
   - Double-click a table to verify data displays correctly
   - Common naming: Access may add `dbo_` prefix; tables link with their PostgreSQL names

### Step 5: Import VBA Module

1. **Open VBA Editor**
   - Press `Alt + F11` or
   - Ribbon: Database Tools → **Visual Basic**

2. **Import the Module**
   - In VBA Editor: **File** → **Import File...**
   - Navigate to project folder
   - Select `src/vba_tng_app/generate_tng_app.vba`
   - Click **Open**

3. **Verify Module Loaded**
   - In Project Explorer (left pane), expand **Modules**
   - You should see **Module1** (or it might be named `generate_tng_app`)
   - Double-click to view the code

4. **Check Table Constants** (Optional)
   - Scroll to top of module
   - Verify constants match your linked table names:
     ```vba
     Public Const T_COURT_CASE As String = "court_case"
     Public Const T_PERSON As String = "person"
     ' etc...
     ```
   - If Access prefixed tables with `dbo_`, update constants:
     ```vba
     Public Const T_COURT_CASE As String = "dbo_court_case"
     ```

### Step 6: Run the Form Generator

1. **Open Immediate Window**
   - In VBA Editor: View → **Immediate Window** (or press `Ctrl + G`)

2. **Run the Main Function**
   - Type in Immediate Window:
     ```vba
     BuildForms_DigiDiggie_TNG
     ```
   - Press `Enter`

3. **Monitor Progress**
   - The script will create forms sequentially
   - Progress messages appear in Immediate Window:
     ```
     Starting form generation...
     Creating: Court Case forms
     Creating: Person forms
     Creating: Community forms
     ...
     Successfully created 25 forms.
     ```
   - **Time**: Takes 30-60 seconds depending on your system

4. **Handle Errors (if any)**
   - If errors occur, the script continues and reports which forms succeeded
   - Check error messages in the message box
   - Common issues:
     - Table name mismatches → Update constants
     - Missing permissions → Check PostgreSQL user privileges
     - Existing forms → Delete old forms first

### Step 7: Verify Generated Forms

1. **Close VBA Editor**
   - Click X or press `Alt + Q`

2. **View Forms in Navigation Pane**
   - You should see newly created forms:
     - **Main Forms**: `frm_CourtCase`, `frm_Person`, `frm_Community`, `frm_Ruling`, `frm_Role`
     - **Subforms**: `sfrm_CourtCaseEntry`, `sfrm_PersonEntry_ByPerson`, etc.
     - **Lookup Forms**: `frm_Source`, `frm_Parish`, `frm_RoleType`, etc.
     - **Picker**: `frm_PlacenamePicker`

3. **Test a Form**
   - Double-click `frm_CourtCase` to open
   - Verify:
     - ✅ Fields display correctly
     - ✅ Combo boxes load data
     - ✅ Navigation buttons work (First, Previous, Next, Last, New)
     - ✅ Subforms show related records
     - ✅ Can edit and save records

4. **Test CRUD Operations**
   - **Read**: Navigate through existing records
   - **Create**: Click "New" button, enter data, move to another record (auto-saves)
   - **Update**: Edit a field, navigate away (auto-saves)
   - **Delete**: Select record, press `Delete` key, confirm

---

## Running the VBA Form Generator

### Understanding the Script

The VBA script `generate_tng_app.vba` contains:

- **Main entry point**: `BuildForms_DigiDiggie_TNG()` - Call this to generate all forms
- **Table constants**: Define table names for easy reference
- **Helper functions**: Create controls (text boxes, combos, subforms, navigation buttons)
- **Form creation functions**: Specialized routines for each entity (Court Case, Person, etc.)

### Generated Form Features

All main forms include:

1. **Navigation Buttons** (footer)
   - |< First | < Previous | Next > | Last >| | + New

2. **Smart Combo Boxes**
   - Person combos sorted by surname, given_name, patronymic
   - Other combos sorted alphabetically
   - Linked to primary tables with key/display field pairs

3. **Required Field Indicators**
   - Red labels for mandatory fields

4. **Validation**
   - Year fields: Accepts only 1600-2100
   - Tooltips on fields

5. **Subforms**
   - Parent-child relationships (e.g., Court Case → Entries)
   - Continuous forms for easy data entry

6. **Specialized Controls**
   - Placename picker (button-launched form for large placename table)
   - Memo fields for descriptions
   - Read-only calculated fields (e.g., `full_name` for persons)

### Regenerating Forms

If you need to recreate forms (e.g., after schema changes):

1. **Delete Existing Forms** (optional but recommended)
   - Select forms in Navigation Pane
   - Press `Delete`
   - Confirm deletion

2. **Re-run Generator**
   - Open Immediate Window (`Ctrl + G`)
   - Type: `BuildForms_DigiDiggie_TNG`
   - Press `Enter`

**Note**: The script automatically deletes and recreates forms, but manual deletion ensures a clean slate.

### Customizing Generated Forms

After generation, you can customize:

1. **Layout Adjustments**
   - Open form in Design View (right-click → Design View)
   - Move/resize controls
   - Save changes

2. **Add Custom Logic**
   - Add VBA event handlers (e.g., `OnLoad`, `AfterUpdate`)
   - Create calculated fields
   - Add custom validation

3. **Visual Styling**
   - Apply themes: Design → Themes
   - Change colors, fonts
   - Add logos or headers

**Warning**: If you re-run `BuildForms_DigiDiggie_TNG()`, custom changes will be lost. Consider:
- Only run generator once for initial creation
- Make customizations afterward
- Or modify the VBA script itself for persistent changes

---

## Post-Setup Configuration

### 1. Set Startup Form (Optional)

Create a main menu/switchboard:

1. **Create Navigation Form**
   - Create → Navigation → Horizontal Tabs
   - Drag main forms to tabs
   - Save as `frm_Main`

2. **Set as Startup Form**
   - File → Options → Current Database
   - Display Form: Select `frm_Main`
   - Click OK
   - Restart database

### 2. Configure Form Properties

For better UX, set these properties on main forms:

| Property            | Recommended Value                 |
| ------------------- | --------------------------------- |
| Record Selectors    | Yes                               |
| Navigation Buttons  | No (using custom buttons)         |
| Scroll Bars         | Vertical Only                     |
| Pop Up              | No                                |
| Modal               | No                                |
| Allow Edits         | Yes                               |
| Allow Deletions     | Yes (or No for read-only)         |
| Allow Additions     | Yes                               |
| Data Entry          | No                                |

### 3. Set Up Security (Optional)

For production use:

1. **User-Level Security**
   - Implement Access workgroup security
   - Create users with specific permissions
   - Restrict design view access

2. **PostgreSQL Security**
   - Create read-only users for querying
   - Create edit users with INSERT/UPDATE/DELETE
   - Never embed passwords in DSN (use Windows Authentication if possible)

3. **Database Password**
   - File → Info → Encrypt with Password
   - Choose strong password
   - Distribute to authorized users only

### 4. Performance Optimization

1. **Indexing**
   - Ensure PostgreSQL has indexes on foreign keys
   - Create composite indexes for common queries:
     ```sql
     CREATE INDEX idx_person_names ON person(surname, given_name, patronymic);
     CREATE INDEX idx_court_case_year ON court_case(case_year);
     ```

2. **Compact & Repair**
   - Database Tools → Compact and Repair Database
   - Do this regularly (or on close)

3. **Split Database** (Advanced)
   - Create front-end (forms, queries, VBA) and back-end (linked tables)
   - Users get individual front-ends, all connect to same PostgreSQL
   - Benefits: Multi-user editing, easier updates

### 5. Backup Strategy

**PostgreSQL Database:**
```bash
# Backup
pg_dump -U postgres -d digidiggie -n digidiggie_tng -F c -f backup_tng.dump

# Restore
pg_restore -U postgres -d digidiggie -c backup_tng.dump
```

**Access Database:**
- Regular file copies of `.accdb`
- Use version control for VBA code (export modules)
- Document customizations

---

## Troubleshooting

### Issue: "Can't find table 'table_name'"

**Cause**: Table not linked or name mismatch

**Solution**:
1. Check Navigation Pane for linked tables (🌐 icon)
2. If missing, relink using External Data → ODBC
3. If table names have prefixes (e.g., `dbo_`), update VBA constants:
   ```vba
   Public Const T_COURT_CASE As String = "dbo_court_case"
   ```

### Issue: "This recordset is not updateable"

**Cause**: Missing primary key or insufficient permissions

**Solution**:
1. **Relink table** and specify unique record identifier
2. **Check PostgreSQL permissions**:
   ```sql
   GRANT SELECT, INSERT, UPDATE, DELETE ON ALL TABLES IN SCHEMA digidiggie_tng TO your_user;
   ```
3. Ensure table has primary key in PostgreSQL

### Issue: "ODBC--call failed"

**Cause**: Connection timeout or network issue

**Solution**:
1. Test DSN connection in ODBC Administrator
2. Verify PostgreSQL is running: `services.msc` → `postgresql-x64-XX`
3. Check firewall allows port 5432
4. For remote servers, verify network connectivity
5. Increase timeout in DSN settings

### Issue: VBA compile error when importing

**Cause**: References not set

**Solution**:
1. VBA Editor → Tools → References
2. Ensure checked:
   - ✅ Visual Basic For Applications
   - ✅ Microsoft Access XX.0 Object Library
   - ✅ Microsoft DAO 3.6 Object Library
3. Click OK, close and reopen

### Issue: Forms not generating

**Cause**: Various (see error message)

**Solutions**:
- **"Object already exists"**: Delete existing forms first
- **"Permission denied"**: Check PostgreSQL user permissions
- **Partial generation**: Check which forms succeeded in Immediate Window, manually debug failed ones

### Issue: Geometry column errors

**Cause**: PostGIS geometry type not supported by Access

**Solution**:
- Access cannot display `geometry` columns directly
- Use queries or views that convert to text:
  ```sql
  CREATE VIEW placename_view AS
  SELECT placename_id, name, 
         ST_X(geom) AS longitude, 
         ST_Y(geom) AS latitude,
         ST_AsText(geom) AS geom_wkt
  FROM placename;
  ```
- Link the view instead of raw table

### Issue: Slow form loading

**Cause**: Large datasets, missing indexes

**Solutions**:
1. **Add PostgreSQL indexes** on foreign keys and commonly filtered columns
2. **Limit default records**: Set form `Filter` property:
   ```vba
   Private Sub Form_Load()
       Me.Filter = "case_year >= Year(Date) - 5"
       Me.FilterOn = True
   End Sub
   ```
3. Use **search forms** instead of loading all records
4. Consider **Pass-Through Queries** for complex operations

---

## Additional Resources

- **PostgreSQL Linking**: [LINKED-DATABASE.md](./LINKED-DATABASE.md)
- **UCanAccess Setup**: [UCANACCESS.md](./UCANACCESS.md) (for Java-based Access reading)
- **Schema Documentation**: [src/schema/digidiggie_en_tng.sql](../src/schema/digidiggie_en_tng.sql)
- **VBA Source Code**: [src/vba_tng_app/generate_tng_app.vba](../src/vba_tng_app/generate_tng_app.vba)
- **Improvements Tracking**: [src/vba_tng_app/improvements.md](../src/vba_tng_app/improvements.md)

## Support

For issues or questions about this setup:
1. Check troubleshooting section above
2. Review VBA code comments for specific form logic
3. Consult PostgreSQL and MS Access documentation
4. Check project README and TODO for known issues

---

## Appendix: Automated Setup Script

For advanced users who want to automate the entire setup:

### AutomatedSetup.vba

Save this as a separate module and run `AutomatedFullSetup()`:

```vba
Option Compare Database
Option Explicit

' Automated setup script for DigiDiggie TNG Access App
Public Sub AutomatedFullSetup()
    On Error GoTo ErrorHandler
    
    Dim success As Boolean
    
    ' Step 1: Link all tables
    MsgBox "Step 1: Linking PostgreSQL tables...", vbInformation
    success = LinkAllTngTables()
    If Not success Then
        MsgBox "Failed to link tables. Setup aborted.", vbCritical
        Exit Sub
    End If
    
    ' Step 2: Import and run form generator
    MsgBox "Step 2: Generating forms...", vbInformation
    ' Assumes generate_tng_app module is already imported
    Call BuildForms_DigiDiggie_TNG
    
    MsgBox "Setup complete! Your DigiDiggie TNG app is ready.", vbInformation, "Success"
    Exit Sub
    
ErrorHandler:
    MsgBox "Setup error: " & Err.Description, vbCritical
End Sub

Public Function LinkAllTngTables() As Boolean
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim strDSN As String
    Dim arrTables As Variant
    Dim i As Integer
    Dim strTable As String
    
    Set db = CurrentDb
    
    ' Configure DSN connection string
    ' Update with your DSN name
    strDSN = "ODBC;DSN=DigiDiggie_TNG;DATABASE=digidiggie;"
    
    ' List all tables in digidiggie_tng schema
    arrTables = Array("audit_log", "community", "court_case", "court_case_entry", _
                     "land_rights_status", "legal_source", "outcome_type", "parish", _
                     "person", "person_entry", "person_outcome", "person_relationship", _
                     "placename", "relationship_type", "role", "role_type", _
                     "ruling", "ruling_type", "source")
    
    ' Link each table
    For i = LBound(arrTables) To UBound(arrTables)
        strTable = CStr(arrTables(i))
        
        ' Delete existing link if present
        On Error Resume Next
        db.TableDefs.Delete strTable
        On Error GoTo ErrorHandler
        
        ' Create new linked table
        Set tdf = db.CreateTableDef(strTable)
        tdf.Connect = strDSN
        tdf.SourceTableName = strTable
        db.TableDefs.Append tdf
        
        Debug.Print "Linked: " & strTable
    Next i
    
    Set tdf = Nothing
    Set db = Nothing
    
    LinkAllTngTables = True
    Exit Function
    
ErrorHandler:
    Debug.Print "Error linking table " & strTable & ": " & Err.Description
    LinkAllTngTables = False
End Function
```

**Usage**:
1. Import this VBA module
2. Update DSN name if different
3. Run `AutomatedFullSetup()` from Immediate Window

This automates Steps 4-6 of the setup process.

---

**Document Version**: 1.0  
**Last Updated**: February 2026  
**Author**: DigiDiggie Project Team
