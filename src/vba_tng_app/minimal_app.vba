Option Compare Database
Option Explicit

'==============================================================================
' modBuildPilotForms
' Purpose: Generate minimal court case data entry forms for DigiDiggie TNG
' Usage: Run BuildAllForms() to create all forms and queries
'==============================================================================

'------------------------------------------------------------------------------
' Main Entry Point
'------------------------------------------------------------------------------
Public Sub BuildAllForms()
On Error GoTo ErrHandler
    
    Debug.Print "Starting BuildAllForms..."
    
    ' Delete existing forms if they exist
    DeleteFormsIfExist
    
    ' Create saved queries
    CreateQueries
    
    ' Create forms in order (subforms first, then parent forms)
    Create_sfrmCourtCaseEntries
    Create_sfrmPersonEntryByEntry
    Create_sfrmPersonOutcomes
    Create_sfrmRuling
    Create_frmCourtCaseEntryDetail
    Create_frmCourtCase
    Create_frmPlacenameSearch
    Create_frmPersonSearch
    Create_frmPerson
    ' Create_frmLookups ' Optional
    
    MsgBox "All forms created successfully!" & vbCrLf & _
           "Open frmCourtCase to begin data entry.", vbInformation, "Build Complete"
    
    Debug.Print "BuildAllForms completed successfully."
    Exit Sub
    
ErrHandler:
    MsgBox "Error in BuildAllForms: " & Err.Description, vbCritical
    Debug.Print "Error in BuildAllForms: " & Err.Description
End Sub

'------------------------------------------------------------------------------
' Helper: Delete existing forms to avoid conflicts
'------------------------------------------------------------------------------
Private Sub DeleteFormsIfExist()
On Error Resume Next
    Dim formNames As Variant
    Dim i As Integer
    
    formNames = Array("frmCourtCase", "sfrmCourtCaseEntries", "frmCourtCaseEntryDetail", _
                      "sfrmPersonEntryByEntry", "sfrmRuling", "sfrmPersonOutcomes", _
                      "frmPlacenameSearch", "frmPersonSearch", "frmPerson", "frmLookups")
    
    For i = LBound(formNames) To UBound(formNames)
        DoCmd.DeleteObject acForm, formNames(i)
    Next i
    
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Create Saved Queries
'------------------------------------------------------------------------------
Private Sub CreateQueries()
On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    
    Set db = CurrentDb
    
    ' Delete existing queries
    On Error Resume Next
    db.QueryDefs.Delete "qPlacenameSearch"
    db.QueryDefs.Delete "qPersonSearch"
    On Error GoTo ErrHandler
    
    ' qPlacenameSearch
    Set qdf = db.CreateQueryDef("qPlacenameSearch")
    qdf.SQL = "PARAMETERS [pSearch] Text ( 255 ); " & _
              "SELECT TOP 200 placename_id, placename, parish_name, serial_number " & _
              "FROM placename " & _
              "WHERE (placename LIKE [pSearch]) OR (parish_name LIKE [pSearch]) " & _
              "ORDER BY placename;"
    
    ' qPersonSearch
    Set qdf = db.CreateQueryDef("qPersonSearch")
    qdf.SQL = "PARAMETERS [pSearch] Text ( 255 ); " & _
              "SELECT TOP 200 person_id, full_name, birth_year, death_year " & _
              "FROM person " & _
              "WHERE full_name LIKE [pSearch] " & _
              "ORDER BY full_name;"
    
    Debug.Print "Queries created successfully."
    Exit Sub
    
ErrHandler:
    MsgBox "Error in CreateQueries: " & Err.Description, vbCritical
    Debug.Print "Error in CreateQueries: " & Err.Description
End Sub

'------------------------------------------------------------------------------
' Create Form: frmCourtCase (Main workspace)
'------------------------------------------------------------------------------
Private Sub Create_frmCourtCase()
On Error GoTo ErrHandler
    Dim frm As Form
    Dim ctl As Control
    Dim yPos As Integer
    Dim strFormName As String
    
    Set frm = CreateForm()
    strFormName = frm.Name
    frm.RecordSource = "court_case"
    frm.Caption = "Court Case"
    frm.DefaultView = 0 ' Single Form
    frm.NavigationButtons = True
    frm.RecordSelectors = True
    frm.AllowAdditions = True
    frm.AllowEdits = True
    frm.AllowDeletions = True
    
    yPos = 200
    
    ' cboSource
    Set ctl = CreateControl(frm.Name, acComboBox, acDetail, "", "", 2000, yPos, 4000, 300)
    ctl.Name = "cboSource"
    ctl.ControlSource = "source_id"
    ctl.RowSourceType = "Table/Query"
    ctl.RowSource = "SELECT source_id, source_abbreviation, source_name FROM source ORDER BY source_abbreviation;"
    ctl.ColumnCount = 3
    ctl.ColumnWidths = "0cm;3cm;5cm"
    ctl.BoundColumn = 1
    ctl.LimitToList = True
    CreateLabel frm.Name, "lblSource", "Source:", 200, yPos, 1600, 300
    
    yPos = yPos + 500
    
    ' txtReferenceNumber
    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, "", "", 2000, yPos, 4000, 300)
    ctl.Name = "txtReferenceNumber"
    ctl.ControlSource = "reference_number"
    CreateLabel frm.Name, "lblReferenceNumber", "Reference #:", 200, yPos, 1600, 300
    
    yPos = yPos + 500
    
    ' txtCaseYear
    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, "", "", 2000, yPos, 2000, 300)
    ctl.Name = "txtCaseYear"
    ctl.ControlSource = "case_year"
    CreateLabel frm.Name, "lblCaseYear", "Case Year:", 200, yPos, 1600, 300
    
    yPos = yPos + 500
    
    ' txtDistrictCourtName
    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, "", "", 2000, yPos, 6000, 300)
    ctl.Name = "txtDistrictCourtName"
    ctl.ControlSource = "district_court_name"
    CreateLabel frm.Name, "lblDistrictCourtName", "District Court:", 200, yPos, 1600, 300
    
    yPos = yPos + 700
    
    ' cmdNewCase button
    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, "", "", 200, yPos, 2000, 400)
    ctl.Name = "cmdNewCase"
    ctl.Caption = "New Case"
    ctl.OnClick = "=frmCourtCase_cmdNewCase_Click()"
    
    ' cmdOpenEntryDetail button
    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, "", "", 2400, yPos, 2500, 400)
    ctl.Name = "cmdOpenEntryDetail"
    ctl.Caption = "Open Entry Detail"
    ctl.OnClick = "=frmCourtCase_cmdOpenEntryDetail_Click()"
    
    yPos = yPos + 600
    
    ' sfrmCourtCaseEntries subform
    Set ctl = CreateControl(frm.Name, acSubform, acDetail, "", "", 200, yPos, 10000, 2500)
    ctl.Name = "sfrmCourtCaseEntries"
    ctl.SourceObject = "Form.sfrmCourtCaseEntries"
    ctl.LinkMasterFields = "court_case_id"
    ctl.LinkChildFields = "court_case_id"
    CreateLabel frm.Name, "lblEntries", "Court Case Entries:", 200, yPos - 300, 3000, 300
    
    yPos = yPos + 2700
    
    ' sfrmRuling subform
    Set ctl = CreateControl(frm.Name, acSubform, acDetail, "", "", 200, yPos, 10000, 2000)
    ctl.Name = "sfrmRuling"
    ctl.SourceObject = "Form.sfrmRuling"
    ctl.LinkMasterFields = "court_case_id"
    ctl.LinkChildFields = "court_case_id"
    CreateLabel strFormName, "lblRuling", "Ruling:", 200, yPos - 300, 3000, 300
    
    DoCmd.Close acForm, strFormName, acSaveYes
    DoCmd.Rename "frmCourtCase", acForm, strFormName
    
    Debug.Print "frmCourtCase created."
    Exit Sub
    
ErrHandler:
    MsgBox "Error in Create_frmCourtCase: " & Err.Description, vbCritical
End Sub

'------------------------------------------------------------------------------
' Create Form: sfrmCourtCaseEntries (Subform grid)
'------------------------------------------------------------------------------
Private Sub Create_sfrmCourtCaseEntries()
On Error GoTo ErrHandler
    Dim frm As Form
    Dim ctl As Control
    Dim strFormName As String
    
    Set frm = CreateForm()
    strFormName = frm.Name
    frm.RecordSource = "court_case_entry"
    frm.Caption = "Court Case Entries"
    frm.DefaultView = 1 ' Continuous Forms (datasheet-like)
    frm.NavigationButtons = False
    frm.RecordSelectors = True
    frm.AllowAdditions = True
    frm.AllowEdits = True
    
    ' entry_year
    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, "", "", 200, 200, 1200, 300)
    ctl.Name = "txtEntryYear"
    ctl.ControlSource = "entry_year"
    CreateLabel frm.Name, "lblEntryYear", "Year", 200, 50, 1200, 300
    
    ' season_id
    Set ctl = CreateControl(frm.Name, acComboBox, acDetail, "", "", 1500, 200, 1500, 300)
    ctl.Name = "cboSeasonId"
    ctl.ControlSource = "season_id"
    ctl.RowSource = "SELECT season_id, season_name FROM season ORDER BY season_name;"
    ctl.ColumnCount = 2
    ctl.ColumnWidths = "0cm;3cm"
    ctl.BoundColumn = 1
    ctl.LimitToList = True
    CreateLabel frm.Name, "lblSeasonId", "Season", 1500, 50, 1500, 300
    
    ' land_use_id
    Set ctl = CreateControl(frm.Name, acComboBox, acDetail, "", "", 3100, 200, 2000, 300)
    ctl.Name = "cboLandUseId"
    ctl.ControlSource = "land_use_id"
    ctl.RowSource = "SELECT land_use_id, description FROM land_use ORDER BY description;"
    ctl.ColumnCount = 2
    ctl.ColumnWidths = "0cm;4cm"
    ctl.BoundColumn = 1
    ctl.LimitToList = True
    CreateLabel frm.Name, "lblLandUseId", "Land Use", 3100, 50, 2000, 300
    
    ' placename_id (numeric only)
    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, "", "", 5200, 200, 1200, 300)
    ctl.Name = "txtPlacenameId"
    ctl.ControlSource = "placename_id"
    CreateLabel frm.Name, "lblPlacenameId", "Placename ID", 5200, 50, 1200, 300
    
    ' original_placename
    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, "", "", 6500, 200, 2500, 300)
    ctl.Name = "txtOriginalPlacename"
    ctl.ControlSource = "original_placename"
    CreateLabel frm.Name, "lblOriginalPlacename", "Original Placename", 6500, 50, 2500, 300
    
    ' cmdEntryDetail button
    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, "", "", 9100, 200, 1500, 300)
    ctl.Name = "cmdEntryDetail"
    ctl.Caption = "Detail..."
    ctl.OnClick = "=sfrmCourtCaseEntries_cmdEntryDetail_Click()"
    
    DoCmd.Close acForm, strFormName, acSaveYes
    DoCmd.Rename "sfrmCourtCaseEntries", acForm, strFormName
    
    Debug.Print "sfrmCourtCaseEntries created."
    Exit Sub
    
ErrHandler:
    MsgBox "Error in Create_sfrmCourtCaseEntries: " & Err.Description, vbCritical
End Sub

'------------------------------------------------------------------------------
' Create Form: frmCourtCaseEntryDetail (Popup detail)
'------------------------------------------------------------------------------
Private Sub Create_frmCourtCaseEntryDetail()
On Error GoTo ErrHandler
    Dim frm As Form
    Dim ctl As Control
    Dim yPos As Integer
    Dim strFormName As String
    
    Set frm = CreateForm()
    strFormName = frm.Name
    frm.RecordSource = "court_case_entry"
    frm.Caption = "Court Case Entry Detail"
    frm.DefaultView = 0 ' Single Form
    frm.PopUp = True
    frm.Modal = True
    frm.NavigationButtons = True
    frm.RecordSelectors = True
    
    yPos = 200
    
    ' entry_year
    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, "", "", 2000, yPos, 2000, 300)
    ctl.Name = "txtEntryYear"
    ctl.ControlSource = "entry_year"
    CreateLabel frm.Name, "lblEntryYear", "Entry Year:", 200, yPos, 1600, 300
    
    yPos = yPos + 500
    
    ' season_id
    Set ctl = CreateControl(frm.Name, acComboBox, acDetail, "", "", 2000, yPos, 3000, 300)
    ctl.Name = "cboSeasonId"
    ctl.ControlSource = "season_id"
    ctl.RowSource = "SELECT season_id, season_name FROM season ORDER BY season_name;"
    ctl.ColumnCount = 2
    ctl.ColumnWidths = "0cm;5cm"
    ctl.BoundColumn = 1
    ctl.LimitToList = True
    CreateLabel frm.Name, "lblSeasonId", "Season:", 200, yPos, 1600, 300
    
    yPos = yPos + 500
    
    ' land_use_id
    Set ctl = CreateControl(frm.Name, acComboBox, acDetail, "", "", 2000, yPos, 4000, 300)
    ctl.Name = "cboLandUseId"
    ctl.ControlSource = "land_use_id"
    ctl.RowSource = "SELECT land_use_id, description FROM land_use ORDER BY description;"
    ctl.ColumnCount = 2
    ctl.ColumnWidths = "0cm;7cm"
    ctl.BoundColumn = 1
    ctl.LimitToList = True
    CreateLabel frm.Name, "lblLandUseId", "Land Use:", 200, yPos, 1600, 300
    
    yPos = yPos + 500
    
    ' txtPlacenameId (bound)
    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, "", "", 2000, yPos, 1500, 300)
    ctl.Name = "txtPlacenameId"
    ctl.ControlSource = "placename_id"
    CreateLabel frm.Name, "lblPlacenameId", "Placename ID:", 200, yPos, 1600, 300
    
    ' txtPlacenameDisplay (unbound)
    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, "", "", 3600, yPos, 4000, 300)
    ctl.Name = "txtPlacenameDisplay"
    ctl.Enabled = False
    ctl.Locked = True
    
    ' cmdPickPlacename
    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, "", "", 7700, yPos, 1500, 300)
    ctl.Name = "cmdPickPlacename"
    ctl.Caption = "Pick..."
    ctl.OnClick = "=frmCourtCaseEntryDetail_cmdPickPlacename_Click()"
    
    yPos = yPos + 500
    
    ' original_placename
    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, "", "", 2000, yPos, 6000, 300)
    ctl.Name = "txtOriginalPlacename"
    ctl.ControlSource = "original_placename"
    CreateLabel frm.Name, "lblOriginalPlacename", "Original Placename:", 200, yPos, 1600, 300
    
    yPos = yPos + 500
    
    ' curated_text
    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, "", "", 2000, yPos, 6000, 800)
    ctl.Name = "txtCuratedText"
    ctl.ControlSource = "curated_text"
    CreateLabel frm.Name, "lblCuratedText", "Curated Text:", 200, yPos, 1600, 300
    
    yPos = yPos + 1000
    
    ' sfrmPersonEntryByEntry subform
    Set ctl = CreateControl(frm.Name, acSubform, acDetail, "", "", 200, yPos, 9000, 2500)
    ctl.Name = "sfrmPersonEntryByEntry"
    ctl.SourceObject = "Form.sfrmPersonEntryByEntry"
    ctl.LinkMasterFields = "court_case_entry_id"
    ctl.LinkChildFields = "court_case_entry_id"
    CreateLabel strFormName, "lblPersonEntries", "People in Entry:", 200, yPos - 300, 3000, 300
    
    DoCmd.Close acForm, strFormName, acSaveYes
    DoCmd.Rename "frmCourtCaseEntryDetail", acForm, strFormName
    
    Debug.Print "frmCourtCaseEntryDetail created."
    Exit Sub
    
ErrHandler:
    MsgBox "Error in Create_frmCourtCaseEntryDetail: " & Err.Description, vbCritical
End Sub

'------------------------------------------------------------------------------
' Create Form: sfrmPersonEntryByEntry (Grid: people in entry)
'------------------------------------------------------------------------------
Private Sub Create_sfrmPersonEntryByEntry()
On Error GoTo ErrHandler
    Dim frm As Form
    Dim ctl As Control
    Dim strFormName As String
    
    Set frm = CreateForm()
    strFormName = frm.Name
    frm.RecordSource = "person_entry"
    frm.Caption = "Person Entries"
    frm.DefaultView = 1 ' Continuous Forms
    frm.NavigationButtons = False
    frm.RecordSelectors = True
    frm.AllowAdditions = True
    frm.AllowEdits = True
    
    ' person_id (numeric only, use picker)
    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, "", "", 200, 200, 1200, 300)
    ctl.Name = "txtPersonId"
    ctl.ControlSource = "person_id"
    CreateLabel frm.Name, "lblPersonId", "Person ID", 200, 50, 1200, 300
    
    ' cmdPickPerson
    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, "", "", 1500, 200, 1200, 300)
    ctl.Name = "cmdPickPerson"
    ctl.Caption = "Pick..."
    ctl.OnClick = "=sfrmPersonEntryByEntry_cmdPickPerson_Click()"
    
    ' community_id
    Set ctl = CreateControl(frm.Name, acComboBox, acDetail, "", "", 2800, 200, 2000, 300)
    ctl.Name = "cboCommunityId"
    ctl.ControlSource = "community_id"
    ctl.RowSource = "SELECT community_id, community_name FROM community ORDER BY community_name;"
    ctl.ColumnCount = 2
    ctl.ColumnWidths = "0cm;4cm"
    ctl.BoundColumn = 1
    ctl.LimitToList = True
    CreateLabel frm.Name, "lblCommunityId", "Community", 2800, 50, 2000, 300
    
    ' land_rights_status_id
    Set ctl = CreateControl(frm.Name, acComboBox, acDetail, "", "", 4900, 200, 2000, 300)
    ctl.Name = "cboLandRightsStatusId"
    ctl.ControlSource = "land_rights_status_id"
    ctl.RowSource = "SELECT land_rights_status_id, land_rights_status FROM land_rights_status ORDER BY land_rights_status;"
    ctl.ColumnCount = 2
    ctl.ColumnWidths = "0cm;4cm"
    ctl.BoundColumn = 1
    ctl.LimitToList = True
    CreateLabel frm.Name, "lblLandRightsStatusId", "Land Rights", 4900, 50, 2000, 300
    
    ' role_id
    Set ctl = CreateControl(frm.Name, acComboBox, acDetail, "", "", 7000, 200, 2000, 300)
    ctl.Name = "cboRoleId"
    ctl.ControlSource = "role_id"
    ctl.RowSource = "SELECT role_id, role_name FROM role ORDER BY role_name;"
    ctl.ColumnCount = 2
    ctl.ColumnWidths = "0cm;4cm"
    ctl.BoundColumn = 1
    ctl.LimitToList = True
    CreateLabel strFormName, "lblRoleId", "Role", 7000, 50, 2000, 300
    
    DoCmd.Close acForm, strFormName, acSaveYes
    DoCmd.Rename "sfrmPersonEntryByEntry", acForm, strFormName
    
    Debug.Print "sfrmPersonEntryByEntry created."
    Exit Sub
    
ErrHandler:
    MsgBox "Error in Create_sfrmPersonEntryByEntry: " & Err.Description, vbCritical
End Sub

'------------------------------------------------------------------------------
' Create Form: sfrmRuling (Single: 0/1 ruling per case)
'------------------------------------------------------------------------------
Private Sub Create_sfrmRuling()
On Error GoTo ErrHandler
    Dim frm As Form
    Dim ctl As Control
    Dim yPos As Integer
    Dim strFormName As String
    
    Set frm = CreateForm()
    strFormName = frm.Name
    frm.RecordSource = "ruling"
    frm.Caption = "Ruling"
    frm.DefaultView = 0 ' Single Form
    frm.NavigationButtons = False
    frm.RecordSelectors = False
    frm.AllowAdditions = True
    frm.AllowEdits = True
    
    yPos = 200
    
    ' cmdCreateRuling button
    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, "", "", 200, yPos, 2000, 400)
    ctl.Name = "cmdCreateRuling"
    ctl.Caption = "Create Ruling"
    ctl.OnClick = "=sfrmRuling_cmdCreateRuling_Click()"
    
    yPos = yPos + 600
    
    ' ruling_year
    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, "", "", 1500, yPos, 1500, 300)
    ctl.Name = "txtRulingYear"
    ctl.ControlSource = "ruling_year"
    CreateLabel frm.Name, "lblRulingYear", "Year:", 200, yPos, 1200, 300
    
    yPos = yPos + 500
    
    ' ruling_type_id
    Set ctl = CreateControl(frm.Name, acComboBox, acDetail, "", "", 1500, yPos, 3000, 300)
    ctl.Name = "cboRulingTypeId"
    ctl.ControlSource = "ruling_type_id"
    ctl.RowSource = "SELECT ruling_type_id, ruling_type FROM ruling_type ORDER BY ruling_type;"
    ctl.ColumnCount = 2
    ctl.ColumnWidths = "0cm;5cm"
    ctl.BoundColumn = 1
    ctl.LimitToList = True
    CreateLabel frm.Name, "lblRulingTypeId", "Type:", 200, yPos, 1200, 300
    
    yPos = yPos + 500
    
    ' legal_source_id
    Set ctl = CreateControl(frm.Name, acComboBox, acDetail, "", "", 1500, yPos, 3000, 300)
    ctl.Name = "cboLegalSourceId"
    ctl.ControlSource = "legal_source_id"
    ctl.RowSource = "SELECT legal_source_id, legal_source_name FROM legal_source ORDER BY legal_source_name;"
    ctl.ColumnCount = 2
    ctl.ColumnWidths = "0cm;5cm"
    ctl.BoundColumn = 1
    ctl.LimitToList = True
    CreateLabel frm.Name, "lblLegalSourceId", "Legal Source:", 200, yPos, 1200, 300
    
    yPos = yPos + 500
    
    ' description
    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, "", "", 1500, yPos, 6000, 600)
    ctl.Name = "txtDescription"
    ctl.ControlSource = "description"
    CreateLabel frm.Name, "lblDescription", "Description:", 200, yPos, 1200, 300
    
    yPos = yPos + 800
    
    ' sfrmPersonOutcomes subform (embedded in ruling)
    Set ctl = CreateControl(frm.Name, acSubform, acDetail, "", "", 200, yPos, 9000, 2000)
    ctl.Name = "sfrmPersonOutcomes"
    ctl.SourceObject = "Form.sfrmPersonOutcomes"
    ctl.LinkMasterFields = "ruling_id"
    ctl.LinkChildFields = "ruling_id"
    CreateLabel strFormName, "lblPersonOutcomes", "Person Outcomes:", 200, yPos - 300, 3000, 300
    
    DoCmd.Close acForm, strFormName, acSaveYes
    DoCmd.Rename "sfrmRuling", acForm, strFormName
    
    Debug.Print "sfrmRuling created."
    Exit Sub
    
ErrHandler:
    MsgBox "Error in Create_sfrmRuling: " & Err.Description, vbCritical
End Sub

'------------------------------------------------------------------------------
' Create Form: sfrmPersonOutcomes (Grid outcomes for ruling)
'------------------------------------------------------------------------------
Private Sub Create_sfrmPersonOutcomes()
On Error GoTo ErrHandler
    Dim frm As Form
    Dim ctl As Control
    Dim strFormName As String
    
    Set frm = CreateForm()
    strFormName = frm.Name
    frm.RecordSource = "person_outcome"
    frm.Caption = "Person Outcomes"
    frm.DefaultView = 1 ' Continuous Forms
    frm.NavigationButtons = False
    frm.RecordSelectors = True
    frm.AllowAdditions = True
    frm.AllowEdits = True
    
    ' person_id
    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, "", "", 200, 200, 1200, 300)
    ctl.Name = "txtPersonId"
    ctl.ControlSource = "person_id"
    CreateLabel frm.Name, "lblPersonId", "Person ID", 200, 50, 1200, 300
    
    ' cmdPickPersonOutcomePerson
    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, "", "", 1500, 200, 1200, 300)
    ctl.Name = "cmdPickPersonOutcomePerson"
    ctl.Caption = "Pick..."
    ctl.OnClick = "=sfrmPersonOutcomes_cmdPickPerson_Click()"
    
    ' outcome_type_id
    Set ctl = CreateControl(frm.Name, acComboBox, acDetail, "", "", 2800, 200, 2500, 300)
    ctl.Name = "cboOutcomeTypeId"
    ctl.ControlSource = "outcome_type_id"
    ctl.RowSource = "SELECT outcome_type_id, outcome_type_name FROM outcome_type ORDER BY outcome_type_name;"
    ctl.ColumnCount = 2
    ctl.ColumnWidths = "0cm;5cm"
    ctl.BoundColumn = 1
    ctl.LimitToList = True
    CreateLabel frm.Name, "lblOutcomeTypeId", "Outcome Type", 2800, 50, 2500, 300
    
    ' description
    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, "", "", 5400, 200, 4000, 300)
    ctl.Name = "txtDescription"
    ctl.ControlSource = "description"
    CreateLabel strFormName, "lblDescription", "Description", 5400, 50, 4000, 300
    
    DoCmd.Close acForm, strFormName, acSaveYes
    DoCmd.Rename "sfrmPersonOutcomes", acForm, strFormName
    
    Debug.Print "sfrmPersonOutcomes created."
    Exit Sub
    
ErrHandler:
    MsgBox "Error in Create_sfrmPersonOutcomes: " & Err.Description, vbCritical
End Sub

'------------------------------------------------------------------------------
' Create Form: frmPlacenameSearch (Popup picker)
'------------------------------------------------------------------------------
Private Sub Create_frmPlacenameSearch()
On Error GoTo ErrHandler
    Dim frm As Form
    Dim ctl As Control
    Dim yPos As Integer
    Dim strFormName As String
    
    Set frm = CreateForm()
    strFormName = frm.Name
    frm.RecordSource = ""
    frm.Caption = "Search Placename"
    frm.DefaultView = 0 ' Single Form
    frm.PopUp = True
    frm.Modal = True
    frm.NavigationButtons = False
    frm.RecordSelectors = False
    
    yPos = 200
    
    ' txtSearch
    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, "", "", 1500, yPos, 4000, 300)
    ctl.Name = "txtSearch"
    CreateLabel frm.Name, "lblSearch", "Search:", 200, yPos, 1200, 300
    
    ' cmdSearch
    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, "", "", 5600, yPos, 1500, 300)
    ctl.Name = "cmdSearch"
    ctl.Caption = "Search"
    ctl.OnClick = "=frmPlacenameSearch_cmdSearch_Click()"
    
    yPos = yPos + 500
    
    ' lstResults
    Set ctl = CreateControl(frm.Name, acListBox, acDetail, "", "", 200, yPos, 8000, 4000)
    ctl.Name = "lstResults"
    ctl.RowSourceType = "Table/Query"
    ctl.RowSource = ""
    ctl.ColumnCount = 4
    ctl.ColumnWidths = "1.5cm;4cm;4cm;2cm"
    ctl.ColumnHeads = True
    
    yPos = yPos + 4200
    
    ' cmdSelect
    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, "", "", 200, yPos, 2000, 400)
    ctl.Name = "cmdSelect"
    ctl.Caption = "Select"
    ctl.OnClick = "=frmPlacenameSearch_cmdSelect_Click()"
    
    ' cmdCancel
    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, "", "", 2400, yPos, 2000, 400)
    ctl.Name = "cmdCancel"
    ctl.Caption = "Cancel"
    ctl.OnClick = "=frmPlacenameSearch_cmdCancel_Click()"
    
    DoCmd.Close acForm, strFormName, acSaveYes
    DoCmd.Rename "frmPlacenameSearch", acForm, strFormName
    
    Debug.Print "frmPlacenameSearch created."
    Exit Sub
    
ErrHandler:
    MsgBox "Error in Create_frmPlacenameSearch: " & Err.Description, vbCritical
End Sub

'------------------------------------------------------------------------------
' Create Form: frmPersonSearch (Popup picker)
'------------------------------------------------------------------------------
Private Sub Create_frmPersonSearch()
On Error GoTo ErrHandler
    Dim frm As Form
    Dim ctl As Control
    Dim yPos As Integer
    Dim strFormName As String
    
    Set frm = CreateForm()
    strFormName = frm.Name
    frm.RecordSource = ""
    frm.Caption = "Search Person"
    frm.DefaultView = 0 ' Single Form
    frm.PopUp = True
    frm.Modal = True
    frm.NavigationButtons = False
    frm.RecordSelectors = False
    
    yPos = 200
    
    ' txtSearch
    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, "", "", 1500, yPos, 4000, 300)
    ctl.Name = "txtSearch"
    CreateLabel frm.Name, "lblSearch", "Search:", 200, yPos, 1200, 300
    
    ' cmdSearch
    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, "", "", 5600, yPos, 1500, 300)
    ctl.Name = "cmdSearch"
    ctl.Caption = "Search"
    ctl.OnClick = "=frmPersonSearch_cmdSearch_Click()"
    
    yPos = yPos + 500
    
    ' lstResults
    Set ctl = CreateControl(frm.Name, acListBox, acDetail, "", "", 200, yPos, 8000, 4000)
    ctl.Name = "lstResults"
    ctl.RowSourceType = "Table/Query"
    ctl.RowSource = ""
    ctl.ColumnCount = 4
    ctl.ColumnWidths = "1.5cm;5cm;2cm;2cm"
    ctl.ColumnHeads = True
    
    yPos = yPos + 4200
    
    ' cmdSelect
    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, "", "", 200, yPos, 2000, 400)
    ctl.Name = "cmdSelect"
    ctl.Caption = "Select"
    ctl.OnClick = "=frmPersonSearch_cmdSelect_Click()"
    
    ' cmdNewPerson
    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, "", "", 2400, yPos, 2000, 400)
    ctl.Name = "cmdNewPerson"
    ctl.Caption = "New Person"
    ctl.OnClick = "=frmPersonSearch_cmdNewPerson_Click()"
    
    ' cmdCancel
    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, "", "", 4600, yPos, 2000, 400)
    ctl.Name = "cmdCancel"
    ctl.Caption = "Cancel"
    ctl.OnClick = "=frmPersonSearch_cmdCancel_Click()"
    
    DoCmd.Close acForm, strFormName, acSaveYes
    DoCmd.Rename "frmPersonSearch", acForm, strFormName
    
    Debug.Print "frmPersonSearch created."
    Exit Sub
    
ErrHandler:
    MsgBox "Error in Create_frmPersonSearch: " & Err.Description, vbCritical
End Sub

'------------------------------------------------------------------------------
' Create Form: frmPerson (Basic person editor)
'------------------------------------------------------------------------------
Private Sub Create_frmPerson()
On Error GoTo ErrHandler
    Dim frm As Form
    Dim ctl As Control
    Dim yPos As Integer
    Dim strFormName As String
    
    Set frm = CreateForm()
    strFormName = frm.Name
    frm.RecordSource = "person"
    frm.Caption = "Person"
    frm.DefaultView = 0 ' Single Form
    frm.NavigationButtons = True
    frm.RecordSelectors = True
    frm.AllowAdditions = True
    frm.AllowEdits = True
    
    yPos = 200
    
    ' given_name
    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, "", "", 2000, yPos, 4000, 300)
    ctl.Name = "txtGivenName"
    ctl.ControlSource = "given_name"
    CreateLabel frm.Name, "lblGivenName", "Given Name:", 200, yPos, 1600, 300
    
    yPos = yPos + 500
    
    ' patronymic
    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, "", "", 2000, yPos, 4000, 300)
    ctl.Name = "txtPatronymic"
    ctl.ControlSource = "patronymic"
    CreateLabel frm.Name, "lblPatronymic", "Patronymic:", 200, yPos, 1600, 300
    
    yPos = yPos + 500
    
    ' surname
    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, "", "", 2000, yPos, 4000, 300)
    ctl.Name = "txtSurname"
    ctl.ControlSource = "surname"
    CreateLabel frm.Name, "lblSurname", "Surname:", 200, yPos, 1600, 300
    
    yPos = yPos + 500
    
    ' birth_year
    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, "", "", 2000, yPos, 1500, 300)
    ctl.Name = "txtBirthYear"
    ctl.ControlSource = "birth_year"
    CreateLabel frm.Name, "lblBirthYear", "Birth Year:", 200, yPos, 1600, 300
    
    yPos = yPos + 500
    
    ' death_year
    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, "", "", 2000, yPos, 1500, 300)
    ctl.Name = "txtDeathYear"
    ctl.ControlSource = "death_year"
    CreateLabel frm.Name, "lblDeathYear", "Death Year:", 200, yPos, 1600, 300
    
    yPos = yPos + 500
    
    ' community_name
    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, "", "", 2000, yPos, 4000, 300)
    ctl.Name = "txtCommunityName"
    ctl.ControlSource = "community_name"
    CreateLabel frm.Name, "lblCommunityName", "Community:", 200, yPos, 1600, 300
    
    yPos = yPos + 500
    
    ' note
    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, "", "", 2000, yPos, 6000, 800)
    ctl.Name = "txtNote"
    ctl.ControlSource = "note"
    CreateLabel strFormName, "lblNote", "Note:", 200, yPos, 1600, 300
    
    DoCmd.Close acForm, strFormName, acSaveYes
    DoCmd.Rename "frmPerson", acForm, strFormName
    
    Debug.Print "frmPerson created."
    Exit Sub
    
ErrHandler:
    MsgBox "Error in Create_frmPerson: " & Err.Description, vbCritical
End Sub

'------------------------------------------------------------------------------
' Helper: Create a label control
'------------------------------------------------------------------------------
Private Function CreateLabel(formName As String, labelName As String, _
                            caption As String, left As Integer, top As Integer, _
                            width As Integer, height As Integer) As Control
    Dim ctl As Control
    Set ctl = CreateControl(formName, acLabel, acDetail, "", "", left, top, width, height)
    ctl.Name = labelName
    ctl.Caption = caption
    Set CreateLabel = ctl
End Function

'==============================================================================
' EVENT HANDLERS - These are called by form button OnClick events
'==============================================================================

'------------------------------------------------------------------------------
' frmCourtCase: New Case button
'------------------------------------------------------------------------------
Public Function frmCourtCase_cmdNewCase_Click()
On Error GoTo ErrHandler
    DoCmd.GoToRecord , , acNewRec
    Exit Function
ErrHandler:
    MsgBox "Error in frmCourtCase_cmdNewCase_Click: " & Err.Description, vbCritical
End Function

'------------------------------------------------------------------------------
' frmCourtCase: Open Entry Detail button
'------------------------------------------------------------------------------
Public Function frmCourtCase_cmdOpenEntryDetail_Click()
On Error GoTo ErrHandler
    Dim entryId As Variant
    
    ' Get selected entry from subform
    If Not IsNull(Forms!frmCourtCase!sfrmCourtCaseEntries.Form!court_case_entry_id) Then
        entryId = Forms!frmCourtCase!sfrmCourtCaseEntries.Form!court_case_entry_id
        DoCmd.OpenForm "frmCourtCaseEntryDetail", , , "court_case_entry_id=" & entryId, , acDialog
    Else
        MsgBox "Please select an entry first.", vbInformation
    End If
    
    Exit Function
ErrHandler:
    MsgBox "Error in frmCourtCase_cmdOpenEntryDetail_Click: " & Err.Description, vbCritical
End Function

'------------------------------------------------------------------------------
' sfrmCourtCaseEntries: Entry Detail button
'------------------------------------------------------------------------------
Public Function sfrmCourtCaseEntries_cmdEntryDetail_Click()
On Error GoTo ErrHandler
    Dim entryId As Variant
    
    entryId = Screen.ActiveControl.Parent!court_case_entry_id
    If Not IsNull(entryId) Then
        DoCmd.OpenForm "frmCourtCaseEntryDetail", , , "court_case_entry_id=" & entryId, , acDialog
    End If
    
    Exit Function
ErrHandler:
    MsgBox "Error in sfrmCourtCaseEntries_cmdEntryDetail_Click: " & Err.Description, vbCritical
End Function

'------------------------------------------------------------------------------
' frmCourtCaseEntryDetail: Pick Placename button
'------------------------------------------------------------------------------
Public Function frmCourtCaseEntryDetail_cmdPickPlacename_Click()
On Error GoTo ErrHandler
    DoCmd.OpenForm "frmPlacenameSearch", , , , , acDialog, _
        "caller=frmCourtCaseEntryDetail;target=txtPlacenameId"
    
    ' Update display field after picker closes
    UpdatePlacenameDisplay "frmCourtCaseEntryDetail"
    
    Exit Function
ErrHandler:
    MsgBox "Error in frmCourtCaseEntryDetail_cmdPickPlacename_Click: " & Err.Description, vbCritical
End Function

'------------------------------------------------------------------------------
' sfrmPersonEntryByEntry: Pick Person button
'------------------------------------------------------------------------------
Public Function sfrmPersonEntryByEntry_cmdPickPerson_Click()
On Error GoTo ErrHandler
    DoCmd.OpenForm "frmPersonSearch", , , , , acDialog, _
        "caller=sfrmPersonEntryByEntry;target=txtPersonId"
    Exit Function
ErrHandler:
    MsgBox "Error in sfrmPersonEntryByEntry_cmdPickPerson_Click: " & Err.Description, vbCritical
End Function

'------------------------------------------------------------------------------
' sfrmPersonOutcomes: Pick Person button
'------------------------------------------------------------------------------
Public Function sfrmPersonOutcomes_cmdPickPerson_Click()
On Error GoTo ErrHandler
    DoCmd.OpenForm "frmPersonSearch", , , , , acDialog, _
        "caller=sfrmPersonOutcomes;target=txtPersonId"
    Exit Function
ErrHandler:
    MsgBox "Error in sfrmPersonOutcomes_cmdPickPerson_Click: " & Err.Description, vbCritical
End Function

'------------------------------------------------------------------------------
' sfrmRuling: Create Ruling button
'------------------------------------------------------------------------------
Public Function sfrmRuling_cmdCreateRuling_Click()
On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim courtCaseId As Variant
    Dim defaultRulingType As Variant
    
    Set db = CurrentDb
    
    ' Get parent court_case_id
    courtCaseId = Forms!frmCourtCase!court_case_id
    If IsNull(courtCaseId) Then
        MsgBox "Please save the court case first.", vbInformation
        Exit Function
    End If
    
    ' Check if ruling already exists
    Set rs = db.OpenRecordset("SELECT ruling_id FROM ruling WHERE court_case_id = " & courtCaseId)
    If Not rs.EOF Then
        MsgBox "A ruling already exists for this case.", vbInformation
        rs.Close
        Exit Function
    End If
    rs.Close
    
    ' Get default ruling_type_id (first one)
    defaultRulingType = DMin("ruling_type_id", "ruling_type")
    If IsNull(defaultRulingType) Then
        MsgBox "No ruling types defined. Please add at least one ruling type.", vbExclamation
        Exit Function
    End If
    
    ' Insert new ruling
    Set rs = db.OpenRecordset("ruling", dbOpenDynaset)
    rs.AddNew
    rs!court_case_id = courtCaseId
    rs!ruling_type_id = defaultRulingType
    rs.Update
    rs.Close
    
    ' Requery the ruling form
    Forms!frmCourtCase!sfrmRuling.Form.Requery
    
    MsgBox "Ruling created.", vbInformation
    
    Exit Function
ErrHandler:
    MsgBox "Error in sfrmRuling_cmdCreateRuling_Click: " & Err.Description, vbCritical
End Function

'------------------------------------------------------------------------------
' frmPlacenameSearch: Search button
'------------------------------------------------------------------------------
Public Function frmPlacenameSearch_cmdSearch_Click()
On Error GoTo ErrHandler
    Dim searchTerm As String
    Dim qdf As DAO.QueryDef
    Dim db As DAO.Database
    
    Set db = CurrentDb
    Set qdf = db.QueryDefs("qPlacenameSearch")
    
    searchTerm = "*" & Nz(Forms!frmPlacenameSearch!txtSearch, "") & "*"
    qdf.Parameters("pSearch") = searchTerm
    
    Forms!frmPlacenameSearch!lstResults.RowSource = "qPlacenameSearch"
    Forms!frmPlacenameSearch!lstResults.Requery
    
    Exit Function
ErrHandler:
    MsgBox "Error in frmPlacenameSearch_cmdSearch_Click: " & Err.Description, vbCritical
End Function

'------------------------------------------------------------------------------
' frmPlacenameSearch: Select button
'------------------------------------------------------------------------------
Public Function frmPlacenameSearch_cmdSelect_Click()
On Error GoTo ErrHandler
    Dim selectedId As Variant
    Dim callerForm As String
    Dim targetControl As String
    
    ' Get selected placename_id
    If IsNull(Forms!frmPlacenameSearch!lstResults) Then
        MsgBox "Please select a placename.", vbInformation
        Exit Function
    End If
    
    selectedId = Forms!frmPlacenameSearch!lstResults
    
    ' Parse OpenArgs
    ParseOpenArgs Forms!frmPlacenameSearch.OpenArgs, callerForm, targetControl
    
    ' Write back to caller
    If callerForm <> "" And targetControl <> "" Then
        Forms(callerForm).Controls(targetControl).Value = selectedId
    End If
    
    DoCmd.Close acForm, "frmPlacenameSearch"
    
    Exit Function
ErrHandler:
    MsgBox "Error in frmPlacenameSearch_cmdSelect_Click: " & Err.Description, vbCritical
End Function

'------------------------------------------------------------------------------
' frmPlacenameSearch: Cancel button
'------------------------------------------------------------------------------
Public Function frmPlacenameSearch_cmdCancel_Click()
    DoCmd.Close acForm, "frmPlacenameSearch"
End Function

'------------------------------------------------------------------------------
' frmPersonSearch: Search button
'------------------------------------------------------------------------------
Public Function frmPersonSearch_cmdSearch_Click()
On Error GoTo ErrHandler
    Dim searchTerm As String
    Dim qdf As DAO.QueryDef
    Dim db As DAO.Database
    
    Set db = CurrentDb
    Set qdf = db.QueryDefs("qPersonSearch")
    
    searchTerm = "*" & Nz(Forms!frmPersonSearch!txtSearch, "") & "*"
    qdf.Parameters("pSearch") = searchTerm
    
    Forms!frmPersonSearch!lstResults.RowSource = "qPersonSearch"
    Forms!frmPersonSearch!lstResults.Requery
    
    Exit Function
ErrHandler:
    MsgBox "Error in frmPersonSearch_cmdSearch_Click: " & Err.Description, vbCritical
End Function

'------------------------------------------------------------------------------
' frmPersonSearch: Select button
'------------------------------------------------------------------------------
Public Function frmPersonSearch_cmdSelect_Click()
On Error GoTo ErrHandler
    Dim selectedId As Variant
    Dim callerForm As String
    Dim targetControl As String
    
    ' Get selected person_id
    If IsNull(Forms!frmPersonSearch!lstResults) Then
        MsgBox "Please select a person.", vbInformation
        Exit Function
    End If
    
    selectedId = Forms!frmPersonSearch!lstResults
    
    ' Parse OpenArgs
    ParseOpenArgs Forms!frmPersonSearch.OpenArgs, callerForm, targetControl
    
    ' Write back to caller
    If callerForm <> "" And targetControl <> "" Then
        Forms(callerForm).Controls(targetControl).Value = selectedId
    End If
    
    DoCmd.Close acForm, "frmPersonSearch"
    
    Exit Function
ErrHandler:
    MsgBox "Error in frmPersonSearch_cmdSelect_Click: " & Err.Description, vbCritical
End Function

'------------------------------------------------------------------------------
' frmPersonSearch: New Person button
'------------------------------------------------------------------------------
Public Function frmPersonSearch_cmdNewPerson_Click()
On Error GoTo ErrHandler
    ' Open person form in data entry mode
    DoCmd.OpenForm "frmPerson", , , , acFormAdd, acDialog
    
    ' Requery results after closing
    Forms!frmPersonSearch!lstResults.Requery
    
    Exit Function
ErrHandler:
    MsgBox "Error in frmPersonSearch_cmdNewPerson_Click: " & Err.Description, vbCritical
End Function

'------------------------------------------------------------------------------
' frmPersonSearch: Cancel button
'------------------------------------------------------------------------------
Public Function frmPersonSearch_cmdCancel_Click()
    DoCmd.Close acForm, "frmPersonSearch"
End Function

'==============================================================================
' HELPER FUNCTIONS
'==============================================================================

'------------------------------------------------------------------------------
' Parse OpenArgs: "caller=formName;target=controlName"
'------------------------------------------------------------------------------
Private Sub ParseOpenArgs(openArgs As Variant, callerForm As String, targetControl As String)
On Error GoTo ErrHandler
    Dim parts() As String
    Dim i As Integer
    
    callerForm = ""
    targetControl = ""
    
    If IsNull(openArgs) Or openArgs = "" Then Exit Sub
    
    parts = Split(openArgs, ";")
    For i = LBound(parts) To UBound(parts)
        If InStr(parts(i), "caller=") > 0 Then
            callerForm = Replace(parts(i), "caller=", "")
        ElseIf InStr(parts(i), "target=") > 0 Then
            targetControl = Replace(parts(i), "target=", "")
        End If
    Next i
    
    Exit Sub
ErrHandler:
    Debug.Print "Error in ParseOpenArgs: " & Err.Description
End Sub

'------------------------------------------------------------------------------
' Update placename display field after selection
'------------------------------------------------------------------------------
Private Sub UpdatePlacenameDisplay(formName As String)
On Error Resume Next
    Dim placenameId As Variant
    Dim displayText As String
    
    placenameId = Forms(formName)!txtPlacenameId
    If Not IsNull(placenameId) Then
        displayText = Nz(DLookup("placename", "placename", "placename_id=" & placenameId), "")
        Forms(formName)!txtPlacenameDisplay = displayText
    End If
End Sub
