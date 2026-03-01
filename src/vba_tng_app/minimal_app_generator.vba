Option Compare Database
Option Explicit

'==============================================================================
' modMinimalAppGenerator
' Purpose: Generate minimal court Case data entry forms For DigiDiggie TNG
' Usage: Run BuildAllForms() To create all forms And queries
' Note: This module contains ONLY form generation code.
'       Runtime event handlers are in modMinimalAppRuntime module.
'==============================================================================

'------------------------------------------------------------------------------
' Main Entry Point
'------------------------------------------------------------------------------
Public Sub BuildAllForms()
    On Error GoTo ErrHandler

        Debug.Print "Starting BuildAllForms..."

        ' Delete existing forms If they exist
        DeleteFormsIfExist

        ' Create saved queries
        CreateQueries

        ' Create forms in order (subforms first, Then parent forms)
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
        "Open frmCourtCase To begin data entry.", vbInformation, "Build Complete"

        Debug.Print "BuildAllForms completed successfully."
     Exit Sub

ErrHandler:
        MsgBox "Error in BuildAllForms: " & Err.Description, vbCritical
        Debug.Print "Error in BuildAllForms: " & Err.Description
End Sub

'------------------------------------------------------------------------------
' Public Method: Delete ALL forms in the database
'------------------------------------------------------------------------------
Public Sub DeleteAllForms()
    On Error Resume Next
    Dim db As DAO.Database
    Dim obj As AccessObject
    Dim formName As String
    Dim deletedCount As Integer
    Dim errorCount As Integer
    
    Set db = CurrentDb
    deletedCount = 0
    errorCount = 0
    
    Debug.Print "Starting DeleteAllForms..."
    
    ' Loop through all forms in reverse order (to avoid index shifting issues)
    Dim i As Integer
    For i = CurrentProject.AllForms.Count - 1 To 0 Step -1
        formName = CurrentProject.AllForms(i).Name

        If Not (formName Like "X*") Then
            ' Try to delete the form
            On Error Resume Next
            DoCmd.Close acForm, formName, acSaveNo
            DoCmd.DeleteObject acForm, formName
            
            If Err.Number = 0 Then
                deletedCount = deletedCount + 1
                Debug.Print "Deleted form: " & formName
            Else
                errorCount = errorCount + 1
                Debug.Print "Error deleting form: " & formName & " (" & Err.Description & ")"
                Err.Clear
            End If
            On Error GoTo 0
        Else
            Debug.Print "Skipping form: " & formName
        End If
    Next i
    
    Debug.Print "DeleteAllForms completed. Deleted: " & deletedCount & ", Errors: " & errorCount
    
    If errorCount > 0 Then
        Debug.Print "Some forms could not be deleted. Please check the error log above."
        MsgBox "Deleted " & deletedCount & " form(s)." & vbCrLf & _
            IIf(errorCount > 0, "Errors: " & errorCount & " form(s) could not be deleted.", ""), _
            vbInformation, "Delete All Forms"
    End If
End Sub

'------------------------------------------------------------------------------
' Helper: Delete existing forms To avoid conflicts
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
            qdf.sql = "PARAMETERS [pSearch] Text ( 255 ); " & _
            "Select TOP 200 placename_id, placename, parish_name, serial_number " & _
            "FROM placename " & _
            "WHERE (placename LIKE [pSearch]) Or (parish_name LIKE [pSearch]) " & _
            "ORDER BY placename;"

        ' qPersonSearch (enhanced with community context)
        Set qdf = db.CreateQueryDef("qPersonSearch")
        qdf.sql = "PARAMETERS [pSearch] Text ( 255 ); " & _
            "Select TOP 200 person_id, full_name, birth_year, community_name " & _
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
        frm.caption = "Court Case"
        frm.DefaultView = 0 ' Single Form
        frm.NavigationButtons = True
        frm.RecordSelectors = True
        frm.AllowAdditions = True
        frm.AllowEdits = True
        frm.AllowDeletions = True
        frm.AutoResize = False ' Allow manual resizing
        frm.AutoCenter = True
        frm.OnCurrent = "=frmCourtCase_OnCurrent()" ' Update button visibility

        yPos = 200

        ' cboSource
        Set ctl = CreateControl(frm.Name, acComboBox, acDetail, "", "", 2000, yPos, 4000, 300)
        ctl.Name = "cboSource"
        ctl.ControlSource = "source_id"
        ctl.RowSourceType = "Table/Query"
        ctl.RowSource = "Select source_id, source_abbreviation, source_name FROM source ORDER BY source_abbreviation;"
        ctl.ColumnCount = 3
        ctl.ColumnWidths = "0cm;3cm;5cm"
        ctl.BoundColumn = 1
        ctl.LimitToList = True
        ctl.HorizontalAnchor = acHorizontalAnchorBoth ' Stretch horizontally
        CreateLabel frm.Name, "lblSource", "Source:", 200, yPos, 1600, 300

        yPos = yPos + 500

        ' txtReferenceNumber
        Set ctl = CreateControl(frm.Name, acTextBox, acDetail, "", "", 2000, yPos, 4000, 300)
        ctl.Name = "txtReferenceNumber"
        ctl.ControlSource = "reference_number"
        ctl.HorizontalAnchor = acHorizontalAnchorBoth ' Stretch horizontally
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
        ctl.HorizontalAnchor = acHorizontalAnchorBoth ' Stretch horizontally
        CreateLabel frm.Name, "lblDistrictCourtName", "District Court:", 200, yPos, 1600, 300

        yPos = yPos + 700

        ' cmdNewCase button
        Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, "", "", 200, yPos, 2000, 400)
        ctl.Name = "cmdNewCase"
        ctl.caption = "New Case"
        ctl.OnClick = "=frmCourtCase_cmdNewCase_Click()"

        ' cmdOpenEntryDetail button
        Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, "", "", 2400, yPos, 2500, 400)
        ctl.Name = "cmdOpenEntryDetail"
        ctl.caption = "Open Entry Detail"
        ctl.OnClick = "=frmCourtCase_cmdOpenEntryDetail_Click()"

        ' cmdCreateRuling button (visible only when no ruling exists)
        Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, "", "", 5100, yPos, 2000, 400)
        ctl.Name = "cmdCreateRuling"
        ctl.caption = "Create Ruling"
        ctl.OnClick = "=frmCourtCase_cmdCreateRuling_Click()"
        ctl.Visible = True ' Default To visible (no ruling initially)

        ' cmdPrevious button
        Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, "", "", 7300, yPos, 1200, 400)
        ctl.Name = "cmdPrevious"
        ctl.caption = "< Previous"
        ctl.OnClick = "=frmCourtCase_cmdPrevious_Click()"

        ' cmdNext button
        Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, "", "", 8600, yPos, 1200, 400)
        ctl.Name = "cmdNext"
        ctl.caption = "Next >"
        ctl.OnClick = "=frmCourtCase_cmdNext_Click()"

        yPos = yPos + 800
        Debug.Print "Creating tab control at yPos: " & yPos

        ' Tab Control
        Set ctl = CreateControl(frm.Name, acTabCtl, acDetail, "", "", 200, yPos, 10000, 3500)
        ctl.Name = "tabMain"
        ctl.HorizontalAnchor = acHorizontalAnchorBoth ' Stretch horizontally
        ctl.VerticalAnchor = acVerticalAnchorBoth ' Stretch vertically
        Debug.Print "Tab control created. Top position: " & ctl.Top
        
        ' Set tab page names and captions
        ctl.Pages(0).Name = "pageEntries"
        ctl.Pages(0).Caption = "Court Case Entries"
        ctl.Pages(1).Name = "pageRuling"
        ctl.Pages(1).Caption = "Ruling"

        ' sfrmCourtCaseEntries subform (on first tab page)
        Set ctl = CreateControl(frm.Name, acSubform, acDetail, "pageEntries", "", 300, 500, 9600, 2800)
        ctl.Name = "sfrmCourtCaseEntries"
        ctl.SourceObject = "Form.sfrmCourtCaseEntries"
        ctl.LinkMasterFields = "court_case_id"
        ctl.LinkChildFields = "court_case_id"
        ctl.HorizontalAnchor = acHorizontalAnchorBoth ' Stretch horizontally
        ctl.VerticalAnchor = acVerticalAnchorBoth ' Stretch vertically

        ' sfrmRuling subform (on second tab page)
        Set ctl = CreateControl(frm.Name, acSubform, acDetail, "pageRuling", "", 300, 500, 9600, 2800)
        ctl.Name = "sfrmRuling"
        ctl.SourceObject = "Form.sfrmRuling"
        ctl.LinkMasterFields = "court_case_id"
        ctl.LinkChildFields = "court_case_id"
        ctl.HorizontalAnchor = acHorizontalAnchorBoth ' Stretch horizontally
        ctl.VerticalAnchor = acVerticalAnchorBoth ' Stretch vertically

        ' Set Detail section height to accommodate all controls
        frm.Section(acDetail).Height = yPos + 3700  ' Tab control height + margin
        
        ' CRITICAL: Reset tab control position after creating subforms
        ' (Access may reposition it when subforms are added to pages)
        frm.Controls("tabMain").Top = yPos
        frm.Controls("tabMain").Left = 200
        
        Debug.Print "Tab control position corrected to: " & frm.Controls("tabMain").Top

        DoCmd.Close acForm, strFormName, acSaveYes
        DoCmd.Rename "frmCourtCase", acForm, strFormName

        Debug.Print "frmCourtCase created and saved."
     Exit Sub

ErrHandler:
        MsgBox "Error in Create_frmCourtCase: " & Err.Description, vbCritical
        Debug.Print "ERROR in Create_frmCourtCase: " & Err.Description
End Sub

'------------------------------------------------------------------------------
' Create Form: sfrmCourtCaseEntries (Subform grid)
'------------------------------------------------------------------------------
Private Sub Create_sfrmCourtCaseEntries()
    On Error GoTo ErrHandler
        Dim frm As Form
        Dim ctl As Control
        Dim strFormName As String
        Dim xPos As Integer

        Set frm = CreateForm()
        strFormName = frm.Name
        ' Join With placename table To display placename text
        frm.RecordSource = "Select cce.court_case_entry_id, cce.court_case_id, cce.entry_year, " & _
            "cce.curated_text, cce.original_placename, cce.season_id, cce.land_use_id, " & _
            "cce.placename_id, p.placename " & _
            "FROM court_case_entry As cce " & _
            "LEFT JOIN placename As p ON cce.placename_id = p.placename_id"
        frm.caption = "Court Case Entries"
        frm.DefaultView = 2 ' Datasheet (grid view)
        frm.NavigationButtons = False
        frm.RecordSelectors = True
        frm.AllowAdditions = True
        frm.AllowEdits = True
        frm.OnLoad = "=sfrmCourtCaseEntries_OnLoad()" ' Auto-size columns

        xPos = 0 ' Start position for columns

        ' entry_year
        Set ctl = CreateControl(frm.Name, acTextBox, acDetail, "", "", xPos, 0, 800, 300)
        ctl.Name = "txtEntryYear"
        ctl.ControlSource = "entry_year"
        SetDatasheetCaptionSafe ctl, "Year"
        CreateLabel frm.Name, "lblEntryYear", "Year", xPos, 0, 800, 300
        xPos = xPos + 800

        ' season_id
        Set ctl = CreateControl(frm.Name, acComboBox, acDetail, "", "", xPos, 0, 1200, 300)
        ctl.Name = "cboSeasonId"
        ctl.ControlSource = "season_id"
        SetDatasheetCaptionSafe ctl, "Season"
        ctl.RowSource = "Select season_id, season_name FROM season ORDER BY season_name;"
        ctl.ColumnCount = 2
        ctl.ColumnWidths = "0cm;3cm"
        ctl.BoundColumn = 1
        ctl.LimitToList = True
        CreateLabel frm.Name, "lblSeasonId", "Season", xPos, 0, 1200, 300
        xPos = xPos + 1200

        ' land_use_id
        Set ctl = CreateControl(frm.Name, acComboBox, acDetail, "", "", xPos, 0, 1800, 300)
        ctl.Name = "cboLandUseId"
        ctl.ControlSource = "land_use_id"
        SetDatasheetCaptionSafe ctl, "Land Use"
        ctl.RowSource = "Select land_use_id, description FROM land_use ORDER BY description;"
        ctl.ColumnCount = 2
        ctl.ColumnWidths = "0cm;4cm"
        ctl.BoundColumn = 1
        ctl.LimitToList = True
        CreateLabel frm.Name, "lblLandUseId", "Land Use", xPos, 0, 1800, 300
        xPos = xPos + 1800

        ' placename (from joined table - display name)
        Set ctl = CreateControl(frm.Name, acTextBox, acDetail, "", "", xPos, 0, 2500, 300)
        ctl.Name = "txtPlacename"
        ctl.ControlSource = "placename"
        SetDatasheetCaptionSafe ctl, "Placename"
        ctl.Locked = True
        ctl.Enabled = False
        CreateLabel frm.Name, "lblPlacename", "Placename", xPos, 0, 2500, 300
        xPos = xPos + 2500

        ' original_placename
        Set ctl = CreateControl(frm.Name, acTextBox, acDetail, "", "", xPos, 0, 2000, 300)
        ctl.Name = "txtOriginalPlacename"
        ctl.ControlSource = "original_placename"
        SetDatasheetCaptionSafe ctl, "Original Placename"
        CreateLabel frm.Name, "lblOriginalPlacename", "Original Placename", xPos, 0, 2000, 300
        xPos = xPos + 2000

        ' cmdEntryDetail button
        Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, "", "", xPos, 0, 1000, 300)
        ctl.Name = "cmdEntryDetail"
        ctl.caption = "Detail..."
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
        frm.caption = "Court Case Entry Detail"
        frm.DefaultView = 0 ' Single Form
        frm.PopUp = True
        frm.Modal = True
        frm.NavigationButtons = True
        frm.RecordSelectors = True
        frm.AutoResize = False ' Allow manual resizing
        frm.AutoCenter = True
        frm.OnLoad = "=frmCourtCaseEntryDetail_OnLoad()" ' Auto-populate placename display

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
        ctl.RowSource = "Select season_id, season_name FROM season ORDER BY season_name;"
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
        ctl.RowSource = "Select land_use_id, description FROM land_use ORDER BY description;"
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
        ctl.caption = "Pick..."
        ctl.OnClick = "=frmCourtCaseEntryDetail_cmdPickPlacename_Click()"

        yPos = yPos + 500

        ' original_placename
        Set ctl = CreateControl(frm.Name, acTextBox, acDetail, "", "", 2000, yPos, 6000, 300)
        ctl.Name = "txtOriginalPlacename"
        ctl.ControlSource = "original_placename"
        ctl.HorizontalAnchor = acHorizontalAnchorBoth ' Stretch horizontally
        CreateLabel frm.Name, "lblOriginalPlacename", "Original Placename:", 200, yPos, 1600, 300

        yPos = yPos + 500

        ' curated_text
        Set ctl = CreateControl(frm.Name, acTextBox, acDetail, "", "", 2000, yPos, 6000, 800)
        ctl.Name = "txtCuratedText"
        ctl.ControlSource = "curated_text"
        ctl.HorizontalAnchor = acHorizontalAnchorBoth ' Stretch horizontally
        ctl.VerticalAnchor = acVerticalAnchorBoth ' Stretch vertically
        CreateLabel frm.Name, "lblCuratedText", "Curated Text:", 200, yPos, 1600, 300

        yPos = yPos + 1000

        ' sfrmPersonEntryByEntry subform
        Set ctl = CreateControl(frm.Name, acSubform, acDetail, "", "", 200, yPos, 9000, 2500)
        ctl.Name = "sfrmPersonEntryByEntry"
        ctl.SourceObject = "Form.sfrmPersonEntryByEntry"
        ctl.LinkMasterFields = "court_case_entry_id"
        ctl.LinkChildFields = "court_case_entry_id"
        ctl.HorizontalAnchor = acHorizontalAnchorBoth ' Stretch horizontally
        ctl.VerticalAnchor = acVerticalAnchorBoth ' Stretch vertically
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
        Dim xPos As Integer

        Set frm = CreateForm()
        strFormName = frm.Name
        frm.RecordSource = "person_entry"
        frm.caption = "Person Entries"
        frm.DefaultView = 2 ' Datasheet (grid view)
        frm.NavigationButtons = False
        frm.RecordSelectors = True
        frm.AllowAdditions = True
        frm.AllowEdits = True
        frm.OnLoad = "=sfrmPersonEntryByEntry_OnLoad()" ' Auto-size columns

        xPos = 0 ' Start position for columns

        ' person_id (numeric only, use picker)
        Set ctl = CreateControl(frm.Name, acTextBox, acDetail, "", "", xPos, 0, 1000, 300)
        ctl.Name = "txtPersonId"
        ctl.ControlSource = "person_id"
        SetDatasheetCaptionSafe ctl, "Person ID"
        CreateLabel frm.Name, "lblPersonId", "Person ID", xPos, 0, 1000, 300
        xPos = xPos + 1000

        ' cmdPickPerson
        Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, "", "", xPos, 0, 800, 300)
        ctl.Name = "cmdPickPerson"
        ctl.caption = "Pick..."
        ctl.OnClick = "=sfrmPersonEntryByEntry_cmdPickPerson_Click()"
        xPos = xPos + 800

        ' community_id
        Set ctl = CreateControl(frm.Name, acComboBox, acDetail, "", "", xPos, 0, 1800, 300)
        ctl.Name = "cboCommunityId"
        ctl.ControlSource = "community_id"
        SetDatasheetCaptionSafe ctl, "Community"
        ctl.RowSource = "Select community_id, community_name FROM community ORDER BY community_name;"
        ctl.ColumnCount = 2
        ctl.ColumnWidths = "0cm;4cm"
        ctl.BoundColumn = 1
        ctl.LimitToList = True
        CreateLabel frm.Name, "lblCommunityId", "Community", xPos, 0, 1800, 300
        xPos = xPos + 1800

        ' land_rights_status_id
        Set ctl = CreateControl(frm.Name, acComboBox, acDetail, "", "", xPos, 0, 1800, 300)
        ctl.Name = "cboLandRightsStatusId"
        ctl.ControlSource = "land_rights_status_id"
        SetDatasheetCaptionSafe ctl, "Land Rights"
        ctl.RowSource = "Select land_rights_status_id, land_rights_status FROM land_rights_status ORDER BY land_rights_status;"
        ctl.ColumnCount = 2
        ctl.ColumnWidths = "0cm;4cm"
        ctl.BoundColumn = 1
        ctl.LimitToList = True
        CreateLabel frm.Name, "lblLandRightsStatusId", "Land Rights", xPos, 0, 1800, 300
        xPos = xPos + 1800

        ' role_id
        Set ctl = CreateControl(frm.Name, acComboBox, acDetail, "", "", xPos, 0, 1500, 300)
        ctl.Name = "cboRoleId"
        ctl.ControlSource = "role_id"
        SetDatasheetCaptionSafe ctl, "Role"
        ctl.RowSource = "Select role_id, role_name FROM role ORDER BY role_name;"
        ctl.ColumnCount = 2
        ctl.ColumnWidths = "0cm;4cm"
        ctl.BoundColumn = 1
        ctl.LimitToList = True
        CreateLabel frm.Name, "lblRoleId", "Role", xPos, 0, 1500, 300

        DoCmd.Close acForm, strFormName, acSaveYes
        DoCmd.Rename "sfrmPersonEntryByEntry", acForm, strFormName

        Debug.Print "sfrmPersonEntryByEntry created."
     Exit Sub

ErrHandler:
        MsgBox "Error in Create_sfrmPersonEntryByEntry: " & Err.Description, vbCritical
End Sub

'------------------------------------------------------------------------------
' Create Form: sfrmRuling (Single: 0/1 ruling per Case)
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
        frm.caption = "Ruling"
        frm.DefaultView = 0 ' Single Form
        frm.NavigationButtons = False
        frm.RecordSelectors = False
        frm.AllowAdditions = False ' Prevent adding multiple rulings
        frm.AllowEdits = True
        frm.OnCurrent = "=sfrmRuling_OnCurrent()" ' Dynamic UI updates

        yPos = 200

        ' lblNoRuling - shown when no ruling exists
        Set ctl = CreateControl(frm.Name, acLabel, acDetail, "", "", 200, yPos, 6000, 400)
        ctl.Name = "lblNoRuling"
        ctl.caption = "No ruling yet. Click 'Create Ruling' button above To add one."
        ctl.ForeColor = RGB(128, 128, 128) ' Gray text

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
        ctl.RowSource = "Select ruling_type_id, ruling_type FROM ruling_type ORDER BY ruling_type;"
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
        ctl.RowSource = "Select legal_source_id, legal_source_name FROM legal_source ORDER BY legal_source_name;"
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
        ctl.HorizontalAnchor = acHorizontalAnchorBoth ' Stretch horizontally
        CreateLabel frm.Name, "lblDescription", "Description:", 200, yPos, 1200, 300

        yPos = yPos + 800

        ' sfrmPersonOutcomes subform (embedded in ruling)
        Set ctl = CreateControl(frm.Name, acSubform, acDetail, "", "", 200, yPos, 9000, 2000)
        ctl.Name = "sfrmPersonOutcomes"
        ctl.SourceObject = "Form.sfrmPersonOutcomes"
        ctl.LinkMasterFields = "ruling_id"
        ctl.LinkChildFields = "ruling_id"
        ctl.HorizontalAnchor = acHorizontalAnchorBoth ' Stretch horizontally
        ctl.VerticalAnchor = acVerticalAnchorBoth ' Stretch vertically
        CreateLabel strFormName, "lblPersonOutcomes", "Person Outcomes:", 200, yPos - 300, 3000, 300

        DoCmd.Close acForm, strFormName, acSaveYes
        DoCmd.Rename "sfrmRuling", acForm, strFormName

        Debug.Print "sfrmRuling created."
     Exit Sub

ErrHandler:
        MsgBox "Error in Create_sfrmRuling: " & Err.Description, vbCritical
End Sub

'------------------------------------------------------------------------------
' Create Form: sfrmPersonOutcomes (Grid outcomes For ruling)
'------------------------------------------------------------------------------
Private Sub Create_sfrmPersonOutcomes()
    On Error GoTo ErrHandler
        Dim frm As Form
        Dim ctl As Control
        Dim strFormName As String
        Dim xPos As Integer

        Set frm = CreateForm()
        strFormName = frm.Name
        frm.RecordSource = "person_outcome"
        frm.caption = "Person Outcomes"
        frm.DefaultView = 2 ' Datasheet (grid view)
        frm.NavigationButtons = False
        frm.RecordSelectors = True
        frm.AllowAdditions = True
        frm.AllowEdits = True
        frm.OnLoad = "=sfrmPersonOutcomes_OnLoad()" ' Auto-size columns

        xPos = 0 ' Start position for columns

        ' person_id
        Set ctl = CreateControl(frm.Name, acTextBox, acDetail, "", "", xPos, 0, 1000, 300)
        ctl.Name = "txtPersonId"
        ctl.ControlSource = "person_id"
        SetDatasheetCaptionSafe ctl, "Person ID"
        CreateLabel frm.Name, "lblPersonId", "Person ID", xPos, 0, 1000, 300
        xPos = xPos + 1000

        ' cmdPickPersonOutcomePerson
        Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, "", "", xPos, 0, 800, 300)
        ctl.Name = "cmdPickPersonOutcomePerson"
        ctl.caption = "Pick..."
        ctl.OnClick = "=sfrmPersonOutcomes_cmdPickPerson_Click()"
        xPos = xPos + 800

        ' outcome_type_id
        Set ctl = CreateControl(frm.Name, acComboBox, acDetail, "", "", xPos, 0, 2200, 300)
        ctl.Name = "cboOutcomeTypeId"
        ctl.ControlSource = "outcome_type_id"
        SetDatasheetCaptionSafe ctl, "Outcome Type"
        ctl.RowSource = "Select outcome_type_id, outcome_type_name FROM outcome_type ORDER BY outcome_type_name;"
        ctl.ColumnCount = 2
        ctl.ColumnWidths = "0cm;5cm"
        ctl.BoundColumn = 1
        ctl.LimitToList = True
        CreateLabel frm.Name, "lblOutcomeTypeId", "Outcome Type", xPos, 0, 2200, 300
        xPos = xPos + 2200

        ' description
        Set ctl = CreateControl(frm.Name, acTextBox, acDetail, "", "", xPos, 0, 3500, 300)
        ctl.Name = "txtDescription"
        ctl.ControlSource = "description"
        SetDatasheetCaptionSafe ctl, "Description"
        CreateLabel frm.Name, "lblDescription", "Description", xPos, 0, 3500, 300

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
        frm.caption = "Search Placename"
        frm.DefaultView = 0 ' Single Form
        frm.PopUp = True
        frm.Modal = True
        frm.NavigationButtons = False
        frm.RecordSelectors = False
        frm.OnLoad = "=frmPlacenameSearch_OnLoad()" ' Show initial results

        yPos = 200

        ' txtSearch
        Set ctl = CreateControl(frm.Name, acTextBox, acDetail, "", "", 1500, yPos, 4000, 300)
        ctl.Name = "txtSearch"
        CreateLabel frm.Name, "lblSearch", "Search:", 200, yPos, 1200, 300

        ' cmdSearch
        Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, "", "", 5600, yPos, 1500, 300)
        ctl.Name = "cmdSearch"
        ctl.caption = "Search"
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
        ctl.caption = "Select"
        ctl.OnClick = "=frmPlacenameSearch_cmdSelect_Click()"

        ' cmdCancel
        Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, "", "", 2400, yPos, 2000, 400)
        ctl.Name = "cmdCancel"
        ctl.caption = "Cancel"
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
        frm.caption = "Search Person"
        frm.DefaultView = 0 ' Single Form
        frm.PopUp = True
        frm.Modal = True
        frm.NavigationButtons = False
        frm.RecordSelectors = False
        frm.OnLoad = "=frmPersonSearch_OnLoad()" ' Show initial results

        yPos = 200

        ' txtSearch
        Set ctl = CreateControl(frm.Name, acTextBox, acDetail, "", "", 1500, yPos, 4000, 300)
        ctl.Name = "txtSearch"
        CreateLabel frm.Name, "lblSearch", "Search:", 200, yPos, 1200, 300

        ' cmdSearch
        Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, "", "", 5600, yPos, 1500, 300)
        ctl.Name = "cmdSearch"
        ctl.caption = "Search"
        ctl.OnClick = "=frmPersonSearch_cmdSearch_Click()"

        yPos = yPos + 500

        ' lstResults
        Set ctl = CreateControl(frm.Name, acListBox, acDetail, "", "", 200, yPos, 8000, 4000)
        ctl.Name = "lstResults"
        ctl.RowSourceType = "Table/Query"
        ctl.RowSource = ""
        ctl.ColumnCount = 4
        ctl.ColumnWidths = "1.5cm;4cm;1.5cm;3cm" ' Adjusted for community context
        ctl.ColumnHeads = True

        yPos = yPos + 4200

        ' cmdSelect
        Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, "", "", 200, yPos, 2000, 400)
        ctl.Name = "cmdSelect"
        ctl.caption = "Select"
        ctl.OnClick = "=frmPersonSearch_cmdSelect_Click()"

        ' cmdNewPerson
        Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, "", "", 2400, yPos, 2000, 400)
        ctl.Name = "cmdNewPerson"
        ctl.caption = "New Person"
        ctl.OnClick = "=frmPersonSearch_cmdNewPerson_Click()"

        ' cmdCancel
        Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, "", "", 4600, yPos, 2000, 400)
        ctl.Name = "cmdCancel"
        ctl.caption = "Cancel"
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
        frm.caption = "Person"
        frm.DefaultView = 0 ' Single Form
        frm.NavigationButtons = True
        frm.RecordSelectors = True
        frm.AllowAdditions = True
        frm.AllowEdits = True
        frm.PopUp = True
        frm.Modal = True
        frm.OnClose = "=frmPerson_OnClose()" ' Return new person_id to caller

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
' Helper: Create a Label control
'------------------------------------------------------------------------------
Private Function CreateLabel(formName As String, labelName As String, _
    caption As String, leftPos As Integer, topPos As Integer, _
    widthSize As Integer, heightSize As Integer) As Control
    Dim ctl As Control
    Set ctl = CreateControl(formName, acLabel, acDetail, "", "", leftPos, topPos, widthSize, heightSize)
    ctl.Name = labelName
    ctl.caption = caption
    Set CreateLabel = ctl
End Function

'------------------------------------------------------------------------------
' Helper: Create a Label in Form Header section (for continuous form column headers)
'------------------------------------------------------------------------------
Private Function CreateHeaderLabel(formName As String, labelName As String, _
    caption As String, leftPos As Integer, topPos As Integer, _
    widthSize As Integer, heightSize As Integer) As Control
    Dim ctl As Control
    On Error GoTo CreateFailed
    Set ctl = CreateControl(formName, acLabel, acHeader, "", "", leftPos, topPos, widthSize, heightSize)
    ctl.Name = labelName
    ctl.caption = caption
    Set CreateHeaderLabel = ctl
    Exit Function

CreateFailed:
    Err.Clear
    Set CreateHeaderLabel = Nothing
End Function

'------------------------------------------------------------------------------
' Helper: Enable Form Header section safely
'------------------------------------------------------------------------------
Private Sub EnableFormHeaderSafe(frm As Form, headerHeight As Integer)
    On Error Resume Next
    DoCmd.SelectObject acForm, frm.Name, True

    Err.Clear
    Dim currentHeaderHeight As Long
    currentHeaderHeight = frm.Section(acHeader).Height

    If Err.Number <> 0 Then
        Err.Clear
        DoCmd.RunCommand acCmdFormHdrFtr
    End If

    Err.Clear
    frm.Section(acHeader).Height = headerHeight
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Helper: Set datasheet caption safely (create property if needed)
'------------------------------------------------------------------------------
Private Sub SetDatasheetCaptionSafe(ByVal ctl As Control, ByVal captionText As String)
    On Error Resume Next
    ctl.DatasheetCaption = captionText   ' use direct property when available
    Err.Clear
    If Err.Number <> 0 Then
        ' Some controls/contexts don't support it. Don't attempt Append.
        Err.Clear
    End If
    On Error GoTo 0
End Sub