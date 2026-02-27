Option Compare Database
Option Explicit

'================================================================================
' DigiDiggie TNG – Access Form Generator (with)
'  (1) Placename picker/search form (no giant combo)
'  (2) Meaningful court_case display in combos (qry_CourtCaseDisplay)
'
' ASSUMPTION: PostgreSQL tables are linked into Access (ODBC).
' IMPORTANT: Update the T_* constants to match your linked table names EXACTLY.
'================================================================================

'========================
' CONFIG: Linked table names in Access
'========================
Private Const T_COMMUNITY As String = "community"
Private Const T_PARISH As String = "parish"
Private Const T_SOURCE As String = "source"
Private Const T_COURT_CASE As String = "court_case"
Private Const T_COURT_CASE_ENTRY As String = "court_case_entry"
Private Const T_SEASON As String = "season"
Private Const T_LAND_USE As String = "land_use"
Private Const T_PLACENAME As String = "placename"
Private Const T_PERSON As String = "person"
Private Const T_PERSON_ENTRY As String = "person_entry"
Private Const T_LAND_RIGHTS_STATUS As String = "land_rights_status"
Private Const T_ROLE_TYPE As String = "role_type"
Private Const T_ROLE As String = "role"
Private Const T_RULING As String = "ruling"
Private Const T_RULING_TYPE As String = "ruling_type"
Private Const T_LEGAL_SOURCE As String = "legal_source"
Private Const T_OUTCOME_TYPE As String = "outcome_type"
Private Const T_PERSON_OUTCOME As String = "person_outcome"
Private Const T_RELATIONSHIP_TYPE As String = "relationship_type"
Private Const T_PERSON_RELATIONSHIP As String = "person_relationship"

'========================
' CONFIG: Form Layout Constants
'========================
Private Const FORM_LABEL_LEFT As Long = 300
Private Const FORM_CONTROL_LEFT As Long = 2800
Private Const FORM_LABEL_WIDTH As Long = 2400
Private Const FORM_CONTROL_WIDTH As Long = 5200
Private Const FORM_ROW_HEIGHT As Long = 360
Private Const FORM_ROW_SPACING As Long = 450
Private Const FORM_START_TOP As Long = 600

'================================================================================
' ENTRY POINT
'================================================================================
Public Sub BuildForms_DigiDiggie_TNG()
    On Error GoTo EH
    
    Dim createdForms As Collection
    Set createdForms = New Collection
    Dim currentForm As String
    
    ' (2) Ensure display query exists for court case combos
    On Error Resume Next
    currentForm = "Query: qry_CourtCaseDisplay"
    EnsureCourtCaseDisplayQuery
    If Err.Number = 0 Then createdForms.Add currentForm
    On Error GoTo EH

    ' (1) Create placename picker/search dialog
    On Error Resume Next
    currentForm = "frm_PlacenamePicker"
    CreatePlacenamePickerForm
    If Err.Number = 0 Then createdForms.Add currentForm
    On Error GoTo EH

    ' --- Lookup / reference table forms (simple editors)
    On Error Resume Next
    currentForm = "frm_Parish": CreateSimpleTableForm "frm_Parish", T_PARISH, "parish_id"
    If Err.Number = 0 Then createdForms.Add currentForm
    On Error Resume Next
    currentForm = "frm_Source": CreateSimpleTableForm "frm_Source", T_SOURCE, "source_id"
    If Err.Number = 0 Then createdForms.Add currentForm
    On Error Resume Next
    currentForm = "frm_Season": CreateSimpleTableForm "frm_Season", T_SEASON, "season_id"
    If Err.Number = 0 Then createdForms.Add currentForm
    On Error Resume Next
    currentForm = "frm_LandUse": CreateSimpleTableForm "frm_LandUse", T_LAND_USE, "land_use_id"
    If Err.Number = 0 Then createdForms.Add currentForm
    On Error Resume Next
    currentForm = "frm_LandRightsStatus": CreateSimpleTableForm "frm_LandRightsStatus", T_LAND_RIGHTS_STATUS, "land_rights_status_id"
    If Err.Number = 0 Then createdForms.Add currentForm
    On Error Resume Next
    currentForm = "frm_RoleType": CreateSimpleTableForm "frm_RoleType", T_ROLE_TYPE, "role_type_id"
    If Err.Number = 0 Then createdForms.Add currentForm
    On Error Resume Next
    currentForm = "frm_Role": CreateRoleForm
    If Err.Number = 0 Then createdForms.Add currentForm
    On Error Resume Next
    currentForm = "frm_RulingType": CreateSimpleTableForm "frm_RulingType", T_RULING_TYPE, "ruling_type_id"
    If Err.Number = 0 Then createdForms.Add currentForm
    On Error Resume Next
    currentForm = "frm_LegalSource": CreateSimpleTableForm "frm_LegalSource", T_LEGAL_SOURCE, "legal_source_id"
    If Err.Number = 0 Then createdForms.Add currentForm
    On Error Resume Next
    currentForm = "frm_OutcomeType": CreateSimpleTableForm "frm_OutcomeType", T_OUTCOME_TYPE, "outcome_type_id"
    If Err.Number = 0 Then createdForms.Add currentForm
    On Error Resume Next
    currentForm = "frm_RelationshipType": CreateSimpleTableForm "frm_RelationshipType", T_RELATIONSHIP_TYPE, "relationship_type_id"
    If Err.Number = 0 Then createdForms.Add currentForm
    On Error GoTo EH

    ' Placename table might be large; keep a datasheet for browsing/editing
    On Error Resume Next
    currentForm = "frm_Placename": CreateDatasheetForm "frm_Placename", T_PLACENAME
    If Err.Number = 0 Then createdForms.Add currentForm
    On Error GoTo EH

    ' --- Core entity forms
    On Error Resume Next
    currentForm = "Community forms": CreateCommunityForm
    If Err.Number = 0 Then createdForms.Add currentForm
    On Error Resume Next
    currentForm = "CourtCase forms": CreateCourtCaseForms
    If Err.Number = 0 Then createdForms.Add currentForm
    On Error Resume Next
    currentForm = "CourtCaseEntry forms": CreateCourtCaseEntry_WithPersonEntry
    If Err.Number = 0 Then createdForms.Add currentForm
    On Error Resume Next
    currentForm = "Person forms": CreatePersonForms
    If Err.Number = 0 Then createdForms.Add currentForm
    On Error Resume Next
    currentForm = "Ruling forms": CreateRulingForms
    If Err.Number = 0 Then createdForms.Add currentForm
    On Error GoTo EH

    MsgBox "Successfully created " & createdForms.Count & " form(s)/object(s).", vbInformation, "DigiDiggie TNG"
    Exit Sub
    
EH:
    MsgBox "Error creating forms: " & Err.Description & vbCrLf & vbCrLf & _
           "Last attempted: " & currentForm & vbCrLf & _
           "Successfully created: " & createdForms.Count & " form(s)/object(s)", vbCritical, "Error"
End Sub

'================================================================================
' GENERIC HELPERS
'================================================================================
Private Function ObjectExists(ByVal objType As AcObjectType, ByVal objName As String) As Boolean
    On Error GoTo EH
    ObjectExists = (SysCmd(acSysCmdGetObjectState, objType, objName) <> 0)
    Exit Function
EH:
    Debug.Print "Error checking object '" & objName & "': " & Err.Description
    ObjectExists = False
End Function

Private Sub DeleteIfExists(ByVal objType As AcObjectType, ByVal objName As String)
    If ObjectExists(objType, objName) Then
        DoCmd.DeleteObject objType, objName
    End If
End Sub

Private Sub CreateOrReplaceQuery(ByVal qName As String, ByVal sql As String)
    On Error Resume Next
    CurrentDb.QueryDefs.Delete qName
    On Error GoTo 0
    CurrentDb.CreateQueryDef qName, sql
End Sub

Private Sub CreateSimpleTableForm(ByVal formName As String, ByVal tableName As String, ByVal pkField As String)
    DeleteIfExists acForm, formName

    Dim frm As Form
    DoCmd.CreateForm
    Set frm = Screen.ActiveForm

    frm.RecordSource = tableName
    frm.Caption = formName
    frm.DefaultView = acNormal

    LayoutFieldsAsTextboxes frm, tableName, pkField, True

    AddNavigationButtons formName

    DoCmd.Save acForm, formName
    DoCmd.Close acForm, frm.Name, acSaveYes
End Sub

Private Sub CreateDatasheetForm(ByVal formName As String, ByVal tableName As String)
    DeleteIfExists acForm, formName

    Dim frm As Form
    DoCmd.CreateForm
    Set frm = Screen.ActiveForm

    frm.RecordSource = tableName
    frm.Caption = formName
    frm.DefaultView = acDatasheet

    DoCmd.Save acForm, formName
    DoCmd.Close acForm, frm.Name, acSaveYes
End Sub

Private Sub LayoutFieldsAsTextboxes(ByVal frm As Form, ByVal tableName As String, ByVal pkField As String, ByVal hidePK As Boolean)
    Dim db As DAO.Database: Set db = CurrentDb
    Dim tdf As DAO.TableDef: Set tdf = db.TableDefs(tableName)
    Dim fld As DAO.Field

    Dim topPos As Long: topPos = 600
    Dim lblLeft As Long: lblLeft = 300
    Dim txtLeft As Long: txtLeft = 2800
    Dim lblWidth As Long: lblWidth = 2400
    Dim txtWidth As Long: txtWidth = 5200
    Dim rowHeight As Long: rowHeight = 360

    For Each fld In tdf.Fields
        If hidePK And fld.Name = pkField Then
            ' skip PK field
        Else
            CreateControl frm.Name, acLabel, acDetail, , fld.Name & ":", lblLeft, topPos, lblWidth, rowHeight
            Dim tb As Control
            Set tb = CreateControl(frm.Name, acTextBox, acDetail, , , txtLeft, topPos, txtWidth, rowHeight)
            tb.ControlSource = fld.Name
            topPos = topPos + rowHeight + 120
        End If
    Next fld
End Sub

Private Sub AddText(ByVal formName As String, ByVal fieldName As String, ByVal topPos As Long, _
                    Optional ByVal widthTwips As Long = 5200, Optional ByVal heightTwips As Long = 360, _
                    Optional ByVal isRequired As Boolean = False)
    Dim lblCaption As String
    lblCaption = fieldName & ":"
    If isRequired Then lblCaption = lblCaption & " *"
    
    Dim lbl As Control
    Set lbl = CreateControl(formName, acLabel, acDetail, , lblCaption, FORM_LABEL_LEFT, topPos, FORM_LABEL_WIDTH, FORM_ROW_HEIGHT)
    If isRequired Then lbl.ForeColor = RGB(255, 0, 0)
    
    Dim tb As Control
    Set tb = CreateControl(formName, acTextBox, acDetail, , , FORM_CONTROL_LEFT, topPos, widthTwips, heightTwips)
    tb.ControlSource = fieldName
    tb.ControlTipText = "Enter " & fieldName
End Sub

Private Sub AddMemo(ByVal formName As String, ByVal fieldName As String, ByVal topPos As Long, _
                    Optional ByVal heightTwips As Long = 900)
    CreateControl formName, acLabel, acDetail, , fieldName & ":", FORM_LABEL_LEFT, topPos, FORM_LABEL_WIDTH, FORM_ROW_HEIGHT
    Dim tb As Control
    Set tb = CreateControl(formName, acTextBox, acDetail, , , FORM_CONTROL_LEFT, topPos, FORM_CONTROL_WIDTH, heightTwips)
    tb.ControlSource = fieldName
    tb.EnterKeyBehavior = True
    tb.ScrollBars = 2
    tb.ControlTipText = "Enter " & fieldName
End Sub

Private Sub AddYearField(ByVal formName As String, ByVal fieldName As String, ByVal topPos As Long, _
                         Optional ByVal widthTwips As Long = 1000, Optional ByVal isRequired As Boolean = False)
    Dim lblCaption As String
    lblCaption = fieldName & ":"
    If isRequired Then lblCaption = lblCaption & " *"
    
    Dim lbl As Control
    Set lbl = CreateControl(formName, acLabel, acDetail, , lblCaption, FORM_LABEL_LEFT, topPos, FORM_LABEL_WIDTH, FORM_ROW_HEIGHT)
    If isRequired Then lbl.ForeColor = RGB(255, 0, 0)
    
    Dim tb As Control
    Set tb = CreateControl(formName, acTextBox, acDetail, , , FORM_CONTROL_LEFT, topPos, widthTwips, FORM_ROW_HEIGHT)
    tb.ControlSource = fieldName
    tb.ValidationRule = ">=1600 And <=2100"
    tb.ValidationText = "Year must be between 1600 and 2100"
    tb.ControlTipText = "Enter year (1600-2100)"
End Sub

Private Sub AddCombo(ByVal formName As String, ByVal boundField As String, _
                     ByVal lookupTable As String, ByVal keyField As String, ByVal displayField As String, _
                     ByVal leftPos As Long, ByVal topPos As Long, _
                     Optional ByVal labelCaption As String = "")
    If labelCaption = "" Then labelCaption = boundField

    CreateControl formName, acLabel, acDetail, , labelCaption & ":", 300, topPos, 2400, 360

    Dim cbo As Control
    Set cbo = CreateControl(formName, acComboBox, acDetail, , , leftPos, topPos, 5200, 360)
    cbo.ControlSource = boundField
    cbo.BoundColumn = 1
    cbo.ColumnCount = 2
    cbo.ColumnWidths = "0;7"
    cbo.LimitToList = True
    cbo.RowSourceType = "Table/Query"
    
    ' Intelligent sorting based on table/field
    Dim orderBy As String
    If (lookupTable = T_PERSON And displayField = "full_name") Then
        ' Sort persons by surname, given_name, patronymic for better UX
        orderBy = "ORDER BY surname, given_name, patronymic"
    Else
        ' Default: sort by display field
        orderBy = "ORDER BY [" & displayField & "]"
    End If
    
    cbo.RowSource = "SELECT [" & keyField & "], [" & displayField & "] FROM [" & lookupTable & "] " & orderBy & ";"
End Sub

' (2) Combo sourced from a query (for prettier court case display)
Private Sub AddComboFromQuery(ByVal formName As String, ByVal boundField As String, _
                              ByVal queryName As String, ByVal keyField As String, ByVal displayField As String, _
                              ByVal leftPos As Long, ByVal topPos As Long, Optional ByVal labelCaption As String = "")
    If labelCaption = "" Then labelCaption = boundField

    CreateControl formName, acLabel, acDetail, , labelCaption & ":", 300, topPos, 2400, 360

    Dim cbo As Control
    Set cbo = CreateControl(formName, acComboBox, acDetail, , , leftPos, topPos, 5200, 360)
    cbo.ControlSource = boundField
    cbo.BoundColumn = 1
    cbo.ColumnCount = 2
    cbo.ColumnWidths = "0;7"
    cbo.LimitToList = True
    cbo.RowSourceType = "Table/Query"
    cbo.RowSource = "SELECT [" & keyField & "], [" & displayField & "] FROM [" & queryName & "];"
End Sub

Private Sub AddSubform(ByVal mainFormName As String, ByVal subformName As String, _
                       ByVal linkMaster As String, ByVal linkChild As String, _
                       ByVal leftPos As Long, ByVal topPos As Long, ByVal widthTwips As Long, ByVal heightTwips As Long, _
                       Optional ByVal caption As String = "")
    If caption <> "" Then
        CreateControl mainFormName, acLabel, acDetail, , caption, leftPos, topPos - 330, widthTwips, 300
    End If

    Dim sf As Control
    Set sf = CreateControl(mainFormName, acSubform, acDetail, , subformName, leftPos, topPos, widthTwips, heightTwips)
    sf.SourceObject = "Form." & subformName
    sf.LinkMasterFields = linkMaster
    sf.LinkChildFields = linkChild
End Sub

Private Sub AddNavigationButtons(ByVal formName As String)
    ' Add standard navigation buttons to form footer
    Const BTN_TOP As Long = 100
    Const BTN_WIDTH As Long = 1200
    Const BTN_HEIGHT As Long = 360
    Const BTN_SPACING As Long = 100
    
    Dim btnLeft As Long
    btnLeft = 300
    
    ' First Record button
    Dim btnFirst As Control
    Set btnFirst = CreateControl(formName, acCommandButton, acFooter, , , btnLeft, BTN_TOP, BTN_WIDTH, BTN_HEIGHT)
    btnFirst.Caption = "|< First"
    btnFirst.OnClick = "=DoCmd.GoToRecord(,,acFirst)"
    btnLeft = btnLeft + BTN_WIDTH + BTN_SPACING
    
    ' Previous Record button
    Dim btnPrev As Control
    Set btnPrev = CreateControl(formName, acCommandButton, acFooter, , , btnLeft, BTN_TOP, BTN_WIDTH, BTN_HEIGHT)
    btnPrev.Caption = "< Previous"
    btnPrev.OnClick = "=DoCmd.GoToRecord(,,acPrevious)"
    btnLeft = btnLeft + BTN_WIDTH + BTN_SPACING
    
    ' Next Record button
    Dim btnNext As Control
    Set btnNext = CreateControl(formName, acCommandButton, acFooter, , , btnLeft, BTN_TOP, BTN_WIDTH, BTN_HEIGHT)
    btnNext.Caption = "Next >"
    btnNext.OnClick = "=DoCmd.GoToRecord(,,acNext)"
    btnLeft = btnLeft + BTN_WIDTH + BTN_SPACING
    
    ' Last Record button
    Dim btnLast As Control
    Set btnLast = CreateControl(formName, acCommandButton, acFooter, , , btnLeft, BTN_TOP, BTN_WIDTH, BTN_HEIGHT)
    btnLast.Caption = "Last >|"
    btnLast.OnClick = "=DoCmd.GoToRecord(,,acLast)"
    btnLeft = btnLeft + BTN_WIDTH + BTN_SPACING + 300
    
    ' New Record button
    Dim btnNew As Control
    Set btnNew = CreateControl(formName, acCommandButton, acFooter, , , btnLeft, BTN_TOP, BTN_WIDTH, BTN_HEIGHT)
    btnNew.Caption = "+ New"
    btnNew.OnClick = "=DoCmd.GoToRecord(,,acNewRec)"
End Sub

'================================================================================
' (2) COURT CASE DISPLAY QUERY
'================================================================================
Private Sub EnsureCourtCaseDisplayQuery()
    Dim sql As String
    sql = "SELECT cc.court_case_id, " & _
          "Trim(Nz(s.source_abbreviation,'') & ' ' & " & _
          "Nz(cc.reference_number,'') & " & _
          "IIf(IsNull(cc.case_year),'',' (' & cc.case_year & ')')) AS court_case_display " & _
          "FROM [" & T_COURT_CASE & "] AS cc " & _
          "INNER JOIN [" & T_SOURCE & "] AS s ON cc.source_id = s.source_id " & _
          "ORDER BY s.source_abbreviation, cc.case_year, cc.reference_number;"
    CreateOrReplaceQuery "qry_CourtCaseDisplay", sql
End Sub

'================================================================================
' (1) PLACENAME PICKER – public functions used by generated controls
'================================================================================
Public Function PickPlacename(ByVal targetControlName As String) As Boolean
    ' Called from any form/subform button with: =PickPlacename("placename_id")
    On Error GoTo EH

    Dim callerFormName As String
    callerFormName = Screen.ActiveForm.Name

    If TempVars.Exists("PlacenameSelectedId") Then TempVars.Remove "PlacenameSelectedId"
    If TempVars.Exists("PlacenameTargetForm") Then TempVars.Remove "PlacenameTargetForm"
    If TempVars.Exists("PlacenameTargetControl") Then TempVars.Remove "PlacenameTargetControl"

    TempVars.Add "PlacenameTargetForm", callerFormName
    TempVars.Add "PlacenameTargetControl", targetControlName

    DoCmd.OpenForm "frm_PlacenamePicker", WindowMode:=acDialog

    If TempVars.Exists("PlacenameSelectedId") Then
        Forms(callerFormName).Controls(targetControlName).Value = TempVars!PlacenameSelectedId
        TempVars.Remove "PlacenameSelectedId"
        PickPlacename = True
    Else
        PickPlacename = False
    End If

    Exit Function
EH:
    PickPlacename = False
End Function

Public Function PlacenamePicker_Search() As Boolean
    On Error GoTo EH
    Dim f As Form
    Set f = Forms("frm_PlacenamePicker")

    Dim q As String
    q = Nz(f.Controls("txtSearch").Value, "")

    Dim sql As String
    sql = "SELECT placename_id, placename, parish_name, northing, easting " & _
          "FROM [" & f.Controls("txtTable").Value & "] "

    If Len(Trim$(q)) > 0 Then
        q = Replace(q, "'", "''")
        sql = sql & "WHERE placename LIKE '*" & q & "*' " & _
                    "OR parish_name LIKE '*" & q & "*' "
    End If

    sql = sql & "ORDER BY placename;"

    f.Controls("lstResults").RowSource = sql
    f.Controls("lstResults").Requery

    PlacenamePicker_Search = True
    Exit Function
EH:
    PlacenamePicker_Search = False
End Function

Public Function PlacenamePicker_UseSelected() As Boolean
    On Error GoTo EH
    Dim f As Form
    Set f = Forms("frm_PlacenamePicker")

    Dim idValue As Variant
    idValue = f.Controls("lstResults").Value

    If IsNull(idValue) Then
        PlacenamePicker_UseSelected = False
        Exit Function
    End If

    If TempVars.Exists("PlacenameSelectedId") Then TempVars.Remove "PlacenameSelectedId"
    TempVars.Add "PlacenameSelectedId", CLng(idValue)

    DoCmd.Close acForm, "frm_PlacenamePicker", acSaveNo
    PlacenamePicker_UseSelected = True
    Exit Function
EH:
    PlacenamePicker_UseSelected = False
End Function

Public Function PlacenamePicker_Cancel() As Boolean
    On Error Resume Next
    If TempVars.Exists("PlacenameSelectedId") Then TempVars.Remove "PlacenameSelectedId"
    DoCmd.Close acForm, "frm_PlacenamePicker", acSaveNo
    PlacenamePicker_Cancel = True
End Function

Private Sub CreatePlacenamePickerForm()
    DeleteIfExists acForm, "frm_PlacenamePicker"

    DoCmd.CreateForm
    Dim frm As Form: Set frm = Screen.ActiveForm
    frm.Caption = "Pick placename"
    frm.DefaultView = acNormal

    ' hidden textbox storing table name used by picker
    Dim tbTable As Control
    Set tbTable = CreateControl(frm.Name, acTextBox, acDetail, , , 100, 100, 0, 0)
    tbTable.Name = "txtTable"
    tbTable.Visible = False
    tbTable.Value = T_PLACENAME

    ' Search label + textbox
    CreateControl frm.Name, acLabel, acDetail, , "Search (placename or parish):", 300, 300, 3000, 300
    Dim tbSearch As Control
    Set tbSearch = CreateControl(frm.Name, acTextBox, acDetail, , , 3300, 300, 4500, 300)
    tbSearch.Name = "txtSearch"
    tbSearch.AfterUpdate = "=PlacenamePicker_Search()"

    ' Results listbox
    Dim lb As Control
    Set lb = CreateControl(frm.Name, acListBox, acDetail, , , 300, 750, 7500, 3000)
    lb.Name = "lstResults"
    lb.ColumnCount = 5
    lb.BoundColumn = 1
    lb.ColumnWidths = "0;6;4;2;2"
    lb.RowSourceType = "Table/Query"
    lb.RowSource = "SELECT placename_id, placename, parish_name, northing, easting FROM [" & T_PLACENAME & "] ORDER BY placename;"

    ' Use selected
    Dim btnUse As Control
    Set btnUse = CreateControl(frm.Name, acCommandButton, acDetail, , "Use selected", 300, 3900, 1600, 400)
    btnUse.Name = "cmdUse"
    btnUse.OnClick = "=PlacenamePicker_UseSelected()"

    ' Cancel
    Dim btnCancel As Control
    Set btnCancel = CreateControl(frm.Name, acCommandButton, acDetail, , "Cancel", 2000, 3900, 1200, 400)
    btnCancel.Name = "cmdCancel"
    btnCancel.OnClick = "=PlacenamePicker_Cancel()"

    DoCmd.Save acForm, "frm_PlacenamePicker"
    DoCmd.Close acForm, frm.Name, acSaveYes
End Sub

'================================================================================
' SCHEMA-SPECIFIC FORMS
'================================================================================

' community(parish_id -> parish)
Private Sub CreateCommunityForm()
    DeleteIfExists acForm, "frm_Community"

    DoCmd.CreateForm
    Dim frm As Form: Set frm = Screen.ActiveForm
    frm.RecordSource = T_COMMUNITY
    frm.Caption = "Community"

    ' community_name
    CreateControl frm.Name, acLabel, acDetail, , "community_name:", 300, 600, 2400, 360
    Dim tb As Control
    Set tb = CreateControl(frm.Name, acTextBox, acDetail, , , 2800, 600, 5200, 360)
    tb.ControlSource = "community_name"

    ' parish_id combo
    AddCombo frm.Name, "parish_id", T_PARISH, "parish_id", "parish", 2800, 1080, "parish"

    AddNavigationButtons frm.Name

    DoCmd.Save acForm, "frm_Community"
    DoCmd.Close acForm, frm.Name, acSaveYes
End Sub

' court_case(source_id -> source) + subform court_case_entry
Private Sub CreateCourtCaseForms()
    CreateCourtCaseEntrySubform

    DeleteIfExists acForm, "frm_CourtCase"

    DoCmd.CreateForm
    Dim frm As Form: Set frm = Screen.ActiveForm
    frm.RecordSource = T_COURT_CASE
    frm.Caption = "Court Case"

    Dim topPos As Long: topPos = 600

    AddCombo frm.Name, "source_id", T_SOURCE, "source_id", "source_name", 2800, topPos, "source"
    topPos = topPos + 480

    AddText frm.Name, "reference_number", topPos, 2000, 360
    topPos = topPos + 480

    AddText frm.Name, "district_court_name", topPos
    topPos = topPos + 480

    AddYearField frm.Name, "case_year", topPos
    topPos = topPos + 700

    AddMemo frm.Name, "source_text", topPos, 1200
    topPos = topPos + 1500

    AddSubform frm.Name, "sfrm_CourtCaseEntry", "court_case_id", "court_case_id", 300, topPos, 8000, 2600, "Case entries"

    AddNavigationButtons frm.Name

    DoCmd.Save acForm, "frm_CourtCase"
    DoCmd.Close acForm, frm.Name, acSaveYes
End Sub

Private Sub CreateCourtCaseEntrySubform()
    DeleteIfExists acForm, "sfrm_CourtCaseEntry"

    DoCmd.CreateForm
    Dim frm As Form: Set frm = Screen.ActiveForm
    frm.RecordSource = T_COURT_CASE_ENTRY
    frm.Caption = "Court Case Entries"
    frm.DefaultView = acContinuousForms

    Dim topPos As Long: topPos = 300

    ' Hide court_case_id (link field)
    Dim tbCaseId As Control
    Set tbCaseId = CreateControl(frm.Name, acTextBox, acDetail, , , 100, topPos, 0, 0)
    tbCaseId.ControlSource = "court_case_id"
    tbCaseId.Visible = False

    ' entry_year (with validation)
    CreateControl frm.Name, acLabel, acDetail, , "entry_year:", 300, topPos, 1400, 300
    Dim tbY As Control
    Set tbY = CreateControl(frm.Name, acTextBox, acDetail, , , 1800, topPos, 900, 300)
    tbY.ControlSource = "entry_year"
    tbY.ValidationRule = ">=1600 And <=2100"
    tbY.ValidationText = "Year must be between 1600 and 2100"
    tbY.ControlTipText = "Enter year (1600-2100)"

    ' season
    AddCombo frm.Name, "season_id", T_SEASON, "season_id", "season_name", 3500, topPos, "season"

    topPos = topPos + 480

    ' land use
    AddCombo frm.Name, "land_use_id", T_LAND_USE, "land_use_id", "description", 2800, topPos, "land_use"
    topPos = topPos + 480

    ' original_placename
    CreateControl frm.Name, acLabel, acDetail, , "original_placename:", 300, topPos, 2400, 300
    Dim tbOP As Control
    Set tbOP = CreateControl(frm.Name, acTextBox, acDetail, , , 2800, topPos, 5200, 300)
    tbOP.ControlSource = "original_placename"
    topPos = topPos + 480

    ' placename_id textbox + picker button (1)
    CreateControl frm.Name, acLabel, acDetail, , "placename_id:", 300, topPos, 2400, 300
    Dim tbPID As Control
    Set tbPID = CreateControl(frm.Name, acTextBox, acDetail, , , 2800, topPos, 1200, 300)
    tbPID.ControlSource = "placename_id"

    Dim btnPick As Control
    Set btnPick = CreateControl(frm.Name, acCommandButton, acDetail, , "Pick…", 4100, topPos, 900, 300)
    btnPick.OnClick = "=PickPlacename(""placename_id"")"

    topPos = topPos + 480

    ' curated_text
    CreateControl frm.Name, acLabel, acDetail, , "curated_text:", 300, topPos, 2400, 300
    Dim tbCT As Control
    Set tbCT = CreateControl(frm.Name, acTextBox, acDetail, , , 2800, topPos, 5200, 900)
    tbCT.ControlSource = "curated_text"
    tbCT.EnterKeyBehavior = True
    tbCT.ScrollBars = 2

    DoCmd.Save acForm, "sfrm_CourtCaseEntry"
    DoCmd.Close acForm, frm.Name, acSaveYes
End Sub

' court_case_entry as main form + person_entry subform
Private Sub CreateCourtCaseEntry_WithPersonEntry()
    CreatePersonEntrySubform

    DeleteIfExists acForm, "frm_CourtCaseEntry"

    DoCmd.CreateForm
    Dim frm As Form: Set frm = Screen.ActiveForm
    frm.RecordSource = T_COURT_CASE_ENTRY
    frm.Caption = "Court Case Entry"

    Dim topPos As Long: topPos = 600

    ' (2) court_case_id as meaningful display combo
    AddComboFromQuery frm.Name, "court_case_id", "qry_CourtCaseDisplay", "court_case_id", "court_case_display", 2800, topPos, "court case"
    topPos = topPos + 480

    AddText frm.Name, "entry_year", topPos, 1000, 360
    topPos = topPos + 480

    AddCombo frm.Name, "season_id", T_SEASON, "season_id", "season_name", 2800, topPos, "season"
    topPos = topPos + 480

    AddCombo frm.Name, "land_use_id", T_LAND_USE, "land_use_id", "description", 2800, topPos, "land_use"
    topPos = topPos + 480

    AddText frm.Name, "original_placename", topPos
    topPos = topPos + 480

    ' placename_id textbox + picker button (1)
    CreateControl frm.Name, acLabel, acDetail, , "placename_id:", 300, topPos, 2400, 360
    Dim tbPID As Control
    Set tbPID = CreateControl(frm.Name, acTextBox, acDetail, , , 2800, topPos, 1200, 360)
    tbPID.ControlSource = "placename_id"

    Dim btnPick As Control
    Set btnPick = CreateControl(frm.Name, acCommandButton, acDetail, , "Pick…", 4100, topPos, 900, 360)
    btnPick.OnClick = "=PickPlacename(""placename_id"")"

    topPos = topPos + 480

    AddMemo frm.Name, "curated_text", topPos, 1200
    topPos = topPos + 1500

    AddSubform frm.Name, "sfrm_PersonEntry", "court_case_entry_id", "court_case_entry_id", 300, topPos, 8000, 2600, "Persons in this entry"

    DoCmd.Save acForm, "frm_CourtCaseEntry"
    DoCmd.Close acForm, frm.Name, acSaveYes
End Sub

Private Sub CreatePersonEntrySubform()
    DeleteIfExists acForm, "sfrm_PersonEntry"

    DoCmd.CreateForm
    Dim frm As Form: Set frm = Screen.ActiveForm
    frm.RecordSource = T_PERSON_ENTRY
    frm.Caption = "Person Entry"
    frm.DefaultView = acContinuousForms

    Dim topPos As Long: topPos = 300

    ' Link field hidden
    Dim tbLink As Control
    Set tbLink = CreateControl(frm.Name, acTextBox, acDetail, , , 100, topPos, 0, 0)
    tbLink.ControlSource = "court_case_entry_id"
    tbLink.Visible = False

    AddCombo frm.Name, "person_id", T_PERSON, "person_id", "full_name", 2800, topPos, "person"
    topPos = topPos + 480

    AddCombo frm.Name, "community_id", T_COMMUNITY, "community_id", "community_name", 2800, topPos, "community"
    topPos = topPos + 480

    AddCombo frm.Name, "land_rights_status_id", T_LAND_RIGHTS_STATUS, "land_rights_status_id", "land_rights_status", 2800, topPos, "land_rights_status"
    topPos = topPos + 480

    AddCombo frm.Name, "role_id", T_ROLE, "role_id", "role_name", 2800, topPos, "role"
    topPos = topPos + 480

    AddMemo frm.Name, "curated_text", topPos, 900

    DoCmd.Save acForm, "sfrm_PersonEntry"
    DoCmd.Close acForm, frm.Name, acSaveYes
End Sub

' person main form + subforms
Private Sub CreatePersonForms()
    CreatePersonEntrySubform_ForPerson
    CreatePersonRelationshipSubform

    DeleteIfExists acForm, "frm_Person"

    DoCmd.CreateForm
    Dim frm As Form: Set frm = Screen.ActiveForm
    frm.RecordSource = T_PERSON
    frm.Caption = "Person"

    Dim topPos As Long: topPos = 600

    AddText frm.Name, "given_name", topPos: topPos = topPos + 480
    AddText frm.Name, "patronymic", topPos: topPos = topPos + 480
    AddText frm.Name, "surname", topPos: topPos = topPos + 480
    AddYearField frm.Name, "birth_year", topPos: topPos = topPos + 480
    AddYearField frm.Name, "death_year", topPos: topPos = topPos + 480
    AddText frm.Name, "community_name", topPos: topPos = topPos + 480

    AddMemo frm.Name, "note", topPos, 900
    topPos = topPos + 1200

    ' full_name generated (locked)
    CreateControl frm.Name, acLabel, acDetail, , "full_name:", 300, topPos, 2400, 360
    Dim tbFN As Control
    Set tbFN = CreateControl(frm.Name, acTextBox, acDetail, , , 2800, topPos, 5200, 360)
    tbFN.ControlSource = "full_name"
    tbFN.Locked = True
    tbFN.Enabled = True
    topPos = topPos + 800

    AddSubform frm.Name, "sfrm_PersonEntry_ByPerson", "person_id", "person_id", 300, topPos, 8000, 2200, "Entries for this person"
    topPos = topPos + 2500

    AddSubform frm.Name, "sfrm_PersonRelationship", "person_id", "person_1_id", 300, topPos, 8000, 2200, "Relationships (as person_1)"

    AddNavigationButtons frm.Name

    DoCmd.Save acForm, "frm_Person"
    DoCmd.Close acForm, frm.Name, acSaveYes
End Sub

Private Sub CreatePersonEntrySubform_ForPerson()
    DeleteIfExists acForm, "sfrm_PersonEntry_ByPerson"

    DoCmd.CreateForm
    Dim frm As Form: Set frm = Screen.ActiveForm
    frm.RecordSource = T_PERSON_ENTRY
    frm.Caption = "Person Entries"
    frm.DefaultView = acContinuousForms

    Dim topPos As Long: topPos = 300

    ' Hide person_id (link)
    Dim tbPID As Control
    Set tbPID = CreateControl(frm.Name, acTextBox, acDetail, , , 100, topPos, 0, 0)
    tbPID.ControlSource = "person_id"
    tbPID.Visible = False

    ' Show court_case_entry_id as number (simple)
    CreateControl frm.Name, acLabel, acDetail, , "court_case_entry_id:", 300, topPos, 2400, 300
    Dim tbCCE As Control
    Set tbCCE = CreateControl(frm.Name, acTextBox, acDetail, , , 2800, topPos, 1200, 300)
    tbCCE.ControlSource = "court_case_entry_id"
    topPos = topPos + 480

    AddCombo frm.Name, "community_id", T_COMMUNITY, "community_id", "community_name", 2800, topPos, "community"
    topPos = topPos + 480

    AddCombo frm.Name, "land_rights_status_id", T_LAND_RIGHTS_STATUS, "land_rights_status_id", "land_rights_status", 2800, topPos, "land_rights_status"
    topPos = topPos + 480

    AddCombo frm.Name, "role_id", T_ROLE, "role_id", "role_name", 2800, topPos, "role"
    topPos = topPos + 480

    DoCmd.Save acForm, "sfrm_PersonEntry_ByPerson"
    DoCmd.Close acForm, frm.Name, acSaveYes
End Sub

Private Sub CreatePersonRelationshipSubform()
    DeleteIfExists acForm, "sfrm_PersonRelationship"

    DoCmd.CreateForm
    Dim frm As Form: Set frm = Screen.ActiveForm
    frm.RecordSource = T_PERSON_RELATIONSHIP
    frm.Caption = "Person Relationships"
    frm.DefaultView = acContinuousForms

    Dim topPos As Long: topPos = 300

    ' Hide person_1_id (link)
    Dim tbP1 As Control
    Set tbP1 = CreateControl(frm.Name, acTextBox, acDetail, , , 100, topPos, 0, 0)
    tbP1.ControlSource = "person_1_id"
    tbP1.Visible = False

    AddCombo frm.Name, "person_2_id", T_PERSON, "person_id", "full_name", 2800, topPos, "related person"
    topPos = topPos + 480

    AddCombo frm.Name, "relationship_type_id", T_RELATIONSHIP_TYPE, "relationship_type_id", "relationship_type_name", 2800, topPos, "relationship_type"
    topPos = topPos + 480

    AddMemo frm.Name, "description", topPos, 600

    DoCmd.Save acForm, "sfrm_PersonRelationship"
    DoCmd.Close acForm, frm.Name, acSaveYes
End Sub

' ruling + person_outcome subform
Private Sub CreateRulingForms()
    CreatePersonOutcomeSubform

    DeleteIfExists acForm, "frm_Ruling"

    DoCmd.CreateForm
    Dim frm As Form: Set frm = Screen.ActiveForm
    frm.RecordSource = T_RULING
    frm.Caption = "Ruling"

    Dim topPos As Long: topPos = 600

    ' (2) court_case_id meaningful display combo
    AddComboFromQuery frm.Name, "court_case_id", "qry_CourtCaseDisplay", "court_case_id", "court_case_display", 2800, topPos, "court case"
    topPos = topPos + 480

    AddCombo frm.Name, "ruling_type_id", T_RULING_TYPE, "ruling_type_id", "ruling_type", 2800, topPos, "ruling_type"
    topPos = topPos + 480

    AddCombo frm.Name, "legal_source_id", T_LEGAL_SOURCE, "legal_source_id", "legal_source_name", 2800, topPos, "legal_source"
    topPos = topPos + 480

    AddYearField frm.Name, "ruling_year", topPos
    topPos = topPos + 480

    AddMemo frm.Name, "description", topPos, 900
    topPos = topPos + 1200

    AddSubform frm.Name, "sfrm_PersonOutcome", "ruling_id", "ruling_id", 300, topPos, 8000, 2400, "Outcomes for persons"

    AddNavigationButtons frm.Name

    DoCmd.Save acForm, "frm_Ruling"
    DoCmd.Close acForm, frm.Name, acSaveYes
End Sub

Private Sub CreatePersonOutcomeSubform()
    DeleteIfExists acForm, "sfrm_PersonOutcome"

    DoCmd.CreateForm
    Dim frm As Form: Set frm = Screen.ActiveForm
    frm.RecordSource = T_PERSON_OUTCOME
    frm.Caption = "Person Outcomes"
    frm.DefaultView = acContinuousForms

    Dim topPos As Long: topPos = 300

    ' Hide ruling_id (link)
    Dim tbR As Control
    Set tbR = CreateControl(frm.Name, acTextBox, acDetail, , , 100, topPos, 0, 0)
    tbR.ControlSource = "ruling_id"
    tbR.Visible = False

    AddCombo frm.Name, "person_id", T_PERSON, "person_id", "full_name", 2800, topPos, "person"
    topPos = topPos + 480

    AddCombo frm.Name, "outcome_type_id", T_OUTCOME_TYPE, "outcome_type_id", "outcome_type_name", 2800, topPos, "outcome_type"
    topPos = topPos + 480

    AddMemo frm.Name, "description", topPos, 600

    DoCmd.Save acForm, "sfrm_PersonOutcome"
    DoCmd.Close acForm, frm.Name, acSaveYes
End Sub

' role(role_type_id -> role_type)
Private Sub CreateRoleForm()
    DeleteIfExists acForm, "frm_Role"
    
    DoCmd.CreateForm
    Dim frm As Form: Set frm = Screen.ActiveForm
    frm.RecordSource = T_ROLE
    frm.Caption = "Roles"
    frm.DefaultView = acNormal
    
    Dim topPos As Long: topPos = 600
    
    ' role_id (PK, hidden)
    Dim tbPK As Control
    Set tbPK = CreateControl(frm.Name, acTextBox, acDetail, , , 300, topPos, 0, 0)
    tbPK.ControlSource = "role_id"
    tbPK.Visible = False
    
    ' role_name
    AddText frm.Name, "role_name", topPos
    topPos = topPos + 450
    
    ' role_type_id (FK -> role_type)
    AddCombo frm.Name, "role_type_id", T_ROLE_TYPE, "role_type_id", "role_type_name", 2800, topPos, "Role Type"
    topPos = topPos + 450
    
    ' description
    AddMemo frm.Name, "description", topPos, 900
    
    AddNavigationButtons frm.Name
    
    DoCmd.Save acForm, "frm_Role"
    DoCmd.Close acForm, frm.Name, acSaveYes
End Sub
