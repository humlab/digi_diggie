Option Compare Database
Option Explicit

'================================================================================
' CONFIG: Change these to match your linked-table names in Access
' (Check the Navigation Pane for exact names.)
'================================================================================
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

'================================================================================
' ENTRY POINT
'================================================================================
Public Sub BuildForms_DigiDiggie_TNG()
    ' --- Lookup / reference tables (simple editors)
    CreateSimpleTableForm "frm_Parish", T_PARISH, "parish_id"
    CreateSimpleTableForm "frm_Source", T_SOURCE, "source_id"
    CreateSimpleTableForm "frm_Season", T_SEASON, "season_id"
    CreateSimpleTableForm "frm_LandUse", T_LAND_USE, "land_use_id"
    CreateSimpleTableForm "frm_LandRightsStatus", T_LAND_RIGHTS_STATUS, "land_rights_status_id"
    CreateSimpleTableForm "frm_RoleType", T_ROLE_TYPE, "role_type_id"
    CreateSimpleTableForm "frm_Role", T_ROLE, "role_id"
    CreateSimpleTableForm "frm_RulingType", T_RULING_TYPE, "ruling_type_id"
    CreateSimpleTableForm "frm_LegalSource", T_LEGAL_SOURCE, "legal_source_id"
    CreateSimpleTableForm "frm_OutcomeType", T_OUTCOME_TYPE, "outcome_type_id"
    CreateSimpleTableForm "frm_RelationshipType", T_RELATIONSHIP_TYPE, "relationship_type_id"

    ' Placename can be large; make it datasheet for browsing/editing (or read-only).
    CreateDatasheetForm "frm_Placename", T_PLACENAME

    ' --- Core entity forms
    CreateCommunityForm
    CreateCourtCaseForms
    CreateCourtCaseEntry_WithPersonEntry
    CreatePersonForms
    CreateRulingForms

    MsgBox "Finished generating TNG forms.", vbInformation
End Sub

'================================================================================
' GENERIC HELPERS
'================================================================================
Private Function ObjectExists(ByVal objType As AcObjectType, ByVal objName As String) As Boolean
    On Error GoTo EH
    ObjectExists = (SysCmd(acSysCmdGetObjectState, objType, objName) <> 0)
    Exit Function
EH:
    ObjectExists = False
End Function

Private Sub DeleteIfExists(ByVal objType As AcObjectType, ByVal objName As String)
    If ObjectExists(objType, objName) Then
        DoCmd.DeleteObject objType, objName
    End If
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
    cbo.ColumnWidths = "0;6" ' hide key, show display
    cbo.LimitToList = True
    cbo.RowSourceType = "Table/Query"
    cbo.RowSource = "SELECT [" & keyField & "], [" & displayField & "] FROM [" & lookupTable & "] ORDER BY [" & displayField & "];"
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

    ' community_name textbox
    CreateControl frm.Name, acLabel, acDetail, , "community_name:", 300, 600, 2400, 360
    Dim tb As Control
    Set tb = CreateControl(frm.Name, acTextBox, acDetail, , , 2800, 600, 5200, 360)
    tb.ControlSource = "community_name"

    ' parish_id combo
    AddCombo frm.Name, "parish_id", T_PARISH, "parish_id", "parish", 2800, 1080, "parish"

    DoCmd.Save acForm, "frm_Community"
    DoCmd.Close acForm, frm.Name, acSaveYes
End Sub

' court_case(source_id -> source) + subform court_case_entry
Private Sub CreateCourtCaseForms()
    ' Subform first
    CreateCourtCaseEntrySubform

    DeleteIfExists acForm, "frm_CourtCase"

    DoCmd.CreateForm
    Dim frm As Form: Set frm = Screen.ActiveForm
    frm.RecordSource = T_COURT_CASE
    frm.Caption = "Court Case"

    Dim topPos As Long: topPos = 600

    AddCombo frm.Name, "source_id", T_SOURCE, "source_id", "source_name", 2800, topPos, "source"
    topPos = topPos + 480

    ' reference_number
    CreateControl frm.Name, acLabel, acDetail, , "reference_number:", 300, topPos, 2400, 360
    Dim tbRef As Control
    Set tbRef = CreateControl(frm.Name, acTextBox, acDetail, , , 2800, topPos, 2000, 360)
    tbRef.ControlSource = "reference_number"
    topPos = topPos + 480

    ' district_court_name
    CreateControl frm.Name, acLabel, acDetail, , "district_court_name:", 300, topPos, 2400, 360
    Dim tbDC As Control
    Set tbDC = CreateControl(frm.Name, acTextBox, acDetail, , , 2800, topPos, 5200, 360)
    tbDC.ControlSource = "district_court_name"
    topPos = topPos + 480

    ' case_year
    CreateControl frm.Name, acLabel, acDetail, , "case_year:", 300, topPos, 2400, 360
    Dim tbYear As Control
    Set tbYear = CreateControl(frm.Name, acTextBox, acDetail, , , 2800, topPos, 1000, 360)
    tbYear.ControlSource = "case_year"
    topPos = topPos + 700

    ' source_text
    CreateControl frm.Name, acLabel, acDetail, , "source_text:", 300, topPos, 2400, 360
    Dim tbST As Control
    Set tbST = CreateControl(frm.Name, acTextBox, acDetail, , , 2800, topPos, 5200, 1200)
    tbST.ControlSource = "source_text"
    tbST.EnterKeyBehavior = True
    tbST.ScrollBars = 2
    topPos = topPos + 1500

    ' Subform for court_case_entry (linked by court_case_id)
    AddSubform frm.Name, "sfrm_CourtCaseEntry", "court_case_id", "court_case_id", 300, topPos, 8000, 2600, "Case entries"

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

    ' Keep court_case_id hidden (link field)
    Dim tbCaseId As Control
    Set tbCaseId = CreateControl(frm.Name, acTextBox, acDetail, , , 100, topPos, 0, 0)
    tbCaseId.ControlSource = "court_case_id"
    tbCaseId.Visible = False

    ' entry_year
    CreateControl frm.Name, acLabel, acDetail, , "entry_year:", 300, topPos, 1400, 300
    Dim tbY As Control
    Set tbY = CreateControl(frm.Name, acTextBox, acDetail, , , 1800, topPos, 900, 300)
    tbY.ControlSource = "entry_year"

    ' season_id lookup
    AddCombo frm.Name, "season_id", T_SEASON, "season_id", "season_name", 3500, topPos, "season"

    topPos = topPos + 480

    ' land_use_id lookup
    AddCombo frm.Name, "land_use_id", T_LAND_USE, "land_use_id", "description", 2800, topPos, "land_use"

    topPos = topPos + 480

    ' original_placename
    CreateControl frm.Name, acLabel, acDetail, , "original_placename:", 300, topPos, 2400, 300
    Dim tbOP As Control
    Set tbOP = CreateControl(frm.Name, acTextBox, acDetail, , , 2800, topPos, 5200, 300)
    tbOP.ControlSource = "original_placename"

    topPos = topPos + 480

    ' placename_id: WARNING: placename can be huge. Keep as textbox by default.
    ' If your placename table is small enough, switch to combo:
    ' AddCombo frm.Name, "placename_id", T_PLACENAME, "placename_id", "placename", 2800, topPos, "placename"
    CreateControl frm.Name, acLabel, acDetail, , "placename_id:", 300, topPos, 2400, 300
    Dim tbPID As Control
    Set tbPID = CreateControl(frm.Name, acTextBox, acDetail, , , 2800, topPos, 1200, 300)
    tbPID.ControlSource = "placename_id"

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

' court_case_entry as main form + person_entry subform (useful for editing participants)
Private Sub CreateCourtCaseEntry_WithPersonEntry()
    CreatePersonEntrySubform

    DeleteIfExists acForm, "frm_CourtCaseEntry"

    DoCmd.CreateForm
    Dim frm As Form: Set frm = Screen.ActiveForm
    frm.RecordSource = T_COURT_CASE_ENTRY
    frm.Caption = "Court Case Entry"

    Dim topPos As Long: topPos = 600

    AddCombo frm.Name, "court_case_id", T_COURT_CASE, "court_case_id", "court_case_id", 2800, topPos, "court_case_id"
    topPos = topPos + 480

    ' entry_year
    CreateControl frm.Name, acLabel, acDetail, , "entry_year:", 300, topPos, 2400, 360
    Dim tbY As Control
    Set tbY = CreateControl(frm.Name, acTextBox, acDetail, , , 2800, topPos, 1000, 360)
    tbY.ControlSource = "entry_year"
    topPos = topPos + 480

    AddCombo frm.Name, "season_id", T_SEASON, "season_id", "season_name", 2800, topPos, "season"
    topPos = topPos + 480

    AddCombo frm.Name, "land_use_id", T_LAND_USE, "land_use_id", "description", 2800, topPos, "land_use"
    topPos = topPos + 480

    ' original_placename
    CreateControl frm.Name, acLabel, acDetail, , "original_placename:", 300, topPos, 2400, 360
    Dim tbOP As Control
    Set tbOP = CreateControl(frm.Name, acTextBox, acDetail, , , 2800, topPos, 5200, 360)
    tbOP.ControlSource = "original_placename"
    topPos = topPos + 480

    ' placename_id as textbox (see note above)
    CreateControl frm.Name, acLabel, acDetail, , "placename_id:", 300, topPos, 2400, 360
    Dim tbPID As Control
    Set tbPID = CreateControl(frm.Name, acTextBox, acDetail, , , 2800, topPos, 1200, 360)
    tbPID.ControlSource = "placename_id"
    topPos = topPos + 480

    ' curated_text
    CreateControl frm.Name, acLabel, acDetail, , "curated_text:", 300, topPos, 2400, 360
    Dim tbCT As Control
    Set tbCT = CreateControl(frm.Name, acTextBox, acDetail, , , 2800, topPos, 5200, 1200)
    tbCT.ControlSource = "curated_text"
    tbCT.EnterKeyBehavior = True
    tbCT.ScrollBars = 2
    topPos = topPos + 1500

    ' Subform for participants/person_entry
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

    ' Link field hidden (court_case_entry_id)
    Dim tbLink As Control
    Set tbLink = CreateControl(frm.Name, acTextBox, acDetail, , , 100, topPos, 0, 0)
    tbLink.ControlSource = "court_case_entry_id"
    tbLink.Visible = False

    ' person_id -> person.full_name
    AddCombo frm.Name, "person_id", T_PERSON, "person_id", "full_name", 2800, topPos, "person"
    topPos = topPos + 480

    ' community_id -> community.community_name
    AddCombo frm.Name, "community_id", T_COMMUNITY, "community_id", "community_name", 2800, topPos, "community"
    topPos = topPos + 480

    ' land_rights_status_id -> land_rights_status.land_rights_status
    AddCombo frm.Name, "land_rights_status_id", T_LAND_RIGHTS_STATUS, "land_rights_status_id", "land_rights_status", 2800, topPos, "land_rights_status"
    topPos = topPos + 480

    ' role_id -> role.role_name
    AddCombo frm.Name, "role_id", T_ROLE, "role_id", "role_name", 2800, topPos, "role"
    topPos = topPos + 480

    ' curated_text
    CreateControl frm.Name, acLabel, acDetail, , "curated_text:", 300, topPos, 2400, 300
    Dim tbCT As Control
    Set tbCT = CreateControl(frm.Name, acTextBox, acDetail, , , 2800, topPos, 5200, 900)
    tbCT.ControlSource = "curated_text"
    tbCT.EnterKeyBehavior = True
    tbCT.ScrollBars = 2

    DoCmd.Save acForm, "sfrm_PersonEntry"
    DoCmd.Close acForm, frm.Name, acSaveYes
End Sub

' person main form + subforms: person_entry, person_relationship (as person_1)
Private Sub CreatePersonForms()
    CreatePersonEntrySubform_ForPerson
    CreatePersonRelationshipSubform

    DeleteIfExists acForm, "frm_Person"

    DoCmd.CreateForm
    Dim frm As Form: Set frm = Screen.ActiveForm
    frm.RecordSource = T_PERSON
    frm.Caption = "Person"

    Dim topPos As Long: topPos = 600

    ' given_name, patronymic, surname
    AddText frm.Name, "given_name", topPos: topPos = topPos + 480
    AddText frm.Name, "patronymic", topPos: topPos = topPos + 480
    AddText frm.Name, "surname", topPos: topPos = topPos + 480

    AddText frm.Name, "birth_year", topPos: topPos = topPos + 480
    AddText frm.Name, "death_year", topPos: topPos = topPos + 480
    AddText frm.Name, "community_name", topPos: topPos = topPos + 480

    ' note
    CreateControl frm.Name, acLabel, acDetail, , "note:", 300, topPos, 2400, 360
    Dim tbN As Control
    Set tbN = CreateControl(frm.Name, acTextBox, acDetail, , , 2800, topPos, 5200, 900)
    tbN.ControlSource = "note"
    tbN.EnterKeyBehavior = True
    tbN.ScrollBars = 2
    topPos = topPos + 1200

    ' full_name (generated) - show as locked textbox
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

    DoCmd.Save acForm, "frm_Person"
    DoCmd.Close acForm, frm.Name, acSaveYes
End Sub

Private Sub AddText(ByVal formName As String, ByVal fieldName As String, ByVal topPos As Long)
    CreateControl formName, acLabel, acDetail, , fieldName & ":", 300, topPos, 2400, 360
    Dim tb As Control
    Set tb = CreateControl(formName, acTextBox, acDetail, , , 2800, topPos, 5200, 360)
    tb.ControlSource = fieldName
End Sub

Private Sub CreatePersonEntrySubform_ForPerson()
    ' re-use person_entry but linked by person_id (not court_case_entry_id)
    DeleteIfExists acForm, "sfrm_PersonEntry_ByPerson"

    DoCmd.CreateForm
    Dim frm As Form: Set frm = Screen.ActiveForm
    frm.RecordSource = T_PERSON_ENTRY
    frm.Caption = "Person Entries"
    frm.DefaultView = acContinuousForms

    Dim topPos As Long: topPos = 300

    ' Hide person_id (link field)
    Dim tbPID As Control
    Set tbPID = CreateControl(frm.Name, acTextBox, acDetail, , , 100, topPos, 0, 0)
    tbPID.ControlSource = "person_id"
    tbPID.Visible = False

    ' court_case_entry_id (show as number; you could make a combo if you prefer)
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

    ' Hide person_1_id (link field)
    Dim tbP1 As Control
    Set tbP1 = CreateControl(frm.Name, acTextBox, acDetail, , , 100, topPos, 0, 0)
    tbP1.ControlSource = "person_1_id"
    tbP1.Visible = False

    ' person_2_id -> person.full_name
    AddCombo frm.Name, "person_2_id", T_PERSON, "person_id", "full_name", 2800, topPos, "related person"
    topPos = topPos + 480

    ' relationship_type_id -> relationship_type.relationship_type_name
    AddCombo frm.Name, "relationship_type_id", T_RELATIONSHIP_TYPE, "relationship_type_id", "relationship_type_name", 2800, topPos, "relationship_type"
    topPos = topPos + 480

    ' description
    CreateControl frm.Name, acLabel, acDetail, , "description:", 300, topPos, 2400, 300
    Dim tbD As Control
    Set tbD = CreateControl(frm.Name, acTextBox, acDetail, , , 2800, topPos, 5200, 600)
    tbD.ControlSource = "description"
    tbD.EnterKeyBehavior = True
    tbD.ScrollBars = 2

    DoCmd.Save acForm, "sfrm_PersonRelationship"
    DoCmd.Close acForm, frm.Name, acSaveYes
End Sub

' ruling (1:1 with court_case via unique court_case_id) + person_outcome subform
Private Sub CreateRulingForms()
    CreatePersonOutcomeSubform

    DeleteIfExists acForm, "frm_Ruling"

    DoCmd.CreateForm
    Dim frm As Form: Set frm = Screen.ActiveForm
    frm.RecordSource = T_RULING
    frm.Caption = "Ruling"

    Dim topPos As Long: topPos = 600

    AddCombo frm.Name, "court_case_id", T_COURT_CASE, "court_case_id", "court_case_id", 2800, topPos, "court_case_id"
    topPos = topPos + 480

    AddCombo frm.Name, "ruling_type_id", T_RULING_TYPE, "ruling_type_id", "ruling_type", 2800, topPos, "ruling_type"
    topPos = topPos + 480

    AddCombo frm.Name, "legal_source_id", T_LEGAL_SOURCE, "legal_source_id", "legal_source_name", 2800, topPos, "legal_source"
    topPos = topPos + 480

    AddText frm.Name, "ruling_year", topPos
    topPos = topPos + 480

    ' description
    CreateControl frm.Name, acLabel, acDetail, , "description:", 300, topPos, 2400, 360
    Dim tbDesc As Control
    Set tbDesc = CreateControl(frm.Name, acTextBox, acDetail, , , 2800, topPos, 5200, 900)
    tbDesc.ControlSource = "description"
    tbDesc.EnterKeyBehavior = True
    tbDesc.ScrollBars = 2
    topPos = topPos + 1200

    AddSubform frm.Name, "sfrm_PersonOutcome", "ruling_id", "ruling_id", 300, topPos, 8000, 2400, "Outcomes for persons"

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

    ' Hide ruling_id (link field)
    Dim tbR As Control
    Set tbR = CreateControl(frm.Name, acTextBox, acDetail, , , 100, topPos, 0, 0)
    tbR.ControlSource = "ruling_id"
    tbR.Visible = False

    ' person_id -> person.full_name
    AddCombo frm.Name, "person_id", T_PERSON, "person_id", "full_name", 2800, topPos, "person"
    topPos = topPos + 480

    ' outcome_type_id -> outcome_type_name
    AddCombo frm.Name, "outcome_type_id", T_OUTCOME_TYPE, "outcome_type_id", "outcome_type_name", 2800, topPos, "outcome_type"
    topPos = topPos + 480

    ' description
    CreateControl frm.Name, acLabel, acDetail, , "description:", 300, topPos, 2400, 300
    Dim tbD As Control
    Set tbD = CreateControl(frm.Name, acTextBox, acDetail, , , 2800, topPos, 5200, 600)
    tbD.ControlSource = "description"
    tbD.EnterKeyBehavior = True
    tbD.ScrollBars = 2

    DoCmd.Save acForm, "sfrm_PersonOutcome"
    DoCmd.Close acForm, frm.Name, acSaveYes
End Sub

Private Sub CreatePlacenamePickerForm()
    DeleteIfExists acForm, "frm_PlacenamePicker"

    DoCmd.CreateForm
    Dim frm As Form: Set frm = Screen.ActiveForm
    frm.Caption = "Pick placename"
    frm.DefaultView = acNormal

    ' Hidden textbox that stores table name used by picker (lets you change it centrally)
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
    ' AfterUpdate is reliable; OnChange is “live” but can be noisy.
    tbSearch.AfterUpdate = "=PlacenamePicker_Search()"

    ' Results listbox
    Dim lb As Control
    Set lb = CreateControl(frm.Name, acListBox, acDetail, , , 300, 750, 7500, 3000)
    lb.Name = "lstResults"
    lb.ColumnCount = 5
    lb.BoundColumn = 1
    lb.ColumnWidths = "0;6;4;2;2" ' hide id
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
