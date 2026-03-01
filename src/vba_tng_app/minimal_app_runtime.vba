Option Compare Database
Option Explicit

'==============================================================================
' modMinimalAppRuntime
' Purpose: Runtime event handlers For DigiDiggie TNG minimal forms
' Note: This module contains ONLY runtime code (event handlers).
'       Form generation code is in modMinimalAppGenerator module.
'==============================================================================

'==============================================================================
' EVENT HANDLERS - Called by form controls
'==============================================================================

'------------------------------------------------------------------------------
' frmCourtCase: OnCurrent event - Update button visibility
'------------------------------------------------------------------------------
Public Function frmCourtCase_OnCurrent()
    On Error Resume Next
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim courtCaseId As Variant
    Dim hasRuling As Boolean

    Set db = CurrentDb

    ' Get current court Case ID
    courtCaseId = Forms!frmCourtCase!court_case_id

    If IsNull(courtCaseId) Then
        ' New/unsaved record - hide button Until saved
        Forms!frmCourtCase!cmdCreateRuling.Visible = False
     Exit Function
    End If

    ' Check If ruling exists For this Case
    Set rs = db.OpenRecordset("Select ruling_id FROM ruling WHERE court_case_id = " & courtCaseId, dbOpenSnapshot)
    hasRuling = Not rs.EOF
    rs.Close

    ' Show button only when no ruling exists
    Forms!frmCourtCase!cmdCreateRuling.Visible = Not hasRuling

    ' Requery the ruling subform To ensure it's in sync
    Forms!frmCourtCase!sfrmRuling.Form.Requery
End Function

'------------------------------------------------------------------------------
' frmCourtCase: New Case button
'------------------------------------------------------------------------------
Public Function frmCourtCase_cmdNewCase_Click()
    On Error Goto ErrHandler
        DoCmd.GoToRecord , , acNewRec
     Exit Function
 ErrHandler:
        MsgBox "Error in frmCourtCase_cmdNewCase_Click: " & Err.Description, vbCritical
End Function

'------------------------------------------------------------------------------
' frmCourtCase: Previous button
'------------------------------------------------------------------------------
Public Function frmCourtCase_cmdPrevious_Click()
    On Error Goto ErrHandler
        DoCmd.GoToRecord , , acPrevious
     Exit Function
 ErrHandler:
        ' Silently ignore errors (e.g., already at first record)
        On Error Resume Next
End Function

'------------------------------------------------------------------------------
' frmCourtCase: Next button
'------------------------------------------------------------------------------
Public Function frmCourtCase_cmdNext_Click()
    On Error Goto ErrHandler
        DoCmd.GoToRecord , , acNext
     Exit Function
 ErrHandler:
        ' Silently ignore errors (e.g., already at last record)
        On Error Resume Next
End Function

'------------------------------------------------------------------------------
' frmCourtCase: Open Entry Detail button
'------------------------------------------------------------------------------
Public Function frmCourtCase_cmdOpenEntryDetail_Click()
    On Error Goto ErrHandler
        Dim entryId As Variant

        ' Get selected entry from subform
        If Not IsNull(Forms!frmCourtCase!sfrmCourtCaseEntries.Form!court_case_entry_id) Then
            entryId = Forms!frmCourtCase!sfrmCourtCaseEntries.Form!court_case_entry_id
            DoCmd.OpenForm "frmCourtCaseEntryDetail", , , "court_case_entry_id=" & entryId, , acDialog
        Else
            MsgBox "Please Select an entry first.", vbInformation
        End If

     Exit Function
 ErrHandler:
        MsgBox "Error in frmCourtCase_cmdOpenEntryDetail_Click: " & Err.Description, vbCritical
End Function

'------------------------------------------------------------------------------
' sfrmCourtCaseEntries: Entry Detail button
'------------------------------------------------------------------------------
Public Function sfrmCourtCaseEntries_cmdEntryDetail_Click()
    On Error Goto ErrHandler
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
' frmCourtCaseEntryDetail: OnLoad event - Auto-populate placename display
'------------------------------------------------------------------------------
Public Function frmCourtCaseEntryDetail_OnLoad()
    On Error Resume Next
    ' Auto-populate placename display field on form load
    UpdatePlacenameDisplay "frmCourtCaseEntryDetail"
End Function

'------------------------------------------------------------------------------
' frmCourtCaseEntryDetail: Pick Placename button
'------------------------------------------------------------------------------
Public Function frmCourtCaseEntryDetail_cmdPickPlacename_Click()
    On Error Goto ErrHandler
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
    On Error Goto ErrHandler
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
    On Error Goto ErrHandler
        DoCmd.OpenForm "frmPersonSearch", , , , , acDialog, _
        "caller=sfrmPersonOutcomes;target=cboPersonId"
     Exit Function
 ErrHandler:
        MsgBox "Error in sfrmPersonOutcomes_cmdPickPerson_Click: " & Err.Description, vbCritical
End Function

'------------------------------------------------------------------------------
' frmCourtCase: Create Ruling button
'------------------------------------------------------------------------------
Public Function frmCourtCase_cmdCreateRuling_Click()
    On Error Goto ErrHandler
        Dim db As DAO.Database
        Dim rs As DAO.Recordset
        Dim courtCaseId As Variant
        Dim defaultRulingType As Variant

        Set db = CurrentDb

        ' Get parent court_case_id
        courtCaseId = Forms!frmCourtCase!court_case_id
        If IsNull(courtCaseId) Then
            MsgBox "Please save the court Case first.", vbInformation
         Exit Function
        End If

        ' Check If ruling already exists
        Set rs = db.OpenRecordset("Select ruling_id FROM ruling WHERE court_case_id = " & courtCaseId)
        If Not rs.EOF Then
            MsgBox "A ruling already exists For this Case.", vbInformation
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

        ' Insert New ruling
        Set rs = db.OpenRecordset("ruling", dbOpenDynaset)
        rs.AddNew
        rs!court_case_id = courtCaseId
        rs!ruling_type_id = defaultRulingType
        rs.Update
        rs.Close

        ' Requery the ruling form
        Forms!frmCourtCase!sfrmRuling.Form.Requery

        ' Hide the Create Ruling button now that a ruling exists
        Forms!frmCourtCase!cmdCreateRuling.Visible = False

        MsgBox "Ruling created.", vbInformation

     Exit Function
 ErrHandler:
        MsgBox "Error in frmCourtCase_cmdCreateRuling_Click: " & Err.Description, vbCritical
End Function

'------------------------------------------------------------------------------
' sfrmRuling: OnCurrent event - Handle dynamic UI
'------------------------------------------------------------------------------
Public Function sfrmRuling_OnCurrent()
    On Error Resume Next
    Dim hasRuling As Boolean

    ' Check If ruling record exists (recordset has data)
    hasRuling = Not (Forms!frmCourtCase!sfrmRuling.Form.Recordset.EOF And _
    Forms!frmCourtCase!sfrmRuling.Form.Recordset.BOF)

    ' Show/hide Create Ruling button on parent form
    Forms!frmCourtCase!cmdCreateRuling.Visible = Not hasRuling

    ' Show/hide "No ruling" Label in subform
    Forms!frmCourtCase!sfrmRuling.Form!lblNoRuling.Visible = Not hasRuling

    ' Show/hide data entry controls in subform
    Forms!frmCourtCase!sfrmRuling.Form!txtRulingYear.Visible = hasRuling
    Forms!frmCourtCase!sfrmRuling.Form!lblRulingYear.Visible = hasRuling
    Forms!frmCourtCase!sfrmRuling.Form!cboRulingTypeId.Visible = hasRuling
    Forms!frmCourtCase!sfrmRuling.Form!lblRulingTypeId.Visible = hasRuling
    Forms!frmCourtCase!sfrmRuling.Form!cboLegalSourceId.Visible = hasRuling
    Forms!frmCourtCase!sfrmRuling.Form!lblLegalSourceId.Visible = hasRuling
    Forms!frmCourtCase!sfrmRuling.Form!txtDescription.Visible = hasRuling
    Forms!frmCourtCase!sfrmRuling.Form!lblDescription.Visible = hasRuling
    Forms!frmCourtCase!sfrmRuling.Form!sfrmPersonOutcomes.Visible = hasRuling
    Forms!frmCourtCase!sfrmRuling.Form!lblPersonOutcomes.Visible = hasRuling
    Forms!frmCourtCase!sfrmRuling.Form!cmdAddOutcome.Visible = hasRuling
    Forms!frmCourtCase!sfrmRuling.Form!cmdDeleteOutcome.Visible = hasRuling

    If hasRuling Then
        Forms!frmCourtCase!sfrmRuling.Form!sfrmPersonOutcomes.Form!cboPersonId.Requery
    End If
End Function

'------------------------------------------------------------------------------
' sfrmRuling: Add outcome button
'------------------------------------------------------------------------------
Public Function sfrmRuling_cmdAddOutcome_Click()
    On Error Goto ErrHandler
        If IsNull(Forms!frmCourtCase!sfrmRuling.Form!ruling_id) Then
            MsgBox "Please create/save a ruling first.", vbInformation
            Exit Function
        End If

        Forms!frmCourtCase!sfrmRuling.Form!sfrmPersonOutcomes.SetFocus
        DoCmd.GoToRecord , , acNewRec

     Exit Function
ErrHandler:
        MsgBox "Error in sfrmRuling_cmdAddOutcome_Click: " & Err.Description, vbCritical
End Function

'------------------------------------------------------------------------------
' sfrmRuling: Delete selected outcome button
'------------------------------------------------------------------------------
Public Function sfrmRuling_cmdDeleteOutcome_Click()
    On Error Goto ErrHandler
        Dim rs As DAO.Recordset

        With Forms!frmCourtCase!sfrmRuling.Form!sfrmPersonOutcomes.Form
            Set rs = .RecordsetClone
            If rs.BOF And rs.EOF Then
                MsgBox "No person outcome record is selected.", vbInformation
                rs.Close
                Exit Function
            End If
            rs.Close

            If .NewRecord Then
                MsgBox "No person outcome record is selected.", vbInformation
                Exit Function
            End If

            If MsgBox("Delete selected person outcome?", vbQuestion + vbYesNo, "Confirm Delete") = vbYes Then
                .SetFocus
                DoCmd.RunCommand acCmdDeleteRecord
            End If
        End With

     Exit Function
ErrHandler:
        MsgBox "Error in sfrmRuling_cmdDeleteOutcome_Click: " & Err.Description, vbCritical
End Function

'------------------------------------------------------------------------------
' frmPlacenameSearch: OnLoad event - Show initial results
'------------------------------------------------------------------------------
Public Function frmPlacenameSearch_OnLoad()
    On Error Resume Next
    ' Show top 50 placenames by default
    Forms!frmPlacenameSearch!lstResults.RowSource = _
        "SELECT TOP 50 placename_id, placename, parish_name, serial_number " & _
        "FROM placename ORDER BY placename"
End Function

'------------------------------------------------------------------------------
' frmPlacenameSearch: Search button
'------------------------------------------------------------------------------
Public Function frmPlacenameSearch_cmdSearch_Click()
    On Error Goto ErrHandler
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
    On Error Goto ErrHandler
        Dim selectedId As Variant
        Dim callerForm As String
        Dim targetControl As String

        ' Get selected placename_id
        If IsNull(Forms!frmPlacenameSearch!lstResults) Then
            MsgBox "Please Select a placename.", vbInformation
         Exit Function
        End If

        selectedId = Forms!frmPlacenameSearch!lstResults

        ' Parse OpenArgs
        ParseOpenArgs Forms!frmPlacenameSearch.OpenArgs, callerForm, targetControl

        ' Write back To caller
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
' frmPersonSearch: OnLoad event - Show initial results
'------------------------------------------------------------------------------
Public Function frmPersonSearch_OnLoad()
    On Error Resume Next
    ' Show top 50 persons by default
    Forms!frmPersonSearch!lstResults.RowSource = _
        "SELECT TOP 50 person_id, full_name, birth_year, community_name " & _
        "FROM person ORDER BY full_name"
End Function

'------------------------------------------------------------------------------
' frmPersonSearch: Search button
'------------------------------------------------------------------------------
Public Function frmPersonSearch_cmdSearch_Click()
    On Error Goto ErrHandler
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
    On Error Goto ErrHandler
        Dim selectedId As Variant
        Dim callerForm As String
        Dim targetControl As String

        ' Get selected person_id
        If IsNull(Forms!frmPersonSearch!lstResults) Then
            MsgBox "Please Select a person.", vbInformation
         Exit Function
        End If

        selectedId = Forms!frmPersonSearch!lstResults

        ' Parse OpenArgs
        ParseOpenArgs Forms!frmPersonSearch.OpenArgs, callerForm, targetControl

        ' Write back To caller
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
    On Error Goto ErrHandler
        ' Open person form in data entry mode, pass through caller info
        DoCmd.OpenForm "frmPerson", , , , acFormAdd, acDialog, Forms!frmPersonSearch.OpenArgs

        ' Requery results after closing
        Forms!frmPersonSearch!lstResults.Requery

     Exit Function
 ErrHandler:
        MsgBox "Error in frmPersonSearch_cmdNewPerson_Click: " & Err.Description, vbCritical
End Function

'------------------------------------------------------------------------------
' frmPerson: OnClose event - Return new person_id to caller
'------------------------------------------------------------------------------
Public Function frmPerson_OnClose()
    On Error Goto ErrHandler
        Dim callerForm As String
        Dim targetControl As String
        Dim newPersonId As Variant
        
        ' Get the newly created person_id (if any)
        If Not IsNull(Forms!frmPerson!person_id) Then
            newPersonId = Forms!frmPerson!person_id
            
            ' Parse OpenArgs to see if we should write back directly
            ParseOpenArgs Forms!frmPerson.OpenArgs, callerForm, targetControl
            
            ' If caller info exists, write person_id directly and close search form
            If callerForm <> "" And targetControl <> "" Then
                ' Write back to original caller
                Forms(callerForm).Controls(targetControl).Value = newPersonId
                
                ' Close the person search form too (we're done)
                On Error Resume Next
                DoCmd.Close acForm, "frmPersonSearch"
                On Error Goto ErrHandler
            End If
        End If
        
     Exit Function
 ErrHandler:
        ' Silently ignore errors (form might already be closed)
        On Error Resume Next
End Function

'------------------------------------------------------------------------------
' frmPersonSearch: Cancel button
'------------------------------------------------------------------------------
Public Function frmPersonSearch_cmdCancel_Click()
    DoCmd.Close acForm, "frmPersonSearch"
End Function

'------------------------------------------------------------------------------
' sfrmCourtCaseEntries: OnLoad event - Auto-size columns to fit content
'------------------------------------------------------------------------------
Public Function sfrmCourtCaseEntries_OnLoad()
    On Error Resume Next
    ' Set column widths to -2 (auto-fit to content)
    Forms!frmCourtCase!sfrmCourtCaseEntries.Form!txtEntryYear.ColumnWidth = -2
    Forms!frmCourtCase!sfrmCourtCaseEntries.Form!cboSeasonId.ColumnWidth = -2
    Forms!frmCourtCase!sfrmCourtCaseEntries.Form!cboLandUseId.ColumnWidth = -2
    Forms!frmCourtCase!sfrmCourtCaseEntries.Form!txtPlacename.ColumnWidth = -2
    Forms!frmCourtCase!sfrmCourtCaseEntries.Form!txtOriginalPlacename.ColumnWidth = -2
    Forms!frmCourtCase!sfrmCourtCaseEntries.Form!cmdEntryDetail.ColumnWidth = -2
End Function

'------------------------------------------------------------------------------
' sfrmPersonEntryByEntry: OnLoad event - Auto-size columns to fit content
'------------------------------------------------------------------------------
Public Function sfrmPersonEntryByEntry_OnLoad()
    On Error Resume Next
    ' Set column widths to -2 (auto-fit to content)
    Forms!frmCourtCaseEntryDetail!sfrmPersonEntryByEntry.Form!txtPersonId.ColumnWidth = -2
    Forms!frmCourtCaseEntryDetail!sfrmPersonEntryByEntry.Form!cmdPickPerson.ColumnWidth = -2
    Forms!frmCourtCaseEntryDetail!sfrmPersonEntryByEntry.Form!cboCommunityId.ColumnWidth = -2
    Forms!frmCourtCaseEntryDetail!sfrmPersonEntryByEntry.Form!cboLandRightsStatusId.ColumnWidth = -2
    Forms!frmCourtCaseEntryDetail!sfrmPersonEntryByEntry.Form!cboRoleId.ColumnWidth = -2
End Function

'------------------------------------------------------------------------------
' sfrmPersonOutcomes: OnLoad event - Auto-size columns to fit content
'------------------------------------------------------------------------------
Public Function sfrmPersonOutcomes_OnLoad()
    On Error Resume Next
    ' Set column widths to -2 (auto-fit to content)
    Forms!frmCourtCase!sfrmRuling.Form!sfrmPersonOutcomes.Form!cboPersonId.Requery
    Forms!frmCourtCase!sfrmRuling.Form!sfrmPersonOutcomes.Form!cboPersonId.ColumnWidth = -2
    Forms!frmCourtCase!sfrmRuling.Form!sfrmPersonOutcomes.Form!cboOutcomeTypeId.ColumnWidth = -2
    Forms!frmCourtCase!sfrmRuling.Form!sfrmPersonOutcomes.Form!txtDescription.ColumnWidth = -2
End Function

'------------------------------------------------------------------------------
' sfrmPersonOutcomes: Requery person list when entering picker control
'------------------------------------------------------------------------------
Public Function sfrmPersonOutcomes_cboPersonId_OnEnter()
    On Error Resume Next
    Screen.ActiveControl.Requery
End Function

'==============================================================================
' HELPER FUNCTIONS
'==============================================================================

'------------------------------------------------------------------------------
' Parse OpenArgs: "caller=formName;target=controlName"
'------------------------------------------------------------------------------
Private Sub ParseOpenArgs(openArgs As Variant, callerForm As String, targetControl As String)
    On Error Goto ErrHandler
        Dim parts() As String
        Dim i As Integer

        callerForm = ""
        targetControl = ""

        If IsNull(openArgs) Or openArgs = "" Then Exit Sub

            parts = Split(openArgs, ";")
            For i = LBound(parts) To UBound(parts)
                If InStr(parts(i), "caller=") > 0 Then
                    callerForm = Replace(parts(i), "caller=", "")
                Elseif InStr(parts(i), "target=") > 0 Then
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
