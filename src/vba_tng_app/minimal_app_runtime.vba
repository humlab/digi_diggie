Option Compare Database
Option Explicit

' Module-level variable used to pass person_id back from frmPersonSearch
' when the caller uses target="_return" (e.g., the Add Person flow).
Private g_SelectedPersonId As Variant

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
        ' New/unsaved record - disable buttons until saved
        Forms!frmCourtCase!cmdCreateRuling.Enabled = False
        Forms!frmCourtCase!cmdSaveCase.Enabled = False
        UpdateOpenEntryDetailButtonState
     Exit Function
    End If

    ' Check If ruling exists For this Case
    Set rs = db.OpenRecordset("Select ruling_id FROM ruling WHERE court_case_id = " & courtCaseId, dbOpenSnapshot)
    hasRuling = Not rs.EOF
    rs.Close

    ' Enable button only when no ruling exists
    Forms!frmCourtCase!cmdCreateRuling.Enabled = Not hasRuling

    ' Disable Save button on navigation (record is clean after navigating)
    Forms!frmCourtCase!cmdSaveCase.Enabled = False

    ' Keep entry-detail action in sync with the linked entry subform selection.
    UpdateOpenEntryDetailButtonState

    ' Requery the ruling subform To ensure it's in sync
    Forms!frmCourtCase!sfrmRuling.Form.Requery
End Function

'------------------------------------------------------------------------------
' frmCourtCase: OnDirty event - Enable Save button when form has unsaved changes
'------------------------------------------------------------------------------
Public Function frmCourtCase_OnDirty()
    On Error Resume Next
    Forms!frmCourtCase!cmdSaveCase.Enabled = True
End Function

'------------------------------------------------------------------------------
' frmCourtCase: Save Case button
'------------------------------------------------------------------------------
Public Function frmCourtCase_cmdSaveCase_Click()
    On Error Goto ErrHandler
        If Forms!frmCourtCase!sfrmRuling.Form.Dirty Then Forms!frmCourtCase!sfrmRuling.Form.Dirty = False
        If Forms!frmCourtCase.Dirty Then RunCommand acCmdSaveRecord
        Forms!frmCourtCase!cmdSaveCase.Enabled = False
        frmCourtCase_OnCurrent
     Exit Function
 ErrHandler:
        MsgBox "Error in frmCourtCase_cmdSaveCase_Click: " & Err.Description, vbCritical
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
        Dim frmEntries As Form
        Dim entryId As Variant

        Set frmEntries = Forms!frmCourtCase!sfrmCourtCaseEntries.Form

        If frmEntries.NewRecord Or IsNull(frmEntries!court_case_entry_id) Then
            MsgBox "Please select a saved entry first.", vbInformation
            Exit Function
        End If

        ' Commit pending edits before opening modal detail form.
        ' This avoids intermittent UI stalls with linked ODBC tables.
        If frmEntries.Dirty Then frmEntries.Dirty = False
        If Forms!frmCourtCase.Dirty Then Forms!frmCourtCase.Dirty = False

        entryId = frmEntries!court_case_entry_id
        OpenCourtCaseEntryDetail CLng(entryId)

     Exit Function
 ErrHandler:
        MsgBox "Error in frmCourtCase_cmdOpenEntryDetail_Click: " & Err.Description, vbCritical
End Function

'------------------------------------------------------------------------------
' sfrmCourtCaseEntries: Entry Detail button
'------------------------------------------------------------------------------
Public Function sfrmCourtCaseEntries_cmdEntryDetail_Click()
    On Error Goto ErrHandler
        Dim frmEntries As Form
        Dim entryId As Variant

        Set frmEntries = Screen.ActiveControl.Parent
        If frmEntries.NewRecord Or IsNull(frmEntries!court_case_entry_id) Then
            MsgBox "Please select a saved entry first.", vbInformation
            Exit Function
        End If

        If frmEntries.Dirty Then frmEntries.Dirty = False
        If Forms!frmCourtCase.Dirty Then Forms!frmCourtCase.Dirty = False

        entryId = frmEntries!court_case_entry_id
        OpenCourtCaseEntryDetail CLng(entryId)

     Exit Function
 ErrHandler:
        MsgBox "Error in sfrmCourtCaseEntries_cmdEntryDetail_Click: " & Err.Description, vbCritical
End Function

'------------------------------------------------------------------------------
' sfrmCourtCaseEntries: OnCurrent event - Sync parent detail button state
'------------------------------------------------------------------------------
Public Function sfrmCourtCaseEntries_OnCurrent()
    On Error Resume Next
    UpdateOpenEntryDetailButtonState
End Function

'------------------------------------------------------------------------------
' frmCourtCaseEntryDetail: OnLoad event - Auto-populate placename display
'------------------------------------------------------------------------------
Public Function frmCourtCaseEntryDetail_OnLoad()
    On Error Resume Next
    EnsureCourtCaseEntryDetailWindow

    ' Auto-populate placename display field on form load
    UpdatePlacenameDisplay "frmCourtCaseEntryDetail"

    ' Limit person subform query to current entry to avoid loading large joined sets.
    ConfigurePersonEntrySubform "frmCourtCaseEntryDetail"
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
' frmCourtCaseEntryDetail: Add person entry button
'------------------------------------------------------------------------------
Public Function frmCourtCaseEntryDetail_cmdAddPersonEntry_Click()
    On Error Goto ErrHandler
        Dim entryId As Variant
        Dim db As DAO.Database
        Dim rs As DAO.Recordset

        ' Ensure the current entry is saved before adding a person to it
        entryId = Forms!frmCourtCaseEntryDetail!court_case_entry_id
        If IsNull(entryId) Then
            MsgBox "Please save the court case entry first.", vbInformation
            Exit Function
        End If

        ' Open person search dialog; result comes back via g_SelectedPersonId
        g_SelectedPersonId = Null
        DoCmd.OpenForm "frmPersonSearch", , , , , acDialog, _
            "caller=frmCourtCaseEntryDetail;target=_return"

        ' If the user cancelled without selecting, do nothing
        If IsNull(g_SelectedPersonId) Then Exit Function

        ' Insert a new person_entry row via DAO so the JOIN-based person_name
        ' resolves correctly after the subsequent requery.
        Set db = CurrentDb
        Set rs = db.OpenRecordset("SELECT * FROM person_entry WHERE 1=0", dbOpenDynaset, dbAppendOnly)
        rs.AddNew
        rs!court_case_entry_id = CLng(entryId)
        rs!person_id = CLng(g_SelectedPersonId)
        rs.Update
        rs.Close

        ' Refresh subform so the new row (with person_name from JOIN) appears
        ConfigurePersonEntrySubform "frmCourtCaseEntryDetail"

     Exit Function
ErrHandler:
        MsgBox "Error in frmCourtCaseEntryDetail_cmdAddPersonEntry_Click: " & Err.Description, vbCritical
End Function

'------------------------------------------------------------------------------
' frmCourtCaseEntryDetail: Delete selected person entry button
'------------------------------------------------------------------------------
Public Function frmCourtCaseEntryDetail_cmdDeletePersonEntry_Click()
    On Error Goto ErrHandler
        Dim subFrm As Form
        Dim rs As DAO.Recordset
        Dim db As DAO.Database
        Dim personEntryId As Variant

        Set subFrm = Forms!frmCourtCaseEntryDetail!sfrmPersonEntryByEntry.Form

        Set rs = subFrm.RecordsetClone
        If rs.BOF And rs.EOF Then
            MsgBox "No person entry record is selected.", vbInformation
            rs.Close
            Exit Function
        End If
        rs.Close

        If subFrm.NewRecord Then
            MsgBox "No person entry record is selected.", vbInformation
            Exit Function
        End If

        If MsgBox("Delete selected person entry?", vbQuestion + vbYesNo, "Confirm Delete") = vbYes Then
            ' Read the PK from a clone synced to the current row, then delete via DAO.
            ' The subform uses a JOIN-based RecordSource so its recordset is non-deletable directly.
            Set rs = subFrm.RecordsetClone
            rs.Bookmark = subFrm.Bookmark
            personEntryId = rs!person_entry_id
            rs.Close

            Set db = CurrentDb
            db.Execute "DELETE FROM person_entry WHERE person_entry_id = " & CLng(personEntryId), dbFailOnError
            ConfigurePersonEntrySubform "frmCourtCaseEntryDetail"
        End If

     Exit Function
ErrHandler:
        MsgBox "Error in frmCourtCaseEntryDetail_cmdDeletePersonEntry_Click: " & Err.Description, vbCritical
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

        MsgBox "Ruling created.", vbInformation

        ' Switch to the ruling subform tab and move focus there,
        ' which also moves focus away from cmdCreateRuling (required before disabling it)
        Forms!frmCourtCase!tabMain.Value = 1
        Forms!frmCourtCase!sfrmRuling.Form!txtRulingYear.SetFocus
        Forms!frmCourtCase!cmdCreateRuling.Enabled = False

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

    ' Enable/disable Create Ruling button on parent form
    Forms!frmCourtCase!cmdCreateRuling.Enabled = Not hasRuling

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
        Dim rulingId As Variant
        Dim db As DAO.Database
        Dim rs As DAO.Recordset

        ' Get parent ruling_id
        rulingId = Forms!frmCourtCase!sfrmRuling.Form!ruling_id
        If IsNull(rulingId) Then
            MsgBox "Please create/save a ruling first.", vbInformation
            Exit Function
        End If

        ' Insert new person_outcome with default outcome_type_id (first one)
        Set db = CurrentDb
        Set rs = db.OpenRecordset("person_outcome", dbOpenDynaset)
        rs.AddNew
        rs!ruling_id = CLng(rulingId)
        rs!outcome_type_id = DMin("outcome_type_id", "outcome_type")
        rs.Update
        rs.Close

        ' Requery subform to show new record
        Forms!frmCourtCase!sfrmRuling.Form!sfrmPersonOutcomes.Form.Requery

        ' Set focus to the new record (the last one) in the person outcomes subform for better UX
        Forms!frmCourtCase!sfrmRuling.Form!sfrmPersonOutcomes.SetFocus
        DoCmd.GoToRecord , , acLast

     Exit Function
ErrHandler:
        MsgBox "Error in sfrmRuling_cmdAddOutcome_Click: " & Err.Description, vbCritical
End Function

'------------------------------------------------------------------------------
' sfrmRuling: Delete selected outcome button
'------------------------------------------------------------------------------
Public Function sfrmRuling_cmdDeleteOutcome_Click()
    On Error Goto ErrHandler
        Dim subFrm As Form
        Dim rs As DAO.Recordset
        Dim db As DAO.Database
        Dim personOutcomeId As Variant

        Set subFrm = Forms!frmCourtCase!sfrmRuling.Form!sfrmPersonOutcomes.Form

        Set rs = subFrm.RecordsetClone
        If rs.BOF And rs.EOF Then
            MsgBox "No person outcome record is selected.", vbInformation
            rs.Close
            Exit Function
        End If
        rs.Close

        If subFrm.NewRecord Then
            MsgBox "No person outcome record is selected.", vbInformation
            Exit Function
        End If

        If MsgBox("Delete selected person outcome?", vbQuestion + vbYesNo, "Confirm Delete") = vbYes Then
            ' Read the PK from a clone synced to the current row, then delete via DAO.
            ' The subform uses a JOIN-based RecordSource so its recordset is non-deletable directly.
            Set rs = subFrm.RecordsetClone
            rs.Bookmark = subFrm.Bookmark
            personOutcomeId = rs!person_outcome_id
            rs.Close

            Set db = CurrentDb
            db.Execute "DELETE FROM person_outcome WHERE person_outcome_id = " & CLng(personOutcomeId), dbFailOnError
            subFrm.Requery
        End If

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
        "SELECT TOP 50 placename_id, placename, parish_name AS parish_display, serial_number " & _
        "FROM placename " & _
        "ORDER BY placename"
End Function

'------------------------------------------------------------------------------
' frmPlacenameSearch: Search button
'------------------------------------------------------------------------------
Public Function frmPlacenameSearch_cmdSearch_Click()
    On Error Goto ErrHandler
        Dim searchTerm As String
        Dim sSql As String
        Dim db As DAO.Database

        Set db = CurrentDb

        searchTerm = "*" & Nz(Forms!frmPlacenameSearch!txtSearch, "") & "*"
        sSql = "SELECT " & _
                "placename_id, " & _
                "placename, " & _
                "parish_name AS parish_display, " & _
                "serial_number " & _
                "FROM placename " & _
                "WHERE (placename LIKE '" & Replace(searchTerm, "'", "''") & "') " & _
                "   OR (parish_name LIKE '" & Replace(searchTerm, "'", "''") & "') " & _
                "ORDER BY placename;"

        Forms!frmPlacenameSearch!lstResults.RowSource = sSql
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
        Dim sSql As String
        Dim db As DAO.Database

        Set db = CurrentDb

        searchTerm = "*" & Nz(Forms!frmPersonSearch!txtSearch, "") & "*"
        sSql = "SELECT " & _
                    "person.person_id, " & _
                    "person.full_name, " & _
                    "person.birth_year, " & _
                    "person.community_name " & _
               "FROM person " & _
               "WHERE (((person.[full_name]) LIKE '" & Replace(searchTerm, "'", "''") & "')) " & _
               "ORDER BY person.full_name;"

        Forms!frmPersonSearch!lstResults.RowSource = sSql
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
            If targetControl = "_return" Then
                ' Return via module-level variable (used by Add Person flow)
                g_SelectedPersonId = selectedId
            Else
                Forms(callerForm).Controls(targetControl).Value = selectedId
            End If
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
    UpdateOpenEntryDetailButtonState
End Function

'------------------------------------------------------------------------------
' sfrmPersonEntryByEntry: OnLoad event - Auto-size columns to fit content
'------------------------------------------------------------------------------
Public Function sfrmPersonEntryByEntry_OnLoad()
    On Error Resume Next
    ' Use fixed widths to avoid expensive auto-fit scans on linked ODBC tables.
    Forms!frmCourtCaseEntryDetail!sfrmPersonEntryByEntry.Form!txtPersonId.ColumnWidth = 0
    Forms!frmCourtCaseEntryDetail!sfrmPersonEntryByEntry.Form!txtPersonName.ColumnWidth = 2300
    Forms!frmCourtCaseEntryDetail!sfrmPersonEntryByEntry.Form!cmdPickPerson.ColumnWidth = 900
    Forms!frmCourtCaseEntryDetail!sfrmPersonEntryByEntry.Form!cboCommunityId.ColumnWidth = 1800
    Forms!frmCourtCaseEntryDetail!sfrmPersonEntryByEntry.Form!cboLandRightsStatusId.ColumnWidth = 1800
    Forms!frmCourtCaseEntryDetail!sfrmPersonEntryByEntry.Form!cboRoleId.ColumnWidth = 1500
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
' Enable the parent entry-detail button only when a saved subform row is current
'------------------------------------------------------------------------------
Private Sub UpdateOpenEntryDetailButtonState()
    On Error GoTo ErrHandler
    Dim frmEntries As Form
    Dim hasSelectedEntry As Boolean

    Forms!frmCourtCase!cmdOpenEntryDetail.Enabled = False

    If IsNull(Forms!frmCourtCase!court_case_id) Then Exit Sub

    Set frmEntries = Forms!frmCourtCase!sfrmCourtCaseEntries.Form
    hasSelectedEntry = Not frmEntries.NewRecord And Not IsNull(frmEntries!court_case_entry_id)

    Forms!frmCourtCase!cmdOpenEntryDetail.Enabled = hasSelectedEntry
    Exit Sub

ErrHandler:
    Forms!frmCourtCase!cmdOpenEntryDetail.Enabled = False
End Sub

'------------------------------------------------------------------------------
' Open entry detail safely for a known entry ID
'------------------------------------------------------------------------------
Private Sub OpenCourtCaseEntryDetail(entryId As Long)
    On Error GoTo ErrHandler
    DoCmd.Hourglass True
    ' Form itself is already Popup+Modal; avoid acDialog to reduce UI deadlock risk.
    DoCmd.OpenForm "frmCourtCaseEntryDetail", , , "court_case_entry_id=" & entryId, , acWindowNormal
    EnsureCourtCaseEntryDetailWindow

Cleanup:
    DoCmd.Hourglass False
    Exit Sub

ErrHandler:
    DoCmd.Hourglass False
    Err.Raise Err.Number, "OpenCourtCaseEntryDetail", Err.Description
End Sub

'------------------------------------------------------------------------------
' Configure person-entry subform to load only rows for the current entry
'------------------------------------------------------------------------------
Private Sub ConfigurePersonEntrySubform(parentFormName As String)
    On Error GoTo ErrHandler
    Dim entryId As Variant
    Dim sqlText As String

    entryId = Forms(parentFormName)!court_case_entry_id
    If IsNull(entryId) Then Exit Sub

    sqlText = "SELECT pe.person_entry_id, pe.court_case_entry_id, pe.person_id, " & _
        "p.full_name AS person_name, pe.community_id, pe.land_rights_status_id, pe.role_id " & _
        "FROM person_entry AS pe " & _
        "LEFT JOIN person AS p ON pe.person_id = p.person_id " & _
        "WHERE pe.court_case_entry_id = " & CLng(entryId)

    Forms(parentFormName)!sfrmPersonEntryByEntry.Form.RecordSource = sqlText
    Forms(parentFormName)!sfrmPersonEntryByEntry.Form.Requery
    Exit Sub

ErrHandler:
    Debug.Print "Error in ConfigurePersonEntrySubform: " & Err.Description
End Sub

'------------------------------------------------------------------------------
' Ensure detail popup opens with a usable size/position
'------------------------------------------------------------------------------
Private Sub EnsureCourtCaseEntryDetailWindow()
    On Error GoTo ErrHandler
    Dim frm As Form

    Set frm = Forms!frmCourtCaseEntryDetail

    ' Twips (1 cm ≈ 567 twips): width ~17.6 cm, height ~13.8 cm
    If frm.WindowWidth < 9000 Or frm.InsideWidth < 8500 Then
        DoCmd.MoveSize 800, 500, 10000, 7800
    End If

    Exit Sub

ErrHandler:
    Debug.Print "Error in EnsureCourtCaseEntryDetailWindow: " & Err.Description
End Sub

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
