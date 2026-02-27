'========================================
' PLACENAME PICKER - public API
'========================================

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

    ' Open dialog picker (modal)
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
    ' Bound to txtSearch AfterUpdate / OnChange (expression)
    On Error GoTo EH
    Dim f As Form
    Set f = Forms("frm_PlacenamePicker")

    Dim q As String
    q = Nz(f.Controls("txtSearch").Value, "")

    Dim sql As String
    sql = "SELECT placename_id, placename, parish_name, northing, easting " & _
          "FROM [" & f.Controls("txtTable").Value & "] "

    If Len(Trim$(q)) > 0 Then
        ' Escape single quotes
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
    ' Bound to cmdUse OnClick (expression)
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
