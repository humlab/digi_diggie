Option Explicit
Option Compare Database

Public Sub ImportOrReplaceModule()
    Dim dbPath As String
    Dim modulePath As String
    Dim moduleName As String

    ' Get the folder of the current Access database
    dbPath = CurrentProject.Path
    moduleName = "rename"                         ' Module name inside Access (no .bas)
    modulePath = dbPath & "\" & moduleName & ".bas"

    On Error Resume Next
    ' Try to delete any existing module with this name
    DoCmd.DeleteObject acModule, moduleName
    On Error GoTo 0

    ' Import the .bas file from the same folder
    Application.LoadFromText acModule, moduleName, modulePath

    MsgBox "Module '" & moduleName & "' imported from: " & modulePath, vbInformation
End Sub


' --- Public entry points (use RunCode in a macro) ---
Public Sub CreateShadowTables()
    CloseAllOpenTables

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim candidates As Collection
    Dim tName As String, shadowName As String

    Set db = CurrentDb
    Set candidates = New Collection

    ' Collect user table names (exclude system/hidden/linked and shadows)
    For Each tdf In db.TableDefs
        If IsUserTable(tdf) Then candidates.Add tdf.Name
    Next

    ' Create / refresh shadow tables as "original_<name>"
    On Error GoTo EH
    Dim i As Long
    For i = 1 To candidates.Count
        tName = candidates(i)
        shadowName = "original_" & tName

        ' If a previous shadow exists, delete it
        If TableExists(shadowName) Then
            DoCmd.DeleteObject acTable, shadowName
        End If

        ' Copy the table (structure + data + indexes) to the shadow name
        DoCmd.CopyObject , shadowName, acTable, tName
    Next

    MsgBox "Shadow tables created/updated successfully.", vbInformation
    Exit Sub

EH:
    MsgBox "CreateShadowTables failed on table '" & tName & "'." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation
End Sub


Public Sub RestoreFromShadowTables()
    CloseAllOpenTables

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim candidates As Collection
    Dim tName As String, shadowName As String

    Set db = CurrentDb
    Set candidates = New Collection

    ' Collect user table names (exclude system/hidden/linked and shadows)
    For Each tdf In db.TableDefs
        If IsUserTable(tdf) Then candidates.Add tdf.Name
    Next

    On Error GoTo EH

    ' Replace each user table with its shadow
    Dim i As Long
    For i = 1 To candidates.Count
        tName = candidates(i)
        shadowName = "original_" & tName

        If TableExists(shadowName) Then
            ' Delete the current table (this also drops relationships on it)
            If TableExists(tName) Then
                DoCmd.DeleteObject acTable, tName
            End If

            ' Copy shadow back to the original name
            DoCmd.CopyObject , tName, acTable, shadowName
        Else
            ' No shadow for this table—skip gracefully
        End If
    Next

    MsgBox "All tables restored from shadows.", vbInformation
    Exit Sub

EH:
    MsgBox "RestoreFromShadowTables failed on table '" & tName & "'." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation
End Sub


' --- Helpers ---

Private Sub CloseAllOpenTables()
    ' Closes any open table datasheets in the UI, if present
    On Error Resume Next

    Dim accObj As Access.AccessObject
    For Each accObj In CurrentData.AllTables
        If SysCmd(acSysCmdGetObjectState, acTable, accObj.Name) <> 0 Then
            DoCmd.Close acTable, accObj.Name, acSaveNo
        End If
    Next

    On Error GoTo 0
End Sub


Private Function IsUserTable(tdf As DAO.TableDef) As Boolean
    ' Exclude system, hidden, and linked tables, and any "original_" shadow

    Dim isSystem As Boolean
    Dim isHidden As Boolean
    Dim isLinked As Boolean
    Dim isShadow As Boolean
    Dim nm As String

    nm = tdf.Name
    isShadow = (Left$(nm, 9) = "original_")
    isLinked = (Len(tdf.Connect) > 0)
    isSystem = (Left$(nm, 4) = "MSys") Or ((tdf.Attributes And dbSystemObject) <> 0)
    isHidden = ((tdf.Attributes And dbHiddenObject) <> 0)

    IsUserTable = (Not isSystem) And (Not isHidden) And (Not isLinked) And (Not isShadow)
End Function


Private Function TableExists(tableName As String) As Boolean
    On Error GoTo NoTable
    Dim t As DAO.TableDef
    Set t = CurrentDb.TableDefs(tableName)
    TableExists = True
    Exit Function
NoTable:
    TableExists = False
End Function


