Option Compare Database


Private Sub CloseTableIfOpen(ByVal tableName As String)
    On Error Resume Next
    If SysCmd(acSysCmdGetObjectState, acTable, tableName) <> 0 Then
        DoCmd.Close acTable, tableName, acSaveNo
        Debug.Print "Closed open table: " & tableName
    End If
    On Error GoTo 0
End Sub


Public Function RenameTables() As Boolean
    On Error GoTo ErrorHandler
    Dim db As DAO.Database
    Dim table As DAO.TableDef
    Set db = CurrentDb
    
    For Each table In db.TableDefs
        If Mid(table.Name, 1, 6) = "public" Then
            CloseTableIfOpen table.Name
            table.Name = Replace(table.Name, "public_", "")
            Debug.Print "Renamed table: " & table.Name
        End If
    Next table

    RenameTables = True
    Debug.Print "Successfully renamed tables."
    
    Exit Function
ErrorHandler:
    MsgBox "Error renaming tables: " & Err.Description, vbCritical
    RenameTables = False
End Function



Public Function PostLinkedTableUpdates() As Boolean
    On Error GoTo ErrorHandler
    Dim isOK As Boolean

    ' Rename tables to remove "public_" prefix
    isOK = RenameTables()
    
    PostLinkedTableUpdates = isOK
    Exit Function
ErrorHandler:
    Debug.Print "error: PostLinkedTableUpdates failed: " & Err.Description
    PostLinkedTableUpdates = False
End Function
