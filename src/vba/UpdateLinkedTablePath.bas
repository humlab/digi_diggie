Public Sub UpdateLinkedTablePath(tableName As String, newPath As String)
    ' Updates the link path (Connect property) for a linked table.

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim oldConnect As String
    
    On Error GoTo ErrHandler
    Set db = CurrentDb
    Set tdf = db.TableDefs(tableName)
    
    If Len(tdf.Connect) = 0 Then
        MsgBox "Table '" & tableName & "' is not linked.", vbExclamation
        Exit Sub
    End If
    
    oldConnect = tdf.Connect
    Debug.Print "Old Connect: " & oldConnect
    
    ' For Access backend (.accdb / .mdb)
    ' Example connect: ";DATABASE=C:\OldPath\backend.accdb"
    If InStr(1, oldConnect, ";DATABASE=", vbTextCompare) > 0 Then
        tdf.Connect = ";DATABASE=" & newPath
    Else
        ' For ODBC links (SQL Server, etc.)
        tdf.Connect = newPath
    End If
    
    tdf.RefreshLink
    Debug.Print "Updated " & tableName & " to " & tdf.Connect
    MsgBox "Updated link for '" & tableName & "' to:" & vbCrLf & tdf.Connect, vbInformation
    
ExitHere:
    Exit Sub
    
ErrHandler:
    MsgBox "Error updating link for '" & tableName & "': " & Err.Description, vbCritical
    Resume ExitHere
End Sub


Sub UpdatePath()
    ' Update the linked table path for "Ortnamn_ny" table
    ' Assumes the linked database is in the same folder as the current database
    
    Dim linkedDbPath As String
    Dim tableName As String
    linkedDbPath = CurrentProject.Path
    tableName = "Ortnamn_ny"

    Call UpdateLinkedTablePath(tableName, linkedDbPath & "\Ortnamn.accdb")
End Sub