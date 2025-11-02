Option Explicit
Option Compare Database

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


'------------------------------------------------------------
' Finds a TableDef in CurrentDb by name, ignoring case.
' Returns:
'   - TableDef object if found
'   - Nothing if not found
'------------------------------------------------------------
Public Function FindTableNameIgnoreCase(tableName As String) As String
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef

    Set db = CurrentDb

    For Each tdf In db.TableDefs
        ' Compare ignoring case
        If StrComp(tdf.Name, tableName, vbTextCompare) = 0 Then
            Let FindTableNameIgnoreCase = tdf.Name
            Exit Function
        End If
    Next tdf

    ' If no match, return Nothing
    Let FindTableNameIgnoreCase = ""
End Function


'------------------------------------------------------------
' Finds a field name in CurrentDb by table name and field name,
' ignoring case sensitivity.
'
' Parameters:
'   tableName - name of the table (case-insensitive)
'   fieldName - name of the field/column (case-insensitive)
'
' Returns:
'   String containing the field name if found, empty string otherwise.
'------------------------------------------------------------
Public Function FindFieldNameIgnoreCase(tableName As String, fieldName As String) As String
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field

    Set db = CurrentDb

    ' --- Find table ignoring case ---
    For Each tdf In db.TableDefs
        If StrComp(tdf.Name, tableName, vbTextCompare) = 0 Then
            ' --- Found table: now find field ignoring case ---
            For Each fld In tdf.Fields
                If StrComp(fld.Name, fieldName, vbTextCompare) = 0 Then
                    Let FindFieldNameIgnoreCase = fld.Name
                    Exit Function
                End If
            Next fld
            ' Field not found in that table
            Exit Function
        End If
    Next tdf

    ' Not found: return empty string
    Let FindFieldNameIgnoreCase = ""
End Function

'------------------------------------------------------------
' Safely gets a Field object from a TableDef by field name,
' ignoring case sensitivity.
' Parameters:
'   tdf - TableDef object
'   field_name - name of the field/column (case-insensitive)
' Returns:
'   Field object if found, Nothing otherwise.
'------------------------------------------------------------
Public Function GetFieldSafe(tdf As DAO.TableDef, field_name As String) As DAO.Field
    Dim fld As DAO.Field
    For Each fld In tdf.Fields
        If StrComp(fld.Name, field_name, vbTextCompare) = 0 Then
            Set GetFieldSafe = fld
            Exit Function
        End If
    Next fld
    Set GetFieldSafe = Nothing
End Function

'------------------------------------------------------------
' Safely renames a field in a TableDef.
' Parameters:
'   tdf - TableDef object
'   oldName - current name of the field
'   newName - new name for the field
' Returns:
'   True if rename succeeded, False otherwise.
'------------------------------------------------------------
Public Function RenameFieldSafe(tdf As DAO.TableDef, oldName As String, newName As String) As Boolean
    On Error Resume Next
    tdf.Fields(oldName).Name = newName
    If Err.Number <> 0 Then
        Debug.Print "Error renaming [" & oldName & "]: " & Err.Description
        RenameFieldSafe = False
        Err.Clear
    Else
        RenameFieldSafe = True
    End If
    On Error GoTo 0
End Function

