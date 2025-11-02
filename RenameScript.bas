Option Compare Database
Option Explicit

Public Sub RemoveFieldRowSource(tableName As String, fieldName As String)
    On Error GoTo EH

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field

    Set db = CurrentDb
    Set tdf = db.TableDefs(tableName)
    Set fld = tdf.Fields(fieldName)

    ' Remove lookup-related properties (if they exist)
    On Error Resume Next
    fld.Properties.Delete "RowSource"
    fld.Properties.Delete "RowSourceType"
    fld.Properties.Delete "BoundColumn"
    fld.Properties.Delete "ColumnCount"
    fld.Properties.Delete "ColumnHeads"
    fld.Properties.Delete "ColumnWidths"
    fld.Properties.Delete "ListRows"
    fld.Properties.Delete "ListWidth"
    fld.Properties.Delete "LimitToList"
    On Error GoTo EH

    MsgBox "RowSource removed from " & tableName & "." & fieldName, vbInformation
    Exit Sub

EH:
    MsgBox "Error: " & Err.Description, vbExclamation
End Sub

Public Sub SetFieldRowSource(tableName As String, fieldName As String, sqlSource As String)
    On Error GoTo EH

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field

    Set db = CurrentDb
    Set tdf = db.TableDefs(tableName)
    Set fld = tdf.Fields(fieldName)

    ' --- Set lookup display mode ---
    ' RowSourceType "Table/Query" means the RowSource is SQL.
    ' If using a value list instead, set RowSourceType = "Value List"
    On Error Resume Next

    ' If properties donâ€™t exist yet, create them.
    If PropertyExists(fld, "RowSourceType") Then
        fld.Properties("RowSourceType") = "Table/Query"
    Else
        fld.Properties.Append fld.CreateProperty("RowSourceType", dbText, "Table/Query")
    End If

    If PropertyExists(fld, "RowSource") Then
        fld.Properties("RowSource") = sqlSource
    Else
        fld.Properties.Append fld.CreateProperty("RowSource", dbMemo, sqlSource)
    End If

    ' Optional but recommended:
    ' Specify which column is stored and how many are shown.
    AddOrReplaceProperty fld, "BoundColumn", dbInteger, 1
    AddOrReplaceProperty fld, "ColumnCount", dbInteger, 2
    AddOrReplaceProperty fld, "ColumnWidths", dbText, "0cm;5cm" ' hide ID, show Name

    On Error GoTo EH
    MsgBox "RowSource added to field " & tableName & "." & fieldName, vbInformation
    Exit Sub

EH:
    MsgBox "Error setting RowSource: " & Err.Description, vbExclamation
End Sub


' --- Helper: create or update any property ---
Private Sub AddOrReplaceProperty(obj As Object, propName As String, propType As DAO.DataTypeEnum, propValue As Variant)
    On Error Resume Next
    If PropertyExists(obj, propName) Then
        obj.Properties(propName) = propValue
    Else
        obj.Properties.Append obj.CreateProperty(propName, propType, propValue)
    End If
    On Error GoTo 0
End Sub

' --- Helper: check if a property exists ---
Private Function PropertyExists(obj As Object, propName As String) As Boolean
    Dim p As DAO.property
    On Error Resume Next
    Set p = obj.Properties(propName)
    PropertyExists = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
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


Public Function VerifyInputData() As Boolean
    On Error GoTo ErrorHandler
    Dim rs As DAO.Recordset
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim property As DAO.property
    Dim rowSource As String

    VerifyInputData = True
    For Each tdf In CurrentDb.TableDefs
        If Mid(tdf.Name, 1, 4) <> "MSys" And tdf.Name <> "TranslationMapping" And tdf.Name <> "Ortnamn_ny" Then
            For Each fld In tdf.Fields
                Set rs = CurrentDb.OpenRecordset("SELECT * FROM TranslationMapping WHERE OriginalTable = '" & tdf.Name & "' AND OriginalColumn = '" & fld.Name & "'", dbOpenSnapshot)
                If rs.EOF Then
                    ' Debug.Print "NOT OK: [" & tdf.Name & "] -> [" & fld.Name & "]"
                    VerifyInputData = False
                Else
                    ' Debug.Print "OK: [" & tdf.Name & "] -> [" & fld.Name & "]"
                End If
                On Error Resume Next
                Set property = fld.Properties("RowSource")
                If Err.Number = 0 Then
                    Debug.Print "RowSource;" & tdf.Name & ";" & fld.Name & ";" & property.Value
                End If
                Err.Clear
                On Error GoTo 0
                Set rs = Nothing
            Next fld
        End If
    Next tdf

    Exit Function
ErrorHandler:
    MsgBox "Error verifying input data: " & Err.Description, vbCritical
    VerifyInputData = False
End Function



Public Function FindExpressions() As Boolean
    On Error GoTo ErrorHandler
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field

    FindExpressions = True
    For Each tdf In CurrentDb.TableDefs
        If Mid(tdf.Name, 1, 4) <> "MSys" And tdf.Name <> "TranslationMapping" And tdf.Name <> "Ortnamn_ny" Then
            For Each fld In tdf.Fields
                On Error Resume Next
                If fld.Expression <> "" Then
                    Debug.Print "Expression;" & tdf.Name & ";" & fld.Name & ";" & fld.Expression
                End If
            Next fld
        End If
    Next tdf

    Exit Function
ErrorHandler:
    MsgBox "Error finding expressions: " & Err.Description, vbCritical
    FindExpressions = False
End Function

Public Function ClearExpressions() As Boolean
    On Error GoTo ErrorHandler
    ClearExpressions = True
    CurrentDb.TableDefs("Personer").Fields("Helnamn").Expression = ""
    Exit Function
ErrorHandler:
    MsgBox "Error clearing expressions: " & Err.Description, vbCritical
    ClearExpressions = False
End Function

Public Function SetExpressions() As Boolean
    On Error GoTo ErrorHandler
    Dim rs As DAO.Recordset
    SetExpressions = True
    Set rs = CurrentDb.OpenRecordset("SELECT [TranslatedTable], [TranslatedColumn], [TranslatedExpression] FROM [TranslationMapping] WHERE [TranslatedExpression] <> ''", dbOpenSnapshot)
    Do Until rs.EOF
        On Error Resume Next
        CurrentDb.TableDefs(rs!TranslatedTable).Fields(rs!TranslatedColumn).Expression = rs!TranslatedExpression
        If Err.Number <> 0 Then
            SetExpressions = False
            Debug.Print "Error setting Expression for " & rs!TranslatedTable & "." & rs!TranslatedColumn & ": " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Exit Function
ErrorHandler:
    MsgBox "Error clearing expressions: " & Err.Description, vbCritical
    SetExpressions = False
End Function

Public Function ClearRowSource() As Boolean
    On Error GoTo ErrorHandler
    Dim rs As DAO.Recordset
    ClearRowSource = True
    Set rs = CurrentDb.OpenRecordset("SELECT [OriginalTable], [OriginalColumn] FROM [RowSource]", dbOpenSnapshot)
    Do Until rs.EOF
        On Error Resume Next
        CurrentDb.TableDefs(rs!OriginalTable).Fields(rs!OriginalColumn).Properties("RowSource") = Null
        If Err.Number <> 0 Then
            ClearRowSource = False
            Debug.Print "Error removing RowSource from " & rs!OriginalTable & "." & rs!OriginalColumn & ": " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Debug.Print "Done!"
    Exit Function
ErrorHandler:
    MsgBox "Error removing RowSource: " & Err.Description, vbCritical
    ClearRowSource = False
End Function

Public Function SetRowSource() As Boolean
    On Error GoTo ErrorHandler
    Dim rs As DAO.Recordset
    SetRowSource = True
    Set rs = CurrentDb.OpenRecordset("SELECT [TranslatedTable], [TranslatedColumn], [RowSource] FROM [RowSource]", dbOpenSnapshot)
    Do Until rs.EOF
        On Error Resume Next
        CurrentDb.TableDefs(rs!TranslatedTable).Fields(rs!TranslatedColumn).Properties("RowSource") = rs!rowSource
        If Err.Number <> 0 Then
            SetRowSource = False
            Debug.Print "Error setting RowSource for " & rs!TranslatedTable & "." & rs!TranslatedColumn & ": " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Debug.Print "Done!"
    Exit Function
ErrorHandler:
    MsgBox "Error removing RowSource: " & Err.Description, vbCritical
    SetRowSource = False
End Function


Public Function RenameFields() As Boolean
    On Error GoTo ErrorHandler
    Dim rs As DAO.Recordset
    RenameFields = True
    Set rs = CurrentDb.OpenRecordset("SELECT OriginalTable, OriginalColumn, TranslatedColumn FROM TranslationMapping", dbOpenDynaset)
    Do Until rs.EOF
        On Error Resume Next
        CurrentDb.TableDefs(rs!OriginalTable).Fields(rs!OriginalColumn).Name = rs!TranslatedColumn
        If Err.Number <> 0 Then
            Debug.Print "Error renaming field " & rs!OriginalTable & "." & rs!OriginalColumn & ": " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Debug.Print "Done!"
    Exit Function
ErrorHandler:
    MsgBox "Error renaming fields: " & Err.Description, vbCritical
    RenameFields = False
End Function

Public Sub TranslateTablesAndFields_FromTable()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field

    Dim origTable As String, newTable As String
    Dim origField As String, newField As String
    Dim comment As String
    Dim isOK As Boolean

    Set db = DBEngine(0)(0) ' CurrentDb may have issues in some contexts

    Let isOK = VerifyInputData

    If Not isOK Then
        MsgBox "Input data verification failed. Please check the debug output for details.", vbCritical
        Exit Sub
    End If

    Call RenameFields

    ' Step 1: Rename Tables
    Set rs = db.OpenRecordset("SELECT DISTINCT OriginalTable, TranslatedTable FROM TranslationMapping", dbOpenDynaset)

    If rs.EOF Then
        MsgBox "TranslationMapping table is empty.", vbExclamation
        Exit Sub
    End If

    On Error GoTo 0


    rs.MoveFirst
    Do Until rs.EOF
        origTable = rs!OriginalTable
        newTable = rs!TranslatedTable


        If StrComp(origTable, newTable, vbBinaryCompare) <> 0 Then
            Set tdf = Nothing
            ' Let origTable = FindTableNameIgnoreCase(origTable)
            Set tdf = db.TableDefs(origTable)
            If Not tdf Is Nothing Then
                Debug.Print "Renaming table: " & origTable & " -> " & newTable
                tdf.Name = newTable
            End If
        End If
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing

    ' Step 2: Rename Fields and Add Comments
    Set rs = db.OpenRecordset("SELECT * FROM TranslationMapping", dbOpenDynaset)
    rs.MoveFirst
    Do Until rs.EOF
        newTable = rs!TranslatedTable
        origField = rs!OriginalColumn
        newField = rs!TranslatedColumn
        comment = Nz(rs!comment, "")

        Set tdf = db.TableDefs(newTable)

        Set fld = GetFieldSafe(tdf, origField)



        Debug.Print "Renaming field: " & newTable & "." & origField & " -> " & newField

        If comment <> "" Then
            On Error Resume Next
            fld.Properties("Description") = comment
            If Err.Number <> 0 Then
                Err.Clear
                Dim p As DAO.property
                Set p = fld.CreateProperty("Description", dbText, comment)
                fld.Properties.Append p
            End If
            On Error GoTo 0
        End If

        If RenameFieldSafe(tdf, origField, newField) <> True Then
            Debug.Print "Failed to rename field: " & fld.Name
        End If

        Err.Clear
        rs.MoveNext

    Loop

    rs.Close
    Set rs = Nothing

    ' TODO: Update this!!!

    ' Step 3: Update lookup RowSources (safe, case-insensitive)
    ' For Each tdf In db.TableDefs
    '     ' Skip system/temporary tables
    '     If Left(tdf.Name, 4) <> "MSys" And Left(tdf.Name, 1) <> "~" Then
    '         For Each fld In tdf.Fields
    '             On Error Resume Next
    '             Set prop = fld.Properties("RowSource")
    '             If Err.Number = 0 Then
    '                 If Not IsNull(prop.Value) And prop.Value <> "" Then
    '                     oldSource = prop.Value
    '                     newSource = oldSource

    '                     rs.MoveFirst
    '                     Do Until rs.EOF
    '                         origTable = Nz(rs!OriginalTable, "")
    '                         newTable = Nz(rs!TranslatedTable, "")
    '                         origField = Nz(rs!OriginalColumn, "")
    '                         newField = Nz(rs!TranslatedColumn, "")

    '                         ' Replace whole bracketed names (case-insensitive)
    '                         If origTable <> "" And newTable <> "" Then
    '                             regex.Pattern = "\[" & Replace(origTable, "[", "\[") & "\]"
    '                             newSource = regex.Replace(newSource, "[" & newTable & "]")
    '                         End If
    '                         If origField <> "" And newField <> "" Then
    '                             regex.Pattern = "\[" & Replace(origField, "[", "\[") & "\]"
    '                             newSource = regex.Replace(newSource, "[" & newField & "]")
    '                         End If
    '                         rs.MoveNext
    '                     Loop

    '                     If newSource <> oldSource Then
    '                         prop.Value = newSource
    '                         Debug.Print "Updated RowSource in [" & tdf.Name & "].[" & fld.Name & "]"
    '                     End If
    '                 End If
    '             End If
    '             Err.Clear
    '             On Error GoTo 0
    '         Next fld
    '     End If
    ' Next tdf

    Set db = Nothing

    MsgBox "Table and column translation complete.", vbInformation
End Sub
