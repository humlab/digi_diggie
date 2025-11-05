Option Compare Database
Option Explicit

Private Const dbCalculated As Long = 289

Public Sub RemoveRowSource(tableName As String, fieldName As String)
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim table As DAO.TableDef
    Dim field As DAO.Field

    Set db = CurrentDb
    Set table = db.TableDefs(tableName)
    Set field = table.Fields(fieldName)

    On Error Resume Next
    field.Properties.Delete "RowSource"
    field.Properties.Delete "RowSourceType"
    field.Properties.Delete "BoundColumn"
    field.Properties.Delete "ColumnCount"
    field.Properties.Delete "ColumnHeads"
    field.Properties.Delete "ColumnWidths"
    field.Properties.Delete "ListRows"
    field.Properties.Delete "ListWidth"
    field.Properties.Delete "LimitToList"

    table.Fields.Refresh

    Exit Sub
ErrorHandler:
    Debug.Print "error: failed to remove RowSource from " & tableName & "." & fieldName & ": " & Err.Description
End Sub

Public Sub SetFieldRowSource(tableName As String, fieldName As String, sqlSource As String)
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim table As DAO.TableDef
    Dim field As DAO.Field

    Set db = CurrentDb
    Set table = db.TableDefs(tableName)
    Set field = table.Fields(fieldName)

    ' RowSourceType "Table/Query" means the RowSource is SQL.
    ' If using a value list instead, set RowSourceType = "Value List"
    On Error Resume Next

    ' If properties don't exist yet, create them.
    If PropertyExists(field, "RowSourceType") Then
        field.Properties("RowSourceType") = "Table/Query"
    Else
        field.Properties.Append field.CreateProperty("RowSourceType", dbText, "Table/Query")
    End If

    If PropertyExists(field, "RowSource") Then
        field.Properties("RowSource") = sqlSource
    Else
        field.Properties.Append field.CreateProperty("RowSource", dbMemo, sqlSource)
    End If

    ' Optional but recommended:
    ' Specify which column is stored and how many are shown.
    AddOrReplaceProperty field, "BoundColumn", dbInteger, 1
    AddOrReplaceProperty field, "ColumnCount", dbInteger, 2
    AddOrReplaceProperty field, "ColumnWidths", dbText, "0cm;5cm" ' hide ID, show Name
    
    table.Fields.Refresh

    Exit Sub

ErrorHandler:
    Debug.Print "error: failed setting RowSource for " & tableName & "." & fieldName & ": " & Err.Description
End Sub

Sub SetFieldDescription(tableName As String, fieldName As String, description As String)
    On Error GoTo ErrorHandler
    Dim db As DAO.Database
    Dim td As DAO.TableDef
    Dim fld As DAO.Field

    Set db = CurrentDb
    Set td = db.TableDefs(tableName)
    Set fld = td.Fields(fieldName)

    On Error Resume Next
    fld.Properties("Description") = description
    If Err.Number <> 0 Then
        Err.Clear
        fld.Properties.Append fld.CreateProperty("Description", dbText, description)
    End If
    On Error GoTo 0
    Exit Sub

ErrorHandler:
    Debug.Print "error: failed setting description for " & tableName & "." & fieldName & ": " & Err.Description
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
    Dim property As DAO.property
    On Error Resume Next
    Set property = obj.Properties(propName)
    PropertyExists = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Function SquareBracket(name As String) As String
    If Mid(name, 1, 1) = "[" Then
        SquareBracket = name
    Else
        SquareBracket = "[" & name & "]"
    End If
End Function

Private Function IsDataTable(tableName As String) As Boolean
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM TranslationMapping WHERE OriginalTable='" & SquareBracket(tableName) & "'", dbOpenSnapshot)
    IsDataTable = Not rs.EOF
    rs.Close
    Set rs = Nothing
End Function

Private Function IsDataColumn(tableName As String, columnName As String) As Boolean
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM TranslationMapping WHERE OriginalTable='" & SquareBracket(tableName) & "' AND OriginalColumn='" & SquareBracket(columnName) & "'", dbOpenSnapshot)
    IsDataColumn = Not rs.EOF
    rs.Close
    Set rs = Nothing
End Function

Public Function ShowRowSources() As Boolean
    On Error GoTo ErrorHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim table As DAO.TableDef
    Dim field As DAO.Field
    Dim rowSource As String
    Dim property As DAO.property
    Set db = CurrentDb
    ShowRowSources = True
    For Each table In db.TableDefs
        If IsDataTable(table.Name) Then
            For Each field In table.Fields
                On Error Resume Next
                Set property = field.Properties("RowSource")
                If Err.Number = 0 Then
                    Debug.Print "RowSource;" & table.Name & ";" & field.Name & ";" & property.Value
                End If
                Err.Clear
                On Error GoTo 0
                Set rs = Nothing
            Next field
        End If
    Next table
    Exit Function
ErrorHandler:
    MsgBox "Error verifying original columns: " & Err.Description, vbCritical
    ShowRowSources = False
End Function

Public Function VerifyMappings() As Boolean
    Dim db As DAO.Database
    Dim table As DAO.TableDef
    Dim field As DAO.Field
    Dim property As DAO.property

    Set db = CurrentDb
    VerifyMappings = True
    For Each table In db.TableDefs
        If IsDataTable(table.Name) Then
            For Each field In table.Fields
                If Not IsDataColumn(table.Name, field.Name) Then
                    Debug.Print "error: NOT FOUND [" & table.Name & "] -> [" & field.Name & "]"
                    VerifyMappings = False
                End If
            Next field
        End If
    Next table
    If Not VerifyMappings Then
        Debug.Print "error: some original columns were not found. See above."
    Else
        Debug.Print "info: all original columns verified to be OK!"
    End If
    If Not VerifyMappings Then
        Err.Raise vbObjectError + 513, , "Some original columns were not found."
    End If
End Function

Public Function DropOriginalExpressions() As Boolean
    On Error GoTo ErrorHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    Set db = CurrentDb
    DropOriginalExpressions = True
    Set rs = db.OpenRecordset("SELECT * FROM TranslationMapping WHERE TranslatedExpression <> ''", dbOpenSnapshot)
    Do Until rs.EOF
        On Error Resume Next
        db.Execute "ALTER TABLE " & SquareBracket(rs!OriginalTable) & " DROP COLUMN " & SquareBracket(rs!OriginalColumn) & ";", dbFailOnError
        If Err.Number <> 0 Then
            DropOriginalExpressions = False
            Debug.Print "error: failed dropping computed column " & rs!TranslatedTable & "." & rs!TranslatedColumn & ": " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    db.TableDefs.Refresh
    Debug.Print "info: done dropping original expressions!"
    Exit Function
ErrorHandler:
    MsgBox "error: failed dropping computed column: " & Err.Description, vbCritical
    DropOriginalExpressions = False
End Function

Public Function AddTranslatedExpressions() As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim table As DAO.TableDef
    Dim field As DAO.Field2
    Dim tableName As String, fieldName As String, expr As String

    AddTranslatedExpressions = True
    Set db = CurrentDb


    Set rs = db.OpenRecordset( _
        "SELECT TranslatedTable, TranslatedColumn, TranslatedExpression " & _
        "FROM TranslationMapping WHERE Nz(TranslatedExpression,'')<>''", dbOpenSnapshot)

    Do Until rs.EOF
        tableName = rs!TranslatedTable
        fieldName = rs!TranslatedColumn
        expr = rs!TranslatedExpression

        Set table = db.TableDefs(tableName)

        ' If the field exists, delete it first (no ALTER needed)
        On Error Resume Next
        table.Fields.Delete fieldName
        On Error GoTo ErrorHandler

        ' Create a CALCULATED field and set its expression
        Set field = table.CreateField(fieldName, dbText)   ' Field2
        field.Expression = expr                            ' e.g. [given_name] & " " & [surname]
        table.Fields.Append field
        table.Fields.Refresh

        rs.MoveNext
    Loop

    rs.Close: Set rs = Nothing
    db.TableDefs.Refresh
    Debug.Print "info: done adding translated expressions!"
    Exit Function

ErrorHandler:
    Debug.Print "error: failed adding computed column " & tableName & "." & fieldName & ": " & Err.Number & " - " & Err.Description
    AddTranslatedExpressions = False
End Function


Public Function RemoveOriginalQueries() As Boolean
    On Error GoTo ErrorHandler
    Dim db As DAO.Database
    Dim query As DAO.QueryDef
    Dim queryName As String
    RemoveOriginalQueries = True
    Set db = CurrentDb
    For Each query In db.QueryDefs
        Let queryName = query.Name
        db.QueryDefs.Delete queryName
        Debug.Print "info: deleted query " & queryName
    Next query
    db.QueryDefs.Refresh
    Debug.Print "info: done removing original queries!"
    Exit Function
ErrorHandler:
    MsgBox "Error removing queries: " & Err.Description, vbCritical
    RemoveOriginalQueries = False
End Function

Public Function ClearRowSource() As Boolean
    On Error GoTo ErrorHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    ClearRowSource = True
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT [OriginalTable], [OriginalColumn] FROM [RowSource]", dbOpenSnapshot)
    Do Until rs.EOF
        On Error Resume Next
        db.TableDefs(rs!OriginalTable).Fields(rs!OriginalColumn).Properties("RowSource") = "SELECT 1 AS DummyID, '' AS DummyValue WHERE False;"
        If Err.Number <> 0 Then
            ClearRowSource = False
            Debug.Print "error: failed removing RowSource from " & rs!OriginalTable & "." & rs!OriginalColumn & ": " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    db.TableDefs.Refresh
    Debug.Print "info: done clearing RowSource!"
    Exit Function
ErrorHandler:
    MsgBox "Error removing RowSource: " & Err.Description, vbCritical
    ClearRowSource = False
End Function

Public Function RemoveOriginalRowSources() As Boolean
    On Error GoTo ErrorHandler
    Dim db As DAO.Database
    Dim table As DAO.TableDef
    Dim rs As DAO.Recordset
    RemoveOriginalRowSources = True
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT [OriginalTable], [OriginalColumn] FROM [RowSource]", dbOpenSnapshot)
    Do Until rs.EOF
        On Error Resume Next
        Set table = db.TableDefs(rs!OriginalTable)
        RemoveRowSource table.Name, rs!OriginalColumn
        If Err.Number <> 0 Then
            RemoveOriginalRowSources = False
            Debug.Print "error: failed removing RowSource for " & table.Name & "." & rs!OriginalColumn & ": " & Err.Description
            Err.Clear
        End If
        table.Fields.Refresh
        On Error GoTo 0
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    db.TableDefs.Refresh
    Debug.Print "info: done removing original RowSource!"
    Exit Function
ErrorHandler:
    MsgBox "error: failed removing original RowSource: " & Err.Description, vbCritical
    RemoveOriginalRowSources = False
End Function

Public Function AddTranslatedRowSources() As Boolean
    On Error GoTo ErrorHandler
    Dim db As DAO.Database
    Dim table As DAO.TableDef
    Dim field As DAO.Field
    Dim rs As DAO.Recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT [TranslatedTable], [TranslatedColumn], [RowSource] FROM [RowSource]", dbOpenSnapshot)
    AddTranslatedRowSources = True
    Do Until rs.EOF
        On Error Resume Next
        Set table = db.TableDefs(rs!TranslatedTable)
        Set field = table.Fields(rs!TranslatedColumn)

        SetFieldRowSource table.Name, field.Name, rs!rowSource

        If Err.Number <> 0 Then
            AddTranslatedRowSources = False
            Debug.Print "error: failed setting RowSource for " & table.Name & "." & field.Name & ": " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
        table.Fields.Refresh
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    db.TableDefs.Refresh
    Debug.Print "info: done adding translated RowSource!"
    Exit Function
ErrorHandler:
    MsgBox "Error adding translated RowSource: " & Err.Description, vbCritical
    AddTranslatedRowSources = False
End Function


Public Function RenameFields() As Boolean
    On Error GoTo ErrorHandler
    Dim db As DAO.Database
    Dim table As DAO.TableDef
    Dim field As DAO.Field
    Dim rs As DAO.Recordset
    RenameFields = True
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT OriginalTable, OriginalColumn, TranslatedColumn FROM TranslationMapping", dbOpenSnapshot)
    Do Until rs.EOF
        On Error Resume Next
        Set table = db.TableDefs(rs!OriginalTable)
        Set field = table.Fields(rs!OriginalColumn)
        field.Name = rs!TranslatedColumn
        If Err.Number = 3265 Then ' Item not found in this collection.
            ' Assume already renamed
        ElseIf Err.Number <> 0 Then
            Debug.Print "error: failed renaming field " & rs!OriginalTable & "." & rs!OriginalColumn & ": " & Err.Description
        End If
        Err.Clear
        On Error GoTo 0
        table.Fields.Refresh
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    db.TableDefs.Refresh
    Debug.Print "info: done renaming fields!"
    Exit Function
ErrorHandler:
    MsgBox "Error renaming fields: " & Err.Description, vbCritical
    RenameFields = False
End Function


Public Function DeleteDeprecatedFields() As Boolean
    On Error GoTo ErrorHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim table As DAO.TableDef
    Dim field As DAO.field
    DeleteDeprecatedFields = True
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM TranslationMapping WHERE DeprecateFlag = 'YES'", dbOpenSnapshot)
    Do Until rs.EOF
        On Error Resume Next
        Set table = db.TableDefs(rs!TranslatedTable)
        table.Fields.Delete !TranslatedColumn
        If Err.Number <> 0 Then
            Debug.Print "error: failed deleting field " & rs!TranslatedTable & "." & rs!TranslatedColumn & ": " & Err.description
        End If
        Err.Clear
        On Error GoTo 0
        table.Fields.Refresh
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    db.TableDefs.Refresh
    Debug.Print "info: done deleting deprecated fields!"
    Exit Function
ErrorHandler:
    MsgBox "Error deleting fields: " & Err.description, vbCritical
    DeleteDeprecatedFields = False
End Function



Public Function RenameTables() As Boolean
    On Error GoTo ErrorHandler
    Dim db As DAO.Database
    Dim table As DAO.TableDef
    Dim rs As DAO.Recordset
    RenameTables = True
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT DISTINCT OriginalTable, TranslatedTable FROM TranslationMapping", dbOpenSnapshot)
    Do Until rs.EOF
        On Error Resume Next
        Set table = db.TableDefs(rs!OriginalTable)
        table.Name = rs!TranslatedTable
        If Err.Number <> 0 Then
            Debug.Print "error: failed renaming table " & rs!OriginalTable & " to " & rs!TranslatedTable & ": " & Err.Description
        End If
        Err.Clear
        On Error GoTo 0
        table.Fields.Refresh
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    db.TableDefs.Refresh
    Debug.Print "info: done renaming tables!"
    Exit Function
ErrorHandler:
    MsgBox "Error renaming tables: " & Err.Description, vbCritical
    RenameTables = False
End Function


Public Function AddTranslatedQueries() As Boolean
    On Error GoTo ErrorHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim query As DAO.QueryDef

    AddTranslatedQueries = True
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM QueryDefinitions", dbOpenSnapshot)
    Do Until rs.EOF
        On Error Resume Next
        db.QueryDefs.Delete rs!QueryName
        Err.Clear
        Set query = db.CreateQueryDef(rs!QueryName, rs!SQLText)
        If Err.Number <> 0 Then
            Debug.Print "error: error adding query " & rs!QueryName & ": " & Err.Description
            AddTranslatedQueries = False
            Err.Clear
        End If
        On Error GoTo 0
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    db.QueryDefs.Refresh
    Debug.Print "info: done adding translated queries!"
    Exit Function
ErrorHandler:
    MsgBox "Error adding queries: " & Err.Description, vbCritical
    AddTranslatedQueries = False
End Function

Public Function AddComments() As Boolean
    On Error GoTo ErrorHandler
    Dim rs As DAO.Recordset
    Dim db As DAO.Database
    Set db = CurrentDb
    AddComments = True
    Set rs = db.OpenRecordset("SELECT * FROM TranslationMapping", dbOpenSnapshot)
    Do Until rs.EOF
        SetFieldDescription rs!TranslatedTable, rs!TranslatedColumn, rs!Comment
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Exit Function
ErrorHandler:
    Debug.Print "error: Error adding comments: " & Err.Description
    AddComments = False
End Function

Public Function RemoveImportedObjects() As Boolean
    On Error GoTo ErrorHandler
    Dim db As DAO.Database
    Dim table As DAO.TableDef

    RemoveImportedObjects = True
    Set db = CurrentDb
    db.TableDefs.Delete "TranslationMapping"
    db.TableDefs.Delete "RowSource"
    db.TableDefs.Delete "QueryDefinitions"
    db.TableDefs.Refresh
    Debug.Print "info: done removing imported objects!"
    Exit Function
ErrorHandler:
    MsgBox "Error removing imported objects: " & Err.Description, vbCritical
    RemoveImportedObjects = False
End Function


Private Sub CloseIfOpen(ByVal tableName As String)
    On Error Resume Next
    If SysCmd(acSysCmdGetObjectState, acTable, tableName) <> 0 Then
        DoCmd.Close acTable, tableName, acSaveNo
    End If
    On Error GoTo 0
End Sub

Public Sub CloseOpenTables()
    On Error Resume Next
    Dim rs As DAO.Recordset
    Dim db As DAO.Database
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT DISTINCT OriginalTable FROM TranslationMapping", dbOpenSnapshot)
    Do Until rs.EOF
        CloseIfOpen rs!OriginalTable
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub

Public Sub PrintVersionInfo()
    Debug.Print "Access version: "; Application.Version        ' e.g. 16.0 for 2016/365
    Debug.Print "DB engine:    "; DBEngine.Version             ' e.g. 14.0+ supports calculated fields
    Debug.Print "File format:  "; CurrentProject.FileFormat    ' 12/ACCDB+ is needed; MDB won’t do it
End Sub

Public Function TranslateDatabase() As Boolean
    On Error GoTo ErrorHandler

    Dim isOK As Boolean

    Call PrintVersionInfo
    Call CloseOpenTables

    Let isOK = VerifyMappings

    Let isOK = isOK And ImportTranslationData
    ' Let isOK = isOK And ImportOrReplaceScript ' Can't replace itself!

    ' Let isOK = isOK And ShowRowSources
    Let isOK = isOK And DropOriginalExpressions
    Let isOK = isOK And RemoveOriginalRowSources
    Let isOK = isOK And RemoveOriginalQueries
    Let isOK = isOK And RenameFields
    Let isOK = isOK And RenameTables
    Let isOK = isOK And AddTranslatedRowSources
    Let isOK = isOK And AddTranslatedQueries
    Let isOK = isOK And AddTranslatedExpressions
    Let isOK = isOK And AddComments
    ' Let isOK = isOK And DeleteDeprecatedFields
    ' Let isOK = isOK And RemoveImportedObjects

    TranslateDatabase = isOK
    If isOK Then
        Debug.Print "info: database translation completed successfully!"
    Else
        Debug.Print "error: database translation completed with errors. See above."
    End If
    Exit Function
ErrorHandler:
    Debug.Print "error: Error translating database: " & Err.Description
    TranslateDatabase = False
End Function
