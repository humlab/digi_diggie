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
    Debug.Print "info: RowSource removed from " & tableName & "." & fieldName
    Exit Sub
EH:
    Debug.Print "error: failed to remove RowSource from " & tableName & "." & fieldName & ": " & Err.Description
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

    ' If properties don't exist yet, create them.
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
    Dim rs As DAO.Recordset
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim rowSource As String
    Dim property As DAO.property

    ShowRowSources = True
    For Each tdf In CurrentDb.TableDefs
        If IsDataTable(tdf.Name) Then
            For Each fld In tdf.Fields
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
    MsgBox "Error verifying original columns: " & Err.Description, vbCritical
    ShowRowSources = False
End Function

Public Function VerifyMappings() As Boolean
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim property As DAO.property

    VerifyMappings = True
    For Each tdf In CurrentDb.TableDefs
        If IsDataTable(tdf.Name) Then
            For Each fld In tdf.Fields
                If Not IsDataColumn(tdf.Name, fld.Name) Then
                    Debug.Print "error: NOT FOUND [" & tdf.Name & "] -> [" & fld.Name & "]"
                    VerifyMappings = False
                End If
            Next fld
        End If
    Next tdf
    If Not VerifyMappings Then
        Debug.Print "error: some original columns were not found. See above."
    Else
        Debug.Print "info: all original columns verified to be OK!"
    End If
    If Not VerifyMappings Then
        Err.Raise vbObjectError + 513, , "Some original columns were not found."
    End If
End Function

Public Function UpdateExpressions(Optional sCommand As String = "clear") As Boolean
    On Error GoTo ErrorHandler
    Dim rs As DAO.Recordset

    UpdateExpressions = True
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM TranslationMapping WHERE TranslatedExpression <> ''", dbOpenSnapshot)
    Do Until rs.EOF
        On Error Resume Next
        If sCommand = "clear" Then
            CurrentDb.TableDefs(rs!OriginalTable).Fields(rs!OriginalColumn).Expression = ""
        ElseIf sCommand = "set" Then
            CurrentDb.TableDefs(rs!TranslatedTable).Fields(rs!TranslatedColumn).Expression = rs!TranslatedExpression
        End If
        If Err.Number <> 0 Then
            UpdateExpressions = False
            Debug.Print "error: failed clearing computed expression for " & rs!TranslatedTable & "." & rs!TranslatedColumn & ": " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Exit Function
ErrorHandler:
    MsgBox "error: failed clearing computed expression: " & Err.Description, vbCritical
    UpdateExpressions = False
End Function


Public Function ClearComputedExpressions() As Boolean
    ClearComputedExpressions = UpdateExpressions("clear")
End Function

Public Function AddComputedExpressions() As Boolean
    AddComputedExpressions = UpdateExpressions("set")
End Function

Public RemoveQueries As Boolean
    On Error GoTo ErrorHandler
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef

    RemoveQueries = True
    Set db = CurrentDb
    For Each qdf In db.QueryDefs
        If Left(qdf.Name, 4) = "TRN_" Then
            db.QueryDefs.Delete qdf.Name
            Debug.Print "info: deleted query " & qdf.Name
        End If
    Next qdf
    Exit Function
ErrorHandler:
    MsgBox "Error removing queries: " & Err.Description, vbCritical
    RemoveQueries = False
End Function

Public Function ClearRowSource() As Boolean
    On Error GoTo ErrorHandler
    Dim rs As DAO.Recordset
    ClearRowSource = True
    Set rs = CurrentDb.OpenRecordset("SELECT [OriginalTable], [OriginalColumn] FROM [RowSource]", dbOpenSnapshot)
    Do Until rs.EOF
        On Error Resume Next
        CurrentDb.TableDefs(rs!OriginalTable).Fields(rs!OriginalColumn).Properties("RowSource") = "SELECT 1 AS DummyID, '' AS DummyValue WHERE False;"
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
    Debug.Print "info: Done!"
    Exit Function
ErrorHandler:
    MsgBox "Error removing RowSource: " & Err.Description, vbCritical
    ClearRowSource = False
End Function

Public Function RemoveRowSources() As Boolean
    On Error GoTo ErrorHandler
    Dim rs As DAO.Recordset
    RemoveRowSources = True
    Set rs = CurrentDb.OpenRecordset("SELECT [OriginalTable], [OriginalColumn] FROM [RowSource]", dbOpenSnapshot)
    Do Until rs.EOF
        RemoveFieldRowSource rs!OriginalTable, rs!OriginalColumn
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Debug.Print "info: Done!"
    Exit Function
ErrorHandler:
    MsgBox "error: failed removing RowSource: " & Err.Description, vbCritical
    RemoveRowSources = False
End Function

Public Function AddRowSources() As Boolean
    On Error GoTo ErrorHandler
    Dim rs As DAO.Recordset
    AddRowSources = True
    Set rs = CurrentDb.OpenRecordset("SELECT [TranslatedTable], [TranslatedColumn], [RowSource] FROM [RowSource]", dbOpenSnapshot)
    Do Until rs.EOF
        On Error Resume Next
        CurrentDb.TableDefs(rs!TranslatedTable).Fields(rs!TranslatedColumn).Properties("RowSource") = rs!rowSource
        If Err.Number <> 0 Then
            AddRowSources = False
            Debug.Print "error: failed setting RowSource for " & rs!TranslatedTable & "." & rs!TranslatedColumn & ": " & Err.Description
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
    AddRowSources = False
End Function


Public Function RenameFields() As Boolean
    On Error GoTo ErrorHandler
    Dim rs As DAO.Recordset
    RenameFields = True
    Set rs = CurrentDb.OpenRecordset("SELECT OriginalTable, OriginalColumn, TranslatedColumn FROM TranslationMapping", dbOpenSnapshot)
    Do Until rs.EOF
        On Error Resume Next
        CurrentDb.TableDefs(rs!OriginalTable).Fields(rs!OriginalColumn).Name = rs!TranslatedColumn
        If Err.Number = 3265 Then ' Item not found in this collection.
            ' Assume already renamed
        ElseIf Err.Number <> 0 Then
            Debug.Print "error: failed renaming field " & rs!OriginalTable & "." & rs!OriginalColumn & ": " & Err.Description
        End If
        Err.Clear
        On Error GoTo 0
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Debug.Print "info: Done!"
    Exit Function
ErrorHandler:
    MsgBox "Error renaming fields: " & Err.Description, vbCritical
    RenameFields = False
End Function


Public Function DeleteDeprecatedFields() As Boolean
    On Error GoTo ErrorHandler
    Dim rs As DAO.Recordset
    DeleteDeprecatedFields = True
    Set rs = CurrentDb.OpenRecordset("SELECT OriginalTable, OriginalColumn, TranslatedColumn FROM TranslationMapping WHERE DeprecatedFlag = 'YES'", dbOpenSnapshot)
    Do Until rs.EOF
        On Error Resume Next
        CurrentDb.TableDefs(rs!OriginalTable).Fields(rs!OriginalColumn).Delete
        If Err.Number <> 0 Then
            Debug.Print "error: failed deleting field " & rs!OriginalTable & "." & rs!OriginalColumn & ": " & Err.Description
        End If
        Err.Clear
        On Error GoTo 0
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Debug.Print "info:Done!"
    Exit Function
ErrorHandler:
    MsgBox "Error deleting fields: " & Err.Description, vbCritical
    DeleteDeprecatedFields = False
End Function

Public Function RenameTables() As Boolean
    On Error GoTo ErrorHandler
    Dim rs As DAO.Recordset
    RenameTables = True
    Set rs = CurrentDb.OpenRecordset("SELECT DISTINCT OriginalTable, TranslatedTable FROM TranslationMapping", dbOpenSnapshot)
    Do Until rs.EOF
        On Error Resume Next
        CurrentDb.TableDefs(rs!OriginalTable).Name = rs!TranslatedTable
        If Err.Number <> 0 Then
            Debug.Print "error: failed renaming table " & rs!OriginalTable & " to " & rs!TranslatedTable & ": " & Err.Description
        End If
        Err.Clear
        On Error GoTo 0
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Debug.Print "info: Done!"
    Exit Function
ErrorHandler:
    MsgBox "Error renaming fields: " & Err.Description, vbCritical
    RenameTables = False
End Function

Public Function RemoveQueries() As Boolean
    On Error GoTo ErrorHandler
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef

    RemoveQueries = True
    Set db = CurrentDb
    For Each qdf In db.QueryDefs
        If Left(qdf.Name, 4) = "TRN_" Then
            db.QueryDefs.Delete qdf.Name
            Debug.Print "info: deleted query " & qdf.Name
        End If
    Next qdf
    Exit Function
ErrorHandler:
    MsgBox "Error removing queries: " & Err.Description, vbCritical
    RemoveQueries = False
End Function


Public Function AddQueries() As Boolean
    On Error GoTo ErrorHandler
    Dim rs As DAO.Recordset
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef

    AddQueries = True
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM QueryDefinitions", dbOpenSnapshot)
    Do Until rs.EOF
        On Error Resume Next
        db.QueryDefs.Delete rs!QueryName
        Err.Clear
        Set qdf = db.CreateQueryDef(rs!QueryName, rs!SQLText)
        If Err.Number <> 0 Then
            Debug.Print "error: error adding query " & rs!QueryName & ": " & Err.Description
            AddQueries = False
            Err.Clear
        End If
        On Error GoTo 0
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Debug.Print "info: Done!"
    Exit Function
ErrorHandler:
    MsgBox "Error adding queries: " & Err.Description, vbCritical
    AddQueries = False
End Function

Public Function RemoveImportedObjects() As Boolean
    On Error GoTo ErrorHandler
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef

    RemoveImportedObjects = True
    Set db = CurrentDb
    db.TableDefs.Delete "TranslationMapping"
    db.TableDefs.Delete "RowSource"
    db.TableDefs.Delete "QueryDefinitions"

    Debug.Print "info: deleted mapping tables"
    Exit Function
ErrorHandler:
    MsgBox "Error removing imported objects: " & Err.Description, vbCritical
    RemoveImportedObjects = False
End Function

Public Function TranslateDatabase() As Boolean
    On Error GoTo ErrorHandler

    Dim isOK As Boolean

    Let isOK = VerifyMappings

    Let isOK = isOK And ImportTranslationMappingExcelSheet
    ' Let isOK = isOK And ImportOrReplaceScript ' Can't replace itself!

    Let isOK = isOK And ShowRowSources
    Let isOK = isOK And ClearComputedExpressions
    Let isOK = isOK And RemoveRowSources
    Let isOK = isOK And RemoveQueries
    Let isOK = isOK And DeleteDeprecatedFields
    Let isOK = isOK And RenameFields
    Let isOK = isOK And RenameTables
    Let isOK = isOK And AddRowSources
    Let isOK = isOK And AddComputedExpressions
    Let isOK = isOK And AddQueries

    Let isOK = isOK And RemoveImportedObjects

    TranslateDatabase = isOK
    MsgBox "Table and column translation complete.", vbInformation
    Exit Function
ErrorHandler:
    MsgBox "Error translating database: " & Err.Description, vbCritical
    TranslateDatabase = False
End Function
