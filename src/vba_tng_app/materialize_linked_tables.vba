Option Compare Database
Option Explicit

' ============================================================
' Materialize all linked PostgreSQL tables into local Access tables
' while preserving PK values and recreating FK relations.
'
' Main entry point:
'   MaterializeAllPostgresLinkedTables
'
' Recommended:
'   1. Back up your .accdb first
'   2. Run from Immediate Window:
'        MaterializeAllPostgresLinkedTables
'
' Optional:
'        MaterializeAllPostgresLinkedTables "loc_", False
'
' Parameters:
'   localPrefix      Prefix for created local tables
'   dropIfExists     True = delete existing local table if present
'
' Notes:
'   - Single-column PKs become AutoNumber locally
'   - Existing PK values are preserved during append
'   - FK relations are recreated after data load
'   - Composite PKs/FKs are skipped
' ============================================================

Public Sub MaterializeAllPostgresLinkedTables( _
    Optional ByVal localPrefix As String = "loc_", _
    Optional ByVal dropIfExists As Boolean = False, _
    Optional ByVal showMessageBoxes As Boolean = True)

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim linkedTables As Collection
    Dim tableMap As Object               ' Scripting.Dictionary: linkedName -> localName
    Dim remoteMap As Object              ' Scripting.Dictionary: linkedName -> Array(schema, table)
    Dim connMap As Object                ' Scripting.Dictionary: linkedName -> connect string
    Dim pkMap As Object                  ' Scripting.Dictionary: linkedName -> pk column or ""
    Dim msg As String
    Dim linkedName As Variant
    Dim localName As String
    Dim remoteSchema As String
    Dim remoteTable As String
    Dim connStr As String
    Dim pkCol As String

    Set db = CurrentDb
    Set linkedTables = New Collection
    Set tableMap = CreateObject("Scripting.Dictionary")
    Set remoteMap = CreateObject("Scripting.Dictionary")
    Set connMap = CreateObject("Scripting.Dictionary")
    Set pkMap = CreateObject("Scripting.Dictionary")

    ' 1) Find linked PostgreSQL tables
    For Each tdf In db.TableDefs
        If IsLinkedPostgresTable(tdf) Then
            linkedTables.Add tdf.Name

            connStr = CleanOdbcConnectString(tdf.Connect)

            ParseRemoteTableName tdf.SourceTableName, remoteSchema, remoteTable
            If Len(remoteSchema) = 0 Then remoteSchema = "public"

            localName = localPrefix & tdf.Name

            tableMap.Add tdf.Name, localName
            remoteMap.Add tdf.Name, Array(remoteSchema, remoteTable)
            connMap.Add tdf.Name, connStr
        End If
    Next tdf

    If linkedTables.Count = 0 Then
        MsgBox "No linked PostgreSQL tables were found.", vbInformation
        Exit Sub
    End If

    ' 2) Discover PKs
    For Each linkedName In linkedTables
        remoteSchema = remoteMap(linkedName)(0)
        remoteTable = remoteMap(linkedName)(1)
        connStr = connMap(linkedName)

        pkCol = GetSinglePrimaryKeyColumn(connStr, remoteSchema, remoteTable)
        pkMap.Add linkedName, pkCol
    Next linkedName

    ' 3) Create local tables
    For Each linkedName In linkedTables
        localName = tableMap(linkedName)

        If TableExists(localName) Then
            If dropIfExists Then
                DropTable localName
            Else
                Debug.Print "Skipping existing local table: " & localName
                GoTo ContinueCreateLoop
            End If
        End If

        CreateLocalTableFromLinked _
            linkedTableName:=CStr(linkedName), _
            localTableName:=localName, _
            pkColumnName:=pkMap(linkedName)

        Debug.Print "Created local table: " & localName

ContinueCreateLoop:
    Next linkedName

    ' 4) Append data
    For Each linkedName In linkedTables
        localName = tableMap(linkedName)

        If TableExists(localName) Then
            AppendLinkedTableToLocal CStr(linkedName), localName
            Debug.Print "Copied data: " & linkedName & " -> " & localName
        End If
    Next linkedName

    ' 5) Recreate foreign-key relations locally
    For Each linkedName In linkedTables
        remoteSchema = remoteMap(linkedName)(0)
        remoteTable = remoteMap(linkedName)(1)
        connStr = connMap(linkedName)

        CreateLocalForeignKeysFromPostgres _
            connStr:=connStr, _
            remoteSchema:=remoteSchema, _
            remoteTable:=remoteTable, _
            linkedTableName:=CStr(linkedName), _
            localTableName:=tableMap(linkedName), _
            tableMap:=tableMap, _
            remoteMap:=remoteMap
    Next linkedName

    msg = "Finished materializing PostgreSQL linked tables to local Access tables." & vbCrLf & vbCrLf & _
          "Local prefix: " & localPrefix & vbCrLf & _
          "Tables processed: " & linkedTables.Count & vbCrLf & vbCrLf & _
          "See Immediate Window (Ctrl+G) for details."
    If showMessageBoxes Then
        MsgBox msg, vbInformation
    End If
End Sub

' ============================================================
' Remove all linked tables and remove "loc_" prefix from local tables
' ============================================================
Public Sub UnlinkTablesAndRemovePrefix( _
    Optional ByVal localPrefix As String = "loc_", _
    Optional ByVal showMessageBoxes As Boolean = True)

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim linkedTables As Collection
    Dim localTables As Collection
    Dim tableName As Variant
    Dim newName As String
    Dim msg As String
    Dim linkedCount As Integer
    Dim renamedCount As Integer

    Set db = CurrentDb
    Set linkedTables = New Collection
    Set localTables = New Collection
    linkedCount = 0
    renamedCount = 0

    ' Confirm with user before proceeding
    If showMessageBoxes Then
        If MsgBox("This will:" & vbNewLine & vbNewLine & _
                  "1. Delete ALL linked PostgreSQL tables" & vbNewLine & _
                  "2. Remove '" & localPrefix & "' prefix from local tables" & vbNewLine & vbNewLine & _
                  "This action cannot be undone. Continue?", _
                  vbYesNo + vbExclamation + vbDefaultButton2, _
                  "Unlink Tables and Remove Prefix") = vbNo Then
            MsgBox "Operation cancelled by user.", vbInformation
            Exit Sub
        End If
    End If

    ' 1) Find linked PostgreSQL tables and local tables with prefix
    For Each tdf In db.TableDefs
        If IsLinkedPostgresTable(tdf) Then
            linkedTables.Add tdf.Name
        ElseIf Left(tdf.Name, Len(localPrefix)) = localPrefix Then
            ' Check if it's a user table (not system table)
            If IsUserTable(tdf) Then
                localTables.Add tdf.Name
            End If
        End If
    Next tdf

    ' Show what will be processed
    Debug.Print "=== UNLINK TABLES AND REMOVE PREFIX ==="
    Debug.Print "Found " & linkedTables.Count & " linked PostgreSQL tables to remove"
    Debug.Print "Found " & localTables.Count & " local tables with prefix '" & localPrefix & "' to rename"
    Debug.Print ""

    ' 2) Remove all linked tables
    If linkedTables.Count > 0 Then
        Debug.Print "Removing linked tables:"
        For Each tableName In linkedTables
            Debug.Print "  Removing: " & tableName
            DropTable CStr(tableName)
            linkedCount = linkedCount + 1
        Next tableName
        Debug.Print ""
    End If

    ' 3) Rename local tables to remove prefix
    If localTables.Count > 0 Then
        Debug.Print "Renaming local tables (removing prefix):"
        For Each tableName In localTables
            newName = Mid(CStr(tableName), Len(localPrefix) + 1)
            
            ' Check if target name already exists
            If TableExists(newName) Then
                Debug.Print "  SKIPPED: " & tableName & " -> " & newName & " (target already exists)"
            Else
                Debug.Print "  Renaming: " & tableName & " -> " & newName
                RenameTable CStr(tableName), newName
                renamedCount = renamedCount + 1
            End If
        Next tableName
    End If

    Debug.Print ""
    Debug.Print "Operation completed:"
    Debug.Print "  Linked tables removed: " & linkedCount
    Debug.Print "  Local tables renamed: " & renamedCount

    msg = "Unlink and rename operation completed!" & vbCrLf & vbCrLf & _
          "Linked tables removed: " & linkedCount & vbCrLf & _
          "Local tables renamed: " & renamedCount & vbCrLf & vbCrLf & _
          "See Immediate Window (Ctrl+G) for details."
    If showMessageBoxes Then
        MsgBox msg, vbInformation
    End If
End Sub

Private Function IsUserTable(ByVal tdf As DAO.TableDef) As Boolean
    ' Check if this is a user table (not system, hidden, or linked)
    Dim isSystem As Boolean
    Dim isHidden As Boolean
    Dim isLinked As Boolean
    
    isLinked = (Len(tdf.Connect) > 0)
    isSystem = (Left(tdf.Name, 4) = "MSys") Or ((tdf.Attributes And dbSystemObject) <> 0)
    isHidden = ((tdf.Attributes And dbHiddenObject) <> 0)
    
    IsUserTable = (Not isSystem) And (Not isHidden) And (Not isLinked)
End Function

Private Sub RenameTable(ByVal oldName As String, ByVal newName As String)
    ' Rename a table using DoCmd.Rename
    On Error GoTo ErrorHandler
    
    DoCmd.Rename newName, acTable, oldName
    Exit Sub
    
ErrorHandler:
    Debug.Print "  ERROR renaming " & oldName & " to " & newName & ": " & Err.Description
End Sub

' ============================================================
' Create a local table from a linked table definition
' ============================================================
Private Sub CreateLocalTableFromLinked( _
    ByVal linkedTableName As String, _
    ByVal localTableName As String, _
    ByVal pkColumnName As String)

    Dim db As DAO.Database
    Dim srcTdf As DAO.TableDef
    Dim newTdf As DAO.TableDef
    Dim srcFld As DAO.Field
    Dim newFld As DAO.Field
    Dim idx As DAO.Index
    Dim fldType As Integer

    Set db = CurrentDb
    Set srcTdf = db.TableDefs(linkedTableName)
    Set newTdf = db.CreateTableDef(localTableName)

    For Each srcFld In srcTdf.Fields
        fldType = LocalFieldTypeFromLinkedField(srcFld)

        If StrComp(srcFld.Name, pkColumnName, vbTextCompare) = 0 Then
            ' Make the PK field AutoNumber locally
            Set newFld = newTdf.CreateField(srcFld.Name, dbLong)
            newFld.Attributes = dbAutoIncrField
        Else
            Set newFld = newTdf.CreateField(srcFld.Name, fldType)

            If fldType = dbText Then
                On Error Resume Next
                newFld.Size = srcFld.Size
                newFld.AllowZeroLength = True
                On Error GoTo 0
            End If

            If fldType = dbMemo Then
                On Error Resume Next
                newFld.AllowZeroLength = True
                On Error GoTo 0
            End If
        End If

        newTdf.Fields.Append newFld
    Next srcFld

    db.TableDefs.Append newTdf

    ' Create primary key index if single-column PK exists
    If Len(pkColumnName) > 0 Then
        Set idx = db.TableDefs(localTableName).CreateIndex("pk_" & SafeObjectName(localTableName))
        idx.Primary = True
        idx.Unique = True
        idx.Fields.Append idx.CreateField(pkColumnName)
        db.TableDefs(localTableName).Indexes.Append idx
    End If
End Sub

' ============================================================
' Append all rows from linked table into local table
' ============================================================
Private Sub AppendLinkedTableToLocal( _
    ByVal linkedTableName As String, _
    ByVal localTableName As String)

    Dim db As DAO.Database
    Dim sql As String
    Dim fieldList As String

    Set db = CurrentDb
    fieldList = BuildSharedFieldList(linkedTableName, localTableName)

    sql = "INSERT INTO " & BracketName(localTableName) & " (" & fieldList & ") " & _
          "SELECT " & fieldList & " FROM " & BracketName(linkedTableName) & ";"

    db.Execute sql, dbFailOnError
End Sub

' ============================================================
' Recreate local FK relations from PostgreSQL metadata
' Single-column FKs only
' ============================================================
Private Sub CreateLocalForeignKeysFromPostgres( _
    ByVal connStr As String, _
    ByVal remoteSchema As String, _
    ByVal remoteTable As String, _
    ByVal linkedTableName As String, _
    ByVal localTableName As String, _
    ByVal tableMap As Object, _
    ByVal remoteMap As Object)

    Dim rs As DAO.Recordset
    Dim sql As String
    Dim fkColumn As String
    Dim pkSchema As String
    Dim pkTable As String
    Dim pkColumn As String
    Dim parentLinked As Variant
    Dim parentLocal As String
    Dim relName As String
    Dim db As DAO.Database
    Dim rel As DAO.Relation

    Set db = CurrentDb

    sql = ""
    sql = sql & "SELECT " & vbCrLf
    sql = sql & "  kcu.column_name AS fk_column, " & vbCrLf
    sql = sql & "  ccu.table_schema AS pk_schema, " & vbCrLf
    sql = sql & "  ccu.table_name AS pk_table, " & vbCrLf
    sql = sql & "  ccu.column_name AS pk_column, " & vbCrLf
    sql = sql & "  tc.constraint_name " & vbCrLf
    sql = sql & "FROM information_schema.table_constraints tc " & vbCrLf
    sql = sql & "JOIN information_schema.key_column_usage kcu " & vbCrLf
    sql = sql & "  ON tc.constraint_name = kcu.constraint_name " & vbCrLf
    sql = sql & " AND tc.table_schema = kcu.table_schema " & vbCrLf
    sql = sql & "JOIN information_schema.constraint_column_usage ccu " & vbCrLf
    sql = sql & "  ON ccu.constraint_name = tc.constraint_name " & vbCrLf
    sql = sql & " AND ccu.table_schema = tc.table_schema " & vbCrLf
    sql = sql & "WHERE tc.constraint_type = 'FOREIGN KEY' " & vbCrLf
    sql = sql & "  AND tc.table_schema = " & SqlText(remoteSchema) & vbCrLf
    sql = sql & "  AND tc.table_name = " & SqlText(remoteTable) & vbCrLf
    sql = sql & "ORDER BY tc.constraint_name, kcu.ordinal_position;"

    Set rs = OpenPassThroughRecordset(connStr, sql)

    Do While Not rs.EOF
        fkColumn = Nz(rs!fk_column, "")
        pkSchema = Nz(rs!pk_schema, "")
        pkTable = Nz(rs!pk_table, "")
        pkColumn = Nz(rs!pk_column, "")

        parentLinked = FindLinkedTableByRemoteName(remoteMap, pkSchema, pkTable)

        If Not IsEmpty(parentLinked) Then
            parentLocal = tableMap(parentLinked)

            If TableExists(localTableName) And TableExists(parentLocal) Then
                If FieldExists(localTableName, fkColumn) And FieldExists(parentLocal, pkColumn) Then
                    relName = BuildRelationName(localTableName, parentLocal, fkColumn)

                    If Not RelationExists(relName) Then
                        Set rel = db.CreateRelation(relName, parentLocal, localTableName)
                        rel.Fields.Append rel.CreateField(pkColumn)
                        rel.Fields(pkColumn).ForeignName = fkColumn

                        On Error Resume Next
                        db.Relations.Append rel
                        If Err.Number <> 0 Then
                            Debug.Print "Could not create relation " & relName & ": " & Err.Description
                            Err.Clear
                        Else
                            Debug.Print "Created relation: " & relName
                        End If
                        On Error GoTo 0
                    End If
                End If
            End If
        End If

        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
End Sub

' ============================================================
' Find single-column primary key in PostgreSQL metadata
' Returns "" if none or composite PK
' ============================================================
Private Function GetSinglePrimaryKeyColumn( _
    ByVal connStr As String, _
    ByVal remoteSchema As String, _
    ByVal remoteTable As String) As String

    Dim rs As DAO.Recordset
    Dim sql As String
    Dim countCols As Long
    Dim pkCol As String

    sql = ""
    sql = sql & "SELECT kcu.column_name " & vbCrLf
    sql = sql & "FROM information_schema.table_constraints tc " & vbCrLf
    sql = sql & "JOIN information_schema.key_column_usage kcu " & vbCrLf
    sql = sql & "  ON tc.constraint_name = kcu.constraint_name " & vbCrLf
    sql = sql & " AND tc.table_schema = kcu.table_schema " & vbCrLf
    sql = sql & "WHERE tc.constraint_type = 'PRIMARY KEY' " & vbCrLf
    sql = sql & "  AND tc.table_schema = " & SqlText(remoteSchema) & vbCrLf
    sql = sql & "  AND tc.table_name = " & SqlText(remoteTable) & vbCrLf
    sql = sql & "ORDER BY kcu.ordinal_position;"

    Set rs = OpenPassThroughRecordset(connStr, sql)

    countCols = 0
    pkCol = ""

    Do While Not rs.EOF
        countCols = countCols + 1
        pkCol = Nz(rs.Fields(0).Value, "")
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing

    If countCols = 1 Then
        GetSinglePrimaryKeyColumn = pkCol
    Else
        GetSinglePrimaryKeyColumn = ""
    End If
End Function

' ============================================================
' Open a pass-through recordset against PostgreSQL
' ============================================================
Private Function OpenPassThroughRecordset( _
    ByVal connStr As String, _
    ByVal sql As String) As DAO.Recordset

    Dim qdf As DAO.QueryDef

    Set qdf = CurrentDb.CreateQueryDef("")
    qdf.Connect = connStr
    qdf.ReturnsRecords = True
    qdf.SQL = sql

    Set OpenPassThroughRecordset = qdf.OpenRecordset(dbOpenSnapshot)
End Function

' ============================================================
' Determine whether a TableDef is a linked PostgreSQL table
' ============================================================
Private Function IsLinkedPostgresTable(ByVal tdf As DAO.TableDef) As Boolean
    Dim c As String

    c = Nz(tdf.Connect, "")

    If Len(c) = 0 Then
        IsLinkedPostgresTable = False
        Exit Function
    End If

    If (tdf.Attributes And dbAttachedODBC) = 0 Then
        IsLinkedPostgresTable = False
        Exit Function
    End If

    IsLinkedPostgresTable = _
        (InStr(1, c, "ODBC", vbTextCompare) > 0)
End Function

' ============================================================
' Parse remote source table name into schema + table
' Examples:
'   public.mytable -> public / mytable
'   mytable        -> public / mytable
'   "public"."x"   -> public / x
' ============================================================
Private Sub ParseRemoteTableName( _
    ByVal sourceTableName As String, _
    ByRef remoteSchema As String, _
    ByRef remoteTable As String)

    Dim s As String
    Dim p As Long

    s = Replace(sourceTableName, """", "")

    p = InStr(1, s, ".", vbTextCompare)

    If p > 0 Then
        remoteSchema = Left$(s, p - 1)
        remoteTable = Mid$(s, p + 1)
    Else
        remoteSchema = "public"
        remoteTable = s
    End If
End Sub

' ============================================================
' Remove TABLE=... from Access linked-table ODBC connect string
' so it can be reused in pass-through queries
' ============================================================
Private Function CleanOdbcConnectString(ByVal connectString As String) As String
    Dim parts() As String
    Dim i As Long
    Dim out As String
    Dim part As String

    parts = Split(connectString, ";")
    out = ""

    For i = LBound(parts) To UBound(parts)
        part = Trim$(parts(i))
        If Len(part) > 0 Then
            If InStr(1, part, "TABLE=", vbTextCompare) = 0 And _
               InStr(1, part, "TABLE NAME=", vbTextCompare) = 0 Then
                If Len(out) > 0 Then out = out & ";"
                out = out & part
            End If
        End If
    Next i

    CleanOdbcConnectString = out
End Function

' ============================================================
' Map linked field type to a local Access field type
' ============================================================
Private Function LocalFieldTypeFromLinkedField(ByVal fld As DAO.Field) As Integer
    Select Case fld.Type
        Case dbBoolean
            LocalFieldTypeFromLinkedField = dbBoolean

        Case dbByte
            LocalFieldTypeFromLinkedField = dbByte

        Case dbInteger
            LocalFieldTypeFromLinkedField = dbInteger

        Case dbLong
            LocalFieldTypeFromLinkedField = dbLong

        Case dbSingle
            LocalFieldTypeFromLinkedField = dbSingle

        Case dbDouble
            LocalFieldTypeFromLinkedField = dbDouble

        Case dbCurrency
            LocalFieldTypeFromLinkedField = dbCurrency

        Case dbDate
            LocalFieldTypeFromLinkedField = dbDate

        Case dbText
            LocalFieldTypeFromLinkedField = dbText

        Case dbMemo
            LocalFieldTypeFromLinkedField = dbMemo

        Case dbLongBinary
            LocalFieldTypeFromLinkedField = dbLongBinary

        Case dbGUID
            LocalFieldTypeFromLinkedField = dbGUID

        Case Else
            ' Fallback
            LocalFieldTypeFromLinkedField = dbText
    End Select
End Function

' ============================================================
' Build shared field list between source and target tables
' ============================================================
Private Function BuildSharedFieldList( _
    ByVal sourceTableName As String, _
    ByVal targetTableName As String) As String

    Dim db As DAO.Database
    Dim srcTdf As DAO.TableDef
    Dim tgtTdf As DAO.TableDef
    Dim srcFld As DAO.Field
    Dim s As String

    Set db = CurrentDb
    Set srcTdf = db.TableDefs(sourceTableName)
    Set tgtTdf = db.TableDefs(targetTableName)

    s = ""

    For Each srcFld In srcTdf.Fields
        If FieldExists(targetTableName, srcFld.Name) Then
            If Len(s) > 0 Then s = s & ", "
            s = s & BracketName(srcFld.Name)
        End If
    Next srcFld

    BuildSharedFieldList = s
End Function

' ============================================================
' Find linked table by remote schema/table
' ============================================================
Private Function FindLinkedTableByRemoteName( _
    ByVal remoteMap As Object, _
    ByVal remoteSchema As String, _
    ByVal remoteTable As String) As Variant

    Dim k As Variant

    For Each k In remoteMap.Keys
        If StrComp(remoteMap(k)(0), remoteSchema, vbTextCompare) = 0 And _
           StrComp(remoteMap(k)(1), remoteTable, vbTextCompare) = 0 Then
            FindLinkedTableByRemoteName = k
            Exit Function
        End If
    Next k

    FindLinkedTableByRemoteName = Empty
End Function

' ============================================================
' Helpers
' ============================================================
Private Function TableExists(ByVal tableName As String) As Boolean
    On Error Resume Next
    TableExists = (Len(CurrentDb.TableDefs(tableName).Name) > 0)
    If Err.Number <> 0 Then
        TableExists = False
        Err.Clear
    End If
    On Error GoTo 0
End Function

Private Function FieldExists(ByVal tableName As String, ByVal fieldName As String) As Boolean
    On Error Resume Next
    FieldExists = (Len(CurrentDb.TableDefs(tableName).Fields(fieldName).Name) > 0)
    If Err.Number <> 0 Then
        FieldExists = False
        Err.Clear
    End If
    On Error GoTo 0
End Function

Private Function RelationExists(ByVal relationName As String) As Boolean
    On Error Resume Next
    RelationExists = (Len(CurrentDb.Relations(relationName).Name) > 0)
    If Err.Number <> 0 Then
        RelationExists = False
        Err.Clear
    End If
    On Error GoTo 0
End Function

Private Sub DropTable(ByVal tableName As String)
    On Error Resume Next
    DoCmd.DeleteObject acTable, tableName
    On Error GoTo 0
End Sub

Private Function BracketName(ByVal objectName As String) As String
    BracketName = "[" & Replace(objectName, "]", "]]") & "]"
End Function

Private Function SqlText(ByVal s As String) As String
    SqlText = "'" & Replace(s, "'", "''") & "'"
End Function

Private Function SafeObjectName(ByVal s As String) As String
    Dim t As String
    t = s
    t = Replace(t, " ", "_")
    t = Replace(t, ".", "_")
    t = Replace(t, "-", "_")
    t = Replace(t, "/", "_")
    t = Replace(t, "\", "_")
    SafeObjectName = t
End Function

Private Function BuildRelationName( _
    ByVal childTable As String, _
    ByVal parentTable As String, _
    ByVal childField As String) As String

    BuildRelationName = Left$("rel_" & _
        SafeObjectName(childTable) & "_" & _
        SafeObjectName(parentTable) & "_" & _
        SafeObjectName(childField), 60)
End Function

' ============================================================
' Create all indexes and foreign key constraints for DigiDiggie schema
' ============================================================
Public Sub CreateIndexesAndConstraints( _
    Optional ByVal showMessageBoxes As Boolean = True)
    ' Creates all indexes and foreign key relationships for the DigiDiggie TNG schema
    ' Call this after tables are created to optimize performance and enforce referential integrity
    
    Dim indexCount As Integer
    Dim relationCount As Integer
    Dim msg As String
    
    ' Confirm with user before proceeding
    If showMessageBoxes Then
        If MsgBox("Create all indexes and foreign key constraints for DigiDiggie TNG schema?" & vbNewLine & vbNewLine & _
                  "This will:" & vbNewLine & _
                  "1. Create performance indexes on key fields" & vbNewLine & _
                  "2. Establish foreign key relationships" & vbNewLine & _
                  "3. Enforce referential integrity", _
                  vbYesNo + vbQuestion, _
                  "Create Indexes and Constraints") = vbNo Then
            MsgBox "Operation cancelled by user.", vbInformation
            Exit Sub
        End If
    End If
    
    Debug.Print "=== CREATING INDEXES AND CONSTRAINTS ==="
    Debug.Print "Starting index and constraint creation..."
    Debug.Print ""
    
    ' Create indexes first
    indexCount = CreateAllIndexes()
    
    ' Then create foreign key relationships
    relationCount = CreateAllForeignKeyConstraints()
    
    Debug.Print ""
    Debug.Print "Operation completed:"
    Debug.Print "  Indexes created: " & indexCount
    Debug.Print "  Foreign key relationships created: " & relationCount
    
    msg = "Indexes and constraints created successfully!" & vbCrLf & vbCrLf & _
          "Indexes created: " & indexCount & vbCrLf & _
          "Foreign key relationships: " & relationCount & vbCrLf & vbCrLf & _
          "See Immediate Window (Ctrl+G) for details."
    If showMessageBoxes Then
        MsgBox msg, vbInformation
    End If
End Sub

' ============================================================
' Create all performance indexes
' ============================================================
Private Function CreateAllIndexes() As Integer
    ' Creates all indexes from the DigiDiggie schema
    ' Returns count of indexes successfully created
    
    Dim db As DAO.Database
    Dim successCount As Integer
    
    Set db = CurrentDb
    successCount = 0
    
    Debug.Print "Creating performance indexes..."
    
    ' Community indexes
    If CreateIndex("community", "communities_parish_id_idx", "parish_id", False) Then successCount = successCount + 1
    
    ' Court case indexes
    If CreateIndex("court_case", "court_cases_case_year_idx", "case_year", False) Then successCount = successCount + 1
    If CreateIndex("court_case", "court_cases_reference_number_idx", "reference_number", False) Then successCount = successCount + 1  
    If CreateIndex("court_case", "court_cases_source_id_idx", "source_id", False) Then successCount = successCount + 1
    
    ' Court case entry indexes
    If CreateIndex("court_case_entry", "entries_court_case_id_idx", "court_case_id", False) Then successCount = successCount + 1
    If CreateIndex("court_case_entry", "entries_land_use_id_idx", "land_use_id", False) Then successCount = successCount + 1
    If CreateIndex("court_case_entry", "entries_placename_id_idx", "placename_id", False) Then successCount = successCount + 1
    If CreateIndex("court_case_entry", "entries_season_id_idx", "season_id", False) Then successCount = successCount + 1
    
    ' Ruling indexes
    If CreateIndex("ruling", "rulings_legal_source_id_idx", "legal_source_id", False) Then successCount = successCount + 1
    
    ' Parish indexes
    If CreateIndex("parish", "parishes_parish_idx", "parish", False) Then successCount = successCount + 1
    
    ' Person entry indexes
    If CreateIndex("person_entry", "person_entries_person_id_idx", "person_id", False) Then successCount = successCount + 1
    If CreateIndex("person_entry", "person_entries_community_id_idx", "community_id", False) Then successCount = successCount + 1
    If CreateIndex("person_entry", "person_entries_court_case_entry_id_idx", "court_case_entry_id", False) Then successCount = successCount + 1
    
    ' Person outcome indexes
    If CreateIndex("person_outcome", "person_outcomes_outcome_type_id_idx", "outcome_type_id", False) Then successCount = successCount + 1
    If CreateIndex("person_outcome", "person_outcomes_person_id_idx", "person_id", False) Then successCount = successCount + 1
    If CreateIndex("person_outcome", "person_outcomes_ruling_id_idx", "ruling_id", False) Then successCount = successCount + 1
    
    ' Season indexes
    If CreateIndex("season", "seasons_season_name_idx", "season_name", False) Then successCount = successCount + 1
    
    CreateAllIndexes = successCount
End Function

' ============================================================
' Create all foreign key constraints
' ============================================================
Private Function CreateAllForeignKeyConstraints() As Integer
    ' Creates all foreign key relationships from the DigiDiggie schema
    ' Returns count of relationships successfully created
    
    Dim db As DAO.Database
    Dim successCount As Integer
    
    Set db = CurrentDb
    successCount = 0
    
    Debug.Print ""
    Debug.Print "Creating foreign key relationships..."
    
    ' Level 1: Simple lookups (no dependencies on other foreign tables)
    If CreateFKRelation(db, "community", "parish_id", "parish", "parish_id", "community_parish_fk") Then successCount = successCount + 1
    If CreateFKRelation(db, "court_case", "source_id", "source", "source_id", "court_case_source_fk") Then successCount = successCount + 1
    If CreateFKRelation(db, "role", "role_type_id", "role_type", "role_type_id", "role_role_type_fk") Then successCount = successCount + 1
    If CreateFKRelation(db, "ruling", "ruling_type_id", "ruling_type", "ruling_type_id", "ruling_ruling_type_fk") Then successCount = successCount + 1
    If CreateFKRelation(db, "ruling", "legal_source_id", "legal_source", "legal_source_id", "ruling_legal_source_fk") Then successCount = successCount + 1
    If CreateFKRelation(db, "person_relationship", "relationship_type_id", "relationship_type", "relationship_type_id", "person_rel_type_fk") Then successCount = successCount + 1
    
    ' Level 2: Tables that depend on Level 1
    If CreateFKRelation(db, "ruling", "court_case_id", "court_case", "court_case_id", "ruling_court_case_fk") Then successCount = successCount + 1
    If CreateFKRelation(db, "court_case_entry", "court_case_id", "court_case", "court_case_id", "entry_court_case_fk") Then successCount = successCount + 1
    If CreateFKRelation(db, "court_case_entry", "land_use_id", "land_use", "land_use_id", "entry_land_use_fk") Then successCount = successCount + 1
    If CreateFKRelation(db, "court_case_entry", "season_id", "season", "season_id", "entry_season_fk") Then successCount = successCount + 1
    If CreateFKRelation(db, "court_case_entry", "placename_id", "placename", "placename_id", "entry_placename_fk") Then successCount = successCount + 1
    
    ' Level 3: Tables that depend on Level 2
    If CreateFKRelation(db, "person_entry", "person_id", "person", "person_id", "person_entry_person_fk") Then successCount = successCount + 1
    If CreateFKRelation(db, "person_entry", "community_id", "community", "community_id", "person_entry_community_fk") Then successCount = successCount + 1
    If CreateFKRelation(db, "person_entry", "court_case_entry_id", "court_case_entry", "court_case_entry_id", "person_entry_case_entry_fk") Then successCount = successCount + 1
    If CreateFKRelation(db, "person_entry", "role_id", "role", "role_id", "person_entry_role_fk") Then successCount = successCount + 1
    If CreateFKRelation(db, "person_entry", "land_rights_status_id", "land_rights_status", "land_rights_status_id", "person_entry_land_rights_fk") Then successCount = successCount + 1
    
    ' Level 4: Tables that depend on Level 3  
    If CreateFKRelation(db, "person_outcome", "ruling_id", "ruling", "ruling_id", "person_outcome_ruling_fk") Then successCount = successCount + 1
    If CreateFKRelation(db, "person_outcome", "person_id", "person", "person_id", "person_outcome_person_fk") Then successCount = successCount + 1
    If CreateFKRelation(db, "person_outcome", "outcome_type_id", "outcome_type", "outcome_type_id", "person_outcome_type_fk") Then successCount = successCount + 1
    
    ' Person relationships (self-referencing)
    If CreateFKRelation(db, "person_relationship", "person_1_id", "person", "person_id", "person_rel_person1_fk") Then successCount = successCount + 1
    If CreateFKRelation(db, "person_relationship", "person_2_id", "person", "person_id", "person_rel_person2_fk") Then successCount = successCount + 1
    
    CreateAllForeignKeyConstraints = successCount
End Function

' ============================================================
' Helper function to create a single index
' ============================================================
Private Function CreateIndex( _
    ByVal tableName As String, _
    ByVal indexName As String, _
    ByVal fieldName As String, _
    ByVal isUnique As Boolean) As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim idx As DAO.Index
    
    Set db = CurrentDb
    
    ' Check if table exists
    If Not TableExists(tableName) Then
        Debug.Print "  SKIPPED: Index " & indexName & " (table " & tableName & " does not exist)"
        CreateIndex = False
        Exit Function
    End If
    
    Set tdf = db.TableDefs(tableName)
    
    ' Check if index already exists
    If IndexExists(tableName, indexName) Then
        Debug.Print "  SKIPPED: Index " & indexName & " (already exists)"
        CreateIndex = True
        Exit Function
    End If
    
    ' Create the index
    Set idx = tdf.CreateIndex(indexName)
    idx.Unique = isUnique
    idx.Fields.Append idx.CreateField(fieldName)
    tdf.Indexes.Append idx
    
    Debug.Print "  Created index: " & indexName & " on " & tableName & "." & fieldName
    CreateIndex = True
    Exit Function
    
ErrorHandler:
    Debug.Print "  ERROR creating index " & indexName & ": " & Err.Description
    CreateIndex = False
End Function

' ============================================================
' Helper function to create a foreign key relationship
' ============================================================
Private Function CreateFKRelation( _
    ByVal db As DAO.Database, _
    ByVal childTable As String, _
    ByVal childField As String, _
    ByVal parentTable As String, _
    ByVal parentField As String, _
    ByVal relationName As String) As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim rel As DAO.Relation
    Dim fld As DAO.Field
    
    ' Check if relationship already exists
    If RelationExists(relationName) Then
        Debug.Print "  SKIPPED: Relationship " & relationName & " (already exists)"
        CreateFKRelation = True
        Exit Function
    End If
    
    ' Verify both tables exist
    If Not TableExists(parentTable) Then
        Debug.Print "  SKIPPED: " & relationName & " (parent table " & parentTable & " does not exist)"
        CreateFKRelation = False
        Exit Function
    End If
    
    If Not TableExists(childTable) Then
        Debug.Print "  SKIPPED: " & relationName & " (child table " & childTable & " does not exist)"
        CreateFKRelation = False
        Exit Function
    End If
    
    ' Verify both fields exist
    If Not FieldExists(parentTable, parentField) Then
        Debug.Print "  SKIPPED: " & relationName & " (parent field " & parentTable & "." & parentField & " does not exist)"
        CreateFKRelation = False
        Exit Function
    End If
    
    If Not FieldExists(childTable, childField) Then
        Debug.Print "  SKIPPED: " & relationName & " (child field " & childTable & "." & childField & " does not exist)"
        CreateFKRelation = False
        Exit Function
    End If
    
    ' Create the relationship
    Set rel = db.CreateRelation(relationName, parentTable, childTable)
    
    ' Use minimal attributes - no cascade, basic referential integrity only
    ' This matches the test that succeeded
    rel.Attributes = 0
    
    ' Create the field relationship  
    ' In Access DAO: CreateField uses parent table field, ForeignName is child table field
    Set fld = rel.CreateField(parentField)
    fld.ForeignName = childField
    rel.Fields.Append fld
    
    ' Append the relationship to the database
    db.Relations.Append rel
    
    Debug.Print "  Created FK: " & parentTable & "." & parentField & " to " & childTable & "." & childField
    CreateFKRelation = True
    Exit Function
    
ErrorHandler:
    Debug.Print "  ERROR creating FK " & relationName & ": " & Err.Description
    Debug.Print "    Details: " & childTable & "." & childField & " to " & parentTable & "." & parentField
    Debug.Print "    Error Number: " & Err.Number
    
    ' Additional diagnostics for common issues
    If Err.Number = 3265 Then  ' Item not found in collection
        Debug.Print "    Possible cause: Field name mismatch or table not found"
    ElseIf Err.Number = 3304 Then  ' You must enter a personal identifier
        Debug.Print "    Possible cause: Field data type mismatch"
    ElseIf Err.Number = 3022 Then  ' Changes to table were unsuccessful
        Debug.Print "    Possible cause: Data integrity violation or existing data conflicts"
    End If
    
    CreateFKRelation = False
End Function

' ============================================================
' Test function to diagnose foreign key creation issues  
' ============================================================
Public Sub TestForeignKeyCreation()
    ' Test creating a single, simple foreign key relationship to diagnose issues
    
    Dim db As DAO.Database
    Dim rel As DAO.Relation  
    Dim fld As DAO.Field
    
    Set db = CurrentDb
    
    Debug.Print "=== FOREIGN KEY CREATION DIAGNOSTICS ==="
    Debug.Print ""
    
    ' Test case: community -> parish relationship
    Debug.Print "Testing community -> parish relationship..."
    
    ' 1. Verify tables exist
    Debug.Print "1. Checking tables exist:"
    Debug.Print "   Community table exists: " & TableExists("community")
    Debug.Print "   Parish table exists: " & TableExists("parish")
    
    ' 2. Verify fields exist  
    Debug.Print "2. Checking fields exist:"
    Debug.Print "   community.parish_id exists: " & FieldExists("community", "parish_id")
    Debug.Print "   parish.parish_id exists: " & FieldExists("parish", "parish_id")
    
    ' 3. Check field data types
    Debug.Print "3. Checking field data types:"
    If TableExists("community") And FieldExists("community", "parish_id") Then
        Debug.Print "   community.parish_id type: " & GetFieldDataType("community", "parish_id")
    End If
    If TableExists("parish") And FieldExists("parish", "parish_id") Then  
        Debug.Print "   parish.parish_id type: " & GetFieldDataType("parish", "parish_id")
    End If
    
    ' 4. Check for primary key on parent table
    Debug.Print "4. Checking primary key:"
    Debug.Print "   parish.parish_id is primary key: " & IsFieldPrimaryKey("parish", "parish_id")
    
    ' 5. Check for existing relationship
    Debug.Print "5. Checking existing relationship:"
    Debug.Print "   community_parish_fk exists: " & RelationExists("community_parish_fk")
    
    ' 6. Try creating with minimal attributes
    Debug.Print "6. Attempting to create relationship with minimal attributes..."
    
    On Error GoTo TestError
    
    If Not RelationExists("test_community_parish") Then
        Set rel = db.CreateRelation("test_community_parish", "parish", "community")
        ' No special attributes - just basic relationship
        rel.Attributes = 0
        
        Set fld = rel.CreateField("parish_id")    ' Parent field (PK)
        fld.ForeignName = "parish_id"             ' Child field (FK)
        rel.Fields.Append fld
        
        db.Relations.Append rel
        Debug.Print "   SUCCESS: Basic relationship created!"
        
        ' Clean up test
        db.Relations.Delete "test_community_parish"
        Debug.Print "   Test relationship removed."
    Else
        Debug.Print "   Test relationship already exists - skipping."
    End If
    
    Debug.Print ""
    Debug.Print "Diagnostics complete. Check output above for issues."
    Exit Sub
    
TestError:
    Debug.Print "   ERROR: " & Err.Number & " - " & Err.Description
    
    ' Additional specific diagnostics
    Select Case Err.Number
        Case 3001
            Debug.Print "   Error 3001 usually means:"
            Debug.Print "     - Data type mismatch between fields"
            Debug.Print "     - Missing primary key on parent field"  
            Debug.Print "     - Invalid relationship parameters"
        Case 3125
            Debug.Print "   Error 3125: Field name is not valid"
        Case 3265  
            Debug.Print "   Error 3265: Item not found in collection"
        Case 3029
            Debug.Print "   Error 3029: Not a valid account name or password"
    End Select
End Sub

' ============================================================
' Helper function to get field data type as string
' ============================================================
Private Function GetFieldDataType(ByVal tableName As String, ByVal fieldName As String) As String
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    
    Set db = CurrentDb
    Set tdf = db.TableDefs(tableName)
    Set fld = tdf.Fields(fieldName)
    
    Select Case fld.Type
        Case dbBoolean: GetFieldDataType = "Boolean"
        Case dbByte: GetFieldDataType = "Byte" 
        Case dbInteger: GetFieldDataType = "Integer"
        Case dbLong: GetFieldDataType = "Long"
        Case dbCurrency: GetFieldDataType = "Currency"
        Case dbSingle: GetFieldDataType = "Single"
        Case dbDouble: GetFieldDataType = "Double"
        Case dbDate: GetFieldDataType = "Date/Time"
        Case dbText: GetFieldDataType = "Text(" & fld.Size & ")"
        Case dbLongBinary: GetFieldDataType = "OLE Object"
        Case dbMemo: GetFieldDataType = "Memo"
        Case dbGUID: GetFieldDataType = "Replication ID"
        Case Else: GetFieldDataType = "Unknown(" & fld.Type & ")"
    End Select
    Exit Function
    
ErrorHandler:
    GetFieldDataType = "Error: " & Err.Description
End Function

' ============================================================
' Helper function to check if field is primary key
' ============================================================
Private Function IsFieldPrimaryKey(ByVal tableName As String, ByVal fieldName As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim idx As DAO.Index
    Dim fld As DAO.Field
    
    Set db = CurrentDb
    Set tdf = db.TableDefs(tableName)
    
    ' Check all indexes for primary key
    For Each idx In tdf.Indexes
        If idx.Primary Then
            For Each fld In idx.Fields
                If fld.Name = fieldName Then
                    IsFieldPrimaryKey = True
                    Exit Function
                End If
            Next fld
        End If
    Next idx
    
    IsFieldPrimaryKey = False
    Exit Function
    
ErrorHandler:
    IsFieldPrimaryKey = False
End Function

' ============================================================
' Helper function to check if an index exists
' ============================================================
Private Function IndexExists(ByVal tableName As String, ByVal indexName As String) As Boolean
    On Error GoTo NotFound
    
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim idx As DAO.Index
    
    Set db = CurrentDb
    Set tdf = db.TableDefs(tableName)
    Set idx = tdf.Indexes(indexName)
    
    IndexExists = True
    Exit Function
    
NotFound:
    IndexExists = False
End Function

' ============================================================
' Enable Access's native "Compact on Close" feature for this database session
' ============================================================
Public Sub EnableCompactOnClose()
    ' Turns on the native "Compact on Close" feature for this database session
    Application.SetOption "Auto Compact", True
End Sub

' ============================================================
' Run full materialization pipeline in order:
'   1. MaterializeAllPostgresLinkedTables
'   2. UnlinkTablesAndRemovePrefix
'   3. CreateIndexesAndConstraints
'
' Each step will prompt for confirmation individually.
' At the end, Access's native "Compact on Close" feature will be enabled for this session to optimize the database file after all changes.
' ============================================================
Public Sub RunMaterializationPipeline( _
    Optional ByVal localPrefix As String = "loc_", _
    Optional ByVal dropIfExists As Boolean = False, _
    Optional ByVal showMessageBoxes As Boolean = False)

    If showMessageBoxes Then
        If MsgBox("Run the full materialization pipeline?" & vbNewLine & vbNewLine & _
                  "Steps:" & vbNewLine & _
                  "  1. Materialize all linked PostgreSQL tables" & vbNewLine & _
                  "  2. Unlink PostgreSQL tables and remove prefix" & vbNewLine & _
                  "  3. Create indexes and foreign key constraints" & vbNewLine & vbNewLine & _
                  "Each step will prompt for individual confirmation.", _
                  vbYesNo + vbQuestion, "Run Materialization Pipeline") = vbNo Then
            MsgBox "Pipeline cancelled by user.", vbInformation
            Exit Sub
        End If
    End If

    Debug.Print "=== MATERIALIZATION PIPELINE START ==="

    Debug.Print "Step 1: MaterializeAllPostgresLinkedTables"
    MaterializeAllPostgresLinkedTables localPrefix, dropIfExists, showMessageBoxes

    Debug.Print "Step 2: UnlinkTablesAndRemovePrefix"
    UnlinkTablesAndRemovePrefix localPrefix, showMessageBoxes

    Debug.Print "Step 3: CreateIndexesAndConstraints"
    CreateIndexesAndConstraints showMessageBoxes

    Debug.Print "Enabling Access's native Compact on Close feature for this session..."
    EnableCompactOnClose

    Debug.Print "=== MATERIALIZATION PIPELINE COMPLETE ==="
    If showMessageBoxes Then
        MsgBox "Pipeline complete. See Immediate Window (Ctrl+G) for details.", vbInformation
    End If
End Sub