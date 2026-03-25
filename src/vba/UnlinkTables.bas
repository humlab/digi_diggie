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
    Optional ByVal dropIfExists As Boolean = False)

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
    MsgBox msg, vbInformation
End Sub

' ============================================================
' Remove all linked tables and remove "loc_" prefix from local tables
' ============================================================
Public Sub UnlinkTablesAndRemovePrefix( _
    Optional ByVal localPrefix As String = "loc_")

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
    If MsgBox("This will:" & vbNewLine & vbNewLine & _
              "1. Delete ALL linked PostgreSQL tables" & vbNewLine & _
              "2. Remove '" & localPrefix & "' prefix from local tables" & vbNewLine & vbNewLine & _
              "This action cannot be undone. Continue?", _
              vbYesNo + vbExclamation + vbDefaultButton2, _
              "Unlink Tables and Remove Prefix") = vbNo Then
        MsgBox "Operation cancelled by user.", vbInformation
        Exit Sub
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
    MsgBox msg, vbInformation
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
                        rel.Fields.Append rel.CreateField(fkColumn)
                        rel.Fields(fkColumn).ForeignName = pkColumn

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