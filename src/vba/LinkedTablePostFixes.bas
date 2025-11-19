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


Private Function PropertyExists(obj As Object, propName As String) As Boolean
    On Error GoTo ErrorHandler
    Dim property As DAO.property
    
    Set property = obj.Properties(propName)
    PropertyExists = True
    Exit Function
    
ErrorHandler:
    PropertyExists = False
    Err.Clear
End Function


Private Function AddOrReplaceProperty(obj As Object, propName As String, propType As DAO.DataTypeEnum, propValue As Variant) As Boolean
    On Error GoTo ErrorHandler
    
    If PropertyExists(obj, propName) Then
        obj.Properties(propName) = propValue
    Else
        obj.Properties.Append obj.CreateProperty(propName, propType, propValue)
    End If
    
    AddOrReplaceProperty = True
    Exit Function
    
ErrorHandler:
    Debug.Print "Error setting property " & propName & ": " & Err.Description
    AddOrReplaceProperty = False
End Function


Public Function SetFieldRowSource(tableName As String, fieldName As String, sqlSource As String) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim table As DAO.TableDef
    Dim fld As DAO.field

    Set db = CurrentDb
    Set table = db.TableDefs(tableName)
    Set fld = table.Fields(fieldName)

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
    
    table.Fields.Refresh

    Debug.Print "Set RowSource for " & tableName & "." & fieldName
    SetFieldRowSource = True
    Exit Function

ErrorHandler:
    Debug.Print "error: failed setting RowSource for " & tableName & "." & fieldName & ": " & Err.Description
    SetFieldRowSource = False
End Function


Public Function UpdateRowSources() As Boolean
    On Error GoTo ErrorHandler
    Dim isOK As Boolean
   
    Let isOK = SetFieldRowSource("communities", "parish_id", "SELECT [parishes].[parish_id], [parishes].[parish] FROM [parishes] ORDER BY [parish]")
    Let isOK = SetFieldRowSource("entries", "actor_id", "SELECT [persons].[person_id], [persons].[full_name] FROM [persons] ORDER BY [full_name]")
    Let isOK = SetFieldRowSource("entries", "community_id", "SELECT [communities].[community_id], [communities].[community_name] FROM [communities] ORDER BY [community_name]")
    Let isOK = SetFieldRowSource("entries", "source_id", "SELECT [sources].[source_id], [sources].[source_abbreviation] FROM [sources] ORDER BY [source_abbreviation]")
    Let isOK = SetFieldRowSource("entries", "season_id", "SELECT [seasons].[season_id], [seasons].[season_name] FROM [seasons] ORDER BY [season_name]")
    Let isOK = SetFieldRowSource("entries", "land_use_id", "SELECT [land_use].[land_use_id], [land_use].[type] FROM [land_use] ORDER BY [type]")
    Let isOK = SetFieldRowSource("entries", "winner_id", "SELECT [winners].[winner_id], [winners].[winner_description] FROM [winners] ORDER BY [winner_description]")
    Let isOK = SetFieldRowSource("entries", "legal_source_id", "SELECT [legal_sources].[legal_source_id], [legal_sources].[legal_source_name] FROM [legal_sources] ORDER BY [legal_source_name]")
    Let isOK = SetFieldRowSource("entries", "judgement_id", "SELECT [judgements].[judgement_id], [judgements].[sanction] FROM [judgements] ORDER BY [sanction]")
    Let isOK = SetFieldRowSource("entries", "placename_id", "SELECT [Ortnamn_ny1].[ID], [Ortnamn_ny1].[Kombo] FROM Ortnamn_ny1 ORDER BY [Kombo]")
    Let isOK = SetFieldRowSource("persons", "community_name", "SELECT [communities].[community_name] FROM [communities] ORDER BY [community_name]")
    
    UpdateRowSources = isOK
    If isOK Then
        Debug.Print "Successfully updated all row sources."
    Else
        Debug.Print "Some row sources failed to update."
    End If
    Exit Function
    
ErrorHandler:
    Debug.Print "error: Error updating row source: " & Err.Description
    UpdateRowSources = False
End Function


Public Function PostLinkedTableUpdates() As Boolean
    On Error GoTo ErrorHandler
    Dim isOK As Boolean

    ' Rename tables to remove "public_" prefix
    isOK = RenameTables()
    
    ' After updating linked table paths, we may need to refresh RowSources
    ' in case the linked tables have changed.
    isOK = UpdateRowSources()
    
    PostLinkedTableUpdates = isOK
    Exit Function
ErrorHandler:
    Debug.Print "error: PostLinkedTableUpdates failed: " & Err.Description
    PostLinkedTableUpdates = False
End Function
