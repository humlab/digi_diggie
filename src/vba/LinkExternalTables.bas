Public Function LinkAllTables() As Boolean
    On Error GoTo ErrorHandler
    
    Dim strDSN As String
    Dim arrTables As Variant

    ' Turn off screen updating to speed up process and stop flickering
    Application.Echo False
    
    ' Configure your DSN connection string
    strDSN = "ODBC;DSN=DigiDiggie_PostgreSQL;DATABASE=digidiggie;"
    
    ' Add all your table names
    arrTables = Array("entries", "communities", "persons", "parishes", "sources", "seasons", "land_use", "winners", "legal_sources", "judgements", "placenames", "person_properties", "properties")
    
    ' Link each table
    Dim i As Integer
    For i = LBound(arrTables) To UBound(arrTables)
        Call LinkTable(CStr(arrTables(i)), strDSN)
    Next i
    
    LinkAllTables = True

Exit_Handler:
    ' Turn screen updating back on
    Application.Echo True

    ' Refresh the database window to show linked tables
    Application.RefreshDatabaseWindow

    Exit Function
    
ErrorHandler:
    MsgBox "Error linking tables: " & Err.Description, vbCritical
    LinkAllTables = False
    Resume Exit_Handler
End Function

Private Sub LinkTable(strTableName As String, strDSN As String)
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    
    Set db = CurrentDb
    
    ' Remove existing link if present
    On Error Resume Next
    db.TableDefs.Delete strTableName
    On Error GoTo 0
    
    ' Create new linked table
    Set tdf = db.CreateTableDef(strTableName)
    tdf.Connect = strDSN
    tdf.SourceTableName = strTableName
    db.TableDefs.Append tdf
    
    Set tdf = Nothing
    Set db = Nothing
End Sub