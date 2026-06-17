Option Compare Database
Option Explicit

' Automated setup script for DigiDiggie TNG Access App
Public Sub AutomatedFullSetup()
    On Error GoTo ErrorHandler
    Dim success As Boolean

    ' Step 1: Link tables
    success = LinkAllTngTables()
    If Not success Then
        MsgBox "Table linking failed. Please check the debug output for details.", vbCritical
        Exit Sub
    End If

    ' Step 2: Build forms
    BuildAllForms

    ' Step 3: Run materialization pipeline
    RunMaterializationPipeline

    MsgBox "Automated setup completed successfully!", vbInformation
    Exit Sub    
    
ErrorHandler:
    MsgBox "Setup error: " & Err.Description, vbCritical
End Sub

Public Function LinkAllTngTables() As Boolean
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim strDSN As String
    Dim arrTables As Variant
    Dim i As Integer
    Dim strTable As String
    
    Set db = CurrentDb
    
    ' Configure DSN connection string
    ' Update with your DSN name
    strDSN = "ODBC;DSN=DigiDiggie_TNG;DATABASE=digidiggie;"
    
    ' List all tables in digidiggie_tng schema
    arrTables = Array("community", "court_case", "court_case_entry", "land_use", _
                     "land_rights_status", "legal_source", "outcome_type", "parish", _
                     "person", "person_entry", "person_outcome", "person_relationship", _
                     "placename", "relationship_type", "role", "role_type", _
                     "ruling", "ruling_type", "season", "source")
    
    ' Link each table
    For i = LBound(arrTables) To UBound(arrTables)
        strTable = CStr(arrTables(i))
        
        ' Delete existing link if present
        On Error Resume Next
        db.TableDefs.Delete strTable
        On Error GoTo ErrorHandler
        
        ' Create new linked table
        Set tdf = db.CreateTableDef(strTable)
        tdf.Connect = strDSN
        tdf.SourceTableName = "digidiggie_tng." & strTable  ' Link to server-side schema
        
        On Error Resume Next
        db.TableDefs.Append tdf
        If Err.Number > 0 Then
            Debug.Print "Error linking table " & strTable & ": " & Err.Description
        End If
        On Error GoTo ErrorHandler
        
        Debug.Print "Linked: " & strTable
    Next i
    
    Set tdf = Nothing
    Set db = Nothing
    
    LinkAllTngTables = True
    Debug.Print "All tables linked successfully."
    Exit Function
    
ErrorHandler:
    Debug.Print "Error linking table " & strTable & ": " & Err.Description
    LinkAllTngTables = False
End Function


