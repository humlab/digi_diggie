Option Compare Database
Option Explicit

Public Sub CreateSpecificForms()
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim frm As Form
    Dim ctl As Control
    Dim strTable As String
    Dim strFormName As String
    Dim strTempFormName As String
    Dim strSQL As String
    Dim arrTables As Variant
    Dim i As Integer
    Dim dictLookups As Object
    
    ' Turn off screen updating to speed up process and stop flickering
    Application.Echo False
    
    On Error GoTo ErrorHandler
    
    Set db = CurrentDb
    Set dictLookups = CreateObject("Scripting.Dictionary")
    
    ' ---------------------------------------------------------
    ' 1. Configuration
    ' Map "TableName|FieldName" to the specific RowSource SQL
    ' ---------------------------------------------------------
    AddConfig dictLookups, "communities", "parish_id", "SELECT [parishes].[parish_id], [parishes].[parish] FROM [parishes] ORDER BY [parish]"
    AddConfig dictLookups, "entries", "actor_id", "SELECT [persons].[person_id], [persons].[full_name] FROM [persons] ORDER BY [full_name]"
    AddConfig dictLookups, "entries", "community_id", "SELECT [communities].[community_id], [communities].[community_name] FROM [communities] ORDER BY [community_name]"
    AddConfig dictLookups, "entries", "source_id", "SELECT [sources].[source_id], [sources].[source_abbreviation] FROM [sources] ORDER BY [source_abbreviation]"
    AddConfig dictLookups, "entries", "season_id", "SELECT [seasons].[season_id], [seasons].[season_name] FROM [seasons] ORDER BY [season_name]"
    AddConfig dictLookups, "entries", "land_use_id", "SELECT [land_use].[land_use_id], [land_use].[type] FROM [land_use] ORDER BY [type]"
    AddConfig dictLookups, "entries", "winner_id", "SELECT [winners].[winner_id], [winners].[winner_description] FROM [winners] ORDER BY [winner_description]"
    AddConfig dictLookups, "entries", "legal_source_id", "SELECT [legal_sources].[legal_source_id], [legal_sources].[legal_source_name] FROM [legal_sources] ORDER BY [legal_source_name]"
    AddConfig dictLookups, "entries", "judgement_id", "SELECT [judgements].[judgement_id], [judgements].[sanction] FROM [judgements] ORDER BY [sanction]"
    ' FIXME: Update to use placenames table
    ' AddConfig dictLookups, "entries", "placename_id", "SELECT [Ortnamn_ny1].[ID], [Ortnamn_ny1].[Kombo] FROM Ortnamn_ny1 ORDER BY [Kombo]"
    AddConfig dictLookups, "persons", "community_name", "SELECT [communities].[community_name] FROM [communities] ORDER BY [community_name]"

    ' Define the list of tables to process
    arrTables = Array("communities", "entries", "persons")
    
    ' ---------------------------------------------------------
    ' 2. Form Creation Loop
    ' ---------------------------------------------------------   
    For i = LBound(arrTables) To UBound(arrTables)
        strTable = arrTables(i)
        strFormName = "frm_" & strTable
        
        ' Check if table exists (Using db variable to save overhead)
        If Not TableExists(db, strTable) Then
            Debug.Print "Table not found: " & strTable
            GoTo NextTable
        End If
        
        ' Delete existing form if it exists to avoid errors
        If ObjectExists(acForm, strFormName) Then            
            If CurrentProject.AllForms(strFormName).IsLoaded Then ' Check if it is open first
                DoCmd.Close acForm, strFormName, acSaveNo
            End If
            DoCmd.DeleteObject acForm, strFormName
        End If
        
        ' Create the new form
        Set frm = CreateForm
        strTempFormName = frm.Name ' Likely "Form1"
        
        ' Set Form Properties
        frm.RecordSource = strTable
        frm.DefaultView = 2 ' Datasheet View
        
        Set tdf = db.TableDefs(strTable)
        
        ' Loop through every field in the table to create controls
        Dim ctlTop As Integer
        ctlTop = 0
        
        For Each fld In tdf.Fields
            Dim key As String
            key = strTable & "|" & fld.Name
            
            ' Determine if this field needs a Combo Box or a Text Box
            If dictLookups.Exists(key) Then
                ' Create Combo
                Set ctl = CreateControl(frm.Name, acComboBox, acDetail, , fld.Name, 100, ctlTop, 3000, 300)
                strSQL = dictLookups(key)
                
                With ctl
                    .Name = "cbo_" & fld.Name ' FIXME: remove prefix?
                    .RowSource = strSQL
                    .RowSourceType = "Table/Query"
                    .BoundColumn = 1

                    ' logic to detect if we need to hide the first column (ID)
                    ' If the SQL Selects 2 columns (ID, Name), we usually hide ID.
                    ' If it selects 1 column (Name), we show it.
                    If CountCommasInSelect(strSQL) >= 1 Then
                        .ColumnCount = 2
                        .ColumnWidths = "0;1440" ' 0 width for ID, 1 inch for text
                        .ListWidth = 1440 * 2
                    Else
                        .ColumnCount = 1
                        .ColumnWidths = "1440"
                    End If
                End With
            Else
                ' Create Textbox
                Set ctl = CreateControl(frm.Name, acTextBox, acDetail, , fld.Name, 100, ctlTop, 3000, 300)
                ctl.Name = "txt_" & fld.Name ' FIXME: remove prefix?
            End If
            
            ctlTop = ctlTop + 350
        Next fld
        
        ' Close the temporary form and Save it (it will save as "Form1" or similar)
        DoCmd.Close acForm, strTempFormName, acSaveYes
        
        ' Rename it to the desired name
        DoCmd.Rename strFormName, acForm, strTempFormName
        
        Debug.Print "Created form: " & strFormName
        
NextTable:
    Next i
    
    MsgBox "Forms created successfully!", vbInformation

Exit_Handler:
    Application.Echo True ' Turn screen back on
    Set dictLookups = Nothing
    Set fld = Nothing
    Set tdf = Nothing
    Set db = Nothing
    Exit Sub

ErrorHandler:
    Application.Echo True ' Ensure screen turns on if error occurs
    MsgBox "Error " & Err.Number & ": " & Err.Description
    Resume Exit_Handler
End Sub

' ---- Helpers ----

Sub AddConfig(dict As Object, tbl As String, fld As String, sql As String)
    Dim key As String
    key = tbl & "|" & fld
    If Not dict.Exists(key) Then
        dict.Add key, sql
    End If
End Sub

Function TableExists(dbs As DAO.Database, tblName As String) As Boolean
    Dim tdf As DAO.TableDef
    On Error Resume Next
    Set tdf = dbs.TableDefs(tblName)
    TableExists = (Err.Number = 0)
    On Error GoTo 0
End Function

Function ObjectExists(objType As AcObjectType, objName As String) As Boolean
    ' Use CurrentProject.AllForms to check existence
    ' Note: This collection contains both open and closed forms
    On Error Resume Next
    Dim obj As Object
    If objType = acForm Then
        ' Check if it exists in the AllForms collection
        Set obj = CurrentProject.AllForms(objName)
        ObjectExists = (Err.Number = 0)
    End If
    On Error GoTo 0
End Function

Function CountCommasInSelect(sql As String) As Integer
    ' Heuristic to determine column count. 
    ' Finds the "FROM" keyword and counts commas before it.
    Dim posFrom As Integer
    Dim strSelectPart As String
    posFrom = InStr(1, UCase(sql), " FROM ")
    If posFrom > 0 Then
        strSelectPart = Left(sql, posFrom - 1)
        CountCommasInSelect = Len(strSelectPart) - Len(Replace(strSelectPart, ",", ""))
    Else
        CountCommasInSelect = 0
    End If
End Function