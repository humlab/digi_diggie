Option Compare Database
Option Explicit

' Export result from SQL query to CSV file
Sub ExportQueryCSV(strSQL As String, strFilePath As String)
    Const TEMP_QUERY_NAME As String = "TempExportQuery"
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef

    On Error GoTo ErrorHandler
    Set db = CurrentDb()

    ' Remove existing temp query if present
    On Error Resume Next
    db.QueryDefs.Delete TEMP_QUERY_NAME
    On Error GoTo ErrorHandler

    ' Create temporary query
    Set qdf = db.CreateQueryDef(TEMP_QUERY_NAME, strSQL)

    ' Export using TransferText (no import/export spec required)
    DoCmd.TransferText acExportDelim, , TEMP_QUERY_NAME, strFilePath, True

CleanExit:
    ' Clean up created objects
    On Error Resume Next
    db.QueryDefs.Delete TEMP_QUERY_NAME
    Set qdf = Nothing
    Set db = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Failed to export query to CSV:" & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation
    Resume CleanExit
End Sub

Function ExportPlacenamesCSV() As String
    Dim strSQL As String
    Dim strFilePath As String

    strSQL = "SELECT * FROM Ortnamn_ny " & _
             "INNER JOIN Sockenstad ON Sockenstad.nr = Ortnamn_ny.Sockenstad_nr;"

    strFilePath = CurrentProject.Path & "\placenames.csv"
    Call ExportQueryCSV(strSQL, strFilePath)
    ExportPlacenamesCSV = strFilePath
End Function
