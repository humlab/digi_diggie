Option Compare Database
Option Explicit

' Public entry point:
'   ImportExcelSheet "C:\data\myfile.xlsx", "Sheet1", "Staging_MyData"
Public Function ImportExcelSheet(ByVal filePath As String, _
                                 ByVal sheetName As String, _
                                 ByVal tableName As String) As Boolean
    On Error GoTo EH

    ' 1) Sanity checks
    If Len(Dir(filePath)) = 0 Then
        Err.Raise vbObjectError + 513, , "File not found: " & filePath
    End If

    ' 2) Normalize the Range argument for TransferSpreadsheet
    Dim rangeArg As String
    rangeArg = NormalizeSheetRange(sheetName)  ' e.g., "Sheet1$"

    ' 3) Close & drop existing table (so we truly overwrite)
    CloseTableIfOpen tableName
    If TableExists(tableName) Then
        DoCmd.DeleteObject acTable, tableName
    End If

    ' 4) Determine Excel type from extension and import
    Dim xlType As AcSpreadSheetType
    xlType = GuessSpreadsheetType(filePath)

    DoCmd.TransferSpreadsheet _
        TransferType:=acImport, _
        SpreadsheetType:=xlType, _
        tableName:=tableName, _
        FileName:=filePath, _
        HasFieldNames:=True, _
        Range:=rangeArg

    ImportExcelSheet = True
    Exit Function

EH:
    ImportExcelSheet = False
    MsgBox "ImportExcelSheet failed." & vbCrLf & _
           "File: " & filePath & vbCrLf & _
           "Sheet: " & sheetName & vbCrLf & _
           "Table: " & tableName & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation
End Function

' ---- Helpers ----

' Ensure a valid Range string for TransferSpreadsheet:
' - If caller passed a plain sheet name (e.g., "Sheet1"), make it "Sheet1$"
' - If they already passed "Sheet1$" or "Sheet1$A1:Z999", leave as-is
Private Function NormalizeSheetRange(ByVal sheetName As String) As String
    Dim s As String
    s = Trim$(sheetName)

    ' If it already looks like a proper range, keep it
    If InStr(1, s, "$", vbTextCompare) > 0 Or InStr(1, s, "!", vbTextCompare) > 0 Then
        NormalizeSheetRange = s
        Exit Function
    End If

    ' Otherwise, append "$" so the whole sheet is imported
    NormalizeSheetRange = s & "$"
End Function

' Pick a spreadsheet type based on the file extension.
' - .xlsx / .xlsm => acSpreadsheetTypeExcel12Xml
' - .xls         => acSpreadsheetTypeExcel9
' - .csv         => acSpreadsheetTypeExcel12 (works for CSV via TransferText would be better; included for completeness)
' Note: .xlsb is not supported by TransferSpreadsheet.
Private Function GuessSpreadsheetType(ByVal filePath As String) As AcSpreadSheetType
    Dim ext As String
    ext = LCase$(Mid$(filePath, InStrRev(filePath, ".") + 1))

    Select Case ext
        Case "xlsx", "xlsm"
            GuessSpreadsheetType = acSpreadsheetTypeExcel12Xml
        Case "xls"
            GuessSpreadsheetType = acSpreadsheetTypeExcel9
        Case "csv"
            ' CSV is better handled with DoCmd.TransferText, but this keeps things simple
            GuessSpreadsheetType = acSpreadsheetTypeExcel12
        Case Else
            ' Default to 12Xml; caller can override by editing if needed
            GuessSpreadsheetType = acSpreadsheetTypeExcel12Xml
    End Select
End Function

Private Sub CloseTableIfOpen(ByVal tableName As String)
    On Error Resume Next
    If SysCmd(acSysCmdGetObjectState, acTable, tableName) <> 0 Then
        DoCmd.Close acTable, tableName, acSaveNo
    End If
    On Error GoTo 0
End Sub

Private Function TableExists(ByVal tableName As String) As Boolean
    On Error GoTo NoTable
    Dim t As DAO.TableDef
    Set t = CurrentDb.TableDefs(tableName)
    TableExists = True
    Exit Function
NoTable:
    TableExists = False
End Function


Public Function ImportTranslationMappingExcelSheet() As Boolean
    Dim isOK As Boolean
    isOK = ImportExcelSheet(CurrentProject.Path & "\TranslationMapping.xlsx", "TranslationMapping", "TranslationMapping")
    isOK = isOK And ImportExcelSheet(CurrentProject.Path & "\TranslationMapping.xlsx", "RowSource", "RowSource")
    isOK = isOK And ImportExcelSheet(CurrentProject.Path & "\TranslationMapping.xlsx", "QueryDefinitions", "QueryDefinitions")
    ImportTranslationMappingExcelSheet = isOK
End Function
