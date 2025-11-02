Option Explicit
Option Compare Database

' Helper function to import or replace a module
Private Sub ImportOrReplaceModule(moduleName As String)
    Dim dbPath As String
    Dim modulePath As String

    ' Get the folder of the current Access database
    dbPath = CurrentProject.Path
    modulePath = dbPath & "\" & moduleName & ".bas"

    On Error Resume Next
    ' Try to delete any existing module with this name
    DoCmd.DeleteObject acModule, moduleName
    On Error GoTo 0

    ' Import the .bas file from the same folder
    Application.LoadFromText acModule, moduleName, modulePath

    MsgBox "Module '" & moduleName & "' imported from: " & modulePath, vbInformation
End Sub


Option Compare Database
Option Explicit


Public Function ImportOrReplaceScript() As Boolean
    Call ImportOrReplaceModule("rename")
End Function

