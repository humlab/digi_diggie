Option Explicit
Option Compare Database

'==============================================================================
' DigiDiggie Migration Script Runner
'==============================================================================
' This module provides functions to run SQL migration scripts from file
' Specifically designed for migrate_to_counter_pks.sql
'
' Usage:
'   1. Call RunMigrationFromFile() to execute the migration script
'   2. Make sure you have a backup before running!
'
' Author: Generated for DigiDiggie TNG Project
'==============================================================================

Public Function RunMigrationFromFile() As Boolean
    ' Main function to run the migration script from file
    ' Returns True if successful, False if error occurs
    
    On Error GoTo ErrorHandler
    
    Dim result As Boolean
    Dim scriptPath As String
    
    ' Find the migration script path
    scriptPath = FindSQLFile()
    
    ' Confirm with user before proceeding
    If MsgBox("WARNING: This will modify your database structure!" & vbNewLine & vbNewLine & _
              "Make sure you have a BACKUP before proceeding." & vbNewLine & vbNewLine & _
              "Continue with migration?", _
              vbYesNo + vbExclamation + vbDefaultButton2, _
              "DigiDiggie Migration") = vbNo Then
        MsgBox "Migration cancelled by user.", vbInformation
        RunMigrationFromFile = False
        Exit Function
    End If
    
    ' Run the migration
    result = ExecuteSQLFromFile(scriptPath)
    
    If result Then
        MsgBox "Migration completed successfully!" & vbNewLine & vbNewLine & _
               "Next steps:" & vbNewLine & _
               "1. Verify data integrity" & vbNewLine & _
               "2. Recreate foreign key relationships" & vbNewLine & _
               "3. Compact and repair database", _
               vbInformation, "Migration Complete"
    End If
    
    RunMigrationFromFile = result
    Exit Function
    
ErrorHandler:
    MsgBox "Error in RunMigrationFromFile: " & Err.Description, vbCritical
    RunMigrationFromFile = False
End Function

Public Function ExecuteSQLFromFile(sqlFilePath As String) As Boolean
    ' Executes SQL statements from a file
    ' Properly handles multi-line statements and comments
    
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim fileContent As String
    Dim sqlStatements As Collection
    Dim i As Integer
    Dim stmt As String
    
    ' Check if file exists
    If Dir(sqlFilePath) = "" Then
        MsgBox "SQL file not found: " & sqlFilePath, vbCritical
        ExecuteSQLFromFile = False
        Exit Function
    End If
    
    ' Read file content
    fileContent = ReadTextFile(sqlFilePath)
    If Len(fileContent) = 0 Then
        MsgBox "SQL file is empty or could not be read: " & sqlFilePath, vbCritical
        ExecuteSQLFromFile = False
        Exit Function
    End If
    
    ' Get database reference
    Set db = CurrentDb
    
    ' Parse SQL statements properly
    Set sqlStatements = ParseSQLStatements(fileContent)
    
    ' Execute each statement
    For i = 1 To sqlStatements.Count
        stmt = sqlStatements(i)
        
        If Len(Trim(stmt)) > 0 Then
            DoEvents  ' Allow UI to remain responsive
            
            Debug.Print "Executing: " & Left(stmt, 80) & "..."
            
            ' Execute the statement
            db.Execute stmt, dbFailOnError
        End If
    Next i
    
    ExecuteSQLFromFile = True
    Exit Function
    
ErrorHandler:
    MsgBox "Error executing SQL from file: " & Err.Description & vbNewLine & _
           "Statement: " & Left(stmt, 200) & "...", vbCritical
    ExecuteSQLFromFile = False
End Function

Private Function ReadTextFile(filePath As String) As String
    ' Reads entire text file into a string
    
    On Error GoTo ErrorHandler
    
    Dim fileNum As Integer
    Dim fileContent As String
    
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    fileContent = Input$(LOF(fileNum), fileNum)
    Close #fileNum
    
    ReadTextFile = fileContent
    Exit Function
    
ErrorHandler:
    If fileNum > 0 Then Close #fileNum
    MsgBox "Error reading file: " & Err.Description, vbCritical
    ReadTextFile = ""
End Function

Private Function ParseSQLStatements(sqlContent As String) As Collection
    ' Advanced SQL parsing to handle multi-line statements and comments
    ' Returns a Collection of individual SQL statements ready for execution
    
    Dim statements As New Collection
    Dim lines As Variant
    Dim i As Integer
    Dim currentStatement As String
    Dim currentLine As String
    Dim cleanLine As String
    Dim inMultiLineStatement As Boolean
    
    Debug.Print "=== PARSING DEBUG ==="
    Debug.Print "Content length: " & Len(sqlContent)
    
    ' Handle different line endings (Unix vs Windows)
    If InStr(sqlContent, vbCrLf) > 0 Then
        lines = Split(sqlContent, vbCrLf)
        Debug.Print "Using Windows line endings (CRLF)"
    ElseIf InStr(sqlContent, vbLf) > 0 Then
        lines = Split(sqlContent, vbLf) 
        Debug.Print "Using Unix line endings (LF)"
    Else
        lines = Split(sqlContent, vbNewLine)
        Debug.Print "Using default line endings"
    End If
    
    Debug.Print "Total lines: " & (UBound(lines) + 1)
    
    currentStatement = ""
    inMultiLineStatement = False
    
    For i = 0 To UBound(lines)
        currentLine = Trim(lines(i))
        
        ' Skip empty lines
        If Len(currentLine) = 0 Then
            GoTo NextLine
        End If
        
        ' Skip comment lines (starting with --)
        If Left(currentLine, 2) = "--" Then
            GoTo NextLine
        End If
        
        ' Remove inline comments
        If InStr(currentLine, "--") > 0 Then
            cleanLine = Trim(Left(currentLine, InStr(currentLine, "--") - 1))
        Else
            cleanLine = currentLine
        End If
        
        ' Skip if line became empty after removing comments
        If Len(cleanLine) = 0 Then
            GoTo NextLine
        End If
        
        ' Add line to current statement
        If Len(currentStatement) = 0 Then
            currentStatement = cleanLine
        Else
            currentStatement = currentStatement & " " & cleanLine
        End If
        
        ' Check if statement ends with semicolon
        If Right(Trim(currentStatement), 1) = ";" Then
            ' Remove semicolon and add complete statement to collection
            currentStatement = Trim(Left(currentStatement, Len(currentStatement) - 1))
            If Len(currentStatement) > 0 Then
                statements.Add currentStatement
                Debug.Print "Found statement " & statements.Count & ": " & Left(currentStatement, 50) & "..."
            End If
            currentStatement = ""
        End If
        
NextLine:
    Next i
    
    ' Add any remaining statement (in case file doesn't end with semicolon)
    If Len(Trim(currentStatement)) > 0 Then
        statements.Add Trim(currentStatement)
        Debug.Print "Found final statement " & statements.Count & ": " & Left(currentStatement, 50) & "..."
    End If
    
    Debug.Print "Total statements parsed: " & statements.Count
    Set ParseSQLStatements = statements
End Function

Private Function CleanSQLContent(sqlContent As String) As String
    ' Legacy function - replaced by ParseSQLStatements
    ' Kept for compatibility
    CleanSQLContent = sqlContent
End Function

Public Function BackupDatabase() As Boolean
    ' Creates a backup of the current database before migration
    
    On Error GoTo ErrorHandler
    
    Dim backupPath As String
    Dim timestamp As String
    
    ' Create timestamp for backup filename
    timestamp = Format(Now, "yyyymmdd_hhnnss")
    backupPath = CurrentProject.Path & "\backup_" & timestamp & ".accdb"
    
    ' Create backup using FileCopy
    FileCopy CurrentProject.FullName, backupPath
    
    MsgBox "Backup created successfully: " & backupPath, vbInformation
    BackupDatabase = True
    Exit Function
    
ErrorHandler:
    MsgBox "Error creating backup: " & Err.Description, vbCritical
    BackupDatabase = False
End Function

Public Function RunMigrationWithBackup() As Boolean
    ' Complete migration process with automatic backup
    
    Dim backupResult As Boolean
    Dim migrationResult As Boolean
    
    ' First create backup
    backupResult = BackupDatabase()
    If Not backupResult Then
        MsgBox "Cannot proceed without backup. Migration cancelled.", vbCritical
        RunMigrationWithBackup = False
        Exit Function
    End If
    
    ' Then run migration
    migrationResult = RunMigrationFromFile()
    RunMigrationWithBackup = migrationResult
End Function

Public Function PreviewSQLStatements() As Boolean
    ' Debug function to preview parsed SQL statements without executing
    ' Useful for troubleshooting parsing issues
    
    On Error GoTo ErrorHandler
    
    Dim scriptPath As String
    Dim fileContent As String
    Dim statements As Collection
    Dim i As Integer
    
    ' Try multiple possible paths for the SQL file
    scriptPath = FindSQLFile()
    
    Debug.Print "=== FILE PATH DEBUGGING ==="
    Debug.Print "Database location: " & CurrentProject.Path
    Debug.Print "SQL file path: " & scriptPath
    
    If Dir(scriptPath) = "" Then
        MsgBox "SQL file not found at: " & scriptPath & vbNewLine & vbNewLine & _
               "Database location: " & CurrentProject.Path, vbCritical
        PreviewSQLStatements = False
        Exit Function
    End If

    fileContent = ReadTextFile(scriptPath)
    
    Debug.Print "File size: " & Len(fileContent) & " characters"
    Debug.Print "First 100 characters: " & Left(fileContent, 100)
    
    Set statements = ParseSQLStatements(fileContent)

    Debug.Print "=== PARSED SQL STATEMENTS ==="
    Debug.Print "Total statements found: " & statements.Count
    Debug.Print ""
    
    For i = 1 To statements.Count
        Debug.Print "Statement " & i & ":"
        Debug.Print statements(i)
        Debug.Print "--- End Statement " & i & " ---"
        Debug.Print ""
    Next i
    
    MsgBox "Found " & statements.Count & " SQL statements. Check Immediate Window (Ctrl+G) for details.", vbInformation
    PreviewSQLStatements = True
    Exit Function
    
ErrorHandler:
    MsgBox "Error previewing SQL statements: " & Err.Description, vbCritical
    PreviewSQLStatements = False
End Function

Public Function RunMigrationStepByStep() As Boolean
    ' Execute migration with detailed step-by-step feedback
    ' Better for troubleshooting when full migration fails
    
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim scriptPath As String
    Dim fileContent As String
    Dim statements As Collection
    Dim i As Integer
    Dim stmt As String
    Dim continueOnError As Boolean
    
    scriptPath = FindSQLFile()
    
    ' Confirm with user
    If MsgBox("Run migration step-by-step with detailed feedback?" & vbNewLine & _
              "This will show progress for each SQL statement.", _
              vbYesNo + vbQuestion, "Step-by-Step Migration") = vbNo Then
        RunMigrationStepByStep = False
        Exit Function
    End If
    
    ' Ask if should continue on errors
    continueOnError = (MsgBox("Continue execution if individual statements fail?" & vbNewLine & _
                             "Yes = Continue on errors, No = Stop on first error", _
                             vbYesNo + vbQuestion, "Error Handling") = vbYes)
    
    ' Read and parse SQL
    fileContent = ReadTextFile(scriptPath)
    Set statements = ParseSQLStatements(fileContent)
    Set db = CurrentDb
    
    Debug.Print "=== STEP-BY-STEP MIGRATION ==="
    Debug.Print "Starting migration with " & statements.Count & " statements"
    
    ' Execute each statement with feedback
    For i = 1 To statements.Count
        stmt = statements(i)
        
        Debug.Print "Executing statement " & i & " of " & statements.Count & ":"
        Debug.Print Left(stmt, 100) & "..."
        
        DoEvents
        
        On Error Resume Next
        db.Execute stmt, dbFailOnError
        
        If Err.Number <> 0 Then
            Debug.Print "ERROR: " & Err.Description
            
            If Not continueOnError Then
                MsgBox "Migration failed at statement " & i & ":" & vbNewLine & _
                       Left(stmt, 200) & "..." & vbNewLine & vbNewLine & _
                       "Error: " & Err.Description, vbCritical
                RunMigrationStepByStep = False
                Exit Function
            Else
                Debug.Print "Continuing despite error..."
            End If
        Else
            Debug.Print "SUCCESS"
        End If
        
        On Error GoTo ErrorHandler
        Debug.Print ""
    Next i
    
    Debug.Print "Migration completed!"
    MsgBox "Step-by-step migration completed. Check Immediate Window for details.", vbInformation
    RunMigrationStepByStep = True
    Exit Function
    
ErrorHandler:
    MsgBox "Error in step-by-step migration: " & Err.Description, vbCritical
    RunMigrationStepByStep = False
End Function

Private Function FindSQLFile() As String
    ' Try to locate the SQL file in various possible locations
    ' Returns the first valid path found
    
    Dim basePath As String
    Dim possiblePaths As Variant
    Dim i As Integer
    
    basePath = CurrentProject.Path
    
    ' Remove trailing backslash if present
    If Right(basePath, 1) = "\" Then
        basePath = Left(basePath, Len(basePath) - 1)
    End If
    
    ' Define possible relative paths to the SQL file
    possiblePaths = Array( _
        "\src\schema\migrate_to_counter_pks.sql", _
        "\..\src\schema\migrate_to_counter_pks.sql", _
        "\..\..\src\schema\migrate_to_counter_pks.sql", _
        "\migrate_to_counter_pks.sql", _
        "\..\migrate_to_counter_pks.sql" _
    )
    
    ' Try each possible path
    For i = 0 To UBound(possiblePaths)
        Dim testPath As String
        testPath = basePath & possiblePaths(i)
        
        Debug.Print "Trying path: " & testPath
        
        If Dir(testPath) <> "" Then
            FindSQLFile = testPath
            Debug.Print "Found SQL file at: " & testPath
            Exit Function
        End If
    Next i
    
    ' If no file found, return the default expected path
    FindSQLFile = basePath & "\src\schema\migrate_to_counter_pks.sql"
    Debug.Print "No SQL file found. Using default: " & FindSQLFile
End Function

Public Function TestFilePath() As Boolean
    ' Simple test function to check file path detection
    ' Use this to debug path issues
    
    Dim scriptPath As String
    
    Debug.Print "=== FILE PATH TEST ==="
    Debug.Print "Current database location: " & CurrentProject.Path
    Debug.Print "Current database full name: " & CurrentProject.FullName
    Debug.Print ""
    
    scriptPath = FindSQLFile()
    Debug.Print "Final SQL path: " & scriptPath
    Debug.Print "File exists: " & (Dir(scriptPath) <> "")
    
    If Dir(scriptPath) <> "" Then
        Dim fileSize As Long
        fileSize = FileLen(scriptPath)
        Debug.Print "File size: " & fileSize & " bytes"
    End If
    
    MsgBox "File path test complete. Check Immediate Window (Ctrl+G) for results.", vbInformation
    TestFilePath = True
End Function