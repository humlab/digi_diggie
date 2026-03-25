# RunMigrationScript.bas Usage Instructions

## Overview
This VBA module provides functions to safely execute the `migrate_to_counter_pks.sql` script in MS Access, with built-in safety features and error handling.

## How to Use

### 1. Import the VBA Module
1. Open your Access database
2. Press `Alt + F11` to open VBA Editor
3. Go to `File → Import File...`
4. Select `RunMigrationScript.bas` from the `src/vba/` folder
5. The module will be imported as "RunMigrationScript"

### 2. Run the Migration
Choose one of these options based on your needs:

#### Option A: Run with Automatic Backup (Recommended)
```vba
?RunMigrationWithBackup()
```
- Creates timestamped backup automatically
- Runs the migration script
- Returns True if successful

#### Option B: Run Migration Only
```vba
?RunMigrationFromFile()
```
- Prompts for confirmation before proceeding
- Runs the migration script
- **Make sure you have a manual backup first!**

#### Option C: Step-by-Step Migration (For Troubleshooting)
```vba
?RunMigrationStepByStep()
```
- Executes each SQL statement individually with feedback
- Shows progress in Immediate Window (Ctrl+G)
- Option to continue on errors or stop at first failure
- **Best for debugging when main migration fails**

#### Option D: Preview SQL Statements (Debug Only)
```vba
?PreviewSQLStatements()
```
- Shows all parsed SQL statements without executing
- Displays results in Immediate Window
- **Use this first if migration isn't working**

#### Option E: Create Backup Only
```vba
?BackupDatabase()
```
- Creates a timestamped backup in the database folder
- Useful for manual backup before migration

### 3. Post-Migration Steps
After successful migration:

1. **Verify Data**: Check that all records are present and ID values preserved
2. **Recreate Relationships**: 
   - Go to `Database Tools → Relationships`
   - Add all tables and recreate foreign key relationships
3. **Compact Database**: 
   - Go to `Database Tools → Compact and Repair Database`

## Functions Reference

| Function | Purpose | Returns |
|----------|---------|---------|
| `RunMigrationWithBackup()` | Complete process with backup | Boolean |
| `RunMigrationFromFile()` | Run migration with confirmation | Boolean |
| `RunMigrationStepByStep()` | Execute with detailed feedback | Boolean |
| `PreviewSQLStatements()` | Show parsed SQL without executing | Boolean |
| `ExecuteSQLFromFile(path)` | Execute any SQL file | Boolean |
| `BackupDatabase()` | Create timestamped backup | Boolean |

## Error Handling
- All functions return `True` for success, `False` for errors
- Detailed error messages are displayed via MsgBox
- File not found, permission errors, and SQL errors are handled gracefully
- Migration stops on first error to prevent partial corruption

## File Structure Expected
```
YourDatabase.accdb
├── src/
│   └── schema/
│       └── migrate_to_counter_pks.sql
└── backup_YYYYMMDD_HHMMSS.accdb (created automatically)
```

## Safety Features
- ✅ Automatic backup creation
- ✅ User confirmation before destructive operations  
- ✅ File existence checking
- ✅ SQL parsing and cleanup
- ✅ Transaction-like behavior (fails fast on errors)
- ✅ Progress feedback via DoEvents

## Troubleshooting

### "Tables Not Updating" Issue
If the script runs without errors but tables don't change:

1. **Check SQL Parsing**:
   ```vba
   ?PreviewSQLStatements()
   ```
   - Verify statements are parsed correctly
   - Check Immediate Window (Ctrl+G) for full output

2. **Run Step-by-Step**:
   ```vba
   ?RunMigrationStepByStep()
   ```
   - See exactly which statements succeed/fail
   - Enable "Continue on errors" to see all issues

3. **Check Debug Output**:
   - Open Immediate Window (Ctrl+G)
   - Look for "Executing:" messages
   - Check for error messages

### Common Causes
- **SQL Syntax**: Access uses Jet SQL, not standard SQL
- **Table Dependencies**: Existing relationships prevent table dropping
- **Permissions**: Database may be read-only or locked
- **File Path**: SQL file not found in expected location

### "File not found" Error
- Ensure `migrate_to_counter_pks.sql` exists in `src/schema/` relative to your database
- Check file path spelling and case sensitivity

### "Permission denied" Error
- Make sure Access has write permissions to the database folder
- Close any other applications that might have the database open

### SQL Execution Errors
- Check that all referenced tables exist
- Ensure no other users are connected to the database
- Verify the SQL syntax is compatible with Access/Jet SQL

### Relationship Errors
- Delete existing relationships before migration
- Recreate relationships manually after migration completes

## Manual Alternative
If the VBA script fails, you can still run the SQL manually:
1. Create a backup manually
2. Copy sections of the SQL script
3. Run each table migration individually in Query Design → SQL View
4. Start with tables that have no dependencies (lookup tables)

---
**Important**: Always test on a copy of your database first!