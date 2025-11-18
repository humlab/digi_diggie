# Linking MS Access to PostgreSQL Database on Windows 11

This guide provides step-by-step instructions for setting up linked tables in Microsoft Access that connect to a PostgreSQL database on Windows 11.

## Prerequisites

- Windows 11
- Microsoft Access (2016 or later recommended)
- PostgreSQL database server (local or remote)
- Administrator privileges on Windows 11

## Step 1: Install PostgreSQL ODBC Driver

1. **Download the PostgreSQL ODBC Driver**
   - Visit the official PostgreSQL ODBC driver page: https://www.postgresql.org/ftp/odbc/versions/msi/
   - Download the appropriate version for your system:
     - `psqlodbc_x64.msi` for 64-bit systems (recommended)
     - `psqlodbc_x86.msi` for 32-bit systems
   - **Important**: If you're using 32-bit MS Access on a 64-bit Windows, you need the 32-bit ODBC driver

2. **Install the ODBC Driver**
   - Run the downloaded `.msi` file
   - Follow the installation wizard
   - Accept the license agreement
   - Use default installation settings unless you have specific requirements
   - Complete the installation

3. **Verify Installation**
   - Press `Win + R` to open Run dialog
   - Type `odbcad32.exe` and press Enter
   - For 32-bit ODBC drivers on 64-bit Windows, use: `C:\Windows\SysWOW64\odbcad32.exe`
   - Go to the "Drivers" tab
   - Look for "PostgreSQL ANSI" or "PostgreSQL Unicode" in the list

## Step 2: Create a System DSN (Data Source Name)

1. **Open ODBC Data Source Administrator**
   - Press `Win + X` and select "Run"
   - Type `odbcad32.exe` and press Enter
   - **Note**: Use the correct version:
     - 64-bit: `C:\Windows\System32\odbcad32.exe`
     - 32-bit: `C:\Windows\SysWOW64\odbcad32.exe`

2. **Create a New System DSN**
   - Click on the "System DSN" tab
   - Click "Add..." button
   - Select "PostgreSQL Unicode(x64)" or "PostgreSQL Unicode" from the list
   - Click "Finish"

3. **Configure the PostgreSQL Connection**
   Fill in the following fields:
   
   - **Data Source**: Give your connection a descriptive name (e.g., `DigiDiggie_PostgreSQL`)
   - **Database**: Enter your PostgreSQL database name (e.g., `digidiggie`)
   - **Server**: Enter the server address
     - For local: `localhost` or `127.0.0.1`
     - For remote: IP address or hostname
   - **Port**: Default is `5432` (change if your PostgreSQL uses a different port)
   - **User Name**: PostgreSQL username (e.g., `postgres` or your specific user)
   - **Password**: PostgreSQL password
   - **SSL Mode**: Choose based on your security requirements:
     - `disable` - No SSL (local development)
     - `require` - SSL required (production)
     - `prefer` - Use SSL if available

4. **Test the Connection**
   - Click "Test" button to verify the connection
   - If successful, you'll see "Connection successful"
   - If failed, verify:
     - PostgreSQL service is running
     - Firewall allows connection on port 5432
     - Credentials are correct
     - Network connectivity (for remote servers)

5. **Save the DSN**
   - Click "Save" to save the DSN configuration
   - Click "OK" to close the ODBC Administrator

## Step 3: Link Tables in MS Access

### Method 1: Using External Data Wizard

1. **Open MS Access Database**
   - Open your MS Access database (e.g., `digidiggie_dev.accdb`)

2. **Start the Import/Link Process**
   - Go to the "External Data" tab in the ribbon
   - In the "Import & Link" group, click "ODBC Database"
   - Select "Link to the data source by creating a linked table"
   - Click "OK"

3. **Select the DSN**
   - In the "Select Data Source" dialog
   - Go to the "Machine Data Source" tab
   - Select your DSN (e.g., `DigiDiggie_PostgreSQL`)
   - Click "OK"
   - Enter the password if prompted

4. **Select Tables to Link**
   - A list of available PostgreSQL tables and views will appear
   - Select the tables you want to link (use Ctrl+Click for multiple selections)
   - Common options:
     - **Save password**: Check to avoid re-entering password (security consideration)
     - **Select unique record identifier**: Access will prompt you to choose a primary key if it can't auto-detect

5. **Choose Unique Identifiers**
   - For each table without a detected primary key, Access will ask you to select unique identifier columns
   - Select the appropriate column(s) that uniquely identify each record
   - This is crucial for updating records through the linked table

6. **Complete the Linking**
   - Click "OK" to complete the linking process
   - Linked tables will appear in the Navigation Pane with a globe icon and arrow

### Method 2: Using VBA Code

You can automate the linking process using VBA:

```vba
Sub LinkPostgreSQLTables()
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim strDSN As String
    Dim strTable As String
    Dim arrTables() As Variant
    Dim i As Integer
    
    ' Set your DSN name
    strDSN = "ODBC;DSN=DigiDiggie_PostgreSQL;UID=postgres;PWD=your_password;"
    
    ' List of tables to link
    arrTables = Array("entries", "personer", "byar", "socken", "ortnamn")
    
    Set db = CurrentDb
    
    ' Loop through tables
    For i = LBound(arrTables) To UBound(arrTables)
        strTable = arrTables(i)
        
        ' Delete existing linked table if it exists
        On Error Resume Next
        db.TableDefs.Delete strTable
        On Error GoTo 0
        
        ' Create new linked table
        Set tdf = db.CreateTableDef(strTable)
        tdf.Connect = strDSN
        tdf.SourceTableName = strTable
        db.TableDefs.Append tdf
        
        Debug.Print "Linked table: " & strTable
    Next i
    
    Set tdf = Nothing
    Set db = Nothing
    
    MsgBox "Tables linked successfully!", vbInformation
End Sub
```

## Step 4: Verify Linked Tables

1. **Check Table Links**
   - In Access Navigation Pane, linked tables show a globe icon
   - Double-click a linked table to open it and view data
   - Verify that data displays correctly

2. **Test CRUD Operations**
   - **Create**: Try adding a new record
   - **Read**: Verify existing records display correctly
   - **Update**: Try modifying a record
   - **Delete**: Try deleting a test record (be careful!)

3. **Refresh Linked Tables**
   - If table structure changes in PostgreSQL:
     - Right-click the linked table
     - Select "Linked Table Manager"
     - Check the tables to refresh
     - Click "OK" to refresh

## Step 5: Refresh or Update Linked Tables

### Using Linked Table Manager

1. Go to "External Data" tab
2. Click "Linked Table Manager"
3. Select tables to refresh
4. Click "OK"
5. Re-enter password if needed

### Using VBA to Refresh All Links

```vba
Sub RefreshAllLinkedTables()
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim strConnect As String
    
    Set db = CurrentDb
    
    For Each tdf In db.TableDefs
        If Len(tdf.Connect) > 0 Then
            strConnect = tdf.Connect
            tdf.Connect = ";"
            tdf.Connect = strConnect
            tdf.RefreshLink
            Debug.Print "Refreshed: " & tdf.Name
        End If
    Next tdf
    
    Set tdf = Nothing
    Set db = Nothing
    
    MsgBox "All linked tables refreshed!", vbInformation
End Sub
```

## Troubleshooting

### Connection Issues

**Problem**: "Data source name not found"
- **Solution**: Verify DSN is created in correct ODBC Administrator (32-bit vs 64-bit)
- Match your MS Access version (32-bit Access needs 32-bit ODBC driver)

**Problem**: "Could not connect to server"
- **Solution**: 
  - Verify PostgreSQL service is running: `services.msc` → Find "postgresql-x64-XX"
  - Check firewall settings
  - Verify `pg_hba.conf` allows connections from your IP
  - Test connection with pgAdmin or psql

**Problem**: "Password authentication failed"
- **Solution**:
  - Verify username and password
  - Check PostgreSQL user permissions
  - Ensure user has privileges on the database

### Performance Issues

**Problem**: Slow query performance
- **Solution**:
  - Create indexes on frequently queried columns in PostgreSQL
  - Use Pass-Through queries for complex operations
  - Limit record sets with WHERE clauses
  - Consider creating views in PostgreSQL for complex queries

**Problem**: "ODBC--call failed"
- **Solution**:
  - Increase ODBC timeout settings in DSN configuration
  - Check network latency for remote connections
  - Optimize PostgreSQL queries

### Data Type Incompatibilities

**Problem**: Data types don't map correctly
- **Solution**:
  - PostgreSQL `text` → Access `Long Text`
  - PostgreSQL `boolean` → Access `Yes/No`
  - PostgreSQL `timestamp` → Access `Date/Time`
  - Use CAST in PostgreSQL views if needed

**Problem**: "Cannot update" or "Read-only recordset"
- **Solution**:
  - Ensure table has a primary key
  - Relink table and specify unique record identifier
  - Check user permissions in PostgreSQL (GRANT UPDATE, INSERT, DELETE)

### Linked Table Manager Issues

**Problem**: Tables show as broken links
- **Solution**:
  - Relink using Linked Table Manager
  - Verify PostgreSQL server is accessible
  - Check if table was renamed or deleted in PostgreSQL

## Best Practices

1. **Security**
   - Don't save passwords in DSN for production environments
   - Use least-privilege PostgreSQL users
   - Enable SSL for remote connections
   - Use Windows Authentication when possible (with SSPI)

2. **Performance**
   - Index frequently queried columns
   - Use Pass-Through queries for complex SQL
   - Minimize data transfer with appropriate WHERE clauses
   - Consider local caching for lookup tables

3. **Maintenance**
   - Document your DSN configurations
   - Keep ODBC drivers updated
   - Test links after PostgreSQL schema changes
   - Maintain backup of DSN configurations

4. **Development**
   - Use separate DSNs for development and production
   - Test schema changes in development first
   - Version control your Access database
   - Document any PostgreSQL-specific SQL syntax used

## Additional Resources

- PostgreSQL ODBC Driver Documentation: https://odbc.postgresql.org/
- Microsoft Access ODBC Documentation: https://docs.microsoft.com/en-us/office/client-developer/access/
- PostgreSQL Documentation: https://www.postgresql.org/docs/

## Notes for DigiDiggie Project

For this specific project:
- Database name: `digidiggie`
- Common tables: `entries`, `personer`, `byar`, `socken`, `ortnamn`, etc.
- Consider using Docker PostgreSQL instance (see `docker/docker-compose.yml`)
- If using Docker, ensure port 5432 is exposed and mapped correctly
- Default PostgreSQL Docker credentials may be in environment files

## Automating the Setup

To automate linking for this project, add a VBA module to your Access database:

```vba
Public Function LinkAllTables() As Boolean
    On Error GoTo ErrorHandler
    
    Dim strDSN As String
    Dim arrTables As Variant
    
    ' Configure your DSN connection string
    strDSN = "ODBC;DSN=DigiDiggie_PostgreSQL;DATABASE=digidiggie;"
    
    ' Add all your table names
    arrTables = Array("entries", "personer", "byar", "socken", "ortnamn", _
                     "kommun", "län", "namntyp", "språk", "årstid", _
                     "dom", "källor", "markanvändning", "rättskällor", "vinnare")
    
    ' Link each table
    Dim i As Integer
    For i = LBound(arrTables) To UBound(arrTables)
        Call LinkTable(CStr(arrTables(i)), strDSN)
    Next i
    
    LinkAllTables = True
    Exit Function
    
ErrorHandler:
    MsgBox "Error linking tables: " & Err.Description, vbCritical
    LinkAllTables = False
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
```

Run `LinkAllTables()` function to automatically link all required PostgreSQL tables.
