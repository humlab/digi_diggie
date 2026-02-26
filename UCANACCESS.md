# UCanAccess Setup Guide

UCanAccess is a JDBC driver for MS Access databases that allows reading and writing Access files without needing Microsoft Access installed.

## Download UCanAccess

1. Download UCanAccess from: https://sourceforge.net/projects/ucanaccess/files/
2. Extract the zip file
3. Create a directory `lib/ucanaccess` in the project root
4. Copy all JAR files from the UCanAccess download to `lib/ucanaccess/`

Required JAR files include:
- ucanaccess-x.x.x.jar
- jackcess-x.x.x.jar
- commons-lang3-x.x.x.jar
- commons-logging-x.x.x.jar
- hsqldb-x.x.x.jar

## Directory Structure

```
digi_diggie/
├── lib/
│   └── ucanaccess/
│       ├── ucanaccess-5.0.1.jar
│       ├── jackcess-4.0.4.jar
│       ├── commons-lang3-3.13.0.jar
│       ├── commons-logging-1.2.jar
│       └── hsqldb-2.7.2.jar
└── src/
    └── accdb_to_excel.py
```

## Usage

Once UCanAccess is set up, you can use the script:

```bash
# Export all tables from an Access database
python src/accdb_to_excel.py -d data/database.accdb -o output.xlsx

# Verbose mode
python src/accdb_to_excel.py -d data/database.accdb -o output.xlsx -v

# Export specific tables only
python src/accdb_to_excel.py -d data/database.accdb -o output.xlsx -i Table1 -i Table2

# Exclude certain tables
python src/accdb_to_excel.py -d data/database.accdb -o output.xlsx -e MSysObjects -e TempTable

# Specify custom UCanAccess location
python src/accdb_to_excel.py -d data/database.accdb -o output.xlsx --ucanaccess-jar /path/to/ucanaccess
```

## Alternative: Using pyodbc (Windows only)

If you're on Windows and have MS Access drivers installed, you can use pyodbc instead, which doesn't require UCanAccess. See the commented code in `src/deprecated/code.py` for examples.
