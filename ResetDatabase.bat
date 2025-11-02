@echo off

REM Reset the database by running the reset script in MS Access

REM Determine this script's directory and set an environment variable
set SCRIPT_DIR=%~dp0

echo Resetting database using script located at %SCRIPT_DIR%

set ACCESS_PATH="C:\Program Files\Microsoft Office\root\Office16\MSACCESS.exe"
set DB_ORIGINAL_PATH="%SCRIPT_DIR%\digidiggie_original.accdb"
set DB_DEV_PATH="%SCRIPT_DIR%\digidiggie_dev.accdb"
set LOAD_MACRO="ImportOrReplaceScript"

REM "Kill any running instance of MS Access"
taskkill /IM MSACCESS.EXE /F >nul 2>&1

REM "Copy template (original) database to reset the current dev database"
copy /Y %DB_ORIGINAL_PATH% %DB_DEV_PATH%

REM Add reset_script.bas to the database and run it

%ACCESS_PATH% %DB_DEV_PATH% /x %LOAD_MACRO%
