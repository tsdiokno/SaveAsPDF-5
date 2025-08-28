@echo off
setlocal

:: ---- CONFIG ----
set ADDIN_NAME=SaveAsPDF-5.ppam
set SOURCE_DIR=%~dp0        :: Folder where this .bat and the .ppam are placed
set DEST_DIR=%APPDATA%\Microsoft\AddIns

:: ---- CREATE DESTINATION FOLDER ----
if not exist "%DEST_DIR%" (
    echo Creating AddIns folder: "%DEST_DIR%"
    mkdir "%DEST_DIR%"
)

:: ---- COPY THE ADD-IN ----
echo Copying %ADDIN_NAME% to "%DEST_DIR%"
copy /Y "%SOURCE_DIR%%ADDIN_NAME%" "%DEST_DIR%\"

:: ---- AUTO-ENABLE THE ADD-IN ----
:: Office stores enabled add-ins in HKCU\Software\Microsoft\Office\<version>\PowerPoint\AddIns\<addin_name>
:: Adjust 16.0 if using Office 2013 (15.0), Office 2010 (14.0), etc.
set OFFICE_VER=16.0
set ADDIN_KEY=HKCU\Software\Microsoft\Office\%OFFICE_VER%\PowerPoint\AddIns\SaveAsPDF-5

echo Registering add-in in the registry...
reg add "%ADDIN_KEY%" /v Path /t REG_SZ /d "%DEST_DIR%\%ADDIN_NAME%" /f >nul
reg add "%ADDIN_KEY%" /v AutoLoad /t REG_DWORD /d 1 /f >nul
reg add "%ADDIN_KEY%" /v LoadBehavior /t REG_DWORD /d 3 /f >nul
reg add "%ADDIN_KEY%" /v FriendlyName /t REG_SZ /d "Save As PDF Add-in" /f >nul
reg add "%ADDIN_KEY%" /v Description /t REG_SZ /d "Exports presentations as PDF in same folder" /f >nul

echo.
echo ✅ Deployment complete.
echo The add-in "%ADDIN_NAME%" has been copied and auto-enabled.
echo Please restart PowerPoint to see the new functionality.
echo.

pause
endlocal
