@echo off
setlocal

:: ---- CONFIG ----
set ADDIN_NAME=SaveAsPDF-5.ppam
set SOURCE_DIR=%~dp0
set DEST_DIR=%APPDATA%\Microsoft\AddIns
set FRIENDLY_NAME=Save As PDF Add-in
set DESCRIPTION=Exports presentations as PDF in same folder

:: ---- DETECT OFFICE VERSION ----
set OFFICE_VER=
for %%V in (16.0 15.0 14.0 12.0) do (
    reg query "HKCU\Software\Microsoft\Office\%%V\PowerPoint" >nul 2>&1
    if not errorlevel 1 (
        set OFFICE_VER=%%V
        goto :found
    )
)

:found
if "%OFFICE_VER%"=="" (
    echo ⚠️ Could not detect Office version. Registry operations may be skipped.
)

set ADDIN_KEY=HKCU\Software\Microsoft\Office\%OFFICE_VER%\PowerPoint\AddIns\SaveAsPDF-5
set STARTUP_KEY=HKCU\Software\Microsoft\Office\%OFFICE_VER%\PowerPoint\AddInsStartup\SaveAsPDF-5

:: ---- UNINSTALL MODE ----
if /I "%~1"=="/uninstall" (
    echo 🔴 Uninstalling %ADDIN_NAME% ...

    echo Removing add-in file...
    del /Q "%DEST_DIR%\%ADDIN_NAME%" >nul 2>&1

    echo Removing registry keys...
    reg delete "%ADDIN_KEY%" /f >nul 2>&1
    reg delete "%STARTUP_KEY%" /f >nul 2>&1

    echo ✅ Uninstall complete.
    pause
    exit /b
)

:: ---- NORMAL INSTALL ----
echo 🟢 Installing %ADDIN_NAME% ...

:: Create destination folder
if not exist "%DEST_DIR%" (
    echo Creating AddIns folder: "%DEST_DIR%"
    mkdir "%DEST_DIR%"
)

:: Copy the add-in
echo Copying %ADDIN_NAME% to "%DEST_DIR%"
copy /Y "%SOURCE_DIR%%ADDIN_NAME%" "%DEST_DIR%\" >nul

:: Clean old registry entries
echo Cleaning old registry entries...
reg delete "%ADDIN_KEY%" /f >nul 2>&1
reg delete "%STARTUP_KEY%" /f >nul 2>&1

:: Register the add-in (simulate "checked once")
echo Registering add-in in registry...
reg add "%ADDIN_KEY%" /v Path /t REG_SZ /d "%DEST_DIR%\%ADDIN_NAME%" /f >nul
reg add "%ADDIN_KEY%" /v AutoLoad /t REG_DWORD /d 1 /f >nul
reg add "%ADDIN_KEY%" /v LoadBehavior /t REG_DWORD /d 3 /f >nul
reg add "%ADDIN_KEY%" /v FriendlyName /t REG_SZ /d "%FRIENDLY_NAME%" /f >nul
reg add "%ADDIN_KEY%" /v Description /t REG_SZ /d "%DESCRIPTION%" /f >nul

echo ✅ Install complete. Please restart PowerPoint.
pause
endlocal
