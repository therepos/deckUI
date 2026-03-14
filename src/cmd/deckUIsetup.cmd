@echo off
setlocal EnableExtensions

:: deckUI Setup — one-click install/uninstall
:: Copies deckUI.ppam to %APPDATA%\Microsoft\AddIns
:: No admin rights needed. PowerPoint auto-discovers add-ins in this folder.

set "ADDIN=deckUI.ppam"
set "SRC=%~dp0%ADDIN%"
set "DEST=%APPDATA%\Microsoft\AddIns"
set "DST=%DEST%\%ADDIN%"

echo.
echo  ========================================
echo    deckUI - PowerPoint Add-in Setup
echo  ========================================
echo.

:: Check source file exists
if not exist "%SRC%" (
    echo  ERROR: %ADDIN% not found next to this setup file.
    echo  Place %ADDIN% in the same folder as this script.
    echo.
    pause
    exit /b 1
)

:: Check if already installed
if exist "%DST%" goto :uninstall

:: ---- INSTALL ----
:install

:: Close PowerPoint if running
tasklist /fi "imagename eq POWERPNT.EXE" 2>nul | find /i "POWERPNT.EXE" >nul
if %errorlevel%==0 (
    echo  PowerPoint is running. Please close it first.
    echo.
    pause
    exit /b 1
)

:: Ensure target folder exists
if not exist "%DEST%" mkdir "%DEST%"

:: Copy
copy /y "%SRC%" "%DST%" >nul
if errorlevel 1 (
    echo  ERROR: Failed to copy %ADDIN%.
    echo.
    pause
    exit /b 1
)

echo  Installed to: %DST%
echo.
echo  Next steps:
echo    1. Open PowerPoint
echo    2. File ^> Options ^> Add-Ins
echo    3. At the bottom: Manage = "PowerPoint Add-ins" ^> Go
echo    4. Check "deckUI" and click OK
echo.
echo  You only need to do this once. After that the
echo  DeckUI tab will appear every time you open PowerPoint.
echo.
pause
exit /b 0

:: ---- UNINSTALL ----
:uninstall

echo  deckUI is already installed at:
echo    %DST%
echo.
set /p "ANS=  Uninstall? (Y/N): "
if /i not "%ANS%"=="Y" (
    echo  Cancelled.
    echo.
    pause
    exit /b 0
)

:: Close PowerPoint if running
tasklist /fi "imagename eq POWERPNT.EXE" 2>nul | find /i "POWERPNT.EXE" >nul
if %errorlevel%==0 (
    echo  PowerPoint is running. Please close it first.
    echo.
    pause
    exit /b 1
)

del /f "%DST%" >nul 2>&1
if exist "%DST%" (
    echo  ERROR: Could not remove %ADDIN%.
    echo.
    pause
    exit /b 1
)

echo  deckUI uninstalled.
echo.
pause
exit /b 0
