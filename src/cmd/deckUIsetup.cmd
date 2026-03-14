@echo off
setlocal EnableExtensions

:: deckUI Setup — one-click install/update/uninstall
:: Downloads deckUI.ppam from GitHub and installs to %APPDATA%\Microsoft\AddIns
:: Registers via HKCU registry so PowerPoint loads it automatically.
:: No admin rights needed.

set "ADDIN=deckUI.ppam"
set "ADDIN_NAME=deckUI"
set "DEST=%APPDATA%\Microsoft\AddIns"
set "DST=%DEST%\%ADDIN%"
set "REGKEY=HKCU\Software\Microsoft\Office\16.0\PowerPoint\AddIns\%ADDIN_NAME%"
set "DOWNLOAD_URL=https://raw.githubusercontent.com/therepos/deckUI/main/src/cmd/deckUI.ppam"

echo.
echo  ========================================
echo    deckUI - PowerPoint Add-in Setup
echo  ========================================
echo.

:: Check if already installed
if exist "%DST%" goto :existing

:: ---- FRESH INSTALL ----
:install

:: Close PowerPoint if running
call :checkppt
if %errorlevel%==1 exit /b 1

echo  Downloading %ADDIN% ...
call :download
if %errorlevel%==1 exit /b 1

call :register

echo.
echo  Installed. The DeckUI tab will appear next time
echo  you open PowerPoint.
echo.
pause
exit /b 0

:: ---- ALREADY INSTALLED ----
:existing

echo  deckUI is already installed at:
echo    %DST%
echo.
echo  [U] Update    - download latest version
echo  [R] Uninstall - remove add-in
echo  [C] Cancel
echo.
set /p "ANS=  Choose (U/R/C): "
if /i "%ANS%"=="U" goto :update
if /i "%ANS%"=="R" goto :uninstall
echo  Cancelled.
echo.
pause
exit /b 0

:: ---- UPDATE ----
:update

call :checkppt
if %errorlevel%==1 exit /b 1

echo  Downloading latest %ADDIN% ...
call :download
if %errorlevel%==1 exit /b 1

call :register

echo.
echo  Updated. Restart PowerPoint to load the new version.
echo.
pause
exit /b 0

:: ---- UNINSTALL ----
:uninstall

call :checkppt
if %errorlevel%==1 exit /b 1

reg delete "%REGKEY%" /f >nul 2>&1
del /f "%DST%" >nul 2>&1

if exist "%DST%" (
    echo  ERROR: Could not remove %ADDIN%.
    echo.
    pause
    exit /b 1
)

echo.
echo  deckUI uninstalled.
echo.
pause
exit /b 0

:: ===========================================================
::  SUBROUTINES
:: ===========================================================

:checkppt
tasklist /fi "imagename eq POWERPNT.EXE" 2>nul | find /i "POWERPNT.EXE" >nul
if %errorlevel%==0 (
    echo  PowerPoint is running. Please close it first.
    echo.
    pause
    exit /b 1
)
exit /b 0

:download
if not exist "%DEST%" mkdir "%DEST%"
powershell -NoProfile -Command "try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri '%DOWNLOAD_URL%' -OutFile '%DST%' -UseBasicParsing; exit 0 } catch { Write-Host '  ERROR:' $_.Exception.Message; exit 1 }"
if errorlevel 1 (
    echo  Download failed. Check your internet connection.
    echo.
    pause
    exit /b 1
)
echo  Downloaded successfully.
exit /b 0

:register
reg add "%REGKEY%" /v "Path" /t REG_SZ /d "%DST%" /f >nul 2>&1
reg add "%REGKEY%" /v "AutoLoad" /t REG_DWORD /d 1 /f >nul 2>&1
if errorlevel 1 (
    echo  File copied but could not write registry.
    echo  Enable manually: File ^> Options ^> Add-Ins ^> Go
)
exit /b 0