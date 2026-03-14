@echo off
setlocal EnableExtensions

:: Sanity: required files must be beside this cmd
if not exist "%~dp0deckUI.ppam" (
  echo Missing deckUI.ppam next to deckUIsetup.cmd
  pause
  exit /b 1
)

:: Extract embedded PowerShell payload
set "PAYTAG=::PAYLOAD"
for /f "delims=:" %%A in ('findstr /n /c:"%PAYTAG%" "%~f0"') do set /a LN=%%A+1
set "TMPPS=%TEMP%\deckUI_setup_%RANDOM%.ps1"
more +%LN% "%~f0" > "%TMPPS%"

:: Pass our folder via ENV and PS param; run STA for Office COM if needed
set "SETUP_DIR=%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -Sta -File "%TMPPS%" -SourceDirOverride "%~dp0"
set "rc=%ERRORLEVEL%"
del "%TMPPS%" >nul 2>&1
if not "%rc%"=="0" (
  echo.
  echo Installer reported an error (code %rc%). See messages above.
  pause
)
endlocal
exit /b %rc%

::PAYLOAD
param([string]$SourceDirOverride)

$ErrorActionPreference = 'Stop'
$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'
$Host.UI.RawUI.WindowTitle = 'deckUI setup'

<#
PowerPoint Add-in Installer/Uninstaller (single file)
- Copies deckUI.ppam into PowerPoint's AddIns folder:
  %APPDATA%\Microsoft\AddIns
- No C:\Apps, no Trusted Locations, no COM automation needed.
- After install, enable in PowerPoint: File > Options > Add-Ins >
  Manage: PowerPoint Add-ins > Go > check deckUI
#>

# ===== CONFIG =====
$AddInsDir      = Join-Path $env:APPDATA 'Microsoft\AddIns'
$TargetDir      = $AddInsDir
$FilesToInstall = @('deckUI.ppam')
$AddInFile      = 'deckUI.ppam'
# ==================

# -------- Helpers --------
function Get-ScriptDir {
    if ($env:SETUP_DIR)                 { return $env:SETUP_DIR }
    if ($SourceDirOverride)             { return $SourceDirOverride }
    if ($PSScriptRoot)                  { return $PSScriptRoot }
    if ($PSCommandPath)                 { return (Split-Path -Parent $PSCommandPath) }
    if ($MyInvocation.MyCommand.Path)   { return (Split-Path -Parent $MyInvocation.MyCommand.Path) }
    return (Get-Location).Path
}
$SourceDir = (Get-ScriptDir).TrimEnd('\') + '\'

function Ensure-Dir($p){
    if (-not (Test-Path $p)) {
        New-Item -ItemType Directory -Path $p -Force | Out-Null
    }
}

function Status($msg,[scriptblock]$act,[switch]$Fatal){
    $w = 40
    Write-Host ($msg.PadRight($w)) -NoNewline
    try{
        & $act | Out-Null
        Write-Host "Done" -ForegroundColor Green
    }
    catch{
        Write-Host "Failed" -ForegroundColor Red
        Write-Host ("  " + $_.Exception.Message) -ForegroundColor DarkRed
        if($Fatal){ throw }
    }
}

# Detection (file must exist in AddIns folder)
function Detect-Installed {
    $installed = @()

    $addinsAddin = Join-Path $AddInsDir $AddInFile
    if (Test-Path $addinsAddin) { $installed += $addinsAddin }

    $installed | Select-Object -Unique
}

# Actions
function Install-Addin {
    Status "Installing add-in to AddIns folder" -Fatal {
        Ensure-Dir $AddInsDir

        foreach($f in $FilesToInstall){
            $src = Join-Path $SourceDir $f
            if (-not (Test-Path $src)) {
                throw "Missing file '$f' in $SourceDir"
            }
            $dst = Join-Path $AddInsDir $f
            Copy-Item $src $dst -Force
            try { Unblock-File -Path $dst -ErrorAction SilentlyContinue } catch {}
        }
    }

    Write-Host ""
    Write-Host "To enable the add-in in PowerPoint:" -ForegroundColor Cyan
    Write-Host "  1. Open PowerPoint"
    Write-Host "  2. File > Options > Add-Ins"
    Write-Host "  3. At the bottom, set 'Manage' to 'PowerPoint Add-ins' > Go"
    Write-Host "  4. Click 'Add New...' and select deckUI.ppam"
    Write-Host "     (or check the box if it already appears)"
}

function Uninstall-Addin {
    Status "Removing add-in from AddIns folder" {
        $p = Join-Path $AddInsDir $AddInFile
        if (Test-Path $p) {
            Remove-Item $p -Force -ErrorAction SilentlyContinue
        }
    }
}

# Main
$paths = Detect-Installed
if (-not $paths -or $paths.Count -eq 0) {
    $ans = Read-Host "deckUI is NOT installed. Install now? (Y/N)"
    if ($ans -match '^[Yy]') { Install-Addin }
} else {
    $ans = Read-Host ("deckUI is installed at: " + ($paths -join ', ') + ". Uninstall it? (Y/N)")
    if ($ans -match '^[Yy]') { Uninstall-Addin }
}

Write-Host ""
Write-Host "All done. You can close this window now." -ForegroundColor Yellow

exit 0
