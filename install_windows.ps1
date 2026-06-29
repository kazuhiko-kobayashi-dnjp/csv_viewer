# install_windows.ps1
# Run:  powershell -ExecutionPolicy Bypass -File install_windows.ps1
#
# 1) Installs tkinterdnd2 via pip (for in-window drag and drop)
# 2) Associates .csv with csv_viewer.vbs (no console window)
# No virtualenv. Uses the system Python 3.

param(
    [string]$ScriptDir = $PSScriptRoot
)

$ErrorActionPreference = "Stop"

try {
    $VbsPath  = Join-Path $ScriptDir "csv_viewer.vbs"
    $PyScript = Join-Path $ScriptDir "csv_viewer.py"

    # Find a usable Python: prefer the 'py' launcher, then python/python3
    $PyRunner = $null
    $PyArgsPrefix = @()
    foreach ($cand in @("py", "python", "python3")) {
        $c = Get-Command $cand -ErrorAction SilentlyContinue
        if ($c) {
            $PyRunner = $cand
            if ($cand -eq "py") { $PyArgsPrefix = @("-3") }
            break
        }
    }
    if (-not $PyRunner) {
        throw "Python not found on PATH. Install Python 3 first (and tick 'Add to PATH')."
    }

    if (-not (Test-Path $VbsPath))  { throw "csv_viewer.vbs not found: $VbsPath" }
    if (-not (Test-Path $PyScript)) { throw "csv_viewer.py not found: $PyScript" }

    Write-Host "Python launcher: $PyRunner $($PyArgsPrefix -join ' ')"
    Write-Host "Installing tkinterdnd2 ..."
    & $PyRunner @PyArgsPrefix -m pip install --user tkinterdnd2 | Out-Host

    $WScript = Join-Path $env:WINDIR "System32\wscript.exe"
    $ProgId  = "CSVViewer.csvfile"
    $FriendlyName = "CSV Viewer"
    $Command = "`"$WScript`" `"$VbsPath`" `"%1`""

    New-Item -Path "HKCU:\Software\Classes\$ProgId" -Force | Out-Null
    Set-ItemProperty -Path "HKCU:\Software\Classes\$ProgId" -Name "(Default)" -Value $FriendlyName

    $cmdKey = "HKCU:\Software\Classes\$ProgId\shell\open\command"
    New-Item -Path $cmdKey -Force | Out-Null
    Set-ItemProperty -Path $cmdKey -Name "(Default)" -Value $Command

    $iconKey = "HKCU:\Software\Classes\$ProgId\DefaultIcon"
    New-Item -Path $iconKey -Force | Out-Null
    Set-ItemProperty -Path $iconKey -Name "(Default)" -Value "$WScript,0"

    $extKey = "HKCU:\Software\Classes\.csv"
    New-Item -Path $extKey -Force | Out-Null
    Set-ItemProperty -Path $extKey -Name "(Default)" -Value $ProgId

    $code = @"
using System;
using System.Runtime.InteropServices;
public class Shell {
    [DllImport("shell32.dll")]
    public static extern void SHChangeNotify(int wEventId, uint uFlags, IntPtr dwItem1, IntPtr dwItem2);
}
"@
    Add-Type -TypeDefinition $code -Language CSharp
    [Shell]::SHChangeNotify(0x08000000, 0, [IntPtr]::Zero, [IntPtr]::Zero)

    Write-Host ""
    Write-Host "Done. Double-clicking a .csv file will open it with CSV Viewer."
    Write-Host "Command: $Command"
}
catch {
    Write-Host ""
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    Write-Host ""
    Read-Host "Press Enter to close"
}
