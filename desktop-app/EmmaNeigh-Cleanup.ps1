# EmmaNeigh Complete Cleanup Script
# Run this as Administrator to remove all traces of broken EmmaNeigh installations
# Then install fresh with EmmaNeigh-Setup.exe
#
# To run: Right-click this file > "Run with PowerShell"
# Or open PowerShell as Admin and run: Set-ExecutionPolicy Bypass -Scope Process; .\EmmaNeigh-Cleanup.ps1

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  EmmaNeigh Complete Cleanup Tool" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

$removed = 0

# 1. Kill any running EmmaNeigh processes
Write-Host "[1/6] Stopping EmmaNeigh processes..." -ForegroundColor Yellow
$procs = Get-Process -Name "EmmaNeigh" -ErrorAction SilentlyContinue
if ($procs) {
    $procs | Stop-Process -Force
    Write-Host "  Stopped running EmmaNeigh processes" -ForegroundColor Green
    $removed++
} else {
    Write-Host "  No running processes found" -ForegroundColor Gray
}

# 2. Remove installation directories
Write-Host "[2/6] Removing installation folders..." -ForegroundColor Yellow
$installPaths = @(
    "$env:LOCALAPPDATA\Programs\EmmaNeigh",
    "$env:LOCALAPPDATA\Programs\emmaneigh",
    "$env:ProgramFiles\EmmaNeigh",
    "$env:ProgramFiles\emmaneigh",
    "${env:ProgramFiles(x86)}\EmmaNeigh",
    "${env:ProgramFiles(x86)}\emmaneigh",
    "$env:LOCALAPPDATA\emmaneigh-updater"
)

foreach ($p in $installPaths) {
    if (Test-Path $p) {
        Remove-Item -Path $p -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "  Removed: $p" -ForegroundColor Green
        $removed++
    }
}

# 3. Remove app data
Write-Host "[3/6] Removing app data..." -ForegroundColor Yellow
$dataPaths = @(
    "$env:APPDATA\EmmaNeigh",
    "$env:APPDATA\emmaneigh",
    "$env:LOCALAPPDATA\EmmaNeigh",
    "$env:LOCALAPPDATA\emmaneigh"
)

foreach ($p in $dataPaths) {
    if (Test-Path $p) {
        Remove-Item -Path $p -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "  Removed: $p" -ForegroundColor Green
        $removed++
    }
}

# 4. Clean registry - Uninstall entries
Write-Host "[4/6] Cleaning registry uninstall entries..." -ForegroundColor Yellow
$regPaths = @(
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall",
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
)

foreach ($regPath in $regPaths) {
    if (Test-Path $regPath) {
        Get-ChildItem $regPath -ErrorAction SilentlyContinue | ForEach-Object {
            $displayName = (Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue).DisplayName
            $publisher = (Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue).Publisher
            if ($displayName -like "*EmmaNeigh*" -or $displayName -like "*emmaneigh*") {
                Remove-Item $_.PSPath -Recurse -Force -ErrorAction SilentlyContinue
                Write-Host "  Removed registry key: $($_.Name)" -ForegroundColor Green
                $removed++
            }
        }
    }
}

# 5. Clean NSIS-specific registry entries
Write-Host "[5/6] Cleaning NSIS installer entries..." -ForegroundColor Yellow
$nsisKeys = @(
    "HKCU:\Software\emmaneigh-transaction-app",
    "HKCU:\Software\EmmaNeigh",
    "HKCU:\Software\com.emmaneigh.app",
    "HKLM:\SOFTWARE\emmaneigh-transaction-app",
    "HKLM:\SOFTWARE\EmmaNeigh",
    "HKLM:\SOFTWARE\com.emmaneigh.app"
)

foreach ($key in $nsisKeys) {
    if (Test-Path $key) {
        Remove-Item -Path $key -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "  Removed: $key" -ForegroundColor Green
        $removed++
    }
}

# 6. Remove desktop shortcuts and Start Menu entries
Write-Host "[6/6] Removing shortcuts..." -ForegroundColor Yellow
$shortcutPaths = @(
    "$env:USERPROFILE\Desktop\EmmaNeigh.lnk",
    "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\EmmaNeigh.lnk",
    "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\EmmaNeigh\*",
    "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\EmmaNeigh.lnk",
    "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\EmmaNeigh\*"
)

foreach ($p in $shortcutPaths) {
    $items = Get-Item $p -ErrorAction SilentlyContinue
    if ($items) {
        $items | Remove-Item -Force -ErrorAction SilentlyContinue
        Write-Host "  Removed: $p" -ForegroundColor Green
        $removed++
    }
}

# Summary
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
if ($removed -gt 0) {
    Write-Host "  Cleanup complete! Removed $removed items." -ForegroundColor Green
} else {
    Write-Host "  No EmmaNeigh traces found." -ForegroundColor Green
}
Write-Host "  You can now install EmmaNeigh fresh." -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Press any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
