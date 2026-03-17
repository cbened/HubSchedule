[CmdletBinding()]
param(
    [Parameter()]
    [string]$DllPath = "src/HubSchedule.OutlookAddIn/bin/Release/net48/HubSchedule.OutlookAddIn.dll",

    [Parameter()]
    [string]$RegAsmPath
)

$ErrorActionPreference = "Stop"

function Resolve-RegAsmPath {
    param([string]$ExplicitPath)

    if ($ExplicitPath) {
        if (-not (Test-Path $ExplicitPath)) {
            throw "RegAsm not found at '$ExplicitPath'."
        }

        return (Resolve-Path $ExplicitPath).Path
    }

    $candidates = @(
        "$env:WINDIR/Microsoft.NET/Framework64/v4.0.30319/RegAsm.exe",
        "$env:WINDIR/Microsoft.NET/Framework/v4.0.30319/RegAsm.exe"
    )

    foreach ($candidate in $candidates) {
        if (Test-Path $candidate) {
            return $candidate
        }
    }

    throw "Could not locate RegAsm.exe. Install .NET Framework 4.x developer tools or pass -RegAsmPath."
}

if (Test-Path $DllPath) {
    $dllFullPath = (Resolve-Path $DllPath).Path
    $regasmExe = Resolve-RegAsmPath -ExplicitPath $RegAsmPath

    Write-Host "Unregistering add-in with RegAsm..."
    & $regasmExe $dllFullPath /u
    if ($LASTEXITCODE -ne 0) {
        throw "RegAsm unregister failed with exit code $LASTEXITCODE"
    }
} else {
    Write-Warning "DLL path '$DllPath' not found. Skipping RegAsm /u."
}

$addinRegPath = "HKCU:\Software\Microsoft\Office\Outlook\Addins\HubSchedule.OutlookAddIn"
if (Test-Path $addinRegPath) {
    Remove-Item -Path $addinRegPath -Recurse -Force
    Write-Host "Removed Outlook add-in registry key."
} else {
    Write-Host "Outlook add-in registry key was not present."
}

Write-Host "Uninstall complete."
