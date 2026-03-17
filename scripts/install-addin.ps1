[CmdletBinding()]
param(
    [Parameter()]
    [string]$DllPath = "src/HubSchedule.OutlookAddIn/bin/Release/net48/HubSchedule.OutlookAddIn.dll",

    [Parameter()]
    [string]$RegAsmPath,

    [Parameter()]
    [switch]$SkipRegasm
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

$dllFullPath = Resolve-Path $DllPath -ErrorAction Stop
Write-Host "Using add-in DLL: $dllFullPath"

if (-not $SkipRegasm) {
    $regasmExe = Resolve-RegAsmPath -ExplicitPath $RegAsmPath
    Write-Host "Using RegAsm: $regasmExe"

    & $regasmExe $dllFullPath /codebase /tlb
    if ($LASTEXITCODE -ne 0) {
        throw "RegAsm registration failed with exit code $LASTEXITCODE"
    }
}

$addinRegPath = "Software\\Microsoft\\Office\\Outlook\\Addins\\HubSchedule.OutlookAddIn"
$hkcu = [Microsoft.Win32.Registry]::CurrentUser
$key = $hkcu.CreateSubKey($addinRegPath)

$key.SetValue($null, "HubSchedule Outlook Add-in", [Microsoft.Win32.RegistryValueKind]::String)
$key.SetValue("FriendlyName", "HubSchedule Outlook Add-in", [Microsoft.Win32.RegistryValueKind]::String)
$key.SetValue("Description", "Open HubSpot submissions in classic Outlook", [Microsoft.Win32.RegistryValueKind]::String)
$key.SetValue("LoadBehavior", 3, [Microsoft.Win32.RegistryValueKind]::DWord)
$key.SetValue("CommandLineSafe", 0, [Microsoft.Win32.RegistryValueKind]::DWord)
$key.Close()

Write-Host "Add-in registration complete. Restart Outlook."
