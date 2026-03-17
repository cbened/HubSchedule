[CmdletBinding()]
param(
    [Parameter()]
    [string]$Configuration = "Release",

    [Parameter()]
    [string]$OutputDir = "dist"
)

$ErrorActionPreference = "Stop"

$projectPath = "src/HubSchedule.OutlookAddIn/HubSchedule.OutlookAddIn.csproj"
$dllRelativePath = "src/HubSchedule.OutlookAddIn/bin/$Configuration/net48/HubSchedule.OutlookAddIn.dll"

Write-Host "Building add-in ($Configuration)..."
dotnet build $projectPath -c $Configuration
if ($LASTEXITCODE -ne 0) {
    throw "dotnet build failed with exit code $LASTEXITCODE"
}

if (-not (Test-Path $dllRelativePath)) {
    throw "Build finished but DLL not found at '$dllRelativePath'."
}

$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$packageRoot = Join-Path $OutputDir "HubSchedule-OutlookAddIn-$timestamp"
$packageScripts = Join-Path $packageRoot "scripts"

New-Item -ItemType Directory -Path $packageScripts -Force | Out-Null

Copy-Item $dllRelativePath -Destination $packageRoot -Force
Copy-Item "scripts/install-addin.ps1" -Destination $packageScripts -Force
Copy-Item "scripts/uninstall-addin.ps1" -Destination $packageScripts -Force

$readmePath = Join-Path $packageRoot "README-INSTALL.txt"
@"
HubSchedule Outlook Add-in package

1. Run PowerShell as your normal user (or elevated if required by your environment).
2. From this package directory, run:

   powershell -ExecutionPolicy Bypass -File .\\scripts\\install-addin.ps1 -DllPath .\\HubSchedule.OutlookAddIn.dll

3. Restart Outlook.

To uninstall:

   powershell -ExecutionPolicy Bypass -File .\\scripts\\uninstall-addin.ps1 -DllPath .\\HubSchedule.OutlookAddIn.dll
"@ | Set-Content -Path $readmePath -Encoding UTF8

$zipPath = "$packageRoot.zip"
if (Test-Path $zipPath) {
    Remove-Item $zipPath -Force
}

Compress-Archive -Path "$packageRoot/*" -DestinationPath $zipPath
Write-Host "Created package: $zipPath"
