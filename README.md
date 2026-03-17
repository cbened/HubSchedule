# HubSchedule Classic Outlook Add-in (COM)

This repository is now structured as a **classic Outlook desktop COM add-in** for teams using classic Outlook instead of Office.js web add-ins.

## What changed

- Removed dependency on a sideloaded Office web-add-in manifest and localhost web server.
- Added a .NET Framework COM add-in project that adds an Outlook ribbon toggle.
- The add-in opens a local desktop window that embeds HubSpot submissions using WebView2.

## Project layout

- `HubSchedule.OutlookAddIn.sln` - Visual Studio solution.
- `src/HubSchedule.OutlookAddIn/` - COM add-in source.
  - `Connect.cs` - `IDTExtensibility2` + Ribbon callbacks.
  - `RibbonXml.cs` - Ribbon definition for mail read + appointment compose.
  - `HubScheduleWindow.cs` - Embedded WebView2 UI.

## Prerequisites

- Windows with classic Outlook installed.
- Visual Studio 2022 (or compatible MSBuild for .NET Framework 4.8).
- .NET Framework 4.8 targeting pack.
- Microsoft Edge WebView2 Runtime.

## Build

1. Open `HubSchedule.OutlookAddIn.sln` in Visual Studio.
2. Restore NuGet packages.
3. Build `Release | Any CPU`.

Or from CLI:

```powershell
dotnet build .\src\HubSchedule.OutlookAddIn\HubSchedule.OutlookAddIn.csproj -c Release
```

## Easy install / uninstall scripts

This repo now includes PowerShell helpers so you do not need to manually edit registry keys.

Install for current user:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\install-addin.ps1
```

Uninstall for current user:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\uninstall-addin.ps1
```

Optional parameters:

- `-DllPath <path>` when your DLL is in a non-default location.
- `-RegAsmPath <path>` when `RegAsm.exe` is not auto-detected.

## Create a distributable installer package

To produce a zip package containing the DLL and install/uninstall scripts:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\create-installer-package.ps1
```

The script outputs a timestamped zip file under `dist/` that can be shared with users.

## Notes

- This is desktop Outlook specific and does not run in Outlook on the web/new Outlook.
- If WebView2 fails to initialize, the window shows an installation hint.
