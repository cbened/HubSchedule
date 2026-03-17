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

## Register COM add-in (developer machine)

From a **Developer Command Prompt for Visual Studio**:

```powershell
regasm .\src\HubSchedule.OutlookAddIn\bin\Release\HubSchedule.OutlookAddIn.dll /codebase /tlb
```

Then create these registry keys:

```text
HKCU\Software\Microsoft\Office\Outlook\Addins\HubSchedule.OutlookAddIn
  (Default) = "HubSchedule Outlook Add-in"
  FriendlyName = "HubSchedule Outlook Add-in"
  Description = "Open HubSpot submissions in classic Outlook"
  LoadBehavior (DWORD) = 3
  CommandLineSafe (DWORD) = 0
```

Restart Outlook and enable the add-in if prompted.

## Notes

- This is desktop Outlook specific and does not run in Outlook on the web/new Outlook.
- If WebView2 fails to initialize, the window shows an installation hint.
