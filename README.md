# HubSchedule Outlook Add-in (Local)

A local Outlook task pane add-in that embeds the HubSpot submissions page in a side panel.

## What this project includes

- `manifest.xml`: Outlook add-in manifest for sideloading.
- `server.js`: Local Express server that serves the static add-in files.
- `public/`: Task pane frontend with an embedded HubSpot submissions iframe.

## Prerequisites

- Node.js 18+
- Outlook (desktop/web) with sideloading enabled.

## Setup

1. Install dependencies:

   ```bash
   npm install
   ```

2. Start the local add-in server:

   ```bash
   npm run dev
   ```

3. Trust local cert for HTTPS (if needed in your environment):

   ```bash
   npm run cert
   ```

4. Update `manifest.xml` URLs if your local host/port differ.

## Sideload in Outlook

- **Outlook on the web**: Settings → Manage add-ins → Add from file → select `manifest.xml`.
- **Outlook desktop**: My Add-ins → Add a custom add-in → Add from file.

The add-in command is available when:
- reading an email message
- creating/editing a calendar event as organizer (scheduling pane)

## Notes

- The task pane now loads the HubSpot submissions URL directly in an embedded iframe.
- HubSpot authentication is handled by the HubSpot web app session inside the embedded page.

## Binary-free icon handling

To keep this repo PR-safe in environments that reject binary files, icon assets are stored as Base64 text files (`public/icon-*.png.txt`). Before sideloading in Outlook, convert them back to PNG files and remove the `.txt` suffix.

Example conversion:

```bash
base64 -d public/icon-16.png.txt > public/icon-16.png
base64 -d public/icon-32.png.txt > public/icon-32.png
base64 -d public/icon-64.png.txt > public/icon-64.png
base64 -d public/icon-80.png.txt > public/icon-80.png
base64 -d public/icon-128.png.txt > public/icon-128.png
```
