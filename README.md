# HubSchedule Outlook Add-in (Local)

A local Outlook task pane add-in that shows recent submissions for a specific HubSpot form in a clean side-pane UI.

## What this project includes

- `manifest.xml`: Outlook add-in manifest for sideloading.
- `server.js`: Local Express server that:
  - Serves the task pane UI.
  - Proxies HubSpot submissions API requests.
- `public/`: Task pane frontend (HTML/CSS/JS).

## Prerequisites

- Node.js 18+
- Outlook (desktop/web) with sideloading enabled.
- HubSpot Private App token with form read access.

## Setup

1. Install dependencies:

   ```bash
   npm install
   ```

2. Create `.env` from `.env.example` and fill in values:

   ```bash
   cp .env.example .env
   ```

3. Start the local add-in server:

   ```bash
   npm run dev
   ```

4. Trust local cert for HTTPS (if needed in your environment):

   ```bash
   npm run cert
   ```

5. Update `manifest.xml` URLs if your local host/port differ.

## Environment variables

- `PORT` (default: `3000`)
- `HUBSPOT_PRIVATE_APP_TOKEN` (required)
- `HUBSPOT_FORM_ID` (required)
- `HUBSPOT_PORTAL_ID` (optional, used in display only)
- `SUBMISSION_PAGE_SIZE` (default: `25`)

## Sideload in Outlook

- **Outlook on the web**: Settings → Manage add-ins → Add from file → select `manifest.xml`.
- **Outlook desktop**: My Add-ins → Add a custom add-in → Add from file.

Then open an email and launch **HubSchedule Submissions** from the ribbon.

## Notes

- This project fetches submissions server-side to keep your HubSpot token out of the browser.
- The task pane auto-refreshes every 60 seconds and includes a manual refresh button.


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
