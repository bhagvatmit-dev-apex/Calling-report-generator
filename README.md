# MitsuBridge Calling Report

A static sales calling dashboard for Google Sheets and uploaded Excel/CSV files.

## Deploy with Vercel

1. Import `bhagvatmit-dev-apex/Calling-report-generator` from GitHub in Vercel.
2. Keep the project root directory as `.` and select the **Other** framework preset.
3. Leave the build command and output directory empty; this project is served as static files.
4. Deploy. The root `index.html` forwards visitors to the dashboard page.

Every push to the connected branch will trigger a new deployment.

## Google Sheets refresh

In the dashboard, choose **Upload sheets → Add Google Sheet URL** to configure one or more spreadsheet URLs. Those URLs and the optional connector URL are saved in that browser only; Vercel does not receive them as build-time settings.

If the browser cannot download complete workbooks, deploy `MitsuBridge Sheets Bridge.gs` as a Google Apps Script web app:

1. Open [script.google.com](https://script.google.com/) and create a project.
2. Paste the bridge code into the script editor and save.
3. Add any spreadsheet IDs not already listed in `ALLOWED_SPREADSHEET_IDS`.
4. Deploy as a web app that executes as an account with access to the sheets. Set access to anyone with the link if that is permitted by your organization.
5. Paste the deployed `/exec` URL into the dashboard's optional **Apps Script connector URL** field and save.

The bridge only serves spreadsheet IDs in its allowlist. Anyone with the deployed connector URL can read the allowlisted workbook data, so share that URL accordingly. Do not commit private credentials, service-account keys, or connector secrets.

For exact future disposition-change timestamps, separately install `Disposition History Logger.gs` in each source spreadsheet and run `setupDispositionHistory()` once with editor authorization. It does not reconstruct edits from before installation.
