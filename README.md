# ReplaceMagic SharePoint URL Matcher

Proof of Concept web application for matching ReplaceMagic export rows to files in a SharePoint document library by exact filename.

This PoC intentionally uses exact filename matching only. It does not do fuzzy matching, partial matching, content matching, or path inference.

## Purpose

ReplaceMagic exports old paths and links from Office files. This tool helps map those old values to current SharePoint file URLs so a migration or replacement report can be reviewed and exported.

The generated CSV contains:

- `OldPath`
- `OldURL`
- `SearchFileName`
- `MatchedFileName`
- `NewURL`
- `Status`
- `Remark`

## How The Process Works

1. Open the web UI.
2. Enter Microsoft Graph and SharePoint connection settings, or provide them in `.env`.
3. Upload a ReplaceMagic `.csv` or `.xlsx` export.
4. The backend reads each row with `pandas`.
5. The search filename is built with this order:
   - Use `Filename + "." + Extension` when `Extension` exists.
   - Use `Filename` directly when it already contains an extension.
   - Extract a filename from `LinkTitle` or `Hyperlink` when `Filename` is empty.
6. The backend authenticates to Microsoft Graph with MSAL client credentials.
7. The backend resolves the SharePoint site from host and site path.
8. The backend searches the default site drive.
9. A row is considered matched only when the SharePoint file name is exactly equal to the extracted filename. Comparison is case-insensitive.
10. Results can be exported as CSV.

## Status Values

- `FOUND`: one exact filename match was found.
- `NOT_FOUND`: no exact filename match was found.
- `MULTIPLE_MATCH`: more than one exact filename match was found.
- `ERROR`: row parsing, configuration, authentication, or API error occurred.

## Required Microsoft Graph Permissions

Create an application permission grant for:

- `Sites.Read.All`
- `Files.Read.All`

Admin consent is required for application permissions.

## Create App Registration

1. Go to Microsoft Entra admin center.
2. Open **Applications** > **App registrations**.
3. Select **New registration**.
4. Enter a name such as `ReplaceMagic SharePoint Matcher`.
5. Choose the supported account type for your tenant.
6. Register the app.
7. Copy the **Application (client) ID**.
8. Copy the **Directory (tenant) ID**.
9. Open **Certificates & secrets**.
10. Create a new client secret and copy the value immediately.
11. Open **API permissions**.
12. Add Microsoft Graph application permissions:
    - `Sites.Read.All`
    - `Files.Read.All`
13. Select **Grant admin consent**.

## Configure `.env`

Copy `.env.example` to `.env`:

```powershell
Copy-Item .env.example .env
```

Edit `.env`:

```env
TENANT_ID=00000000-0000-0000-0000-000000000000
CLIENT_ID=00000000-0000-0000-0000-000000000000
CLIENT_SECRET=your-secret-value
SHAREPOINT_HOST=yourtenant.sharepoint.com
SHAREPOINT_SITE_PATH=/sites/TestReplaceMagic
```

You can also enter these values in the UI. UI-entered values are used only for the current running app session and are not saved.

Client secrets are not hardcoded and are not printed by the application.

## Run Locally

Create and activate a virtual environment:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

Install dependencies:

```powershell
pip install -r requirements.txt
```

Start FastAPI:

```powershell
uvicorn app.main:app --reload
```

Open:

```text
http://127.0.0.1:8000
```

## Test With A SharePoint Test Site

1. Create a test SharePoint site, for example `/sites/TestReplaceMagic`.
2. Upload several known files to the default document library.
3. Create or use a ReplaceMagic export with rows that contain matching file names.
4. Run the app locally.
5. Configure:
   - SharePoint host: `yourtenant.sharepoint.com`
   - SharePoint site path: `/sites/TestReplaceMagic`
6. Upload the ReplaceMagic export.
7. Select **Preview File**.
8. Select **Run Matching**.
9. Review the result table and export CSV.

## Move From Test To Production Later

When testing is complete, change only the SharePoint target:

```env
SHAREPOINT_HOST=yourtenant.sharepoint.com
SHAREPOINT_SITE_PATH=/sites/ProductionSite
```

Or enter the production host and site path in the UI for a single run.

Before production use:

- Confirm Graph app permissions are approved for the production tenant.
- Test with a small sample export.
- Review `MULTIPLE_MATCH` rows manually.
- Keep client secrets in `.env` or a proper secret store, not in source control.

## Project Structure

```text
replace-magic-sp-matcher/
  app/
    main.py
    graph_client.py
    parser.py
    matcher.py
    exporter.py
    templates/
      index.html
    static/
      app.js
  uploads/
  outputs/
  .env.example
  requirements.txt
  README.md
```

## Notes And Limitations

- The app stores the latest uploaded rows and results in memory. Restarting the app clears them.
- The app writes uploaded files to `uploads/` and the latest export to `outputs/latest-results.csv`.
- This PoC searches the default site document library drive.
- Exact filename matching can produce `MULTIPLE_MATCH` when duplicate filenames exist in different folders.
