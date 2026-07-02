# Setup Checklist

Use this when you are ready to make the scout live.

## Accounts / Keys

Create or gather:

- Tavily API key
- Gemini API key from Google AI Studio
- Recipient email address
- Google account that owns the Sheet and sends email

## Google Sheet

1. Create a Google Sheet named `Fall 2026 AI Internship Scout`.
2. Open `Extensions -> Apps Script`.
3. Add these files:
   - `Code.js`
   - `SetupHelp.html`
   - `WebDashboard.html`
4. In Apps Script settings, enable/edit manifest and paste `appsscript.json`.

## Script Properties

In `Project Settings -> Script Properties`, add:

```text
TAVILY_API_KEY=...
GEMINI_API_KEY=...
RECIPIENT_EMAIL=you@example.com
```

Only add `SHEET_ID` if the Apps Script project is standalone rather than bound to the Sheet.

## First Run

Run these functions manually from the Apps Script editor:

1. `setupDashboard`
2. Authorize scopes.
3. `sendTestEmail`
4. `runSearchNow`

Then reload the Google Sheet and use the custom `Internship Scout` menu.

## Schedule

After the first manual run works:

1. Open the Sheet.
2. Click `Internship Scout -> Install Daily Trigger`.
3. Confirm future runs appear in the `Runs` tab.

## First-Run Tuning

Start conservative:

```text
maxUrlsPerRun=20
minScore=65
includeRawCandidates=true
```

If runtime and quota look healthy, increase:

```text
maxUrlsPerRun=45
```

## Expected Outputs

- `Opportunities` contains deduped jobs.
- `Runs` contains each search attempt and status.
- `RawCandidates` contains search candidates and snippets.
- New opportunities are emailed once.
- Repeated opportunities update `last_seen_at` but are not re-emailed.

## Troubleshooting

- Missing key error: check Apps Script Project Settings -> Script Properties.
- No email: run `sendTestEmail`, then check Gmail sent mail and spam.
- Timeout: lower `maxUrlsPerRun`.
- Bad JSON from Gemini: rerun; if repeated, lower `maxUrlsPerRun` and inspect Apps Script logs.
- Too many irrelevant jobs: raise `minScore` to `75`.
- Missing good jobs: lower `minScore` or add more queries in `buildSearchQueries_`.

