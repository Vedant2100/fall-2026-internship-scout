# Full-Time AI Job Scout

Google Apps Script + Google Sheets job scout for Vedant's full-time early-career search.

It searches for **full-time new-grad and early-career roles** in:

- AI/ML research
- Research engineering
- Machine learning engineering
- Software engineering
- Data science

It emphasizes San Francisco / Pleasanton / Bay Area / Remote US roles, uses Ashby/Greenhouse/Lever-style job pages plus early-career aggregators, excludes Workday, dedupes results, writes a Google Sheets dashboard, and emails only newly found opportunities.

Internships, co-ops, and part-time roles are no longer accepted by the active search pipeline. Existing internship rows remain in the Sheet as history.

## Why This Exists

The original Dreamer repo was successful as a product prototype, but Dreamer CLI/backend package access is currently unreliable. This version keeps the core product behavior and moves it to a free/low-cost Google-native stack:

- Google Sheets = dashboard + lightweight database
- Apps Script = scheduled runner + manual controls
- Gmail/MailApp = email delivery
- Tavily = search + extraction
- Gemini = structured classification

## Feature Match

Replicated from the Dreamer agent:

- Scheduled search
- Manual search
- Deduplication
- Email only new results
- Dashboard of found opportunities
- Applied/starred tracking
- Search run history
- URL/source filtering
- Visa sponsorship and IITB alumni fields
- Prompt-driven relevance scoring

Different from Dreamer:

- Dashboard is Google Sheets, not React.
- Progress is visible through the `Runs` tab and Apps Script executions, not live UI callbacks.
- Sidekick task is replaced with deterministic search/extract/classify steps.

## Files

- `Code.js` - complete Apps Script implementation.
- `SetupHelp.html` - modal help shown from the Sheet menu.
- `WebDashboard.html` - optional web app view.
- `appsscript.json` - Apps Script manifest and scopes.
- `.clasp.json.example` - template for local `clasp` deployment.

## Required Secrets

Store these in Apps Script **Project Settings -> Script Properties**:

| Key | Purpose |
| --- | --- |
| `TAVILY_API_KEY` | Tavily search/extract API |
| `GEMINI_API_KEY` | Gemini API classification |
| `RECIPIENT_EMAIL` | Email address to send new results to |
| `SHEET_ID` | Optional. Only needed if the script is standalone instead of bound to a Sheet. |

Do not put API keys in Sheet cells.

## Sheet Tabs

The setup creates:

- `Opportunities`
- `Runs`
- `Config`
- `Seen`
- `RawCandidates`

`Opportunities` is the main dashboard. It includes checkboxes for starred/applied state and fields for score, source, visa, IITB alumni, emailed time, and notes.

## Search Inspiration

The core search pattern is based on:

```text
"San Francisco" site:jobs.ashbyhq.com OR site:greenhouse.io OR site:jobs.lever.co "intern" ("data science" OR "AI" OR "software" OR "Research")
```

The implementation expands this into several variants and explicitly excludes Workday:

```text
-workday
```

It also filters Workday-like URLs in code.

## Setup Without Local Tools

This is the fastest path.

1. Create or reuse a Google Sheet named `Full-Time AI Job Scout`.
2. Open `Extensions -> Apps Script`.
3. Copy these repo files into the Apps Script editor:
   - `Code.js`
   - `SetupHelp.html`
   - `WebDashboard.html`
   - `appsscript.json` contents into Project Settings manifest view.
4. In Apps Script, open `Project Settings -> Script Properties`.
5. Add:
   - `TAVILY_API_KEY`
   - `GEMINI_API_KEY`
   - `RECIPIENT_EMAIL`
6. Run `setupDashboard`.
7. Approve scopes.
8. Reload the Google Sheet.
9. Use `Job Scout -> Send Test Email`.
10. Use `Job Scout -> Run Search Now`.
11. Use `Job Scout -> Install Daily Trigger`.

## Setup With clasp

Install clasp:

```bash
npm install -g @google/clasp
clasp login
```

Create or open an Apps Script project attached to your Google Sheet, then copy `.clasp.json.example` to `.clasp.json` and fill in `scriptId`.

Push:

```bash
clasp push
```

Then set Script Properties in the Apps Script UI and run `setupDashboard`.

## Operational Notes

- Apps Script has a 6 minute execution limit per run, so `maxUrlsPerRun` defaults to 45.
- Tavily free tier is 1,000 credits/month. The default daily run should fit if you keep extraction capped.
- Gemini free tier can classify results, but Google states free-tier content may be used to improve products.
- The job is intentionally deterministic and bounded. It is less autonomous than Dreamer Sidekick, but easier to own.

## Manual Controls

After setup, the Google Sheet menu includes:

- `Run Search Now`
- `Setup / Repair Dashboard`
- `Install Daily Trigger`
- `Remove Triggers`
- `Send Test Email`
- `Open Setup Help`

## Tuning

Open the `Config` tab to tune:

- `searchScheduleHour`
- `searchDays` (default: 30, so evergreen job postings are not missed by a one-week freshness filter)
- `maxUrlsPerRun`
- `maxNewEmailItems`
- `minScore`
- `excludedDomains`
- `locationTerms`
- `roleTerms`
- `termTerms`

Start with `maxUrlsPerRun = 20` for the first real run. Increase after confirming quota and runtime.

The classifier uses the candidate's expected December 2026 graduation, F-1/OPT context, CompFly agent-governance/observability experience, prior data-science experience, and research/LLM background. It treats sponsorship as unknown unless the job posting explicitly provides evidence.
