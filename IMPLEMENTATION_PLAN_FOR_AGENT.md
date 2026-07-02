# Implementation Plan For Follow-Up Agent

## Objective

Build and deploy a serious clone of the original Dreamer internship scout using Google Apps Script and Google Sheets.

The new scout must search for **Fall 2026 part-time internships** in AI/ML research, SWE, MLE, and data science. It should take inspiration from this search:

```text
"San Francisco" site:jobs.ashbyhq.com OR site:greenhouse.io OR site:jobs.lever.co "intern" ("data science" OR "AI" OR "software" OR "Research")
```

It must exclude Workday.

## Chosen Architecture

- Google Sheets: dashboard and persistent state
- Apps Script: scheduled runner, manual controls, API calls, email
- Tavily: search and page extraction
- Gemini: structured classification
- MailApp: email results to user

## Deliverables In This Repo

- `Code.js`: full Apps Script implementation
- `SetupHelp.html`: setup modal
- `WebDashboard.html`: optional web view
- `appsscript.json`: manifest/scopes
- `README.md`: user setup instructions

## Functional Requirements

1. Dashboard
   - Store opportunities in a Sheet.
   - Include status, starred, applied, company, role, track, location, term, part-time flag, URL, source, details, visa, IITB alumni, score, timestamps, notes.
   - Add checkbox columns for starred/applied.

2. Scheduled Search
   - Time-based trigger runs daily.
   - Manual search available from Google Sheet menu.
   - Runs must not overlap.

3. Search
   - Use Ashby, Greenhouse, and Lever-focused queries.
   - Search for Fall 2026, part-time, AI/ML research, SWE, MLE, data science internships.
   - Exclude Workday by query and URL filtering.

4. Classification
   - Use Gemini to classify extracted pages into strict JSON.
   - Preserve source URL.
   - Score relevance.
   - Drop low-score/irrelevant rows.

5. Deduplication
   - Normalize URLs.
   - Hash URL + company-role.
   - Do not re-email already-seen opportunities.

6. Email
   - Email only new opportunities.
   - Include top new results and Google Sheet link.
   - Mark emailed rows.

7. Ops
   - Log every run in `Runs`.
   - Capture raw candidates in `RawCandidates`.
   - Write errors to `Runs`.

## Setup Steps For User

1. Create Google Sheet.
2. Open Apps Script.
3. Add files from repo.
4. Add Script Properties:
   - `TAVILY_API_KEY`
   - `GEMINI_API_KEY`
   - `RECIPIENT_EMAIL`
5. Run `setupDashboard`.
6. Authorize.
7. Run `sendTestEmail`.
8. Run `runSearchNow`.
9. Install daily trigger.

## Follow-Up Work

- If Apps Script runtime exceeds 6 minutes, lower `maxUrlsPerRun`.
- If Gemini output occasionally fails JSON parse, add retry with stricter prompt.
- If Tavily misses too much, add specific company career pages and YC company lists as seed URLs.
- If user wants a prettier UI, publish `doGet` as a web app or build a small frontend later.

