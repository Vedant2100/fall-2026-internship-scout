# AGENTS.md

## Cursor Cloud specific instructions

This repository is a **Google Apps Script (GAS) project**, not a locally-runnable
service. The "application" (`Code.js`) runs inside Google's cloud, bound to a
Google Sheet, and calls Apps Script services (`SpreadsheetApp`, `MailApp`,
`UrlFetchApp`, `PropertiesService`, `Utilities`, etc.) that do not exist in a
plain Node.js runtime. Node is only used for the `@google/clasp` dev tooling and
JSON linting.

### Lint / test / build
- Lint (JSON manifest validation): `npm run check` (alias `npm run lint:json`). This is the only automated check defined in `package.json`.
- Syntax-check the app code: `node --check Code.js`.
- There is no build step and no automated unit-test suite.

### Running the app
- You cannot fully run this app locally. Real execution requires deploying to
  Apps Script via `clasp` (`npx clasp login` → OAuth, then copy
  `.clasp.json.example` to `.clasp.json` with a real `scriptId`, then
  `npx clasp push`) plus Script Properties `TAVILY_API_KEY`, `GEMINI_API_KEY`,
  and `RECIPIENT_EMAIL`. See `README.md` / `SETUP_CHECKLIST.md` for the full flow.
- `npx clasp login` needs interactive Google OAuth and is a **user action** —
  it cannot be completed unattended in the VM.
- To exercise the app's **core deterministic logic** without Google (query
  building, URL normalization, Workday exclusion, source labeling, dedup
  fingerprinting, title gating), load `Code.js` into a Node `vm` context with the
  GAS globals shimmed. Only `Logger` and `Utilities` (SHA-256 digest via Node
  `crypto` + web-safe base64) are needed for those functions. This is the
  practical way to smoke-test changes to the pure helpers in the VM.

### Notes
- Secrets belong in Apps Script Script Properties, never in Sheet cells or the repo.
- `.clasp.json` is gitignored; only `.clasp.json.example` is tracked.
