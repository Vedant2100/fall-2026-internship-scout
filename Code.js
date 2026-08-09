/**
 * Full-Time AI Job Scout
 *
 * Google Apps Script + Google Sheets implementation inspired by the original
 * Dreamer AI Startup Internship Scout, but without Dreamer runtime dependencies.
 */

var APP = {
  name: "Full-Time AI Job Scout",
  version: "3.0.0",
  sheets: {
    opportunities: "Opportunities",
    runs: "Runs",
    config: "Config",
    seen: "Seen",
    raw: "RawCandidates"
  },
  props: {
    tavilyKey: "TAVILY_API_KEY",
    geminiKey: "GEMINI_API_KEY",
    recipientEmail: "RECIPIENT_EMAIL",
    sheetId: "SHEET_ID"
  },
  defaults: {
    searchScheduleHour: "8",
    maxUrlsPerRun: "80",
    maxNewEmailItems: "20",
    minScore: "80",
    searchDepth: "basic",
    extractDepth: "basic",
    searchDays: "30",
    geminiModel: "gemini-3.1-flash-lite",
    includeRawCandidates: "true",
    excludedDomains: "workday, myworkdayjobs, wd1.myworkdaysite, wd5.myworkdayjobs, simplify",
    locationTerms: "San Francisco, Pleasanton, Bay Area, Silicon Valley, Remote US, United States",
    roleTerms: "research engineer, research scientist, AI research, machine learning engineer, ML engineer, ML infrastructure, applied scientist, software engineer, SWE, data science, agent engineer, AI agent, agentic, agent evaluation, agent observability, memory governance, LLM reasoning, post-training, reinforcement learning, RL, RLHF, DPO, GRPO, reward model, autonomous agent, tool use, OpenTelemetry",
    termTerms: "full-time, new grad, new graduate, entry level, early career, junior, associate, 0-3 years",
    emailSubjectPrefix: "Full-Time AI Job Scout",
    classifyBatchSize: "7",
    enableInternList: "false",
    enableSimplifyJobs: "false",
    enableNewGradJobs: "true",
    enableNewGradSimplify: "true",
    newGradSearchQueries: "true",
    enableWellfound: "true",
    enableYCombinator: "true",
    enableJobright: "true",
    digestIntervalDays: "2",
    digestMaxItems: "25"
  },
  opportunityHeaders: [
    "id",
    "type",
    "status",
    "starred",
    "applied",
    "company",
    "role",
    "track",
    "location",
    "term",
    "part_time",
    "url",
    "source",
    "details",
    "visa_sponsorship",
    "iitb_alumni",
    "score",
    "found_at",
    "emailed_at",
    "last_seen_at",
    "notes"
  ],
  runHeaders: [
    "run_id",
    "started_at",
    "completed_at",
    "status",
    "queries_run",
    "urls_checked",
    "new_count",
    "emailed",
    "error"
  ],
  configHeaders: ["key", "value", "description"],
  seenHeaders: ["fingerprint", "url", "company_role", "first_seen_at", "last_seen_at"],
  rawHeaders: ["run_id", "query", "title", "url", "source", "snippet", "score"]
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Job Scout")
    .addItem("Run Search Now", "runSearchNow")
    .addItem("Send Digest Now", "sendDigestNow")
    .addItem("Setup / Repair Dashboard", "setupDashboard")
    .addItem("Clear Raw Candidates Only", "clearRawCandidatesOnly")
    .addItem("Clear Existing Data", "clearScoutData")
    .addSeparator()
    .addItem("Install Daily Trigger", "installDailyTrigger")
    .addItem("Install Digest Trigger (every 2 days)", "installDigestTrigger")
    .addItem("Remove Triggers", "removeScoutTriggers")
    .addSeparator()
    .addItem("Send Test Email", "sendTestEmail")
    .addItem("Open Setup Help", "showSetupHelp")
    .addToUi();
}

function clearScoutData() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    "Clear All Scout Data?",
    "This will delete all existing rows from Opportunities, Runs, Seen, and RawCandidates so your scout starts with a clean slate.\n\nYour Config sheet and Script Properties will NOT be affected.\n\nAre you sure you want to proceed?",
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  var ss = getSpreadsheet_();
  var sheetsToClear = [
    APP.sheets.opportunities,
    APP.sheets.runs,
    APP.sheets.seen,
    APP.sheets.raw
  ];

  sheetsToClear.forEach(function (sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      var lastRow = sheet.getLastRow();
      var lastCol = sheet.getLastColumn();
      if (lastRow > 1 && lastCol > 0) {
        sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
      }
    }
  });

  ui.alert("All existing data cleared! You are ready for a fresh search run.");
}

function setupDashboard() {
  var ss = getSpreadsheet_();
  ensureWorkbook_(ss);
  SpreadsheetApp.getUi().alert(
    APP.name + " dashboard is ready.\n\nAdd Script Properties for TAVILY_API_KEY, GEMINI_API_KEY, and RECIPIENT_EMAIL.\n\n(Optional) Add SERPER_API_KEY to use Google Search via Serper.dev, or GSEARCH_API_KEY + GSEARCH_CX to use Google Custom Search as your primary search engine!"
  );
}

function runSearchNow() {
  return runSearch_({ manual: true });
}

function runScheduledSearch() {
  return runSearch_({ manual: false });
}

function installDailyTrigger() {
  var ss = getSpreadsheet_();
  ensureWorkbook_(ss);
  var config = readConfig_(ss);
  removeScoutTriggers();
  ScriptApp.newTrigger("runScheduledSearch")
    .timeBased()
    .everyDays(1)
    .atHour(Number(config.searchScheduleHour || APP.defaults.searchScheduleHour))
    .create();
  SpreadsheetApp.getUi().alert("Daily trigger installed for runScheduledSearch().");
}

function removeScoutTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    var handler = trigger.getHandlerFunction();
    if (handler === "runScheduledSearch" || handler === "runScheduledDigest") {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function installDigestTrigger() {
  var ss = getSpreadsheet_();
  ensureWorkbook_(ss);
  var config = readConfig_(ss);
  var intervalDays = Number(config.digestIntervalDays || APP.defaults.digestIntervalDays);
  // Remove existing digest triggers
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (trigger.getHandlerFunction() === "runScheduledDigest") {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  ScriptApp.newTrigger("runScheduledDigest")
    .timeBased()
    .everyDays(intervalDays)
    .atHour(9)
    .create();
  SpreadsheetApp.getUi().alert("Digest trigger installed! You will receive a top-opportunities recap email every " + intervalDays + " days at 9 AM.");
}

function runScheduledDigest() {
  sendDigestEmail_();
}

function sendDigestNow() {
  sendDigestEmail_();
  SpreadsheetApp.getUi().alert("Digest email sent!");
}

function sendDigestEmail_() {
  var ss = getSpreadsheet_();
  var config = readConfig_(ss);
  var recipient = getRequiredProperty_(APP.props.recipientEmail);
  var minScore = Number(config.minScore || APP.defaults.minScore);
  var maxItems = Number(config.digestMaxItems || APP.defaults.digestMaxItems);
  var sheetUrl = ss.getUrl();

  var oppSheet = ss.getSheetByName(APP.sheets.opportunities);
  if (!oppSheet || oppSheet.getLastRow() <= 1) {
    Logger.log("Digest: No opportunities to send.");
    return;
  }

  var values = oppSheet.getDataRange().getValues();
  var headers = values[0];
  var rows = values.slice(1);

  // Build objects and filter by score >= minScore, status != "Rejected"
  var opps = [];
  rows.forEach(function (row) {
    var obj = {};
    headers.forEach(function (h, i) { obj[h] = row[i]; });
    var score = Number(obj.score || 0);
    var status = String(obj.status || "").toLowerCase();
    if (score >= minScore && status !== "rejected") {
      opps.push(obj);
    }
  });

  // Sort by score descending, then by found_at descending
  opps.sort(function (a, b) {
    var ds = Number(b.score || 0) - Number(a.score || 0);
    if (ds !== 0) return ds;
    var da = new Date(b.found_at || 0);
    var db = new Date(a.found_at || 0);
    return da - db;
  });

  opps = opps.slice(0, maxItems);

  if (opps.length === 0) {
    Logger.log("Digest: No opportunities above minScore " + minScore + ".");
    return;
  }

  // Separate starred, not-yet-applied, and the rest
  var starred = opps.filter(function (o) { return o.starred === true; });
  var unapplied = opps.filter(function (o) { return o.starred !== true && o.applied !== true; });
  var applied = opps.filter(function (o) { return o.applied === true; });

  var subject = config.emailSubjectPrefix + " Digest: Your top " + opps.length + " opportunities";

  var html = [
    "<p>Here is your periodic digest of <strong>" + opps.length + "</strong> top-scored opportunities (score ≥ " + minScore + ").</p>",
    "<p><a href=\"" + escapeHtml_(sheetUrl) + "\">Open dashboard in Google Sheets</a></p>"
  ];
  var text = [
    "Here is your periodic digest of " + opps.length + " top-scored opportunities (score >= " + minScore + ").",
    "Open dashboard: " + sheetUrl,
    ""
  ];

  // Render a section
  function renderSection(title, emoji, list) {
    if (list.length === 0) return;
    html.push("<h3>" + emoji + " " + title + " (" + list.length + ")</h3>");
    html.push("<ol>");
    text.push("=== " + title.toUpperCase() + " (" + list.length + ") ===");
    list.forEach(function (item) {
      var typeTag = item.type === "new_grad" ? "[FULL-TIME] " : "[LEGACY] ";
      var appliedTag = item.applied === true ? " ✅ Applied" : "";
      var starTag = item.starred === true ? " ⭐" : "";
      html.push(
        "<li><strong>" + typeTag + escapeHtml_(String(item.company || "")) + " - " + escapeHtml_(String(item.role || "")) + starTag + appliedTag + "</strong><br>" +
        escapeHtml_(String(item.location || "")) + " | " + escapeHtml_(String(item.track || "")) + " | Score: " + escapeHtml_(String(item.score || "")) + "<br>" +
        escapeHtml_(String(item.details || "")) + "<br>" +
        "<a href=\"" + escapeHtml_(String(item.url || "")) + "\">Apply / view posting</a></li>"
      );
      text.push([
        typeTag + (item.company || "") + " - " + (item.role || "") + starTag + appliedTag,
        "Score: " + (item.score || "") + " | " + (item.location || ""),
        "URL: " + (item.url || ""),
        ""
      ].join("\n"));
    });
    html.push("</ol>");
  }

  renderSection("Starred Opportunities", "⭐", starred);
  renderSection("Not Yet Applied", "🎯", unapplied);
  renderSection("Already Applied", "✅", applied);

  html.push("<hr><p><em>This digest is sent every " + (config.digestIntervalDays || APP.defaults.digestIntervalDays) + " days. Adjust frequency in your Config sheet (digestIntervalDays).</em></p>");

  MailApp.sendEmail({
    to: recipient,
    subject: subject,
    htmlBody: html.join("\n"),
    body: text.join("\n")
  });
  Logger.log("Digest email sent to " + recipient + " with " + opps.length + " opportunities.");
}

function sendTestEmail() {
  var recipient = getRequiredProperty_(APP.props.recipientEmail);
  MailApp.sendEmail({
    to: recipient,
    subject: APP.name + " test email",
    htmlBody: "<p>Your Full-Time AI Job Scout email path works.</p>",
    body: "Your Full-Time AI Job Scout email path works."
  });
  SpreadsheetApp.getUi().alert("Test email sent to " + recipient + ".");
}

function showSetupHelp() {
  var html = HtmlService.createHtmlOutputFromFile("SetupHelp")
    .setWidth(700)
    .setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(html, "Full-Time AI Job Scout Setup");
}

function doGet() {
  var ss = getSpreadsheet_();
  ensureWorkbook_(ss);
  var data = getDashboardSummary_(ss);
  var template = HtmlService.createTemplateFromFile("WebDashboard");
  template.summary = data;
  return template.evaluate()
    .setTitle(APP.name)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function runSearch_(opts) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    throw new Error("Another Job Scout run is already active.");
  }

  var ss = getSpreadsheet_();
  ensureWorkbook_(ss);
  var runId = Utilities.getUuid();
  var startedAt = new Date();
  var runRow = appendRun_(ss, {
    run_id: runId,
    started_at: startedAt,
    completed_at: "",
    status: "running",
    queries_run: 0,
    urls_checked: 0,
    new_count: 0,
    emailed: false,
    error: ""
  });

  try {
    Logger.log("Starting run " + runId + "...");
    var config = readConfig_(ss);
    var tavilyKey = getRequiredProperty_(APP.props.tavilyKey);
    var geminiKey = getRequiredProperty_(APP.props.geminiKey);
    var recipient = getRequiredProperty_(APP.props.recipientEmail);
    var queries = buildSearchQueries_(config);
    var candidates = collectCandidates_(tavilyKey, queries, config);
    var limited = limitCandidates_(candidates, Number(config.maxUrlsPerRun || APP.defaults.maxUrlsPerRun));

    Logger.log("Found " + candidates.length + " unique candidates. Limiting to " + limited.length + " for extraction.");

    var previousRawUrls = getPreviousRawUrls_(ss);
    var evaluatedUrls = getEvaluatedUrls_(ss);
    var newRawGradItems = [];
    var newRawCandidates = [];
    var unevaluatedCandidates = [];
    var activeUrlsMap = {};

    candidates.forEach(function (c) {
      var u = normalizeUrl_(c.url || "");
      if (!u) return;
      activeUrlsMap[u] = true;
      if (isExcludedUrl_(u, config) || isDisqualifiedTitle_(c.title || "")) return;
      
      if (!previousRawUrls[u]) {
        newRawCandidates.push(c);
        previousRawUrls[u] = true;
        if (c.query === "newgrad-jobs.com" || c.query === "SimplifyJobs-NewGrad") {
          newRawGradItems.push(c);
        }
      }

      if (!evaluatedUrls[u]) {
        unevaluatedCandidates.push(c);
        evaluatedUrls[u] = true;
      }
    });

    Logger.log("Detected " + newRawCandidates.length + " brand new raw URLs across all sources (" + unevaluatedCandidates.length + " unseen by Gemini).");

    batchTouchOpportunities_(ss, activeUrlsMap, new Date());

    if (String(config.includeRawCandidates || "true") === "true" || newRawCandidates.length > 0) {
      writeRawCandidates_(ss, runId, newRawCandidates);
    }

    var candidatesToExtract = limitCandidates_(unevaluatedCandidates, Number(config.maxUrlsPerRun || APP.defaults.maxUrlsPerRun));
    Logger.log("Extracting text content from " + candidatesToExtract.length + " page URLs (out of " + unevaluatedCandidates.length + " unevaluated)...");
    var extractedPages = extractCandidatePages_(tavilyKey, candidatesToExtract, config);
    
    Logger.log("Classifying " + extractedPages.length + " pages with Gemini...");
    var classified = classifyPages_(geminiKey, extractedPages, config);
    
    var relevant = classified.filter(function (item) {
      // This scout is now intentionally full-time/new-grad only. Keep legacy
      // internship rows in the Sheet for history, but never add new ones.
      var isFullTime = String(item.part_time || "").toLowerCase() === "no" &&
        !/intern|internship|co-op|part[\s-]time|flexible hours?/i.test(
          String(item.details || "") + " " + String(item.reason || "") + " " + String(item.role || "")
        );
      return item.type === "new_grad" &&
        item.is_relevant === true &&
        isFullTime &&
        Number(item.score || 0) >= Number(config.minScore || APP.defaults.minScore);
    });

    Logger.log("Classification complete. Found " + relevant.length + " relevant options (out of " + classified.length + " total opportunities found on pages).");

    var newItems = upsertOpportunities_(ss, relevant);
    var emailed = false;
    if (newItems.length > 0) {
      Logger.log("Sending search results email to " + recipient + " (LLM items: " + newItems.length + ")...");
      sendResultsEmail_(recipient, ss, newItems, config);
      markEmailed_(ss, newItems);
      emailed = true;
    } else {
      Logger.log("No new LLM-validated opportunities to email in this run.");
    }

    updateRun_(ss, runRow, {
      completed_at: new Date(),
      status: "completed",
      queries_run: queries.length,
      urls_checked: candidatesToExtract.length,
      new_count: newItems.length + newRawGradItems.length,
      emailed: emailed,
      error: ""
    });

    return {
      runId: runId,
      queriesRun: queries.length,
      urlsChecked: limited.length,
      newCount: newItems.length,
      emailed: emailed
    };
  } catch (error) {
    updateRun_(ss, runRow, {
      completed_at: new Date(),
      status: "failed",
      error: error && error.stack ? error.stack : String(error)
    });
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function getSpreadsheet_() {
  var explicitId = PropertiesService.getScriptProperties().getProperty(APP.props.sheetId);
  if (explicitId) {
    return SpreadsheetApp.openById(explicitId);
  }
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) {
    throw new Error("No active spreadsheet. Bind this script to a Sheet or set Script Property SHEET_ID.");
  }
  return active;
}

function ensureWorkbook_(ss) {
  ensureSheet_(ss, APP.sheets.opportunities, APP.opportunityHeaders);
  ensureSheet_(ss, APP.sheets.runs, APP.runHeaders);
  ensureSheet_(ss, APP.sheets.config, APP.configHeaders);
  ensureSheet_(ss, APP.sheets.seen, APP.seenHeaders);
  ensureSheet_(ss, APP.sheets.raw, APP.rawHeaders);
  ensureConfigDefaults_(ss);
  formatDashboard_(ss);
}

function ensureSheet_(ss, name, headers) {
  var sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  var firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  var needsHeaders = firstRow.join("") === "" || firstRow.join("|") !== headers.join("|");
  if (needsHeaders) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  sheet.setFrozenRows(1);
  return sheet;
}

function ensureConfigDefaults_(ss) {
  var sheet = ss.getSheetByName(APP.sheets.config);
  var values = sheet.getDataRange().getValues();
  var existing = {};
  var existingRows = {};
  for (var i = 1; i < values.length; i++) {
    if (values[i][0]) {
      var key = String(values[i][0]);
      existing[key] = true;
      existingRows[key] = i + 1;
    }
  }
  var rows = [
    ["searchScheduleHour", APP.defaults.searchScheduleHour, "Hour of day, script timezone, for daily trigger."],
    ["maxUrlsPerRun", APP.defaults.maxUrlsPerRun, "Maximum unique candidate URLs to extract/classify per run."],
    ["maxNewEmailItems", APP.defaults.maxNewEmailItems, "Maximum new items shown in each email."],
    ["minScore", APP.defaults.minScore, "Minimum Gemini relevance score required to save/email."],
    ["searchDepth", APP.defaults.searchDepth, "Tavily search_depth: basic or advanced."],
    ["extractDepth", APP.defaults.extractDepth, "Tavily extract_depth: basic or advanced."],
    ["searchDays", APP.defaults.searchDays, "How many recent days search providers should consider."],
    ["geminiModel", APP.defaults.geminiModel, "Gemini model used for page classification."],
    ["includeRawCandidates", APP.defaults.includeRawCandidates, "Write raw search candidates to RawCandidates sheet."],
    ["excludedDomains", APP.defaults.excludedDomains, "Comma-separated URL substrings to reject."],
    ["locationTerms", APP.defaults.locationTerms, "Guidance for location filtering."],
    ["roleTerms", APP.defaults.roleTerms, "Guidance for role filtering."],
    ["termTerms", APP.defaults.termTerms, "Guidance for term/timing filtering."],
    ["emailSubjectPrefix", APP.defaults.emailSubjectPrefix, "Email subject prefix."],
    ["enableNewGradJobs", APP.defaults.enableNewGradJobs, "Enable scraping newgrad-jobs.com Airtable sources."],
    ["enableNewGradSimplify", APP.defaults.enableNewGradSimplify, "Enable scraping SimplifyJobs New-Grad GitHub."],
    ["newGradSearchQueries", APP.defaults.newGradSearchQueries, "Enable search engine queries for new grad roles."]
  ];
  rows.forEach(function (row) {
    if (!existing[row[0]]) {
      sheet.appendRow(row);
    }
  });

  // Migrate only values that still equal the old shipped defaults. This keeps
  // deliberate user customizations intact while switching existing sheets to
  // the full-time/new-grad pipeline.
  var migrations = {
    termTerms: ["Fall 2026, part-time, intern, internship, co-op", APP.defaults.termTerms],
    emailSubjectPrefix: ["Fall 2026 AI Internship Scout", APP.defaults.emailSubjectPrefix],
    enableInternList: ["true", APP.defaults.enableInternList],
    enableSimplifyJobs: ["true", APP.defaults.enableSimplifyJobs]
  };
  Object.keys(migrations).forEach(function (key) {
    var rowNumber = existingRows[key];
    if (!rowNumber) return;
    var currentValue = String(sheet.getRange(rowNumber, 2).getValue() || "");
    if (currentValue === migrations[key][0]) {
      sheet.getRange(rowNumber, 2).setValue(migrations[key][1]);
    }
  });
}

function formatDashboard_(ss) {
  var opportunities = ss.getSheetByName(APP.sheets.opportunities);
  var statusCol = APP.opportunityHeaders.indexOf("status") + 1;
  var statusLetter = String.fromCharCode(64 + statusCol);
  var starredCol = APP.opportunityHeaders.indexOf("starred") + 1;
  var appliedCol = APP.opportunityHeaders.indexOf("applied") + 1;
  var starredLetter = String.fromCharCode(64 + starredCol);
  var appliedLetter = String.fromCharCode(64 + appliedCol);
  var scoreCol = APP.opportunityHeaders.indexOf("score") + 1;
  var scoreLetter = String.fromCharCode(64 + scoreCol);

  opportunities.getRange(starredLetter + "2:" + appliedLetter).insertCheckboxes();
  opportunities.autoResizeColumns(1, Math.min(APP.opportunityHeaders.length, 13));
  applyFilterIfMissing_(opportunities, APP.opportunityHeaders.length);

  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["New", "Seen", "Starred", "Applied", "Rejected"], true)
    .setAllowInvalid(true)
    .build();
  opportunities.getRange(statusLetter + "2:" + statusLetter).setDataValidation(statusRule);

  var scoreRange = opportunities.getRange(scoreLetter + "2:" + scoreLetter);
  var rules = opportunities.getConditionalFormatRules();
  var highScoreRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(85)
    .setBackground("#d9ead3")
    .setRanges([scoreRange])
    .build();
  opportunities.setConditionalFormatRules(rules.concat([highScoreRule]));

  applyFilterIfMissing_(ss.getSheetByName(APP.sheets.runs), APP.runHeaders.length);
  applyFilterIfMissing_(ss.getSheetByName(APP.sheets.raw), APP.rawHeaders.length);
}

function applyFilterIfMissing_(sheet, width) {
  if (!sheet.getFilter()) {
    sheet.getRange(1, 1, Math.max(sheet.getMaxRows(), 2), width).createFilter();
  }
}

function readConfig_(ss) {
  var sheet = ss.getSheetByName(APP.sheets.config);
  var values = sheet.getDataRange().getValues();
  var config = {};
  for (var i = 1; i < values.length; i++) {
    if (values[i][0]) {
      config[String(values[i][0])] = values[i][1];
    }
  }
  Object.keys(APP.defaults).forEach(function (key) {
    if (config[key] === undefined || config[key] === "") config[key] = APP.defaults[key];
  });
  return config;
}

function buildSearchQueries_(config) {
  var ats = '(site:jobs.ashbyhq.com OR site:greenhouse.io OR site:jobs.lever.co)';
  var locations = '("San Francisco" OR "Pleasanton" OR "Bay Area" OR "Silicon Valley" OR "Remote US" OR "United States")';
  var fullTime = '("full-time" OR "new grad" OR "entry level" OR "early career" OR junior OR associate OR "emerging talent" OR residency OR fellowship OR "Claude Corps")';
  var technicalRoles = '("research engineer" OR "research scientist" OR "AI research" OR "machine learning engineer" OR "applied scientist" OR "applied AI" OR "agent engineer" OR "AI safety engineer" OR "evals engineer" OR "AI infrastructure" OR "forward-deployed engineer")';
  var frontierTopics = '(agentic OR "AI agent" OR "agent evaluation" OR observability OR "memory governance" OR "LLM reasoning" OR "long horizon" OR "reinforcement learning" OR RLHF OR DPO OR GRPO OR "post-training" OR "reward model" OR RAG OR LLMOps OR "model safety" OR alignment)';

  // ── Core queries: run every day. This pipeline is full-time/new-grad only. ──
  var core = [
    { q: fullTime + ' ' + technicalRoles + ' ' + ats + ' ' + locations + ' -workday -intern -internship -co-op', type: 'new_grad' },
    { q: fullTime + ' ' + frontierTopics + ' ' + ats + ' ' + locations + ' -workday -intern -internship -co-op', type: 'new_grad' },
    { q: '("research engineer" OR "applied scientist" OR "machine learning engineer") ' + fullTime + ' ' + ats + ' ' + locations + ' -workday -intern', type: 'new_grad' },
    { q: '("software engineer" OR "ML engineer" OR "AI engineer") ' + fullTime + ' ' + ats + ' ' + locations + ' -workday -intern', type: 'new_grad' },
    { q: '"new grad" ' + frontierTopics + ' ' + ats + ' ' + locations + ' -workday', type: 'new_grad' },
    { q: '"entry level" ' + technicalRoles + ' ' + ats + ' ' + locations + ' -workday -intern', type: 'new_grad' }
  ];

  // ── Rotating queries: different set each day, all full-time/new-grad ────
  var pool = [
    // Set 0 — Agentic systems, evaluation, and RL
    { q: '"new grad" ("AI" OR "software" OR "agent" OR "reinforcement learning") site:wellfound.com -workday', type: 'new_grad' },
    { q: '"entry level" ("AI agent" OR agentic OR "agent evaluation" OR evals) site:jobs.ashbyhq.com -workday -intern', type: 'new_grad' },
    { q: '"new grad" ("RLHF" OR DPO OR GRPO OR "post-training") site:jobs.ashbyhq.com -workday', type: 'new_grad' },
    { q: '"new grad" ("agent engineer" OR "AI engineer" OR "research engineer") site:greenhouse.io -workday', type: 'new_grad' },
    { q: '"new grad" ("machine learning" OR "deep learning") site:jobs.lever.co ("Remote" OR "United States") -workday', type: 'new_grad' },
    { q: '"new grad" "software engineer" site:jobright.ai ("San Francisco" OR "Remote") -workday', type: 'new_grad' },
    { q: '"full-time" "ML engineer" ("Bay Area" OR "Remote US") -workday -intern', type: 'new_grad' },
    { q: '"new grad" "research engineer" ("San Francisco" OR "Remote") -workday', type: 'new_grad' },

    // Set 1 — Research engineering and ML infrastructure
    { q: '"new grad" ("research engineer" OR "research scientist" OR "applied scientist") site:jobs.lever.co -workday', type: 'new_grad' },
    { q: '"entry level" ("machine learning engineer" OR "ML infrastructure") site:greenhouse.io -workday', type: 'new_grad' },
    { q: '"new graduate" ("AI research" OR "machine learning research") ' + locations + ' -workday -intern', type: 'new_grad' },
    { q: '"new grad" "software" site:jobs.lever.co ("San Francisco" OR "Remote") -workday', type: 'new_grad' },
    { q: '"early career" ("ML engineer" OR "AI engineer" OR "software engineer") ' + locations + ' -workday -intern', type: 'new_grad' },
    { q: '"associate machine learning engineer" ' + locations + ' -workday', type: 'new_grad' },
    { q: '"new grad" "data scientist" ("AI" OR "machine learning") ' + locations + ' -workday', type: 'new_grad' },
    { q: '("new grad" OR "entry level") ("ML platform" OR "ML infrastructure") ' + locations + ' -workday', type: 'new_grad' },

    // Set 2 — Frontier labs, post-training, and LLM systems
    { q: '"new grad" ("LLM" OR "NLP" OR "generative AI" OR "agentic") site:greenhouse.io -workday', type: 'new_grad' },
    { q: '"entry level" ("reinforcement learning" OR RLHF OR DPO OR "LLM reasoning") site:jobs.ashbyhq.com -workday', type: 'new_grad' },
    { q: '"research engineer" ("post-training" OR "reasoning" OR "agent") ' + locations + ' -workday', type: 'new_grad' },
    { q: '"new grad" "machine learning" site:greenhouse.io ("San Francisco" OR "Remote") -workday', type: 'new_grad' },
    { q: '"new grad" ("AI" OR "software" OR "agent" OR "reinforcement learning") site:workatastartup.com -workday', type: 'new_grad' },
    { q: '"full-time" ("agentic" OR "RAG" OR "retrieval" OR "OpenTelemetry") ("AI" OR "ML") -workday -intern', type: 'new_grad' },
    { q: '"new grad" ("AI" OR "research" OR "agent engineer" OR "RL") site:jobs.ashbyhq.com ("Remote" OR "US") -workday', type: 'new_grad' },
    { q: '"entry level" "AI research" ("San Francisco" OR "Bay Area" OR "Remote") -workday', type: 'new_grad' },

    // Set 3 — Broad but still technical early-career coverage
    { q: '"early career" ("deep learning" OR "neural network" OR "agentic" OR "long horizon") ' + locations + ' -workday -intern', type: 'new_grad' },
    { q: '("agent observability" OR "memory governance" OR "reward model" OR "process reward") ("AI" OR "ML") ' + locations + ' -workday -intern', type: 'new_grad' },
    { q: '"new grad" ("AI" OR "ML" OR "RL") site:jobs.ashbyhq.com ("Remote" OR "San Francisco") -workday', type: 'new_grad' },
    { q: '"new grad" site:jobright.ai ("software" OR "AI") ("San Francisco" OR "Remote") -workday', type: 'new_grad' },
    { q: '"new grad" ("software engineer" OR "research") site:workatastartup.com -workday', type: 'new_grad' }
  ];

  // Pick 8 rotating queries based on day-of-year (cycles through all 4 sets)
  var dayOfYear = Math.floor((new Date() - new Date(new Date().getFullYear(), 0, 0)) / 86400000);
  var setIndex = dayOfYear % 4;
  var rotating = pool.slice(setIndex * 8, setIndex * 8 + 8);

  Logger.log("Query rotation: set " + setIndex + " (day " + dayOfYear + "), " + core.length + " full-time core + " + rotating.length + " rotating new-grad queries");
  return core.concat(rotating);
}

function searchDateRestrict_(days) {
  var n = Math.max(1, Number(days) || Number(APP.defaults.searchDays) || 30);
  if (n <= 7) return "d" + Math.ceil(n);
  if (n <= 14) return "w2";
  if (n <= 31) return "m1";
  if (n <= 92) return "m3";
  return "y1";
}

function searchDays_(config) {
  return Math.max(1, Number(config.searchDays || APP.defaults.searchDays) || 30);
}

function searchGoogleCustomSearch_(query, gKey, gCx, config) {
  var url = "https://www.googleapis.com/customsearch/v1?key=" + encodeURIComponent(gKey) +
            "&cx=" + encodeURIComponent(gCx) +
            "&q=" + encodeURIComponent(query) +
            "&num=10&dateRestrict=" + searchDateRestrict_(config && config.searchDays);
  var response = UrlFetchApp.fetch(url, {
    method: "get",
    muteHttpExceptions: true
  });
  var code = response.getResponseCode();
  if (code >= 200 && code < 300) {
    var data = JSON.parse(response.getContentText() || "{}");
    var items = data.items || [];
    return items.map(function(item) {
      return {
        title: item.title || "",
        url: item.link || "",
        snippet: item.snippet || "",
        score: "80"
      };
    });
  }
  throw new Error("GSearch HTTP " + code + ": " + response.getContentText());
}

function searchSerper_(query, apiKey, config) {
  var url = "https://google.serper.dev/search";
  var payload = {
    q: query,
    num: 10,
    tbs: "qdr:" + searchDateRestrict_(config && config.searchDays).replace(/^\d+/, "")
  };
  var response = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    headers: {
      "X-API-KEY": apiKey
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  var code = response.getResponseCode();
  if (code >= 200 && code < 300) {
    var data = JSON.parse(response.getContentText() || "{}");
    var organic = data.organic || [];
    return organic.map(function(item) {
      return {
        title: item.title || "",
        url: item.link || "",
        snippet: item.snippet || "",
        score: "85"
      };
    });
  }
  throw new Error("Serper HTTP " + code + ": " + response.getContentText());
}

function collectCandidates_(tavilyKey, queries, config) {
  var byUrl = {};
  var props = PropertiesService.getScriptProperties();
  var serperKey = props.getProperty("SERPER_API_KEY");
  var gKey = props.getProperty("GSEARCH_API_KEY") || props.getProperty("GOOGLE_SEARCH_API_KEY");
  var gCx = props.getProperty("GSEARCH_CX") || props.getProperty("GOOGLE_SEARCH_CX");

  var requests = [];
  var metaList = [];

  queries.forEach(function (queryObj) {
    var qStr = typeof queryObj === "string" ? queryObj : queryObj.q;
    var qType = typeof queryObj === "string" ? "intern" : (queryObj.type || "intern");

    if (serperKey) {
      requests.push({
        url: "https://google.serper.dev/search",
        method: "post",
        contentType: "application/json",
        headers: { "X-API-KEY": serperKey },
        payload: JSON.stringify({
          q: qStr,
          num: 10,
          tbs: "qdr:" + searchDateRestrict_(config.searchDays).replace(/^\d+/, "")
        }),
        muteHttpExceptions: true
      });
      metaList.push({ engine: "serper", qStr: qStr, qType: qType });
    } else if (gKey && gCx) {
      var gUrl = "https://www.googleapis.com/customsearch/v1?key=" + encodeURIComponent(gKey) +
                 "&cx=" + encodeURIComponent(gCx) +
                 "&q=" + encodeURIComponent(qStr) + "&num=10&dateRestrict=" + searchDateRestrict_(config.searchDays);
      requests.push({
        url: gUrl,
        method: "get",
        muteHttpExceptions: true
      });
      metaList.push({ engine: "gsearch", qStr: qStr, qType: qType });
    } else if (tavilyKey) {
      var tPayload = {
        query: qStr,
        search_depth: String(config.searchDepth || APP.defaults.searchDepth),
        max_results: 10,
        days: searchDays_(config),
        include_answer: false,
        include_raw_content: false
      };
      requests.push({
        url: "https://api.tavily.com/search",
        method: "post",
        contentType: "application/json",
        headers: {
          "Authorization": "Bearer " + tavilyKey,
          "Content-Type": "application/json"
        },
        payload: JSON.stringify(tPayload),
        muteHttpExceptions: true
      });
      metaList.push({ engine: "tavily", qStr: qStr, qType: qType });
    }
  });

  if (requests.length > 0) {
    Logger.log("Parallel executing " + requests.length + " search queries via UrlFetchApp.fetchAll...");
    var responses = UrlFetchApp.fetchAll(requests);
    responses.forEach(function (resp, idx) {
      var meta = metaList[idx];
      var results = [];
      try {
        if (resp.getResponseCode() >= 200 && resp.getResponseCode() < 300) {
          var data = JSON.parse(resp.getContentText() || "{}");
          if (meta.engine === "serper") {
            results = (data.organic || []).map(function(r) { return { title: r.title || "", url: r.link || "", snippet: r.snippet || "", score: "85" }; });
          } else if (meta.engine === "gsearch") {
            results = (data.items || []).map(function(r) { return { title: r.title || "", url: r.link || "", snippet: r.snippet || "", score: "80" }; });
          } else if (meta.engine === "tavily") {
            results = (data.results || []).map(function(r) { return { title: r.title || "", url: r.url || "", snippet: r.content || r.snippet || "", score: r.score || "" }; });
          }
        }
      } catch (e) {
        Logger.log("Error parsing search results for query [" + meta.qStr + "]: " + e.message);
      }

      Logger.log("Query [" + meta.qStr + "] -> Found " + results.length + " candidates using " + meta.engine);

      results.forEach(function (result) {
        var url = normalizeUrl_(result.url || "");
        if (!url || isExcludedUrl_(url, config)) return;
        var source = sourceFromUrl_(url);

        if (!byUrl[url]) {
          byUrl[url] = {
            query: meta.qStr,
            title: result.title || "",
            url: url,
            source: source,
            type: meta.qType,
            snippet: result.snippet || "",
            score: result.score || ""
          };
        }
      });
    });
  }

  // 4. Enrich from intern-list.com (Airtable shared views)
  if (String(config.enableInternList || APP.defaults.enableInternList) === "true") {
    try {
      var internListResults = fetchInternListCandidates_(config);
      Logger.log("intern-list.com: Found " + internListResults.length + " candidates");
      internListResults.forEach(function (result) {
        var url = normalizeUrl_(result.url || "");
        if (!url || isExcludedUrl_(url, config)) return;
        var source = sourceFromUrl_(url);
        if (!byUrl[url]) {
          byUrl[url] = {
            query: "intern-list.com",
            title: result.title || "",
            url: url,
            source: source,
            type: "intern",
            snippet: result.snippet || "",
            score: result.score || "75"
          };
        }
      });
    } catch (e) {
      Logger.log("intern-list.com failed (non-fatal): " + e.message);
    }
  }

  // 5. Enrich from SimplifyJobs GitHub (Internships)
  if (String(config.enableSimplifyJobs || APP.defaults.enableSimplifyJobs) === "true") {
    try {
      var simplifyResults = fetchSimplifyJobsCandidates_(config);
      Logger.log("SimplifyJobs GitHub: Found " + simplifyResults.length + " candidates");
      simplifyResults.forEach(function (result) {
        var url = normalizeUrl_(result.url || "");
        if (!url || isExcludedUrl_(url, config)) return;
        var source = sourceFromUrl_(url);
        if (!byUrl[url]) {
          byUrl[url] = {
            query: "SimplifyJobs",
            title: result.title || "",
            url: url,
            source: source,
            type: "intern",
            snippet: result.snippet || "",
            score: result.score || "75"
          };
        }
      });
    } catch (e) {
      Logger.log("SimplifyJobs GitHub failed (non-fatal): " + e.message);
    }
  }

  // 6. Enrich from newgrad-jobs.com (Airtable shared views)
  if (String(config.enableNewGradJobs || APP.defaults.enableNewGradJobs) === "true") {
    try {
      var newGradJobsResults = fetchNewGradJobsCandidates_(config);
      Logger.log("newgrad-jobs.com: Found " + newGradJobsResults.length + " candidates");
      newGradJobsResults.forEach(function (result) {
        var url = normalizeUrl_(result.url || "");
        if (!url || isExcludedUrl_(url, config)) return;
        var source = sourceFromUrl_(url);
        if (!byUrl[url]) {
          byUrl[url] = {
            query: "newgrad-jobs.com",
            title: result.title || "",
            url: url,
            source: source,
            type: "new_grad",
            snippet: result.snippet || "",
            score: result.score || "80"
          };
        }
      });
    } catch (e) {
      Logger.log("newgrad-jobs.com failed (non-fatal): " + e.message);
    }
  }

  // 7. Enrich from SimplifyJobs New-Grad GitHub
  if (String(config.enableNewGradSimplify || APP.defaults.enableNewGradSimplify) === "true") {
    try {
      var simplifyNewGradResults = fetchNewGradSimplifyJobsCandidates_(config);
      Logger.log("SimplifyJobs New-Grad GitHub: Found " + simplifyNewGradResults.length + " candidates");
      simplifyNewGradResults.forEach(function (result) {
        var url = normalizeUrl_(result.url || "");
        if (!url || isExcludedUrl_(url, config)) return;
        var source = sourceFromUrl_(url);
        if (!byUrl[url]) {
          byUrl[url] = {
            query: "SimplifyJobs-NewGrad",
            title: result.title || "",
            url: url,
            source: source,
            type: "new_grad",
            snippet: result.snippet || "",
            score: result.score || "80"
          };
        }
      });
    } catch (e) {
      Logger.log("SimplifyJobs New-Grad GitHub failed (non-fatal): " + e.message);
    }
  }

  return Object.keys(byUrl).map(function (url) { return byUrl[url]; });
}

function limitCandidates_(candidates, maxUrls) {
  return candidates
    .sort(function (a, b) { return Number(b.score || 0) - Number(a.score || 0); })
    .slice(0, Math.max(1, maxUrls || 45));
}

function extractCandidatePages_(tavilyKey, candidates, config) {
  if (candidates.length === 0) return [];
  var urls = candidates.map(function (c) { return c.url; });
  var byUrl = {};
  candidates.forEach(function (c) { byUrl[c.url] = c; });
  var pages = [];

  var requests = [];
  for (var i = 0; i < urls.length; i += 20) {
    var batch = urls.slice(i, i + 20);
    requests.push({
      url: "https://api.tavily.com/extract",
      method: "post",
      contentType: "application/json",
      headers: { Authorization: "Bearer " + tavilyKey },
      payload: JSON.stringify({
        urls: batch,
        extract_depth: String(config.extractDepth || APP.defaults.extractDepth),
        include_images: false
      }),
      muteHttpExceptions: true
    });
  }

  var responses = UrlFetchApp.fetchAll(requests);
  responses.forEach(function (resp) {
    if (resp.getResponseCode() !== 200) {
      Logger.log("Tavily extract batch failed: HTTP " + resp.getResponseCode() + " - " + resp.getContentText());
      return;
    }
    var data = safeJsonParse_(resp.getContentText()) || {};
    (data.results || []).forEach(function (result) {
      var url = normalizeUrl_(result.url || "");
      var candidate = byUrl[url] || byUrl[normalizeUrl_(result.url || "")] || {};
      pages.push({
        url: url,
        title: candidate.title || "",
        source: candidate.source || sourceFromUrl_(url),
        query: candidate.query || "",
        type: candidate.type || "intern",
        content: truncate_(result.raw_content || result.content || candidate.snippet || "", 18000)
      });
    });
  });

  return pages.filter(function (page) {
    return page.url && page.content && !isExcludedUrl_(page.url, config);
  });
}

function classifyPages_(geminiKey, pages, config) {
  var newGradPages = pages.filter(function(p) { return p.type === "new_grad"; });
  var items = [];

  var legacyPages = pages.filter(function(p) { return p.type !== "new_grad"; });
  if (legacyPages.length > 0) {
    Logger.log("Skipping " + legacyPages.length + " legacy internship pages; this scout is full-time/new-grad only.");
  }
  if (newGradPages.length > 0) {
    Logger.log("=== Starting Classification for NEW GRAD pages (" + newGradPages.length + " total) ===");
    items = items.concat(processBatchLoop_(geminiKey, newGradPages, config, true));
  }
  return items;
}

function processBatchLoop_(geminiKey, pages, config, isNewGrad) {
  var batchSize = Number(config.classifyBatchSize || APP.defaults.classifyBatchSize);
  var items = [];
  var totalBatches = Math.ceil(pages.length / batchSize);
  var label = isNewGrad ? "NEW GRAD" : "INTERN";

  for (var i = 0; i < pages.length; i += batchSize) {
    var batch = pages.slice(i, i + batchSize);
    var batchNum = Math.floor(i / batchSize) + 1;
    Logger.log("Classifying " + label + " batch [" + batchNum + "/" + totalBatches + "] (" + batch.length + " pages)...");

    if (i > 0) {
      Utilities.sleep(200);
    }

    var classified = isNewGrad ?
      classifyNewGradBatchWithGemini_(geminiKey, batch, config) :
      classifyBatchWithGemini_(geminiKey, batch, config);

    if (!classified) {
      Logger.log("  -> Batch classification failed or returned empty.");
      continue;
    }

    var pageResults = Array.isArray(classified.pages) ? classified.pages : [];
    if (pageResults.length === 0 && Array.isArray(classified.opportunities)) {
      pageResults = [{ page_url: batch[0].url, opportunities: classified.opportunities }];
    }

    pageResults.forEach(function (pageResult) {
      var opps = Array.isArray(pageResult.opportunities) ? pageResult.opportunities : [];
      var matchedPage = null;
      batch.forEach(function (p) {
        if (normalizeUrl_(p.url) === normalizeUrl_(pageResult.page_url || "")) matchedPage = p;
      });
      if (!matchedPage && batch.length === 1) matchedPage = batch[0];

      opps.forEach(function (row) {
        row.url = normalizeUrl_(row.url || (matchedPage ? matchedPage.url : ""));
        row.source = row.source || (matchedPage ? matchedPage.source : "");
        row.raw_title = matchedPage ? matchedPage.title : "";
        row.type = isNewGrad ? "new_grad" : "intern";
        Logger.log("  -> Found (" + row.type + "): " + row.company + " - " + row.role + " | Track: " + row.track + " | Relevant: " + row.is_relevant + " | Part-time: " + row.part_time + " | Score: " + row.score + " | Reason: " + (row.reason || "none"));
        if (!isExcludedUrl_(row.url, config)) {
          items.push(row);
        }
      });

      if (opps.length === 0 && matchedPage) {
        Logger.log("  -> No opportunities on: " + matchedPage.url);
      }
    });
  }
  return items;
}

function classifyBatchWithGemini_(geminiKey, pages, config) {
  var promptParts = [
    "You are classifying multiple startup internship job pages for a personal internship scout.",
    "You will receive " + pages.length + " pages below. Classify EACH page independently.",
    "",
    "CANDIDATE PROFILE FOR SCORING ALIGNMENT:",
    "- Education: B.Tech in CS from IIT Bombay (2023), MS in CS at UC Riverside (2025-2027).",
    "- Experience: Research Intern at CompFly AI (San Francisco, CA, Jun 2026-Present) building Agentic Observability (OpenTelemetry nested span tree tracing, tool-call tracking, sub-agent handoffs, token cost normalized) and Agent Memory Governance (context poisoning & cross-session leakage audit).",
    "- Background: 2+ years full-time Data Scientist at Finarb AI (time-series forecasting, LSTMs/TCNs, SQL+Python agentic code-gen workflows, RAG, Docker), plus research in long-horizon LLM agent reasoning (ReAct, CoT, Tree-of-Thought, Buffer-of-Thought, Qwen2.5-7B fine-tuning & MiniGrid sequential decision-making evals).",
    "- HIGHEST PRIORITY FIT (Score 95-100): Agentic AI, AI Agent Engineer, Long-Horizon Reasoning & Planning, Reinforcement Learning (RL, RLHF, RLAIF, DPO, GRPO, PPO), Reward Modeling (PRMs/ORMs), Agent Infrastructure & Observability (OpenTelemetry/tracing), Agent Memory/Context Governance, LLM reasoning/evals (SWE-bench, GAIA, MiniGrid), RAG & GraphRAG, agentic coding tools, or AI research labs.",
    "- SECONDARY FIT (Score 85-94): Applied ML/deep learning, MLOps/infrastructure, or time-series data science forecasting.",
    "- LOW PRIORITY (Score 60-79): Generic SWE, backend SWE, or fullstack roles without any AI/ML focus. You MUST score pure generic SWE roles below 80 UNLESS the company is a frontier AI Lab (e.g., OpenAI, Anthropic, Mistral, Cohere, Scale AI, DeepMind, etc.), in which case SWE roles remain highly relevant (Score 85+).",
    "",
    "Target: STRICTLY part-time or flexible-hour academic-semester internships in Agentic AI, AI/ML research, MLE, SWE, or Data Science.",
    "CRITICAL RULE 1 (Part-time): Do NOT assume co-ops or internships are part-time. US/Canada co-ops are typically full-time 40-hour roles. You MUST only mark `is_relevant` as true if the posting explicitly states it is 'part-time', 'flexible hours', '10-20 hours/week', or designed to be completed concurrently with academic classes. If it is a full-time 40 hr/week position, or is silent on part-time flexibility, set `is_relevant` to false and `part_time` to 'No'.",
    "CRITICAL RULE 2 (Strictly CS / AI / Data Technical Roles): We ONLY want computer science, software engineering, machine learning, AI agents/LLM research, and data science roles. You MUST set `is_relevant` to false, score below 60, and track to 'Other' if the role is:",
    "  (a) Biological, medical, chemical, clinical, genomics, or physical sciences laboratory research (e.g., medical/genomics laboratory interns like Veracyte);",
    "  (b) Hardware, mechanical, electrical, civil, or antenna engineering;",
    "  (c) Product management, IT support, manual QA, finance/investment, design, marketing, or operations.",
    "Prefer San Francisco, Bay Area, Silicon Valley, Remote US, or US roles. Exclude Workday pages.",
    "If the page URL is an aggregator like Jobright (jobright.ai) or Wellfound, you MUST scan the page text for the actual original ATS application link (e.g., boards.greenhouse.io, jobs.lever.co, jobs.ashbyhq.com) and return THAT exact ATS URL in the \"url\" field instead of the aggregator URL. If it's already a direct ATS page, use its URL exactly.",
    "",
    "Return ONLY valid JSON matching this shape:",
    '{"pages":[{"page_url":"THE_EXACT_PAGE_URL","opportunities":[{"is_relevant":true,"company":"","role":"","track":"Agentic AI|AI Research|MLE|SWE|Data Science|Other","location":"","term":"Fall 2026|Unknown|Other","part_time":"Yes|No|Unknown","url":"","source":"Ashby|Greenhouse|Lever|Other","details":"1-2 sentence summary connecting role to candidate profile","visa_sponsorship":"Yes|No|Unknown","iitb_alumni":"Yes|No|Unknown","score":0,"reason":"short filtering rationale"}]}]}',
    "",
    "Scoring: 95+ exact match on Agentic AI / RL / Agent Observability; 85-94 applied ML/DL/AI intern OR generic SWE at a frontier AI lab; below 80 if generic SWE without AI focus (at a non-AI lab), full-time, lab biology/hardware, or irrelevant.",
    "If a page has no part-time or flexible opportunity, return an empty opportunities array for that page.",
    "You MUST return one entry in the pages array for EACH page below, even if opportunities is empty.",
    "",
    "Config guidance:",
    "Locations: " + config.locationTerms,
    "Roles: " + config.roleTerms,
    "Terms: " + config.termTerms
  ];

  pages.forEach(function (page, idx) {
    promptParts.push("");
    promptParts.push("=== PAGE " + (idx + 1) + " of " + pages.length + " ===");
    promptParts.push("Page title: " + page.title);
    promptParts.push("Page URL: " + page.url);
    promptParts.push("Source: " + page.source);
    promptParts.push("Page text:");
    promptParts.push(truncate_(page.content, 12000));
  });

  var prompt = promptParts.join("\n");

  var payload = {
    contents: [{ role: "user", parts: [{ text: prompt }] }],
    generationConfig: {
      responseMimeType: "application/json",
      temperature: 0.1
    }
  };
  var model = String(config.geminiModel || APP.defaults.geminiModel);
  var url = "https://generativelanguage.googleapis.com/v1beta/models/" + encodeURIComponent(model) + ":generateContent?key=" + encodeURIComponent(geminiKey);
  var response = httpPostJson_(url, payload);
  var text = (((response.candidates || [])[0] || {}).content || {}).parts;
  if (!text || !text[0] || !text[0].text) return null;
  return safeJsonParse_(text[0].text);
}

function classifyNewGradBatchWithGemini_(geminiKey, pages, config) {
  var promptParts = [
    "You are classifying multiple full-time early-career AI/ML job pages for a personal job scout.",
    "You will receive " + pages.length + " pages below. Classify EACH page independently.",
    "",
    "CANDIDATE PROFILE FOR SCORING ALIGNMENT:",
    "- Education: B.Tech in CS from IIT Bombay (2023), MS in CS at UC Riverside, expected graduation December 2026.",
    "- Experience: Research Intern at CompFly AI (San Francisco, CA, Jun 2026-Present) building Agentic Observability (OpenTelemetry nested span tree tracing, tool-call tracking, sub-agent handoffs, token cost normalized) and Agent Memory Governance (context poisoning & cross-session leakage audit).",
    "- Background: 2+ years full-time Data Scientist at Finarb AI (time-series forecasting, LSTMs/TCNs, SQL+Python agentic code-gen workflows, RAG, Docker), plus research in long-horizon LLM agent reasoning (ReAct, CoT, Tree-of-Thought, Buffer-of-Thought, Qwen2.5-7B fine-tuning & MiniGrid sequential decision-making evals).",
    "- Work authorization: F-1 student expecting US OPT eligibility. Mark visa_sponsorship Yes only when the posting explicitly says sponsorship is available; otherwise use Unknown, never infer it.",
    "- HIGHEST PRIORITY FIT (Score 95-100): Agentic AI, AI Agent Engineer, Long-Horizon Reasoning & Planning, Reinforcement Learning (RL, RLHF, RLAIF, DPO, GRPO, PPO), Reward Modeling (PRMs/ORMs), Agent Infrastructure & Observability (OpenTelemetry/tracing), Agent Memory/Context Governance, LLM reasoning/evals, RAG & GraphRAG, agentic coding tools, or AI research labs.",
    "- SECONDARY FIT (Score 85-94): Applied ML/deep learning, MLOps/infrastructure, or time-series data science forecasting.",
    "- LOW PRIORITY (Score 60-79): Generic SWE, backend SWE, or fullstack roles without any AI/ML focus. You MUST score pure generic SWE roles below 80 UNLESS the company is a frontier AI Lab (e.g., OpenAI, Anthropic, Mistral, Cohere, Scale AI, DeepMind, etc.), in which case SWE roles remain highly relevant (Score 85+).",
    "",
    "Target: STRICTLY full-time new grad, early-career, or entry-level positions in Agentic AI, AI/ML Research, Research Engineering, MLE, SWE, or Data Science.",
    "CRITICAL RULE 1 (Full-Time Early Career): Accept roles explicitly labeled new grad, new graduate, entry level, early career, junior, or associate, plus technically plausible roles requiring up to 3 years of experience. Reject internships, co-ops, part-time roles, and roles requiring 4+ years or senior/staff/lead experience.",
    "CRITICAL RULE 2 (Strictly CS / AI / Data Technical Roles): We ONLY want computer science, software engineering, machine learning, AI agents/LLM research, and data science roles. You MUST set `is_relevant` to false, score below 60, and track to 'Other' if the role is:",
    "  (a) Biological, medical, chemical, clinical, genomics, or physical sciences laboratory research;",
    "  (b) Hardware, mechanical, electrical, civil, or antenna engineering;",
    "  (c) Product management, IT support, manual QA, finance/investment, design, marketing, or operations.",
    "",
    "PRIORITY TARGETS & SCORING GUIDANCE (must align with Candidate Profile):",
    "- 95+: Agentic AI / AI Agent Engineer / Long-Horizon Reasoning / Reinforcement Learning (RL/RLHF/DPO) / Agent Observability / Research Engineer role at a frontier lab or strong AI startup aligned with the profile.",
    "- 90-94: Research Engineer, Applied Scientist, or MLE early-career role at a strong AI company aligned with the profile.",
    "- 80-89: AI/ML focused technical early-career role at a startup/tech company, OR a generic SWE role at a frontier AI lab.",
    "- Below 80: Generic SWE without AI focus (unless at an AI lab), irrelevant, nontechnical, internship/co-op, senior, or outside the experience range.",
    "",
    "Prefer San Francisco, Bay Area, Silicon Valley, Remote US, or US roles. Exclude Workday pages.",
    "If the page URL is an aggregator like Jobright (jobright.ai) or Wellfound, you MUST scan the page text for the actual original ATS application link (e.g., boards.greenhouse.io, jobs.lever.co, jobs.ashbyhq.com) and return THAT exact ATS URL in the \"url\" field instead of the aggregator URL. If it's already a direct ATS page, use its URL exactly.",
    "",
    "Return ONLY valid JSON matching this shape:",
    '{"pages":[{"page_url":"THE_EXACT_PAGE_URL","opportunities":[{"is_relevant":true,"company":"","role":"","track":"Agentic AI|AI Research|Research Engineering|MLE|SWE|Data Science|Other","location":"","term":"Full-time New Grad 2026|Full-time Early Career|Unknown","part_time":"No","url":"","source":"Ashby|Greenhouse|Lever|Other","details":"1-2 sentence summary connecting role to candidate profile","visa_sponsorship":"Yes|No|Unknown","iitb_alumni":"Yes|No|Unknown","score":0,"reason":"short scoring rationale"}]}]}',
    "",
    "If a page has no relevant new grad opportunity, return an empty opportunities array for that page.",
    "You MUST return one entry in the pages array for EACH page below, even if opportunities is empty.",
    "",
    "Config guidance:",
    "Locations: " + config.locationTerms,
    "Roles: " + config.roleTerms
  ];

  pages.forEach(function (page, idx) {
    promptParts.push("");
    promptParts.push("=== PAGE " + (idx + 1) + " of " + pages.length + " ===");
    promptParts.push("Page title: " + page.title);
    promptParts.push("Page URL: " + page.url);
    promptParts.push("Source: " + page.source);
    promptParts.push("Page text:");
    promptParts.push(truncate_(page.content, 12000));
  });

  var prompt = promptParts.join("\n");

  var payload = {
    contents: [{ role: "user", parts: [{ text: prompt }] }],
    generationConfig: {
      responseMimeType: "application/json",
      temperature: 0.1
    }
  };
  var model = String(config.geminiModel || APP.defaults.geminiModel);
  var url = "https://generativelanguage.googleapis.com/v1beta/models/" + encodeURIComponent(model) + ":generateContent?key=" + encodeURIComponent(geminiKey);
  var response = httpPostJson_(url, payload);
  var text = (((response.candidates || [])[0] || {}).content || {}).parts;
  if (!text || !text[0] || !text[0].text) return null;
  return safeJsonParse_(text[0].text);
}

// ── External Source: intern-list.com (Airtable Shared Views) ───

function discoverAndScrapeAirtables_(homepageUrl, defaultSources, labelPrefix, defaultScore) {
  var sources = defaultSources.slice();
  try {
    var resp = UrlFetchApp.fetch(homepageUrl, { muteHttpExceptions: true });
    if (resp.getResponseCode() === 200) {
      var html = resp.getContentText();
      var regex = /https?:\/\/airtable\.com\/embed\/(app[a-zA-Z0-9]+)\/(shr[a-zA-Z0-9]+)/gi;
      var match;
      var foundMap = {};
      while ((match = regex.exec(html)) !== null) {
        var k = match[1] + "|" + match[2];
        foundMap[k] = true;
      }
      var discovered = Object.keys(foundMap);
      if (discovered.length > 0) {
        sources = discovered.slice(0, 6).map(function(k, idx) {
          var parts = k.split("|");
          return { appId: parts[0], shareId: parts[1], label: "Table " + (idx + 1) };
        });
        Logger.log("  Dynamically discovered " + sources.length + " active Airtable views on " + homepageUrl);
      }
    }
  } catch (e) {
    Logger.log("  Could not check live homepage " + homepageUrl + ", using fallback sources: " + e.message);
  }

  var candidates = [];
  sources.forEach(function (src, idx) {
    if (idx > 0) Utilities.sleep(1500);
    try {
      var rows = fetchAirtableSharedView_(src.appId, src.shareId);
      Logger.log("  " + labelPrefix + " (" + src.label + "): " + rows.length + " rows");
      rows.forEach(function (row) {
        if (row.url) {
          candidates.push({
            title: (row.company || "Unknown") + " - " + (row.role || "Role"),
            url: row.url,
            snippet: (row.role || "Role") + " at " + (row.company || "Unknown") + " | " + (row.location || "US"),
            score: String(defaultScore || "75")
          });
        }
      });
    } catch (e) {
      Logger.log("  " + labelPrefix + " (" + src.label + ") failed: " + e.message);
    }
  });

  return candidates;
}

function fetchInternListCandidates_(config) {
  var sources = [
    { appId: "appLzkCIXi5t8aYf4", shareId: "shrIEKOHYPMwmpheG", label: "SWE" },
    { appId: "appjSXAWiVF4d1HoZ", shareId: "shrf04yGbrK3IebAl", label: "AI/ML" },
    { appId: "app17F0kkWQZhC6HB", shareId: "shrOTtndhc6HSgnYb", label: "Data" },
    { appId: "appbsiP1flCoaXCSm", shareId: "shreRS1cFLbduwBaU", label: "Data Analysis" },
    { appId: "apprzZO4NFGouLji9", shareId: "shrApQMVthWyRpdyu", label: "Other Tech" }
  ];
  return discoverAndScrapeAirtables_("https://intern-list.com", sources, "intern-list.com", "75");
}

function fetchAirtableSharedView_(appId, shareId) {
  // Step 1: Fetch the embed page to get the signed API URL
  var embedUrl = "https://airtable.com/embed/" + appId + "/" + shareId;
  var embedResp = UrlFetchApp.fetch(embedUrl, { muteHttpExceptions: true });
  if (embedResp.getResponseCode() !== 200) {
    throw new Error("Airtable embed fetch failed: HTTP " + embedResp.getResponseCode());
  }
  var html = embedResp.getContentText();

  // Step 2: Extract the prefetch URL with signed accessPolicy
  var urlMatch = html.match(/urlWithParams:\s*"(.*?)"/)
  if (!urlMatch) {
    throw new Error("Could not find urlWithParams in Airtable embed HTML");
  }
  var urlPath = urlMatch[1].replace(/\\u002F/g, "/");
  var apiUrl = "https://airtable.com" + urlPath;

  // Step 3: Fetch the shared view data
  var apiResp = UrlFetchApp.fetch(apiUrl, {
    headers: {
      "x-time-zone": "America/Los_Angeles",
      "X-Requested-With": "XMLHttpRequest",
      "x-airtable-application-id": appId,
      "x-airtable-inter-service-client": "webClient",
      "x-user-locale": "en"
    },
    muteHttpExceptions: true
  });

  if (apiResp.getResponseCode() !== 200) {
    throw new Error("Airtable API HTTP " + apiResp.getResponseCode() + ": " + apiResp.getContentText().substring(0, 200));
  }

  var apiData = JSON.parse(apiResp.getContentText());
  var data = apiData.data || {};
  var rows = data.rows || (data.table || {}).rows || [];
  var columns = (data.table || {}).columns || [];

  // Build column ID -> name map
  var colMap = {};
  columns.forEach(function (col) {
    colMap[col.id] = (col.name || "").toLowerCase();
  });

  // Parse rows into structured objects
  var results = [];
  rows.forEach(function (row) {
    var cells = row.cellValuesByColumnId || {};
    var parsed = { company: "", role: "", location: "", url: "" };
    Object.keys(cells).forEach(function (colId) {
      var name = colMap[colId] || "";
      var val = cells[colId];
      if (name.indexOf("company") !== -1) parsed.company = String(val || "");
      else if (name.indexOf("role") !== -1 || name.indexOf("title") !== -1 || name.indexOf("position") !== -1) parsed.role = String(val || "");
      else if (name.indexOf("location") !== -1) parsed.location = String(val || "");
      else if (name.indexOf("link") !== -1 || name.indexOf("url") !== -1 || name.indexOf("apply") !== -1) {
        if (typeof val === "string") parsed.url = val;
        else if (val && val.url) parsed.url = val.url;
        else if (val && val.label) parsed.url = val.label;
      }
    });
    if (parsed.url) results.push(parsed);
  });

  return results;
}

// ── External Source: SimplifyJobs GitHub ───────────────────────

function parseSimplifyJobsReadme_(md, defaultScore) {
  var candidates = [];
  var lastCompany = "Unknown";

  // 1. Try HTML <tr> table rows
  var trBlocks = md.split("<tr");
  if (trBlocks.length > 1) {
    for (var i = 1; i < trBlocks.length; i++) {
      var r = trBlocks[i];
      if (/Company/i.test(r) || /<th/i.test(r)) continue;
      
      var tds = [];
      var tdRegex = /<td[^>]*>([\s\S]*?)<\/td>/gi;
      var match;
      while ((match = tdRegex.exec(r)) !== null) {
        tds.push(match[1].replace(/<[^>]+>/g, "").trim());
      }
      if (tds.length < 3) continue;

      var rawCompany = tds[0].replace(/🔥/g, "").replace(/\*\*/g, "").replace(/↳/g, "").trim();
      if (rawCompany && rawCompany !== "↳") {
        lastCompany = rawCompany;
      }
      var company = lastCompany;
      var role = tds[1] || "Role";
      var location = tds[2] || "US";

      var urlRegex = /href=[\"\'](https?:\/\/[^\"\']+)[\"\']/gi;
      var urlMatch;
      var applyUrl = null;
      while ((urlMatch = urlRegex.exec(r)) !== null) {
        var u = urlMatch[1];
        if (u.indexOf("simplify.jobs") === -1 && u.indexOf("github.com") === -1 && u.indexOf("swelist.com") === -1) {
          applyUrl = u;
        }
      }
      if (!applyUrl) continue;

      candidates.push({
        title: company + " - " + role,
        url: applyUrl,
        snippet: role + " at " + company + " | " + location,
        score: String(defaultScore || "75")
      });
    }
  }

  // 2. Fall back to Markdown | table rows if HTML blocks yielded nothing
  if (candidates.length === 0) {
    var lines = md.split("\n");
    lines.forEach(function (line) {
      if (!line.match(/^\|/)) return;
      if (line.match(/^\|\s*---/)) return;
      if (line.match(/^\|\s*Company/i)) return;

      var cols = line.split("|").map(function (c) { return c.trim(); }).filter(Boolean);
      if (cols.length < 4) return;

      var rawCompany = cols[0].replace(/\*\*/g, "").replace(/\[([^\]]+)\]\([^)]+\)/g, "$1").trim();
      if (rawCompany && rawCompany !== "↳") lastCompany = rawCompany;
      var company = lastCompany;
      var role = cols[1].replace(/\[([^\]]+)\]\([^)]+\)/g, "$1").trim();
      var location = cols[2].trim();

      var um = line.match(/\[(?:↗️|🔗|Apply|Link)\]\((https?:\/\/[^)]+)\)/i);
      if (!um) {
        um = line.match(/\(https?:\/\/[^)]+\)/g);
        if (um) {
          var lastUrl = um[um.length - 1];
          um = [null, lastUrl.replace(/^\(/, "").replace(/\)$/, "")];
        }
      }
      if (!um) return;
      var applyUrl = um[1];
      if (!applyUrl || applyUrl.indexOf("simplify.jobs") !== -1) return;

      candidates.push({
        title: company + " - " + role,
        url: applyUrl,
        snippet: role + " at " + company + " | " + location,
        score: String(defaultScore || "75")
      });
    });
  }

  return candidates;
}

function fetchSimplifyJobsCandidates_(config) {
  var url = "https://raw.githubusercontent.com/SimplifyJobs/Summer2026-Internships/dev/README-Off-Season.md";
  var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) {
    throw new Error("SimplifyJobs fetch failed: HTTP " + resp.getResponseCode());
  }
  return parseSimplifyJobsReadme_(resp.getContentText(), "75");
}

// ── External Source: newgrad-jobs.com (Airtable Shared Views) ───

function fetchNewGradJobsCandidates_(config) {
  var sources = [
    { appId: "appoxNzAIRReFCzZV", shareId: "shrmDBF1vNPtzNjzl", label: "AI/ML" },
    { appId: "appjDG7vmPOm1pO7S", shareId: "shr763VHjlzPBDCgN", label: "SWE" },
    { appId: "appZ5SmkwkcW7Xd8C", shareId: "shr51y9s2uIRlkvI8", label: "Data Analysis" },
    { appId: "appqYfRGKpLQ8UsdH", shareId: "shrFnvW20reJCEkYZ", label: "Data Engineer" }
  ];
  return discoverAndScrapeAirtables_("https://newgrad-jobs.com", sources, "newgrad-jobs.com", "80");
}

// ── External Source: SimplifyJobs New-Grad GitHub ──────────────

function fetchNewGradSimplifyJobsCandidates_(config) {
  var url = "https://raw.githubusercontent.com/SimplifyJobs/New-Grad-Positions/dev/README.md";
  var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) {
    throw new Error("SimplifyJobs New-Grad fetch failed: HTTP " + resp.getResponseCode());
  }
  return parseSimplifyJobsReadme_(resp.getContentText(), "80");
}

function upsertOpportunities_(ss, items) {
  var oppSheet = ss.getSheetByName(APP.sheets.opportunities);
  var seenSheet = ss.getSheetByName(APP.sheets.seen);
  var maps = loadSeenMaps_(ss);
  var now = new Date();
  var newRows = [];
  var newSeenRows = [];
  var newItems = [];
  var starredCol = APP.opportunityHeaders.indexOf("starred") + 1;

  items.forEach(function (item) {
    var normalizedUrl = normalizeUrl_(item.url || "");
    var companyRole = normalizeCompanyRole_(item.company, item.role);
    var fingerprint = fingerprint_(normalizedUrl, companyRole);
    if (!normalizedUrl || !item.company || !item.role) return;

    // Strict 3-way deduplication: block if fingerprint, URL, OR normalized companyRole was seen before
    if (maps.fingerprints[fingerprint] || maps.urls[normalizedUrl] || maps.companyRoles[companyRole]) {
      Logger.log("Deduplicated: Blocked duplicate opportunity [" + item.company + " - " + item.role + "]");
      return;
    }

    var id = Utilities.getUuid();
    var row = [
      id,
      item.type || "intern",
      "New",
      false,
      false,
      item.company || "",
      item.role || "",
      item.track || "",
      item.location || "",
      item.term || "",
      item.part_time || "Unknown",
      normalizedUrl,
      item.source || sourceFromUrl_(normalizedUrl),
      item.details || item.reason || "",
      item.visa_sponsorship || "Unknown",
      item.iitb_alumni || "Unknown",
      Number(item.score || 0),
      now,
      "",
      now,
      ""
    ];
    newRows.push(row);
    newSeenRows.push([fingerprint, normalizedUrl, companyRole, now, now]);

    // Mark as seen in memory for current batch
    maps.fingerprints[fingerprint] = true;
    maps.urls[normalizedUrl] = true;
    maps.companyRoles[companyRole] = true;

    newItems.push(objectFromHeaders_(APP.opportunityHeaders, row));
  });

  if (newRows.length > 0) {
    oppSheet.getRange(oppSheet.getLastRow() + 1, 1, newRows.length, APP.opportunityHeaders.length).setValues(newRows);
    oppSheet.getRange(oppSheet.getLastRow() - newRows.length + 1, starredCol, newRows.length, 2).insertCheckboxes();
    seenSheet.getRange(seenSheet.getLastRow() + 1, 1, newSeenRows.length, 5).setValues(newSeenRows);
  }

  return newItems;
}

function batchTouchOpportunities_(ss, activeUrlsMap, now) {
  var oppSheet = ss.getSheetByName(APP.sheets.opportunities);
  var seenSheet = ss.getSheetByName(APP.sheets.seen);
  
  if (oppSheet && oppSheet.getLastRow() > 1) {
    var oppValues = oppSheet.getDataRange().getValues();
    var urlIdx = APP.opportunityHeaders.indexOf("url");
    var lastSeenCol = APP.opportunityHeaders.indexOf("last_seen_at");
    var touchedCount = 0;
    for (var i = 1; i < oppValues.length; i++) {
      var u = normalizeUrl_(oppValues[i][urlIdx] || "");
      if (u && activeUrlsMap[u]) {
        oppValues[i][lastSeenCol] = now;
        touchedCount++;
      }
    }
    if (touchedCount > 0) {
      var colValues = oppValues.slice(1).map(function(r) { return [r[lastSeenCol]]; });
      oppSheet.getRange(2, lastSeenCol + 1, colValues.length, 1).setValues(colValues);
      Logger.log("Batched updated last_seen_at for " + touchedCount + " existing opportunities.");
    }
  }

  if (seenSheet && seenSheet.getLastRow() > 1) {
    var seenValues = seenSheet.getDataRange().getValues();
    var touchedSeen = 0;
    for (var j = 1; j < seenValues.length; j++) {
      var su = normalizeUrl_(seenValues[j][1] || "");
      if (su && activeUrlsMap[su]) {
        seenValues[j][4] = now; // col E is last_seen_at
        touchedSeen++;
      }
    }
    if (touchedSeen > 0) {
      var seenColValues = seenValues.slice(1).map(function(r) { return [r[4]]; });
      seenSheet.getRange(2, 5, seenColValues.length, 1).setValues(seenColValues);
    }
  }
}

function loadSeenMaps_(ss) {
  var seenSheet = ss.getSheetByName(APP.sheets.seen);
  var oppSheet = ss.getSheetByName(APP.sheets.opportunities);
  
  var fingerprints = {};
  var urls = {};
  var companyRoles = {};

  if (seenSheet && seenSheet.getLastRow() > 1) {
    var sValues = seenSheet.getDataRange().getValues();
    for (var i = 1; i < sValues.length; i++) {
      var fp = String(sValues[i][0] || "");
      var u = normalizeUrl_(sValues[i][1] || "");
      var cr = String(sValues[i][2] || "");
      if (fp) fingerprints[fp] = true;
      if (u) urls[u] = true;
      if (cr) companyRoles[cr] = true;
    }
  }

  if (oppSheet && oppSheet.getLastRow() > 1) {
    var oValues = oppSheet.getDataRange().getValues();
    var cIdx = APP.opportunityHeaders.indexOf("company");
    var rIdx = APP.opportunityHeaders.indexOf("role");
    var uIdx = APP.opportunityHeaders.indexOf("url");
    for (var j = 1; j < oValues.length; j++) {
      var ou = normalizeUrl_(oValues[j][uIdx] || "");
      var ocr = normalizeCompanyRole_(oValues[j][cIdx], oValues[j][rIdx]);
      if (ou) urls[ou] = true;
      if (ocr) companyRoles[ocr] = true;
    }
  }

  return {
    fingerprints: fingerprints,
    urls: urls,
    companyRoles: companyRoles
  };
}

function sendResultsEmail_(recipient, ss, items, config) {
  var shown = items || [];
  if (shown.length === 0) return; // Only send email if there are LLM-validated items
  
  // Extract top company names for the subject line
  var topCompanies = [];
  shown.forEach(function(item) {
    if (item.company && topCompanies.indexOf(item.company) === -1 && topCompanies.length < 3) {
      topCompanies.push(item.company);
    }
  });
  
  var companySnippet = topCompanies.length > 0 ? " (" + topCompanies.join(", ") + ")" : "";
  var titleSuffix = shown.length === 1 ? "full-time opportunity" : "full-time opportunities";
  
  var subject = config.emailSubjectPrefix + ": " + shown.length + " new " + titleSuffix + companySnippet;
  var sheetUrl = ss.getUrl();
  
  var html = [
    "<div style=\"font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; max-width: 680px; margin: 0 auto; color: #1f2937;\">",
    "<h2 style=\"color: #1e3a8a; border-bottom: 2px solid #e5e7eb; padding-bottom: 8px;\">🤖 Full-Time AI Job Scout Digest</h2>",
    "<p style=\"font-size: 15px; color: #4b5563;\">Found <strong>" + shown.length + "</strong> new LLM-validated full-time opportunities matching your Research Engineer / ML Engineer profile.</p>",
    "<p><a href=\"" + escapeHtml_(sheetUrl) + "\" style=\"display: inline-block; background-color: #2563eb; color: #ffffff; padding: 8px 16px; text-decoration: none; border-radius: 6px; font-weight: 600;\">Open Dashboard in Google Sheets ↗</a></p>",
    "<hr style=\"border: none; border-top: 1px solid #e5e7eb; margin: 20px 0;\">"
  ];
  
  var text = [
    "Your scout found " + shown.length + " new LLM-validated full-time opportunities since the last run.",
    "Open dashboard in Google Sheets: " + sheetUrl,
    "",
    "=== LLM-VALIDATED OPPORTUNITIES (" + shown.length + ") ==="
  ];
  
  shown.forEach(function (item, idx) {
    var scoreNum = Number(item.score || 0);
    var isTopTier = scoreNum >= 95;
    var typeTag = item.type === "new_grad" ? "[FULL-TIME]" : "[LEGACY]";
    
    var badgeHtml = "";
    if (isTopTier) {
      badgeHtml += " <span style=\"background-color: #fef3c7; color: #92400e; padding: 2px 8px; border-radius: 4px; font-size: 12px; font-weight: 700;\">🌟 TOP MATCH (95+)</span>";
    }
    if (String(item.iitb_alumni || "").toLowerCase() === "yes") {
      badgeHtml += " <span style=\"background-color: #dbeafe; color: #1e40af; padding: 2px 8px; border-radius: 4px; font-size: 12px; font-weight: 700;\">🎓 IITB ALUMNI</span>";
    }
    if (item.part_time === "Yes") {
      badgeHtml += " <span style=\"background-color: #dcfce7; color: #166534; padding: 2px 8px; border-radius: 4px; font-size: 12px; font-weight: 700;\">⚡ PART-TIME FLEXIBLE</span>";
    }

    html.push(
      "<div style=\"background-color: #f9fafb; border: 1px solid #e5e7eb; border-radius: 8px; padding: 16px; margin-bottom: 16px;\">" +
        "<div style=\"font-size: 17px; font-weight: 700; color: #111827;\">" +
          escapeHtml_(typeTag) + " " + escapeHtml_(item.company) + " — " + escapeHtml_(item.role) + badgeHtml +
        "</div>" +
        "<div style=\"font-size: 13px; color: #6b7280; margin: 6px 0;\">" +
          "<strong>Location:</strong> " + escapeHtml_(item.location || "Remote US") + " &nbsp;|&nbsp; " +
          "<strong>Track:</strong> " + escapeHtml_(item.track || "Agentic AI") + " &nbsp;|&nbsp; " +
          "<strong>Score:</strong> <span style=\"color: #059669; font-weight: 700;\">" + escapeHtml_(String(item.score || "")) + "/100</span>" +
        "</div>" +
        "<div style=\"font-size: 14px; color: #374151; margin-bottom: 12px; line-height: 1.5;\">" +
          escapeHtml_(item.details || "") +
        "</div>" +
        "<a href=\"" + escapeHtml_(item.url) + "\" style=\"display: inline-block; background-color: #059669; color: #ffffff; padding: 6px 14px; text-decoration: none; border-radius: 4px; font-size: 13px; font-weight: 600;\">Apply Now ↗</a>" +
      "</div>"
    );

    text.push([
      (idx + 1) + ". " + typeTag + " " + item.company + " - " + item.role + (isTopTier ? " [TOP MATCH 95+]" : ""),
      "Location: " + (item.location || ""),
      "Track: " + (item.track || ""),
      "Score: " + (item.score || ""),
      "Details: " + (item.details || ""),
      "Apply URL: " + item.url,
      ""
    ].join("\n"));
  });

  html.push("</div>");

  MailApp.sendEmail({
    to: recipient,
    subject: subject,
    htmlBody: html.join("\n"),
    body: text.join("\n")
  });
}

function markEmailed_(ss, items) {
  var sheet = ss.getSheetByName(APP.sheets.opportunities);
  if (!sheet || sheet.getLastRow() <= 1) return;
  var values = sheet.getDataRange().getValues();
  var ids = {};
  items.forEach(function (item) { ids[item.id] = true; });
  var now = new Date();
  var idIdx = APP.opportunityHeaders.indexOf("id");
  var emailedIdx = APP.opportunityHeaders.indexOf("emailed_at");
  var touched = false;
  for (var i = 1; i < values.length; i++) {
    if (ids[values[i][idIdx]]) {
      values[i][emailedIdx] = now;
      touched = true;
    }
  }
  if (touched) {
    var colValues = values.slice(1).map(function(r) { return [r[emailedIdx]]; });
    sheet.getRange(2, emailedIdx + 1, colValues.length, 1).setValues(colValues);
  }
}

function appendRun_(ss, rowObject) {
  var sheet = ss.getSheetByName(APP.sheets.runs);
  var row = APP.runHeaders.map(function (header) { return rowObject[header]; });
  sheet.appendRow(row);
  return sheet.getLastRow();
}

function updateRun_(ss, rowNumber, patch) {
  var sheet = ss.getSheetByName(APP.sheets.runs);
  APP.runHeaders.forEach(function (header, index) {
    if (Object.prototype.hasOwnProperty.call(patch, header)) {
      sheet.getRange(rowNumber, index + 1).setValue(patch[header]);
    }
  });
}

function getPreviousRawUrls_(ss) {
  var rawSheet = ss.getSheetByName(APP.sheets.raw);
  var seenSheet = ss.getSheetByName(APP.sheets.seen);
  var seenUrls = {};
  if (rawSheet && rawSheet.getLastRow() > 1) {
    var rawValues = rawSheet.getRange(2, 4, rawSheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < rawValues.length; i++) {
      var u1 = normalizeUrl_(rawValues[i][0] || "");
      if (u1) seenUrls[u1] = true;
    }
  }
  if (seenSheet && seenSheet.getLastRow() > 1) {
    var seenValues = seenSheet.getRange(2, 2, seenSheet.getLastRow() - 1, 1).getValues();
    for (var j = 0; j < seenValues.length; j++) {
      var u2 = normalizeUrl_(seenValues[j][0] || "");
      if (u2) seenUrls[u2] = true;
    }
  }
  return seenUrls;
}

function getEvaluatedUrls_(ss) {
  var seenSheet = ss.getSheetByName(APP.sheets.seen);
  var oppSheet = ss.getSheetByName(APP.sheets.opportunities);
  var evaluated = {};
  if (seenSheet && seenSheet.getLastRow() > 1) {
    var seenValues = seenSheet.getRange(2, 2, seenSheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < seenValues.length; i++) {
      var u1 = normalizeUrl_(seenValues[i][0] || "");
      if (u1) evaluated[u1] = true;
    }
  }
  if (oppSheet && oppSheet.getLastRow() > 1) {
    var urlIdx = APP.opportunityHeaders.indexOf("url") + 1;
    var oppValues = oppSheet.getRange(2, urlIdx, oppSheet.getLastRow() - 1, 1).getValues();
    for (var j = 0; j < oppValues.length; j++) {
      var u2 = normalizeUrl_(oppValues[j][0] || "");
      if (u2) evaluated[u2] = true;
    }
  }
  return evaluated;
}

function clearRawCandidatesOnly() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(APP.sheets.raw);
  if (sheet && sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, APP.rawHeaders.length).clearContent();
    SpreadsheetApp.getUi().alert("RawCandidates cleared! Headers and data structure preserved.");
  } else {
    SpreadsheetApp.getUi().alert("RawCandidates is already empty!");
  }
}

function writeRawCandidates_(ss, runId, candidates) {
  if (candidates.length === 0) return;
  var sheet = ss.getSheetByName(APP.sheets.raw);
  var rows = candidates.map(function (c) {
    return [runId, c.query, c.title, c.url, c.source, c.snippet, c.score];
  });
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, APP.rawHeaders.length).setValues(rows);

  // Auto-prune RawCandidates to keep max 2000 rows for high sheet performance
  var maxRows = 2000;
  var totalRows = sheet.getLastRow();
  if (totalRows > maxRows + 1) {
    var deleteCount = totalRows - maxRows;
    sheet.deleteRows(2, deleteCount);
    Logger.log("Auto-pruned " + deleteCount + " old raw candidate log rows.");
  }
}

function getDashboardSummary_(ss) {
  var oppSheet = ss.getSheetByName(APP.sheets.opportunities);
  var rows = oppSheet.getDataRange().getValues().slice(1);
  var opportunities = rows
    .filter(function (row) { return row[0]; })
    .map(function (row) { return objectFromHeaders_(APP.opportunityHeaders, row); });
  opportunities.sort(function (a, b) { return Number(b.score || 0) - Number(a.score || 0); });
  return {
    appName: APP.name,
    sheetUrl: ss.getUrl(),
    total: opportunities.length,
    newCount: opportunities.filter(function (o) { return o.status === "New"; }).length,
    starredCount: opportunities.filter(function (o) { return o.starred === true; }).length,
    top: opportunities.slice(0, 25)
  };
}

function tavilyPost_(path, apiKey, payload) {
  return httpPostJson_("https://api.tavily.com" + path, payload, {
    Authorization: "Bearer " + apiKey
  });
}

function httpPostJson_(url, payload, headers) {
  var maxRetries = 5;
  var waitTime = 1000;
  var response, code, body;

  for (var attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      response = UrlFetchApp.fetch(url, {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        headers: headers || {},
        muteHttpExceptions: true
      });
      code = response.getResponseCode();
      body = response.getContentText();

      if (code >= 200 && code < 300) {
        return safeJsonParse_(body);
      }

      Logger.log("HTTP " + code + " from " + url + " (Attempt " + attempt + "/" + maxRetries + "): " + truncate_(body, 200));

      // Retry on transient status codes: 429 and 5xx (server errors)
      if (code === 429 || code >= 500) {
        if (attempt < maxRetries) {
          var sleepDuration;
          if (code === 429) {
            var match = body.match(/Please retry in (\d+(\.\d+)?)s/i);
            var waitSeconds = match ? (parseFloat(match[1]) + 2) : 15;
            sleepDuration = Math.round(waitSeconds * 1000);
            Logger.log("HTTP 429 (Rate Limit). Sleeping for " + sleepDuration + "ms before retrying...");
          } else {
            var jitter = Math.floor(Math.random() * 500);
            sleepDuration = waitTime + jitter;
            waitTime *= 2;
          }
          Utilities.sleep(sleepDuration);
          continue;
        }
      }

      throw new Error("HTTP " + code + " from " + url + ": " + truncate_(body, 1000));
    } catch (e) {
      if (e.message && e.message.indexOf("HTTP ") === 0) {
        throw e;
      }
      Logger.log("Network fetch error (Attempt " + attempt + "/" + maxRetries + "): " + e.message);
      if (attempt < maxRetries) {
        var jitter = Math.floor(Math.random() * 500);
        var sleepDuration = waitTime + jitter;
        Logger.log("Retrying in " + sleepDuration + "ms...");
        Utilities.sleep(sleepDuration);
        waitTime *= 2;
        continue;
      }
      throw e;
    }
  }
}

function normalizeUrl_(url) {
  if (!url) return "";
  var text = String(url).trim().split("#")[0];
  var parts = text.split("?");
  var base = parts[0].replace(/\/$/, "");
  if (parts.length === 1) return base;
  var drop = {
    utm_source: true,
    utm_medium: true,
    utm_campaign: true,
    utm_term: true,
    utm_content: true,
    gh_src: true,
    "lever-source": true
  };
  var kept = parts.slice(1).join("?").split("&").filter(function (pair) {
    var key = decodeURIComponent(pair.split("=")[0] || "").toLowerCase();
    return key && !drop[key];
  });
  return kept.length ? base + "?" + kept.join("&") : base;
}

function isExcludedUrl_(url, config) {
  var lower = String(url || "").toLowerCase();
  var excluded = String(config.excludedDomains || APP.defaults.excludedDomains)
    .split(",")
    .map(function (x) { return x.trim().toLowerCase(); })
    .filter(Boolean);
  return excluded.some(function (part) { return lower.indexOf(part) !== -1; });
}

function isDisqualifiedTitle_(title) {
  var lower = String(title || "").toLowerCase();
  var badPatterns = [
    /\bsenior\b/, /\bsr\.?\b/, /\bstaff\b/, /\bprincipal\b/, /\bdirector\b/, /\bmanager\b/, /\blead\b/,
    /\bhead of\b/, /\b3\+?\s*years?\b/, /\b5\+?\s*years?\b/, /\bclinical\b/, /\bwet[\s-]lab\b/,
    /\bgenomics\b/, /\bbiotech\b/, /\bpharma(ceutical)?\b/, /\bantenna\b/, /\bmechanical\b/, /\belectrical\b/
  ];
  return badPatterns.some(function (re) { return re.test(lower); });
}

function sourceFromUrl_(url) {
  var lower = String(url || "").toLowerCase();
  if (lower.indexOf("ashbyhq.com") !== -1) return "Ashby";
  if (lower.indexOf("greenhouse.io") !== -1) return "Greenhouse";
  if (lower.indexOf("lever.co") !== -1) return "Lever";
  if (lower.indexOf("wellfound.com") !== -1) return "Wellfound";
  if (lower.indexOf("jobright.ai") !== -1) return "Jobright";
  if (lower.indexOf("workatastartup.com") !== -1) return "YC";
  if (lower.indexOf("linkedin.com") !== -1) return "LinkedIn";
  if (lower.indexOf("indeed.com") !== -1) return "Indeed";
  return "Other";
}

function normalizeCompanyRole_(company, role) {
  var c = String(company || "").toLowerCase().replace(/[^a-z0-9]/g, "").replace(/(inc|llc|corp|ltd)$/, "");
  var r = String(role || "").toLowerCase().replace(/[^a-z0-9]/g, "");
  
  // Normalize common engineering abbreviations
  r = r.replace(/swe|sde/g, "softwareengineer")
       .replace(/mle/g, "machinelearningengineer")
       .replace(/ml/g, "machinelearning")
       .replace(/ai/g, "artificialintelligence")
       .replace(/internship/g, "intern");

  return c + "|||" + r;
}

function fingerprint_(url, companyRole) {
  return Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, normalizeUrl_(url) + "||" + companyRole)
  );
}

function objectFromHeaders_(headers, row) {
  var obj = {};
  headers.forEach(function (header, index) {
    obj[header] = row[index];
  });
  return obj;
}

function getRequiredProperty_(key) {
  var value = PropertiesService.getScriptProperties().getProperty(key);
  if (!value) throw new Error("Missing Script Property: " + key);
  return value;
}

function safeJsonParse_(text) {
  var str = String(text || "").trim();
  // 1. Try raw JSON parsing first
  try {
    return JSON.parse(str);
  } catch (e) {
    // Continue to advanced parsing
  }

  // 2. Try to find the JSON boundary by looking for { ... } or [ ... ]
  var firstBrace = str.indexOf('{');
  var lastBrace = str.lastIndexOf('}');
  var firstBracket = str.indexOf('[');
  var lastBracket = str.lastIndexOf(']');

  var start = -1;
  var end = -1;

  if (firstBrace !== -1 && lastBrace !== -1) {
    if (firstBracket !== -1 && lastBracket !== -1) {
      if (firstBrace < firstBracket) {
        start = firstBrace;
        end = lastBrace;
      } else {
        start = firstBracket;
        end = lastBracket;
      }
    } else {
      start = firstBrace;
      end = lastBrace;
    }
  } else if (firstBracket !== -1 && lastBracket !== -1) {
    start = firstBracket;
    end = lastBracket;
  }

  if (start !== -1 && end !== -1 && end > start) {
    var candidate = str.substring(start, end + 1);
    try {
      return JSON.parse(candidate);
    } catch (error) {
      try {
        var furtherCleaned = candidate
          .replace(/,\s*([\]}])/g, '$1') // remove trailing commas
          .replace(/\/\*[\s\S]*?\*\/|([^\\:]|^)\/\/.*$/gm, '$1'); // strip comments
        return JSON.parse(furtherCleaned);
      } catch (error2) {
        throw new Error("Failed to parse JSON. Error: " + error.message + "\nContent: " + candidate);
      }
    }
  }

  throw new Error("Could not find any JSON object or array in LLM response: " + str);
}

function truncate_(text, maxChars) {
  text = String(text || "");
  return text.length > maxChars ? text.slice(0, maxChars) : text;
}

function escapeHtml_(value) {
  return String(value === undefined || value === null ? "" : value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
