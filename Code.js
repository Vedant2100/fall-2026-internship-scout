/**
 * Fall 2026 AI Internship Scout
 *
 * Google Apps Script + Google Sheets implementation inspired by the original
 * Dreamer AI Startup Internship Scout, but without Dreamer runtime dependencies.
 */

var APP = {
  name: "Fall 2026 AI Internship Scout",
  version: "2.0.0",
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
    maxUrlsPerRun: "45",
    maxNewEmailItems: "20",
    minScore: "65",
    searchDepth: "basic",
    extractDepth: "basic",
    geminiModel: "gemini-3.1-flash-lite",
    includeRawCandidates: "true",
    excludedDomains: "workday, myworkdayjobs, wd1.myworkdaysite, wd5.myworkdayjobs, simplify",
    locationTerms: "San Francisco, Bay Area, Silicon Valley, Remote US, United States",
    roleTerms: "AI, machine learning, ML, MLE, software, SWE, data science, research",
    termTerms: "Fall 2026, part-time, intern, internship, co-op",
    emailSubjectPrefix: "Fall 2026 AI Internship Scout",
    classifyBatchSize: "7",
    enableInternList: "true",
    enableSimplifyJobs: "true",
    enableNewGradJobs: "true",
    enableNewGradSimplify: "true",
    newGradSearchQueries: "true"
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
    .createMenu("Internship Scout")
    .addItem("Run Search Now", "runSearchNow")
    .addItem("Setup / Repair Dashboard", "setupDashboard")
    .addItem("Clear Existing Data", "clearScoutData")
    .addSeparator()
    .addItem("Install Daily Trigger", "installDailyTrigger")
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
    if (handler === "runScheduledSearch") {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function sendTestEmail() {
  var recipient = getRequiredProperty_(APP.props.recipientEmail);
  MailApp.sendEmail({
    to: recipient,
    subject: APP.name + " test email",
    htmlBody: "<p>Your Internship Scout email path works.</p>",
    body: "Your Internship Scout email path works."
  });
  SpreadsheetApp.getUi().alert("Test email sent to " + recipient + ".");
}

function showSetupHelp() {
  var html = HtmlService.createHtmlOutputFromFile("SetupHelp")
    .setWidth(700)
    .setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(html, "Fall 2026 AI Internship Scout Setup");
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
    throw new Error("Another Internship Scout run is already active.");
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

    if (String(config.includeRawCandidates || "true") === "true") {
      writeRawCandidates_(ss, runId, limited);
    }

    Logger.log("Extracting text content from " + limited.length + " page URLs...");
    var extractedPages = extractCandidatePages_(tavilyKey, limited, config);
    
    Logger.log("Classifying " + extractedPages.length + " pages with Gemini...");
    var classified = classifyPages_(geminiKey, extractedPages, config);
    
    var relevant = classified.filter(function (item) {
      if (item.type === "new_grad") {
        return item.is_relevant === true && Number(item.score || 0) >= Number(config.minScore || APP.defaults.minScore);
      }
      var isPartTime = String(item.part_time || "").toLowerCase() === "yes" ||
        /part[\s-]time|flexible|co-op|semester/i.test(String(item.details || "") + " " + String(item.reason || "") + " " + String(item.role || ""));
      return item.is_relevant === true && isPartTime && Number(item.score || 0) >= Number(config.minScore || APP.defaults.minScore);
    });

    Logger.log("Classification complete. Found " + relevant.length + " relevant options (out of " + classified.length + " total opportunities found on pages).");

    var newItems = upsertOpportunities_(ss, relevant);
    var emailed = false;
    if (newItems.length > 0) {
      Logger.log("Sending search results email to " + recipient + " with " + newItems.length + " new items...");
      sendResultsEmail_(recipient, ss, newItems, config);
      markEmailed_(ss, newItems);
      emailed = true;
    } else {
      Logger.log("No new opportunities to email in this run.");
    }

    updateRun_(ss, runRow, {
      completed_at: new Date(),
      status: "completed",
      queries_run: queries.length,
      urls_checked: limited.length,
      new_count: newItems.length,
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
  for (var i = 1; i < values.length; i++) {
    if (values[i][0]) existing[String(values[i][0])] = true;
  }
  var rows = [
    ["searchScheduleHour", APP.defaults.searchScheduleHour, "Hour of day, script timezone, for daily trigger."],
    ["maxUrlsPerRun", APP.defaults.maxUrlsPerRun, "Maximum unique candidate URLs to extract/classify per run."],
    ["maxNewEmailItems", APP.defaults.maxNewEmailItems, "Maximum new items shown in each email."],
    ["minScore", APP.defaults.minScore, "Minimum Gemini relevance score required to save/email."],
    ["searchDepth", APP.defaults.searchDepth, "Tavily search_depth: basic or advanced."],
    ["extractDepth", APP.defaults.extractDepth, "Tavily extract_depth: basic or advanced."],
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
  var baseSites = '(site:jobs.ashbyhq.com OR site:greenhouse.io OR site:jobs.lever.co)';
  var pt = '("part-time" OR "part time" OR "during semester" OR co-op OR "flexible hours" OR "flexible schedule")';
  var queries = [
    { q: '"San Francisco" ' + baseSites + ' "intern" ' + pt + ' ("data science" OR "AI" OR "software" OR "Research") -workday', type: 'intern' },
    { q: '"Fall 2026" "part-time" "AI intern" "San Francisco" site:jobs.ashbyhq.com -workday', type: 'intern' },
    { q: '"Fall 2026" ("part time" OR "flexible hours") "machine learning intern" "Bay Area" site:jobs.lever.co -workday', type: 'intern' },
    { q: '"research intern" ("part-time" OR "flexible") "Fall 2026" "AI" "San Francisco" site:greenhouse.io -workday', type: 'intern' },
    { q: '"software engineer intern" "Fall 2026" "part-time" "AI" site:jobs.ashbyhq.com -workday', type: 'intern' },
    { q: '"MLE intern" ("part-time" OR "flexible schedule") "Fall 2026" "San Francisco" site:jobs.lever.co -workday', type: 'intern' },
    { q: '"machine learning research intern" ("part time" OR "flexible") "Fall 2026" "startup" "San Francisco" -workday', type: 'intern' },
    { q: '"AI research intern" "part-time" "Fall 2026" "Bay Area" -workday', type: 'intern' },
    { q: '"software intern" ("part-time" OR co-op OR flexible) "AI" "Fall 2026" "San Francisco" site:greenhouse.io -workday', type: 'intern' },
    { q: '"data science intern" "Fall 2026" ("part-time" OR flexible) "San Francisco" site:jobs.lever.co -workday', type: 'intern' }
  ];
  if (String(config.newGradSearchQueries || APP.defaults.newGradSearchQueries) === "true") {
    queries.push({ q: '"new grad" OR "entry level" site:jobs.ashbyhq.com ("AI" OR "machine learning" OR "research") ("Bay Area" OR "San Francisco") -workday', type: 'new_grad' });
    queries.push({ q: '"new grad" OR "university grad" site:greenhouse.io ("AI" OR "ML" OR "software engineer" OR "research") -workday', type: 'new_grad' });
    queries.push({ q: '"new grad" site:jobs.lever.co ("research" OR "AI" OR "data science") ("San Francisco" OR "remote") -workday', type: 'new_grad' });
  }
  return queries;
}

function searchGoogleCustomSearch_(query, gKey, gCx) {
  var url = "https://www.googleapis.com/customsearch/v1?key=" + encodeURIComponent(gKey) +
            "&cx=" + encodeURIComponent(gCx) +
            "&q=" + encodeURIComponent(query) +
            "&num=10";
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

function searchSerper_(query, apiKey) {
  var url = "https://google.serper.dev/search";
  var payload = {
    q: query,
    num: 10
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

  queries.forEach(function (queryObj) {
    var qStr = typeof queryObj === "string" ? queryObj : queryObj.q;
    var qType = typeof queryObj === "string" ? "intern" : (queryObj.type || "intern");
    var results = [];
    var usedSearchEngine = false;

    // 1. Try Serper.dev (Google Search API)
    if (serperKey) {
      try {
        results = searchSerper_(qStr, serperKey);
        usedSearchEngine = true;
      } catch (e) {
        Logger.log("Serper Search failed for query [" + qStr + "], trying fallback: " + e.message);
      }
    }

    // 2. Try Google Custom Search (secondary)
    if (!usedSearchEngine && gKey && gCx) {
      try {
        results = searchGoogleCustomSearch_(qStr, gKey, gCx);
        usedSearchEngine = true;
      } catch (e) {
        Logger.log("Google Custom Search failed for query [" + qStr + "], trying fallback: " + e.message);
      }
    }

    // 3. Try Tavily Search (tertiary fallback)
    if (!usedSearchEngine || results.length === 0) {
      try {
        var payload = {
          query: qStr,
          search_depth: String(config.searchDepth || APP.defaults.searchDepth),
          max_results: 10,
          include_answer: false,
          include_raw_content: false
        };
        var response = tavilyPost_("/search", tavilyKey, payload);
        var tResults = response.results || [];
        results = tResults.map(function(r) {
          return {
            title: r.title || "",
            url: r.url || "",
            snippet: r.content || r.snippet || "",
            score: r.score || ""
          };
        });
      } catch (e) {
        Logger.log("Tavily search failed for query [" + qStr + "]: " + e.message);
      }
    }

    var engineName = "Tavily";
    if (usedSearchEngine) {
      engineName = serperKey ? "Serper.dev" : "Google Custom Search";
    }
    Logger.log("Query [" + qStr + "] -> Found " + results.length + " candidates using " + engineName);

    results.forEach(function (result) {
      var url = normalizeUrl_(result.url || "");
      if (!url || isExcludedUrl_(url, config)) return;
      var source = sourceFromUrl_(url);
      if (source === "Other") return; // Keep ONLY direct Greenhouse, Ashby, and Lever application pages

      if (!byUrl[url]) {
        byUrl[url] = {
          query: qStr,
          title: result.title || "",
          url: url,
          source: source,
          type: qType,
          snippet: result.snippet || "",
          score: result.score || ""
        };
      }
    });
  });

  // 4. Enrich from intern-list.com (Airtable shared views)
  if (String(config.enableInternList || APP.defaults.enableInternList) === "true") {
    try {
      var internListResults = fetchInternListCandidates_(config);
      Logger.log("intern-list.com: Found " + internListResults.length + " candidates");
      internListResults.forEach(function (result) {
        var url = normalizeUrl_(result.url || "");
        if (!url || isExcludedUrl_(url, config)) return;
        var source = sourceFromUrl_(url);
        if (source === "Other") return;
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
        if (source === "Other") return;
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
        if (source === "Other") return;
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
        if (source === "Other") return;
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

  for (var i = 0; i < urls.length; i += 20) {
    var batch = urls.slice(i, i + 20);
    var response = tavilyPost_("/extract", tavilyKey, {
      urls: batch,
      extract_depth: String(config.extractDepth || APP.defaults.extractDepth),
      include_images: false
    });
    (response.results || []).forEach(function (result) {
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
  }

  return pages.filter(function (page) {
    return page.url && page.content && !isExcludedUrl_(page.url, config);
  });
}

function classifyPages_(geminiKey, pages, config) {
  var internPages = pages.filter(function(p) { return p.type !== "new_grad"; });
  var newGradPages = pages.filter(function(p) { return p.type === "new_grad"; });
  var items = [];

  if (internPages.length > 0) {
    Logger.log("=== Starting Classification for INTERN pages (" + internPages.length + " total) ===");
    items = items.concat(processBatchLoop_(geminiKey, internPages, config, false));
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
      Utilities.sleep(5000);
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
    "Target: STRICTLY part-time or flexible-hour academic-semester internships in AI/ML research, MLE, SWE, or Data Science.",
    "CRITICAL RULE 1 (Part-time): Do NOT assume co-ops or internships are part-time. US/Canada co-ops are typically full-time 40-hour roles. You MUST only mark `is_relevant` as true if the posting explicitly states it is 'part-time', 'flexible hours', '10-20 hours/week', or designed to be completed concurrently with academic classes. If it is a full-time 40 hr/week position, or is silent on part-time flexibility, set `is_relevant` to false and `part_time` to 'No'.",
    "CRITICAL RULE 2 (Technical Role): We ONLY want technical engineering/research roles (AI/ML Research, MLE, SWE, Data Science). If a role is design (e.g. strategic design), finance/investment (e.g. market research, private equity), marketing, advisory, security analysis, operations, or project management, set `is_relevant` to false and track to 'Other'.",
    "Prefer San Francisco, Bay Area, Silicon Valley, Remote US, or US roles. Exclude Workday pages.",
    "Use each page's URL exactly if it is a real application/job page.",
    "",
    "Return ONLY valid JSON matching this shape:",
    '{"pages":[{"page_url":"THE_EXACT_PAGE_URL","opportunities":[{"is_relevant":true,"company":"","role":"","track":"AI Research|MLE|SWE|Data Science|Other","location":"","term":"Fall 2026|Unknown|Other","part_time":"Yes|No|Unknown","url":"","source":"Ashby|Greenhouse|Lever|Other","details":"1-2 sentence summary","visa_sponsorship":"Yes|No|Unknown","iitb_alumni":"Yes|No|Unknown","score":0,"reason":"short filtering rationale"}]}]}',
    "",
    "Scoring: 90+ exact part-time or flexible AI/ML/SWE intern match; 75-89 strong part-time/flexible match; below 65 if full-time or irrelevant.",
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
    "You are classifying multiple startup and tech new grad / entry-level job pages for a personal job scout.",
    "You will receive " + pages.length + " pages below. Classify EACH page independently.",
    "",
    "Target: STRICTLY full-time new grad or entry-level (0-2 years experience / university grad) positions in AI/ML Research, MLE, SWE, or Data Science.",
    "CRITICAL RULE 1 (New Grad / Entry Level): We ONLY want full-time new grad or entry-level positions requiring 0 to 2 years of experience. If a posting requires 3+ years of experience, or is an internship/co-op, set `is_relevant` to false.",
    "CRITICAL RULE 2 (Technical Role): We ONLY want technical engineering/research roles (AI/ML Research, MLE, SWE, Data Science). If a role is design, finance/investment, marketing, advisory, security analysis, operations, or project management, set `is_relevant` to false and track to 'Other'.",
    "",
    "PRIORITY TARGETS & SCORING GUIDANCE:",
    "- 95+: AI/ML Research / MLE / SWE role at a Frontier Research Lab (Google DeepMind, FAIR / Meta AI, OpenAI, Anthropic, xAI, SSI, Mistral, Cohere) OR top AI 'Neolab' / startup in Bay Area working on agentic coding, retrieval/RAG, reinforcement learning (RL), post-training, reasoning, or frontier models (e.g., Magic, Cursor / Anysphere, Cognition, Poolside, Factory, Codeium, Augment, E2B, Aider, Perplexity, Glean, Bespoke Labs, Liquid AI, Scale AI, Imbue, Sakana AI, Physical Intelligence, etc.) in San Francisco / Bay Area or Remote US.",
    "- 90-94: SWE / AI / ML / DS new grad role at Big Tech (Google, Apple, Meta, Microsoft, Amazon, Tesla, NVIDIA) in San Francisco / Bay Area.",
    "- 80-89: SWE / ML new grad role at a well-known tech startup or standard tech company in Bay Area or Remote US.",
    "- 70-79: Standard technical new grad role in other US locations.",
    "- Below 65: Irrelevant or requires 3+ years experience.",
    "",
    "Prefer San Francisco, Bay Area, Silicon Valley, Remote US, or US roles. Exclude Workday pages.",
    "Use each page's URL exactly if it is a real application/job page.",
    "",
    "Return ONLY valid JSON matching this shape:",
    '{"pages":[{"page_url":"THE_EXACT_PAGE_URL","opportunities":[{"is_relevant":true,"company":"","role":"","track":"AI Research|MLE|SWE|Data Science|Other","location":"","term":"New Grad 2025/2026","part_time":"No","url":"","source":"Ashby|Greenhouse|Lever|Other","details":"1-2 sentence summary highlighting if it is a Lab/Neolab/Big Tech","visa_sponsorship":"Yes|No|Unknown","iitb_alumni":"Yes|No|Unknown","score":0,"reason":"short scoring rationale based on Lab/Neolab/Big Tech status"}]}]}',
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

function fetchInternListCandidates_(config) {
  var sources = [
    { appId: "appjSXAWiVF4d1HoZ", shareId: "shrf04yGbrK3IebAl", label: "AI/ML" },
    { appId: "appbsiP1flCoaXCSm", shareId: "shreRS1cFLbduwBaU", label: "Data Analysis" },
    { appId: "appzSWTM1QA543oU", shareId: "shrpvJsQjbhk8l9pi", label: "SWE" }
  ];
  var candidates = [];

  sources.forEach(function (src, idx) {
    if (idx > 0) Utilities.sleep(2000);
    try {
      var rows = fetchAirtableSharedView_(src.appId, src.shareId);
      Logger.log("  intern-list.com " + src.label + ": " + rows.length + " rows");
      rows.forEach(function (row) {
        if (row.url) {
          candidates.push({
            title: (row.company || "") + " - " + (row.role || ""),
            url: row.url,
            snippet: (row.role || "") + " at " + (row.company || "") + " | " + (row.location || ""),
            score: "75"
          });
        }
      });
    } catch (e) {
      Logger.log("  intern-list.com " + src.label + " failed: " + e.message);
    }
  });

  return candidates;
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
  var rows = data.rows || [];
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

function fetchSimplifyJobsCandidates_(config) {
  var url = "https://raw.githubusercontent.com/SimplifyJobs/Summer2026-Internships/dev/README-Off-Season.md";
  var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) {
    throw new Error("SimplifyJobs fetch failed: HTTP " + resp.getResponseCode());
  }

  var md = resp.getContentText();
  var candidates = [];

  // Parse markdown table rows: | Company | Role | Location | Link | Date |
  var lines = md.split("\n");
  lines.forEach(function (line) {
    if (!line.match(/^\|/)) return;
    if (line.match(/^\|\s*---/)) return; // header separator
    if (line.match(/^\|\s*Company/i)) return; // header row

    var cols = line.split("|").map(function (c) { return c.trim(); }).filter(Boolean);
    if (cols.length < 4) return;

    var company = cols[0].replace(/\*\*/g, "").replace(/\[([^\]]+)\]\([^)]+\)/g, "$1").trim();
    var role = cols[1].replace(/\[([^\]]+)\]\([^)]+\)/g, "$1").trim();
    var location = cols[2].trim();

    // Extract URL from markdown link in link column or role column
    var urlMatch = line.match(/\[(?:↗️|🔗|Apply|Link)\]\((https?:\/\/[^)]+)\)/i);
    if (!urlMatch) {
      urlMatch = line.match(/\(https?:\/\/[^)]+\)/g);
      if (urlMatch) {
        // Get the last URL (usually the apply link)
        var lastUrl = urlMatch[urlMatch.length - 1];
        urlMatch = [null, lastUrl.replace(/^\(/, "").replace(/\)$/, "")];
      }
    }
    if (!urlMatch) return;

    var applyUrl = urlMatch[1];
    if (!applyUrl || applyUrl.indexOf("simplify.jobs") !== -1) return;

    candidates.push({
      title: company + " - " + role,
      url: applyUrl,
      snippet: role + " at " + company + " | " + location,
      score: "75"
    });
  });

  return candidates;
}

// ── External Source: newgrad-jobs.com (Airtable Shared Views) ───

function fetchNewGradJobsCandidates_(config) {
  var sources = [
    { appId: "appoxNzAIRReFCzZV", shareId: "shrmDBF1vNPtzNjzl", label: "AI/ML" },
    { appId: "appjDG7vmPOm1pO7S", shareId: "shr763VHjlzPBDCgN", label: "SWE" },
    { appId: "appZ5SmkwkcW7Xd8C", shareId: "shr51y9s2uIRlkvI8", label: "Data Analysis" },
    { appId: "appqYfRGKpLQ8UsdH", shareId: "shrFnvW20reJCEkYZ", label: "Data Engineer" }
  ];
  var candidates = [];

  sources.forEach(function (src, idx) {
    if (idx > 0) Utilities.sleep(2000);
    try {
      var rows = fetchAirtableSharedView_(src.appId, src.shareId);
      Logger.log("  newgrad-jobs.com " + src.label + ": " + rows.length + " rows");
      rows.forEach(function (row) {
        if (row.url) {
          candidates.push({
            title: (row.company || "Unknown") + " - " + (row.role || "New Grad"),
            url: row.url,
            snippet: (row.role || "New Grad Role") + " at " + (row.company || "Unknown") + " | " + (row.location || "US"),
            score: "80"
          });
        }
      });
    } catch (e) {
      Logger.log("  newgrad-jobs.com " + src.label + " fetch failed: " + e.message);
    }
  });

  return candidates;
}

// ── External Source: SimplifyJobs New-Grad GitHub ──────────────

function fetchNewGradSimplifyJobsCandidates_(config) {
  var url = "https://raw.githubusercontent.com/SimplifyJobs/New-Grad-Positions/dev/README.md";
  var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) {
    throw new Error("SimplifyJobs New-Grad fetch failed: HTTP " + resp.getResponseCode());
  }

  var md = resp.getContentText();
  var candidates = [];

  // Parse markdown table rows: | Company | Role | Location | Link | Date |
  var lines = md.split("\n");
  lines.forEach(function (line) {
    if (!line.match(/^\|/)) return;
    if (line.match(/^\|\s*---/)) return; // header separator
    if (line.match(/^\|\s*Company/i)) return; // header row

    var cols = line.split("|").map(function (c) { return c.trim(); }).filter(Boolean);
    if (cols.length < 4) return;

    var company = cols[0].replace(/\*\*/g, "").replace(/\[([^\]]+)\]\([^)]+\)/g, "$1").trim();
    var role = cols[1].replace(/\[([^\]]+)\]\([^)]+\)/g, "$1").trim();
    var location = cols[2].trim();

    // Extract URL from markdown link in link column or role column
    var urlMatch = line.match(/\[(?:↗️|🔗|Apply|Link)\]\((https?:\/\/[^)]+)\)/i);
    if (!urlMatch) {
      urlMatch = line.match(/\(https?:\/\/[^)]+\)/g);
      if (urlMatch) {
        var lastUrl = urlMatch[urlMatch.length - 1];
        urlMatch = [null, lastUrl.replace(/^\(/, "").replace(/\)$/, "")];
      }
    }
    if (!urlMatch) return;

    var applyUrl = urlMatch[1];
    if (!applyUrl || applyUrl.indexOf("simplify.jobs") !== -1) return;

    candidates.push({
      title: company + " - " + role,
      url: applyUrl,
      snippet: role + " at " + company + " | " + location,
      score: "80"
    });
  });

  return candidates;
}

function upsertOpportunities_(ss, items) {
  var oppSheet = ss.getSheetByName(APP.sheets.opportunities);
  var seenSheet = ss.getSheetByName(APP.sheets.seen);
  var seen = loadSeen_(ss);
  var now = new Date();
  var newRows = [];
  var newItems = [];
  var starredCol = APP.opportunityHeaders.indexOf("starred") + 1;

  items.forEach(function (item) {
    var normalizedUrl = normalizeUrl_(item.url || "");
    var companyRole = normalizeCompanyRole_(item.company, item.role);
    var fingerprint = fingerprint_(normalizedUrl, companyRole);
    if (!normalizedUrl || !item.company || !item.role) return;

    if (seen[fingerprint]) {
      seenSheet.getRange(seen[fingerprint].row, 5).setValue(now);
      touchExistingOpportunity_(oppSheet, normalizedUrl, companyRole, now);
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
    newItems.push(objectFromHeaders_(APP.opportunityHeaders, row));
    seenSheet.appendRow([fingerprint, normalizedUrl, companyRole, now, now]);
  });

  if (newRows.length > 0) {
    oppSheet.getRange(oppSheet.getLastRow() + 1, 1, newRows.length, APP.opportunityHeaders.length).setValues(newRows);
    oppSheet.getRange(oppSheet.getLastRow() - newRows.length + 1, starredCol, newRows.length, 2).insertCheckboxes();
  }

  return newItems;
}

function touchExistingOpportunity_(sheet, url, companyRole, now) {
  var values = sheet.getDataRange().getValues();
  var urlIdx = APP.opportunityHeaders.indexOf("url");
  var companyIdx = APP.opportunityHeaders.indexOf("company");
  var roleIdx = APP.opportunityHeaders.indexOf("role");
  var lastSeenCol = APP.opportunityHeaders.indexOf("last_seen_at") + 1;
  for (var i = 1; i < values.length; i++) {
    var existingUrl = normalizeUrl_(values[i][urlIdx] || "");
    var existingCompanyRole = normalizeCompanyRole_(values[i][companyIdx], values[i][roleIdx]);
    if (existingUrl === url || existingCompanyRole === companyRole) {
      sheet.getRange(i + 1, lastSeenCol).setValue(now);
      return;
    }
  }
}

function loadSeen_(ss) {
  var sheet = ss.getSheetByName(APP.sheets.seen);
  var values = sheet.getDataRange().getValues();
  var seen = {};
  for (var i = 1; i < values.length; i++) {
    if (values[i][0]) {
      seen[String(values[i][0])] = { row: i + 1 };
    }
  }
  return seen;
}

function sendResultsEmail_(recipient, ss, items, config) {
  var shown = items; // send all
  var hasNewGrad = items.some(function(item) { return item.type === "new_grad"; });
  var titleSuffix = hasNewGrad ? "job & internship opportunity" + (items.length === 1 ? "" : "ies") : "Fall 2026 internship" + (items.length === 1 ? "" : "s");
  var subject = config.emailSubjectPrefix + ": " + items.length + " new " + titleSuffix;
  var sheetUrl = ss.getUrl();
  var html = [
    "<p>Your scout found <strong>" + items.length + "</strong> new " + titleSuffix + ".</p>",
    "<p><a href=\"" + escapeHtml_(sheetUrl) + "\">Open dashboard in Google Sheets</a></p>",
    "<ol>"
  ];
  var text = [
    "Your scout found " + items.length + " new " + titleSuffix + ".",
    "",
    "Dashboard: " + sheetUrl,
    ""
  ];

  shown.forEach(function (item) {
    var typeTag = item.type === "new_grad" ? "[NEW GRAD] " : "[INTERN] ";
    html.push(
      "<li><strong>" + typeTag + escapeHtml_(item.company) + " - " + escapeHtml_(item.role) + "</strong><br>" +
      escapeHtml_(item.location || "") + " | " + escapeHtml_(item.track || "") + " | Score: " + escapeHtml_(String(item.score || "")) + "<br>" +
      escapeHtml_(item.details || "") + "<br>" +
      "<a href=\"" + escapeHtml_(item.url) + "\">Apply / view posting</a></li>"
    );
    text.push([
      typeTag + item.company + " - " + item.role,
      "Location: " + (item.location || ""),
      "Track: " + (item.track || ""),
      "Score: " + (item.score || ""),
      "Details: " + (item.details || ""),
      "URL: " + item.url,
      ""
    ].join("\n"));
  });
  html.push("</ol>");

  MailApp.sendEmail({
    to: recipient,
    subject: subject,
    htmlBody: html.join("\n"),
    body: text.join("\n")
  });
}

function markEmailed_(ss, items) {
  var sheet = ss.getSheetByName(APP.sheets.opportunities);
  var values = sheet.getDataRange().getValues();
  var ids = {};
  items.forEach(function (item) { ids[item.id] = true; });
  var now = new Date();
  var idIdx = APP.opportunityHeaders.indexOf("id");
  var emailedCol = APP.opportunityHeaders.indexOf("emailed_at") + 1;
  for (var i = 1; i < values.length; i++) {
    if (ids[values[i][idIdx]]) {
      sheet.getRange(i + 1, emailedCol).setValue(now);
    }
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

function writeRawCandidates_(ss, runId, candidates) {
  if (candidates.length === 0) return;
  var sheet = ss.getSheetByName(APP.sheets.raw);
  var rows = candidates.map(function (c) {
    return [runId, c.query, c.title, c.url, c.source, c.snippet, c.score];
  });
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, APP.rawHeaders.length).setValues(rows);
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

function sourceFromUrl_(url) {
  var lower = String(url || "").toLowerCase();
  if (lower.indexOf("ashbyhq.com") !== -1) return "Ashby";
  if (lower.indexOf("greenhouse.io") !== -1) return "Greenhouse";
  if (lower.indexOf("lever.co") !== -1) return "Lever";
  return "Other";
}

function normalizeCompanyRole_(company, role) {
  return [company || "", role || ""]
    .join("|||")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim();
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
