/**
 * Fall 2026 AI Internship Scout
 *
 * Google Apps Script + Google Sheets implementation inspired by the original
 * Dreamer AI Startup Internship Scout, but without Dreamer runtime dependencies.
 */

var APP = {
  name: "Fall 2026 AI Internship Scout",
  version: "1.0.0",
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
    emailSubjectPrefix: "Fall 2026 AI Internship Scout"
  },
  opportunityHeaders: [
    "id",
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
    var config = readConfig_(ss);
    var tavilyKey = getRequiredProperty_(APP.props.tavilyKey);
    var geminiKey = getRequiredProperty_(APP.props.geminiKey);
    var recipient = getRequiredProperty_(APP.props.recipientEmail);
    var queries = buildSearchQueries_(config);
    var candidates = collectCandidates_(tavilyKey, queries, config);
    var limited = limitCandidates_(candidates, Number(config.maxUrlsPerRun || APP.defaults.maxUrlsPerRun));

    if (String(config.includeRawCandidates || "true") === "true") {
      writeRawCandidates_(ss, runId, limited);
    }

    var extractedPages = extractCandidatePages_(tavilyKey, limited, config);
    var classified = classifyPages_(geminiKey, extractedPages, config);
    var relevant = classified.filter(function (item) {
      var isPartTime = String(item.part_time || "").toLowerCase() === "yes" ||
        /part[\s-]time|flexible|co-op|semester/i.test(String(item.details || "") + " " + String(item.reason || "") + " " + String(item.role || ""));
      return item.is_relevant === true && isPartTime && Number(item.score || 0) >= Number(config.minScore || APP.defaults.minScore);
    });

    var newItems = upsertOpportunities_(ss, relevant);
    var emailed = false;
    if (newItems.length > 0) {
      sendResultsEmail_(recipient, ss, newItems, config);
      markEmailed_(ss, newItems);
      emailed = true;
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
    ["emailSubjectPrefix", APP.defaults.emailSubjectPrefix, "Email subject prefix."]
  ];
  rows.forEach(function (row) {
    if (!existing[row[0]]) {
      sheet.appendRow(row);
    }
  });
}

function formatDashboard_(ss) {
  var opportunities = ss.getSheetByName(APP.sheets.opportunities);
  opportunities.getRange("C2:D").insertCheckboxes();
  opportunities.autoResizeColumns(1, Math.min(APP.opportunityHeaders.length, 12));
  applyFilterIfMissing_(opportunities, APP.opportunityHeaders.length);

  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["New", "Seen", "Starred", "Applied", "Rejected"], true)
    .setAllowInvalid(true)
    .build();
  opportunities.getRange("B2:B").setDataValidation(statusRule);

  var scoreRange = opportunities.getRange("P2:P");
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
  return [
    '"San Francisco" ' + baseSites + ' "intern" ' + pt + ' ("data science" OR "AI" OR "software" OR "Research") -workday',
    '"Fall 2026" "part-time" "AI intern" "San Francisco" site:jobs.ashbyhq.com -workday',
    '"Fall 2026" ("part time" OR "flexible hours") "machine learning intern" "Bay Area" site:jobs.lever.co -workday',
    '"research intern" ("part-time" OR "flexible") "Fall 2026" "AI" "San Francisco" site:greenhouse.io -workday',
    '"software engineer intern" "Fall 2026" "part-time" "AI" site:jobs.ashbyhq.com -workday',
    '"MLE intern" ("part-time" OR "flexible schedule") "Fall 2026" "San Francisco" site:jobs.lever.co -workday',
    '"machine learning research intern" ("part time" OR "flexible") "Fall 2026" "startup" "San Francisco" -workday',
    '"AI research intern" "part-time" "Fall 2026" "Bay Area" -workday',
    '"software intern" ("part-time" OR co-op OR flexible) "AI" "Fall 2026" "San Francisco" site:greenhouse.io -workday',
    '"data science intern" "Fall 2026" ("part-time" OR flexible) "San Francisco" site:jobs.lever.co -workday'
  ];
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

  queries.forEach(function (query) {
    var results = [];
    var usedSearchEngine = false;

    // 1. Try Serper.dev (Google Search API)
    if (serperKey) {
      try {
        results = searchSerper_(query, serperKey);
        usedSearchEngine = true;
      } catch (e) {
        Logger.log("Serper Search failed for query [" + query + "], trying fallback: " + e.message);
      }
    }

    // 2. Try Google Custom Search (secondary)
    if (!usedSearchEngine && gKey && gCx) {
      try {
        results = searchGoogleCustomSearch_(query, gKey, gCx);
        usedSearchEngine = true;
      } catch (e) {
        Logger.log("Google Custom Search failed for query [" + query + "], trying fallback: " + e.message);
      }
    }

    // 3. Try Tavily Search (tertiary fallback)
    if (!usedSearchEngine || results.length === 0) {
      try {
        var payload = {
          query: query,
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
        Logger.log("Tavily search failed for query [" + query + "]: " + e.message);
      }
    }

    results.forEach(function (result) {
      var url = normalizeUrl_(result.url || "");
      if (!url || isExcludedUrl_(url, config)) return;
      var source = sourceFromUrl_(url);
      if (source === "Other") return; // Keep ONLY direct Greenhouse, Ashby, and Lever application pages

      if (!byUrl[url]) {
        byUrl[url] = {
          query: query,
          title: result.title || "",
          url: url,
          source: source,
          snippet: result.snippet || "",
          score: result.score || ""
        };
      }
    });
  });
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
        content: truncate_(result.raw_content || result.content || candidate.snippet || "", 18000)
      });
    });
  }

  return pages.filter(function (page) {
    return page.url && page.content && !isExcludedUrl_(page.url, config);
  });
}

function classifyPages_(geminiKey, pages, config) {
  var items = [];
  pages.forEach(function (page, index) {
    if (index > 0) {
      Utilities.sleep(5000); // 5-second delay to stay well under the 20 RPM free tier limit
    }
    var classified = classifyPageWithGemini_(geminiKey, page, config);
    if (!classified) return;
    var rows = Array.isArray(classified.opportunities) ? classified.opportunities : [];
    rows.forEach(function (row) {
      row.url = normalizeUrl_(row.url || page.url);
      row.source = row.source || page.source;
      row.raw_title = page.title;
      if (!isExcludedUrl_(row.url, config)) {
        items.push(row);
      }
    });
  });
  return items;
}

function classifyPageWithGemini_(geminiKey, page, config) {
  var prompt = [
    "You are classifying startup internship job pages for a personal internship scout.",
    "",
    "Target: STRICTLY part-time or flexible-hour academic-semester internships in AI/ML research, MLE, SWE, or Data Science.",
    "CRITICAL RULE 1 (Part-time): Do NOT assume co-ops or internships are part-time. US/Canada co-ops are typically full-time 40-hour roles. You MUST only mark `is_relevant` as true if the posting explicitly states it is 'part-time', 'flexible hours', '10-20 hours/week', or designed to be completed concurrently with academic classes. If it is a full-time 40 hr/week position, or is silent on part-time flexibility, set `is_relevant` to false and `part_time` to 'No'.",
    "CRITICAL RULE 2 (Technical Role): We ONLY want technical engineering/research roles (AI/ML Research, MLE, SWE, Data Science). If a role is design (e.g. strategic design), finance/investment (e.g. market research, private equity), marketing, advisory, security analysis, operations, or project management, set `is_relevant` to false and track to 'Other'.",
    "Prefer San Francisco, Bay Area, Silicon Valley, Remote US, or US roles. Exclude Workday pages.",
    "Use the page URL exactly if it is a real application/job page.",
    "",
    "Important search inspiration: \"San Francisco\" site:jobs.ashbyhq.com OR site:greenhouse.io OR site:jobs.lever.co \"intern\" \"part-time\" (\"data science\" OR \"AI\" OR \"software\" OR \"Research\")",
    "",
    "Return ONLY valid JSON matching this shape:",
    "{\"opportunities\":[{\"is_relevant\":true,\"company\":\"\",\"role\":\"\",\"track\":\"AI Research|MLE|SWE|Data Science|Other\",\"location\":\"\",\"term\":\"Fall 2026|Unknown|Other\",\"part_time\":\"Yes|No|Unknown\",\"url\":\"\",\"source\":\"Ashby|Greenhouse|Lever|Other\",\"details\":\"1-2 sentence summary\",\"visa_sponsorship\":\"Yes|No|Unknown\",\"iitb_alumni\":\"Yes|No|Unknown\",\"score\":0,\"reason\":\"short filtering rationale\"}]}",
    "",
    "Scoring: 90+ exact part-time or flexible AI/ML/SWE intern match; 75-89 strong part-time/flexible match; below 65 if full-time or irrelevant.",
    "If no part-time or flexible opportunity exists on the page, return {\"opportunities\":[]}.",
    "",
    "Config guidance:",
    "Locations: " + config.locationTerms,
    "Roles: " + config.roleTerms,
    "Terms: " + config.termTerms,
    "",
    "Page title: " + page.title,
    "Page URL: " + page.url,
    "Source: " + page.source,
    "",
    "Page text:",
    page.content
  ].join("\n");

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

function upsertOpportunities_(ss, items) {
  var oppSheet = ss.getSheetByName(APP.sheets.opportunities);
  var seenSheet = ss.getSheetByName(APP.sheets.seen);
  var seen = loadSeen_(ss);
  var now = new Date();
  var newRows = [];
  var newItems = [];

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
    oppSheet.getRange(oppSheet.getLastRow() - newRows.length + 1, 3, newRows.length, 2).insertCheckboxes();
  }

  return newItems;
}

function touchExistingOpportunity_(sheet, url, companyRole, now) {
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    var existingUrl = normalizeUrl_(values[i][10] || "");
    var existingCompanyRole = normalizeCompanyRole_(values[i][4], values[i][5]);
    if (existingUrl === url || existingCompanyRole === companyRole) {
      sheet.getRange(i + 1, 19).setValue(now);
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
  var subject = config.emailSubjectPrefix + ": " + items.length + " new Fall 2026 internship" + (items.length === 1 ? "" : "s");
  var sheetUrl = ss.getUrl();
  var html = [
    "<p>Your scout found <strong>" + items.length + "</strong> new Fall 2026 AI/ML/SWE internship opportunity" + (items.length === 1 ? "" : "ies") + ".</p>",
    "<p><a href=\"" + escapeHtml_(sheetUrl) + "\">Open dashboard in Google Sheets</a></p>",
    "<ol>"
  ];
  var text = [
    "Your scout found " + items.length + " new Fall 2026 AI/ML/SWE internship opportunities.",
    "",
    "Dashboard: " + sheetUrl,
    ""
  ];

  shown.forEach(function (item) {
    html.push(
      "<li><strong>" + escapeHtml_(item.company) + " - " + escapeHtml_(item.role) + "</strong><br>" +
      escapeHtml_(item.location || "") + " | " + escapeHtml_(item.track || "") + " | Part-time: " + escapeHtml_(item.part_time || "Unknown") + " | Score: " + escapeHtml_(String(item.score || "")) + "<br>" +
      escapeHtml_(item.details || "") + "<br>" +
      "<a href=\"" + escapeHtml_(item.url) + "\">Apply / view posting</a></li>"
    );
    text.push([
      item.company + " - " + item.role,
      "Location: " + (item.location || ""),
      "Track: " + (item.track || ""),
      "Part-time: " + (item.part_time || "Unknown"),
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
  for (var i = 1; i < values.length; i++) {
    if (ids[values[i][0]]) {
      sheet.getRange(i + 1, 18).setValue(now);
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

      console.warn("HTTP " + code + " from " + url + " (Attempt " + attempt + "/" + maxRetries + "): " + truncate_(body, 200));

      // Retry on transient status codes: 429 and 5xx (server errors)
      if (code === 429 || code >= 500) {
        if (attempt < maxRetries) {
          var sleepDuration;
          if (code === 429) {
            var match = body.match(/Please retry in (\d+(\.\d+)?)s/i);
            var waitSeconds = match ? (parseFloat(match[1]) + 2) : 15;
            sleepDuration = Math.round(waitSeconds * 1000);
            console.warn("HTTP 429 (Rate Limit). Sleeping for " + sleepDuration + "ms before retrying...");
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
      console.warn("Network fetch error (Attempt " + attempt + "/" + maxRetries + "): " + e.message);
      if (attempt < maxRetries) {
        var jitter = Math.floor(Math.random() * 500);
        var sleepDuration = waitTime + jitter;
        console.warn("Retrying in " + sleepDuration + "ms...");
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
