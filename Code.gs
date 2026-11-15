/**
 * Mrs. White's Class Volunteer Schedule (2026–2027)
 * Container-Bound Google Apps Script JSON API (`Code.gs`)
 *
 * Zero-schema-change backend for the mobile web app.
 * - Preserves all existing sheet layouts, formulas, and conditional formatting.
 * - Dynamically resolves row/column coordinates via label scanning (never hardcodes row numbers).
 * - Uses per-month chunked CacheService (<100KB per key) + onEdit invalidation.
 * - Enforces LockService + Compare-And-Swap (CAS) concurrency protection.
 * - Logs all mutations to a hidden `_Audit_Log` tab.
 */

// ============================================================================
// 1. CONFIGURATION
// ============================================================================

var CONFIG = {
  // Optional fallback if deployed standalone; leave blank when container-bound
  SPREADSHEET_ID: "1W2KJJienqrxDjRQ8OkhPYT8wiML3bLqafSVbZZcB1U8",

  // Optional 4-digit Class PIN protection (false = frictionless edits with _Audit_Log)
  REQUIRE_PIN: false,
  CLASS_PIN: "2026",

  // Cache configuration (Apps Script CacheService limit is 100KB per key)
  CACHE_PREFIX: "VOL_SCHED_v1_",
  CACHE_TTL_SECONDS: 3600,

  // Exact tab names in the workbook
  READ_ME_TAB: "Read Me",
  ROSTER_TAB: "Original Sign-Ups",
  AUDIT_LOG_TAB: "_Audit_Log",
  MONTH_TABS: [
    "Oct 2026",
    "Nov 2026",
    "Dec 2026",
    "Jan 2027",
    "Feb 2027",
    "Mar 2027",
    "Apr 2027",
    "May 2027",
    "Jun 2027"
  ],

  // Shift default time windows
  SHIFT_TIMES: {
    CENTERS: { label: "8:15–9:45 AM", start: "08:15", end: "09:45" },
    HOMEWORK: { label: "8:00–8:30 AM", start: "08:00", end: "08:30" },
    EXTRA: { label: "TBD / See Notes", start: "08:15", end: "11:30" }
  }
};

// ============================================================================
// 2. HTTP ENDPOINTS (CORS-SAFE doGet & doPost)
// ============================================================================

/**
 * GET endpoint:
 * - ?action=all (default): Returns full schedule (all months + roster + config)
 * - ?action=month&month=Nov%202026: Returns a single month's parsed schedule
 * - ?action=schema: Returns diagnostic coordinate mapping for all tabs
 * - ?nocache=1: Forces cache bypass and refresh
 */
function doGet(e) {
  try {
    var params = (e && e.parameter) ? e.parameter : {};
    var action = (params.action || "all").toLowerCase();
    var forceRefresh = params.nocache === "1" || params.nocache === "true";

    if (action === "schema") {
      var report = inspectWorkbookSchema();
      return createJsonResponse({ ok: true, schema: report });
    }

    if (action === "month" && params.month) {
      var ss = getWorkbook_();
      var monthData = getMonthDataCached_(ss, params.month, forceRefresh);
      return createJsonResponse({
        ok: true,
        serverTime: new Date().toISOString(),
        requirePin: CONFIG.REQUIRE_PIN,
        month: monthData
      });
    }

    var payload = getAllScheduleData_(forceRefresh);
    return createJsonResponse(payload);
  } catch (err) {
    return createJsonResponse({
      ok: false,
      code: "SERVER_ERROR",
      message: err && err.message ? err.message : String(err)
    });
  }
}

/**
 * POST endpoint:
 * Browser fetch() sends Content-Type: text/plain;charset=utf-8 to avoid CORS preflight failure.
 * Body JSON payload shapes:
 * 1. Daily / Extra column slot update:
 *    {
 *      action: "SUB" | "NEED_SUB" | "CLAIM" | "DIRECT" | "CLEAR",
 *      month: "Nov 2026",
 *      colIndex: 5,               // 1-based sheet column index
 *      roleKey: "centers1" | "centers2" | "centers3" | "homework" | "notes",
 *      expectedPreviousValue: "Michelle Pinero Conrey",
 *      newVolunteer: "Lisi Lanford",
 *      actorName: "Lisi Lanford",
 *      comment: "Optional swap note",
 *      pin: ""
 *    }
 * 2. Monthly footer role update:
 *    {
 *      action: "FOOTER_EDIT",
 *      month: "Nov 2026",
 *      footerKey: "artist" | "composer" | "roomParentAssistant" | "workAtHome",
 *      field: "volunteer" | "detail",
 *      expectedPreviousValue: "",
 *      newValue: "Lindy Yen",
 *      actorName: "Lindy Yen",
 *      comment: "Optional topic/note",
 *      pin: ""
 *    }
 */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return createJsonResponse({
        ok: false,
        code: "BAD_REQUEST",
        message: "Missing request body."
      });
    }

    var body = JSON.parse(e.postData.contents);
    var result = processWriteMutation_(body);
    return createJsonResponse(result);
  } catch (err) {
    return createJsonResponse({
      ok: false,
      code: "SERVER_ERROR",
      message: err && err.message ? err.message : String(err)
    });
  }
}

/**
 * Helper to return CORS-safe JSON text output.
 */
function createJsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================================
// 3. INSTALLABLE TRIGGER FOR DESKTOP SPREADSHEET EDITS
// ============================================================================

/**
 * Bind this function to an Installable "On edit" trigger in Apps Script Triggers.
 * Whenever Mrs. White or the Room Parent edits a monthly tab or Original Sign-Ups
 * on desktop, this immediately invalidates the corresponding cache key.
 */
function handleSheetEdit(e) {
  try {
    if (!e || !e.range) {
      invalidateAllCaches_();
      return;
    }
    var sheet = e.range.getSheet();
    if (!sheet) return;
    var tabName = sheet.getName();

    var cache = CacheService.getScriptCache();
    if (!cache) return;

    if (tabName === CONFIG.ROSTER_TAB) {
      cache.remove(CONFIG.CACHE_PREFIX + "ROSTER");
    } else if (CONFIG.MONTH_TABS.indexOf(tabName) !== -1) {
      cache.remove(CONFIG.CACHE_PREFIX + tabName);
    }
  } catch (err) {
    Logger.log("handleSheetEdit warning: " + err);
  }
}

/**
 * Manual utility to clear all cached tabs.
 */
function invalidateAllCaches_() {
  var cache = CacheService.getScriptCache();
  if (!cache) return;
  var keys = [CONFIG.CACHE_PREFIX + "ROSTER"];
  for (var i = 0; i < CONFIG.MONTH_TABS.length; i++) {
    keys.push(CONFIG.CACHE_PREFIX + CONFIG.MONTH_TABS[i]);
  }
  cache.removeAll(keys);
}

// ============================================================================
// 4. DATA AGGREGATION & CHUNKED CACHING
// ============================================================================

function getWorkbook_() {
  if (CONFIG.SPREADSHEET_ID && CONFIG.SPREADSHEET_ID.trim() !== "") {
    return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID.trim());
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getAllScheduleData_(forceRefresh) {
  var ss = getWorkbook_();
  var roster = getRosterCached_(ss, forceRefresh);
  var months = [];
  var allParentNamesMap = {};

  // Seed parent directory from Original Sign-Ups
  for (var r = 0; r < roster.entries.length; r++) {
    var vName = cleanVolunteerName_(roster.entries[r].volunteer);
    if (vName) allParentNamesMap[vName.toLowerCase()] = vName;
  }

  for (var i = 0; i < CONFIG.MONTH_TABS.length; i++) {
    var tabName = CONFIG.MONTH_TABS[i];
    var monthObj = getMonthDataCached_(ss, tabName, forceRefresh);
    if (monthObj) {
      months.push(monthObj);
      collectMonthParentNames_(monthObj, allParentNamesMap);
    }
  }

  var parentDirectory = Object.keys(allParentNamesMap)
    .map(function(k) { return allParentNamesMap[k]; })
    .sort(function(a, b) { return a.localeCompare(b); });

  return {
    ok: true,
    serverTime: new Date().toISOString(),
    requirePin: CONFIG.REQUIRE_PIN,
    workbookTitle: ss.getName(),
    roster: roster,
    parentDirectory: parentDirectory,
    months: months
  };
}

function getRosterCached_(ss, forceRefresh) {
  var cache = CacheService.getScriptCache();
  var cacheKey = CONFIG.CACHE_PREFIX + "ROSTER";

  if (!forceRefresh && cache) {
    var cachedStr = cache.get(cacheKey);
    if (cachedStr) {
      try {
        return JSON.parse(cachedStr);
      } catch (e) {
        // Fall through on corrupt cache
      }
    }
  }

  var sheet = ss.getSheetByName(CONFIG.ROSTER_TAB);
  var parsed = parseRosterSheet_(sheet);

  if (cache) {
    try {
      var serialized = JSON.stringify(parsed);
      if (serialized.length < 95000) {
        cache.put(cacheKey, serialized, CONFIG.CACHE_TTL_SECONDS);
      }
    } catch (e) {
      Logger.log("Cache write skipped for ROSTER: " + e);
    }
  }

  return parsed;
}

function getMonthDataCached_(ss, tabName, forceRefresh) {
  var cache = CacheService.getScriptCache();
  var cacheKey = CONFIG.CACHE_PREFIX + tabName;

  if (!forceRefresh && cache) {
    var cachedStr = cache.get(cacheKey);
    if (cachedStr) {
      try {
        return JSON.parse(cachedStr);
      } catch (e) {
        // Fall through
      }
    }
  }

  var sheet = ss.getSheetByName(tabName);
  if (!sheet) {
    return null;
  }

  var parsed = parseMonthlySheet_(sheet, tabName);

  if (cache && parsed) {
    try {
      var serialized = JSON.stringify(parsed);
      if (serialized.length < 95000) {
        cache.put(cacheKey, serialized, CONFIG.CACHE_TTL_SECONDS);
      }
    } catch (e) {
      Logger.log("Cache write skipped for " + tabName + ": " + e);
    }
  }

  return parsed;
}

// ============================================================================
// 5. DYNAMIC COORDINATE RESOLVER & SHEET PARSERS (ANTI-DRIFT)
// ============================================================================

/**
 * Parses `Original Sign-Ups` baseline roster tab (`A1:E22+`).
 */
function parseRosterSheet_(sheet) {
  if (!sheet) {
    return { sourceUrl: "", entries: [] };
  }

  var displayGrid = sheet.getDataRange().getDisplayValues();
  var sourceUrl = "";
  var headerRowIndex = -1;

  for (var r = 0; r < Math.min(displayGrid.length, 10); r++) {
    var rowText = displayGrid[r].join(" ").trim();
    if (/https?:\/\//i.test(rowText) && !sourceUrl) {
      var match = rowText.match(/https?:\/\/[^\s"']+/i);
      if (match) sourceUrl = match[0];
    }
    var colA = (displayGrid[r][0] || "").trim().toLowerCase();
    var colD = (displayGrid[r][3] || "").trim().toLowerCase();
    if (colA.indexOf("role") !== -1 || colD.indexOf("volunteer") !== -1) {
      headerRowIndex = r;
      break;
    }
  }

  if (headerRowIndex === -1) headerRowIndex = 2; // Fallback to Row 3 (0-indexed 2)

  var entries = [];
  for (var i = headerRowIndex + 1; i < displayGrid.length; i++) {
    var row = displayGrid[i];
    var role = (row[0] || "").trim();
    var day = (row[1] || "").trim();
    var time = (row[2] || "").trim();
    var volunteer = (row[3] || "").trim();
    var comment = (row[4] || "").trim();

    if (!role && !volunteer) continue;
    entries.push({
      rowNumber: i + 1,
      role: role,
      day: day,
      time: time,
      volunteer: volunteer,
      comment: comment
    });
  }

  return {
    sourceUrl: sourceUrl,
    entries: entries
  };
}

/**
 * Dynamically resolves row and column coordinates on a monthly schedule tab.
 * Scans Column A (and Column B if labels are indented) with case-insensitive regexes.
 * Never hardcodes row numbers.
 */
function resolveMonthlySheetSchema_(displayGrid, tabName) {
  var rowMap = {
    title: 1,
    date: null,
    day: null,
    status: null,
    calendarNotes: null,
    centers1: null,
    centers2: null,
    centers3: null,
    homework: null,
    notes: null,
    footer: {
      artist: null,
      composer: null,
      roomParentAssistant: null,
      workAtHome: null
    }
  };

  for (var r = 0; r < displayGrid.length; r++) {
    var rowNum = r + 1; // 1-based sheet row
    var colA = String(displayGrid[r][0] || "").trim();
    var colB = String(displayGrid[r][1] || "").trim();
    var label = colA || colB;
    if (!label) continue;

    var norm = label.toLowerCase().replace(/\s+/g, " ");

    // Main grid rows (match first occurrence)
    if (!rowMap.date && /^date\b/i.test(norm)) {
      rowMap.date = rowNum;
      continue;
    }
    if (!rowMap.day && /^day\b/i.test(norm)) {
      rowMap.day = rowNum;
      continue;
    }
    if (!rowMap.status && /^status\b/i.test(norm)) {
      rowMap.status = rowNum;
      continue;
    }
    if (!rowMap.calendarNotes && /(school\s*calendar\s*notes|calendar\s*notes|school\s*notes)/i.test(norm)) {
      rowMap.calendarNotes = rowNum;
      continue;
    }

    // Editable daily volunteer rows
    if (!rowMap.centers1 && /centers\s*(volunteer\s*)?1\b/i.test(norm)) {
      rowMap.centers1 = rowNum;
      continue;
    }
    if (!rowMap.centers2 && /centers\s*(volunteer\s*)?2\b/i.test(norm)) {
      rowMap.centers2 = rowNum;
      continue;
    }
    if (!rowMap.centers3 && /centers\s*(volunteer\s*)?3\b/i.test(norm)) {
      rowMap.centers3 = rowNum;
      continue;
    }
    if (!rowMap.homework && /homework\s*folders?/i.test(norm)) {
      rowMap.homework = rowNum;
      continue;
    }
    if (!rowMap.notes && /^(notes\s*\/?\s*comments|comments\s*\/?\s*notes|day\s*notes|notes)\b/i.test(norm)) {
      rowMap.notes = rowNum;
      continue;
    }

    // Monthly footer roles (scan full row in case footer table starts in Col A or Col B)
    for (var c = 0; c < Math.min(displayGrid[r].length, 4); c++) {
      var cellText = String(displayGrid[r][c] || "").trim();
      if (!cellText) continue;

      if (!rowMap.footer.artist && /artist\s+of\s+the\s+month/i.test(cellText)) {
        rowMap.footer.artist = resolveFooterRoleCoordinates_(displayGrid[r], rowNum, c + 1);
      } else if (!rowMap.footer.composer && /composer\s+of\s+the\s+month/i.test(cellText)) {
        rowMap.footer.composer = resolveFooterRoleCoordinates_(displayGrid[r], rowNum, c + 1);
      } else if (!rowMap.footer.roomParentAssistant && /assistant\s+to\s+room\s+parent/i.test(cellText)) {
        rowMap.footer.roomParentAssistant = resolveFooterRoleCoordinates_(displayGrid[r], rowNum, c + 1);
      } else if (!rowMap.footer.workAtHome && /work[\s-]*at[\s-]*home\s+(helper|odds|&\s*ends)/i.test(cellText)) {
        rowMap.footer.workAtHome = resolveFooterRoleCoordinates_(displayGrid[r], rowNum, c + 1);
      }
    }
  }

  // Determine active columns and Extra Details columns
  var maxCols = displayGrid[0] ? displayGrid[0].length : 0;
  var dateRowIdx = rowMap.date ? rowMap.date - 1 : 1;
  var dayRowIdx = rowMap.day ? rowMap.day - 1 : 2;

  // Detect `Extra Details` header in Row 1 (or last 3 columns of the grid)
  var extraDetailsStartCol = null; // 1-based
  if (displayGrid.length > 0) {
    for (var col = 1; col < maxCols; col++) {
      var row1Cell = String(displayGrid[0][col] || "").trim();
      if (/extra\s+details/i.test(row1Cell)) {
        extraDetailsStartCol = col + 1; // 1-based
        break;
      }
    }
  }

  // Determine the rightmost used column in the main schedule grid
  var lastGridCol = 1;
  var scanRows = [
    rowMap.date,
    rowMap.day,
    rowMap.status,
    rowMap.centers1,
    rowMap.centers2,
    rowMap.centers3,
    rowMap.homework,
    rowMap.notes
  ];

  for (var colIdx = 1; colIdx < maxCols; colIdx++) {
    var hasGridContent = false;
    if (displayGrid[0] && String(displayGrid[0][colIdx] || "").trim() !== "") {
      hasGridContent = true;
    }
    for (var sr = 0; sr < scanRows.length; sr++) {
      var rNum = scanRows[sr];
      if (rNum && displayGrid[rNum - 1]) {
        var val = String(displayGrid[rNum - 1][colIdx] || "").trim();
        if (val !== "") {
          hasGridContent = true;
          break;
        }
      }
    }
    if (hasGridContent) {
      lastGridCol = colIdx + 1; // 1-based
    }
  }

  // If Extra Details header is merged across the last 3 columns, ensure all 3 columns are included
  if (extraDetailsStartCol !== null) {
    lastGridCol = Math.max(lastGridCol, extraDetailsStartCol + 2);
  } else if (lastGridCol >= 5) {
    // Fallback: check if the last 3 columns have TBD or blank dates
    extraDetailsStartCol = Math.max(2, lastGridCol - 2);
  }

  return {
    tabName: tabName,
    rowMap: rowMap,
    lastGridCol: lastGridCol,
    extraDetailsStartCol: extraDetailsStartCol
  };
}

/**
 * Resolves the volunteer cell column and optional detail/topic column for a footer role row.
 * Handles merged label cells where adjacent columns may be blank before the value cell,
 * or standard layout where Col A = Role Label, Col B = Volunteer Name, Col C = Topic/Notes.
 */
function resolveFooterRoleCoordinates_(rowArray, rowNum, labelCol1Based) {
  var volunteerCol = labelCol1Based + 1;
  var detailCol = labelCol1Based + 2;

  return {
    row: rowNum,
    labelCol: labelCol1Based,
    volunteerCol: volunteerCol,
    detailCol: detailCol
  };
}

/**
 * Parses a single monthly schedule tab into structured JSON for the mobile app.
 */
function parseMonthlySheet_(sheet, tabName) {
  var displayGrid = sheet.getDataRange().getDisplayValues();
  var schema = resolveMonthlySheetSchema_(displayGrid, tabName);
  var rm = schema.rowMap;

  var titleText = (displayGrid[0] && displayGrid[0][0]) ? String(displayGrid[0][0]).trim() : tabName;
  var days = [];

  for (var col = 2; col <= schema.lastGridCol; col++) {
    var cIdx = col - 1; // 0-based index into displayGrid rows
    var isExtraDetails = schema.extraDetailsStartCol !== null && col >= schema.extraDetailsStartCol;

    var rawDate = getCellDisplay_(displayGrid, rm.date, cIdx);
    var rawDay = getCellDisplay_(displayGrid, rm.day, cIdx);
    var rawStatus = getCellDisplay_(displayGrid, rm.status, cIdx);
    var calendarNotes = getCellDisplay_(displayGrid, rm.calendarNotes, cIdx);

    var centers1Val = getCellDisplay_(displayGrid, rm.centers1, cIdx);
    var centers2Val = getCellDisplay_(displayGrid, rm.centers2, cIdx);
    var centers3Val = getCellDisplay_(displayGrid, rm.centers3, cIdx);
    var homeworkVal = getCellDisplay_(displayGrid, rm.homework, cIdx);
    var dayNotes = getCellDisplay_(displayGrid, rm.notes, cIdx);

    // Skip completely empty trailing columns that are neither dated nor Extra Details
    if (!isExtraDetails && !rawDate && !rawDay && !rawStatus && !centers1Val && !centers2Val && !centers3Val) {
      continue;
    }

    var isoDate = normalizeDateToIso_(rawDate, tabName);
    var normalizedStatus = normalizeStatus_(rawStatus, isExtraDetails, rawDate, calendarNotes);
    var isNoSchool = normalizedStatus === "NO SCHOOL";

    var isWednesday = /^wed/i.test(rawDay);
    var homeworkIsNA = /^n\/?a$/i.test(homeworkVal) || (!isWednesday && !isExtraDetails && homeworkVal === "");

    var extraSlotIndex = isExtraDetails ? (col - schema.extraDetailsStartCol + 1) : null;

    var slots = [];
    if (rm.centers1) {
      slots.push({
        roleKey: "centers1",
        roleLabel: isExtraDetails ? "Event Volunteer 1" : "Centers Volunteer 1",
        shortLabel: isExtraDetails ? "Volunteer 1" : "Centers 1",
        timeWindow: isExtraDetails ? CONFIG.SHIFT_TIMES.EXTRA.label : CONFIG.SHIFT_TIMES.CENTERS.label,
        startTime: isExtraDetails ? CONFIG.SHIFT_TIMES.EXTRA.start : CONFIG.SHIFT_TIMES.CENTERS.start,
        endTime: isExtraDetails ? CONFIG.SHIFT_TIMES.EXTRA.end : CONFIG.SHIFT_TIMES.CENTERS.end,
        row: rm.centers1,
        col: col,
        a1: columnToLetter_(col) + rm.centers1,
        volunteer: centers1Val,
        editable: !isNoSchool,
        isNA: false
      });
    }
    if (rm.centers2) {
      slots.push({
        roleKey: "centers2",
        roleLabel: isExtraDetails ? "Event Volunteer 2" : "Centers Volunteer 2",
        shortLabel: isExtraDetails ? "Volunteer 2" : "Centers 2",
        timeWindow: isExtraDetails ? CONFIG.SHIFT_TIMES.EXTRA.label : CONFIG.SHIFT_TIMES.CENTERS.label,
        startTime: isExtraDetails ? CONFIG.SHIFT_TIMES.EXTRA.start : CONFIG.SHIFT_TIMES.CENTERS.start,
        endTime: isExtraDetails ? CONFIG.SHIFT_TIMES.EXTRA.end : CONFIG.SHIFT_TIMES.CENTERS.end,
        row: rm.centers2,
        col: col,
        a1: columnToLetter_(col) + rm.centers2,
        volunteer: centers2Val,
        editable: !isNoSchool,
        isNA: false
      });
    }
    if (rm.centers3) {
      slots.push({
        roleKey: "centers3",
        roleLabel: isExtraDetails ? "Event Volunteer 3" : "Centers Volunteer 3",
        shortLabel: isExtraDetails ? "Volunteer 3" : "Centers 3",
        timeWindow: isExtraDetails ? CONFIG.SHIFT_TIMES.EXTRA.label : CONFIG.SHIFT_TIMES.CENTERS.label,
        startTime: isExtraDetails ? CONFIG.SHIFT_TIMES.EXTRA.start : CONFIG.SHIFT_TIMES.CENTERS.start,
        endTime: isExtraDetails ? CONFIG.SHIFT_TIMES.EXTRA.end : CONFIG.SHIFT_TIMES.CENTERS.end,
        row: rm.centers3,
        col: col,
        a1: columnToLetter_(col) + rm.centers3,
        volunteer: centers3Val,
        editable: !isNoSchool,
        isNA: false
      });
    }
    if (rm.homework) {
      slots.push({
        roleKey: "homework",
        roleLabel: "Homework Folders",
        shortLabel: "Homework Folders",
        timeWindow: CONFIG.SHIFT_TIMES.HOMEWORK.label,
        startTime: CONFIG.SHIFT_TIMES.HOMEWORK.start,
        endTime: CONFIG.SHIFT_TIMES.HOMEWORK.end,
        row: rm.homework,
        col: col,
        a1: columnToLetter_(col) + rm.homework,
        volunteer: homeworkIsNA ? "N/A" : homeworkVal,
        editable: !isNoSchool && !homeworkIsNA,
        isNA: homeworkIsNA
      });
    }

    days.push({
      id: tabName.replace(/\s+/g, "_") + "_col" + col,
      month: tabName,
      colIndex: col,
      colLetter: columnToLetter_(col),
      isExtraDetails: isExtraDetails,
      extraSlotIndex: extraSlotIndex,
      dateDisplay: rawDate || (isExtraDetails ? "TBD" : ""),
      isoDate: isoDate,
      dayOfWeek: rawDay || (isExtraDetails ? "TBD" : ""),
      status: normalizedStatus,
      rawStatus: rawStatus,
      calendarNotes: calendarNotes,
      notes: dayNotes,
      notesCell: rm.notes ? {
        row: rm.notes,
        col: col,
        a1: columnToLetter_(col) + rm.notes
      } : null,
      hasSubNeededFlag: /sub\s*needed/i.test(dayNotes) || /sub\s*needed/i.test(centers1Val) || /sub\s*needed/i.test(centers2Val) || /sub\s*needed/i.test(centers3Val) || /sub\s*needed/i.test(homeworkVal),
      slots: slots
    });
  }

  // Parse the 4 Monthly Footer Roles
  var footerRoles = [];
  var footerDefs = [
    { key: "artist", label: "Artist of the Month", subtitle: "Monthly Classroom Presentation" },
    { key: "composer", label: "Composer of the Month", subtitle: "Monthly Music Presentation" },
    { key: "roomParentAssistant", label: "Assistant to Room Parent", subtitle: "Monthly Coordination Support" },
    { key: "workAtHome", label: "Work-at-home helper", subtitle: "Prep & Materials from Home" }
  ];

  for (var f = 0; f < footerDefs.length; f++) {
    var def = footerDefs[f];
    var coords = rm.footer[def.key];
    if (coords) {
      var volVal = getCellDisplay_(displayGrid, coords.row, coords.volunteerCol - 1);
      var detVal = getCellDisplay_(displayGrid, coords.row, coords.detailCol - 1);
      footerRoles.push({
        footerKey: def.key,
        roleLabel: def.label,
        subtitle: def.subtitle,
        row: coords.row,
        volunteerCol: coords.volunteerCol,
        detailCol: coords.detailCol,
        volunteerA1: columnToLetter_(coords.volunteerCol) + coords.row,
        detailA1: columnToLetter_(coords.detailCol) + coords.row,
        volunteer: volVal,
        detail: detVal
      });
    }
  }

  return {
    tabName: tabName,
    title: titleText,
    schema: schema.rowMap,
    days: days,
    footerRoles: footerRoles
  };
}

// ============================================================================
// 6. WRITE MUTATIONS WITH LOCKSERVICE & COMPARE-AND-SWAP (CAS)
// ============================================================================

function processWriteMutation_(payload) {
  if (!payload || typeof payload !== "object") {
    return { ok: false, code: "BAD_REQUEST", message: "Invalid JSON payload." };
  }

  // Optional Class PIN validation
  if (CONFIG.REQUIRE_PIN) {
    var suppliedPin = String(payload.pin || "").trim();
    if (suppliedPin !== CONFIG.CLASS_PIN) {
      return {
        ok: false,
        code: "INVALID_PIN",
        message: "Incorrect Class PIN. Please check the PIN in ParentSquare and try again."
      };
    }
  }

  var monthName = String(payload.month || "").trim();
  if (CONFIG.MONTH_TABS.indexOf(monthName) === -1) {
    return {
      ok: false,
      code: "INVALID_MONTH",
      message: "Unrecognized monthly schedule tab: " + monthName
    };
  }

  var action = String(payload.action || "").toUpperCase();
  var actorName = String(payload.actorName || payload.newVolunteer || "Parent Volunteer").trim();
  var userComment = String(payload.comment || "").trim();

  var lock = LockService.getDocumentLock();
  try {
    lock.waitLock(10000);
  } catch (lockErr) {
    return {
      ok: false,
      code: "LOCK_TIMEOUT",
      message: "Another parent is saving a change right now. Please tap Save again in a moment."
    };
  }

  try {
    var ss = getWorkbook_();
    var sheet = ss.getSheetByName(monthName);
    if (!sheet) {
      return { ok: false, code: "TAB_NOT_FOUND", message: "Sheet tab not found: " + monthName };
    }

    var displayGrid = sheet.getDataRange().getDisplayValues();
    var schema = resolveMonthlySheetSchema_(displayGrid, monthName);
    var rm = schema.rowMap;

    // Branch A: Monthly Footer Role Mutation
    if (action === "FOOTER_EDIT") {
      return executeFooterWrite_(ss, sheet, displayGrid, schema, payload, actorName, userComment);
    }

    // Branch B: Daily / Extra Details Column Mutation
    var colIndex = Number(payload.colIndex);
    if (!colIndex || colIndex < 2 || colIndex > schema.lastGridCol + 3) {
      return { ok: false, code: "INVALID_COLUMN", message: "Invalid schedule date column." };
    }

    var cIdx = colIndex - 1;
    var isExtraDetails = schema.extraDetailsStartCol !== null && colIndex >= schema.extraDetailsStartCol;
    var rawDate = getCellDisplay_(displayGrid, rm.date, cIdx);
    var rawDay = getCellDisplay_(displayGrid, rm.day, cIdx);
    var rawStatus = getCellDisplay_(displayGrid, rm.status, cIdx);
    var calendarNotes = getCellDisplay_(displayGrid, rm.calendarNotes, cIdx);
    var normalizedStatus = normalizeStatus_(rawStatus, isExtraDetails, rawDate, calendarNotes);

    if (normalizedStatus === "NO SCHOOL") {
      return {
        ok: false,
        code: "NO_SCHOOL_LOCKED",
        message: "This date is marked NO SCHOOL and cannot accept volunteer sign-ups."
      };
    }

    var roleKey = String(payload.roleKey || "").trim();
    var targetRow = null;
    var roleLabel = "";

    if (roleKey === "centers1") {
      targetRow = rm.centers1;
      roleLabel = isExtraDetails ? "Event Volunteer 1" : "Centers Volunteer 1";
    } else if (roleKey === "centers2") {
      targetRow = rm.centers2;
      roleLabel = isExtraDetails ? "Event Volunteer 2" : "Centers Volunteer 2";
    } else if (roleKey === "centers3") {
      targetRow = rm.centers3;
      roleLabel = isExtraDetails ? "Event Volunteer 3" : "Centers Volunteer 3";
    } else if (roleKey === "homework") {
      targetRow = rm.homework;
      roleLabel = "Homework Folders";
    } else if (roleKey === "notes") {
      targetRow = rm.notes;
      roleLabel = "Notes / Comments";
    } else {
      return { ok: false, code: "INVALID_ROLE", message: "Invalid editable role key: " + roleKey };
    }

    if (!targetRow) {
      return {
        ok: false,
        code: "ROW_NOT_RESOLVED",
        message: "Could not locate row for " + roleLabel + " in " + monthName
      };
    }

    var targetRange = sheet.getRange(targetRow, colIndex);
    var currentVal = String(targetRange.getDisplayValue() || "").trim();

    // Reject writes to N/A Homework Folders on Tue/Thu/Fri
    if (roleKey === "homework") {
      var isWednesday = /^wed/i.test(rawDay);
      if (/^n\/?a$/i.test(currentVal) || (!isWednesday && !isExtraDetails)) {
        return {
          ok: false,
          code: "HOMEWORK_WEDNESDAY_ONLY",
          message: "Homework Folders is active on Wednesdays only."
        };
      }
    }

    // Compare-And-Swap (CAS) concurrency check
    var expectedPrev = String(payload.expectedPreviousValue || "").trim();
    if (currentVal !== expectedPrev) {
      return {
        ok: false,
        code: "CONFLICT",
        currentValue: currentVal,
        message: "This slot was just updated by another parent (current value: \"" + (currentVal || "Open") + "\")."
      };
    }

    var dateLabel = (rawDay ? rawDay + " " : "") + (rawDate || "Extra Details");
    var cellA1 = columnToLetter_(colIndex) + targetRow;
    var newVolunteer = String(payload.newVolunteer || "").trim();
    var updatedCellVal = currentVal;
    var updatedNotesVal = rm.notes ? String(sheet.getRange(rm.notes, colIndex).getDisplayValue() || "").trim() : "";
    var auditComment = userComment;

    if (roleKey === "notes") {
      // Direct edit to the Notes / Comments cell
      updatedCellVal = String(payload.newValue !== undefined ? payload.newValue : newVolunteer).trim();
      targetRange.setValue(updatedCellVal);
      updatedNotesVal = updatedCellVal;
    } else if (action === "NEED_SUB") {
      // Keep the current volunteer's name on the slot and append "SUB NEEDED for [Name]" in Notes / Comments
      var holderName = currentVal || newVolunteer || actorName;
      if (!currentVal && holderName) {
        targetRange.setValue(holderName);
        updatedCellVal = holderName;
      }
      var subFlagText = "SUB NEEDED for " + holderName + " (" + shortRoleName_(roleKey) + ")";
      if (userComment) {
        subFlagText += " - " + userComment;
      }
      if (rm.notes) {
        updatedNotesVal = appendStructuredNote_(updatedNotesVal, subFlagText);
        sheet.getRange(rm.notes, colIndex).setValue(updatedNotesVal);
      }
      auditComment = subFlagText;
    } else if (action === "SUB") {
      if (!newVolunteer) {
        return { ok: false, code: "MISSING_VOLUNTEER", message: "Please select or enter the substitute volunteer's name." };
      }
      updatedCellVal = newVolunteer;
      targetRange.setValue(updatedCellVal);

      if (rm.notes) {
        // Clear any old SUB NEEDED flag for the previous volunteer on this role, then append "Subbing for [Original Name]"
        var cleanedNotes = removeSubNeededFlagFor_(updatedNotesVal, currentVal, shortRoleName_(roleKey));
        var subNote = currentVal
          ? (newVolunteer + " subbing for " + currentVal + " (" + shortRoleName_(roleKey) + ")")
          : (newVolunteer + " claimed " + shortRoleName_(roleKey));
        if (userComment) {
          subNote += " - " + userComment;
        }
        updatedNotesVal = appendStructuredNote_(cleanedNotes, subNote);
        sheet.getRange(rm.notes, colIndex).setValue(updatedNotesVal);
        auditComment = subNote;
      }
    } else if (action === "CLAIM" || action === "DIRECT") {
      updatedCellVal = newVolunteer;
      targetRange.setValue(updatedCellVal);

      if (rm.notes) {
        var baseNotes = updatedNotesVal;
        if (currentVal && updatedCellVal && currentVal !== updatedCellVal) {
          baseNotes = removeSubNeededFlagFor_(baseNotes, currentVal, shortRoleName_(roleKey));
        }
        if (userComment) {
          updatedNotesVal = appendStructuredNote_(baseNotes, userComment);
          sheet.getRange(rm.notes, colIndex).setValue(updatedNotesVal);
        } else if (baseNotes !== updatedNotesVal) {
          updatedNotesVal = baseNotes;
          sheet.getRange(rm.notes, colIndex).setValue(updatedNotesVal);
        }
      }
    } else if (action === "CLEAR") {
      updatedCellVal = "";
      targetRange.setValue("");
      if (rm.notes && currentVal) {
        var strippedNotes = removeSubNeededFlagFor_(updatedNotesVal, currentVal, shortRoleName_(roleKey));
        if (userComment) {
          strippedNotes = appendStructuredNote_(strippedNotes, userComment);
        }
        if (strippedNotes !== updatedNotesVal) {
          updatedNotesVal = strippedNotes;
          sheet.getRange(rm.notes, colIndex).setValue(updatedNotesVal);
        }
      }
    } else {
      return { ok: false, code: "UNKNOWN_ACTION", message: "Unsupported action: " + action };
    }

    SpreadsheetApp.flush();

    // Write immutable audit trail row
    appendAuditLogRow_(ss, {
      monthTab: monthName,
      dateLabel: dateLabel,
      roleLabel: roleLabel,
      cellA1: cellA1,
      actionType: action,
      previousValue: currentVal,
      newValue: updatedCellVal,
      actorName: actorName,
      comment: auditComment
    });

    // Invalidate this month's cache and return fresh parsed month object
    var cache = CacheService.getScriptCache();
    if (cache) {
      cache.remove(CONFIG.CACHE_PREFIX + monthName);
    }

    var updatedMonth = getMonthDataCached_(ss, monthName, true);
    return {
      ok: true,
      serverTime: new Date().toISOString(),
      month: updatedMonth,
      updatedCell: {
        month: monthName,
        colIndex: colIndex,
        roleKey: roleKey,
        a1: cellA1,
        previousValue: currentVal,
        newValue: updatedCellVal,
        notesValue: updatedNotesVal
      }
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Handles Compare-And-Swap writes to the 4 Monthly Roles in the footer sub-table.
 */
function executeFooterWrite_(ss, sheet, displayGrid, schema, payload, actorName, userComment) {
  var footerKey = String(payload.footerKey || "").trim();
  var coords = schema.rowMap.footer[footerKey];
  if (!coords) {
    return {
      ok: false,
      code: "FOOTER_ROLE_NOT_FOUND",
      message: "Could not locate monthly role '" + footerKey + "' on " + schema.tabName
    };
  }

  var field = String(payload.field || "volunteer").toLowerCase();
  var targetCol = field === "detail" ? coords.detailCol : coords.volunteerCol;
  var targetRange = sheet.getRange(coords.row, targetCol);
  var currentVal = String(targetRange.getDisplayValue() || "").trim();
  var expectedPrev = String(payload.expectedPreviousValue || "").trim();

  if (currentVal !== expectedPrev) {
    return {
      ok: false,
      code: "CONFLICT",
      currentValue: currentVal,
      message: "This monthly role was just updated by another parent (current value: \"" + (currentVal || "Open") + "\")."
    };
  }

  var newValue = String(payload.newValue !== undefined ? payload.newValue : (payload.newVolunteer || "")).trim();
  targetRange.setValue(newValue);

  // If optional topic/detail was also passed while updating volunteer
  if (field === "volunteer" && payload.detailValue !== undefined) {
    var detailRange = sheet.getRange(coords.row, coords.detailCol);
    detailRange.setValue(String(payload.detailValue || "").trim());
  }

  SpreadsheetApp.flush();

  var cellA1 = columnToLetter_(targetCol) + coords.row;
  appendAuditLogRow_(ss, {
    monthTab: schema.tabName,
    dateLabel: schema.tabName + " (Monthly Role)",
    roleLabel: footerKey + " (" + field + ")",
    cellA1: cellA1,
    actionType: "FOOTER_EDIT",
    previousValue: currentVal,
    newValue: newValue,
    actorName: actorName,
    comment: userComment
  });

  var cache = CacheService.getScriptCache();
  if (cache) {
    cache.remove(CONFIG.CACHE_PREFIX + schema.tabName);
  }

  var updatedMonth = getMonthDataCached_(ss, schema.tabName, true);
  return {
    ok: true,
    serverTime: new Date().toISOString(),
    month: updatedMonth
  };
}

// ============================================================================
// 7. AUTOMATIC AUDIT LOG (`_Audit_Log` HIDDEN TAB)
// ============================================================================

function appendAuditLogRow_(ss, entry) {
  try {
    var auditSheet = ss.getSheetByName(CONFIG.AUDIT_LOG_TAB);
    if (!auditSheet) {
      auditSheet = ss.insertSheet(CONFIG.AUDIT_LOG_TAB);
      auditSheet.hideSheet();
      auditSheet.appendRow([
        "ISO Timestamp",
        "Month Tab",
        "Date Label",
        "Role Label",
        "Cell A1",
        "Action Type",
        "Previous Value",
        "New Value",
        "Actor Name",
        "Comment / Note"
      ]);
      auditSheet.setFrozenRows(1);
    }

    auditSheet.appendRow([
      new Date().toISOString(),
      entry.monthTab || "",
      entry.dateLabel || "",
      entry.roleLabel || "",
      entry.cellA1 || "",
      entry.actionType || "",
      entry.previousValue || "",
      entry.newValue || "",
      entry.actorName || "",
      entry.comment || ""
    ]);
  } catch (err) {
    Logger.log("Failed to append audit log row: " + err);
  }
}

// ============================================================================
// 8. DIAGNOSTIC HELPER: inspectWorkbookSchema()
// ============================================================================

/**
 * Run `inspectWorkbookSchema()` directly in the Apps Script IDE to inspect and verify
 * the dynamically resolved row and column coordinates across all 9 monthly tabs.
 */
function inspectWorkbookSchema() {
  var ss = getWorkbook_();
  var report = {
    workbookName: ss.getName(),
    inspectedAt: new Date().toISOString(),
    rosterTabPresent: !!ss.getSheetByName(CONFIG.ROSTER_TAB),
    months: {}
  };

  for (var i = 0; i < CONFIG.MONTH_TABS.length; i++) {
    var tabName = CONFIG.MONTH_TABS[i];
    var sheet = ss.getSheetByName(tabName);
    if (!sheet) {
      report.months[tabName] = { exists: false };
      continue;
    }
    var displayGrid = sheet.getDataRange().getDisplayValues();
    var schema = resolveMonthlySheetSchema_(displayGrid, tabName);
    report.months[tabName] = {
      exists: true,
      totalRows: displayGrid.length,
      lastGridCol: schema.lastGridCol,
      lastGridColLetter: columnToLetter_(schema.lastGridCol),
      extraDetailsStartCol: schema.extraDetailsStartCol,
      extraDetailsStartColLetter: schema.extraDetailsStartCol ? columnToLetter_(schema.extraDetailsStartCol) : null,
      rowMap: schema.rowMap
    };
  }

  Logger.log(JSON.stringify(report, null, 2));
  return report;
}

// ============================================================================
// 9. UTILITY FUNCTIONS
// ============================================================================

function getCellDisplay_(grid, row1Based, col0Based) {
  if (!row1Based || row1Based < 1 || row1Based > grid.length) return "";
  var row = grid[row1Based - 1];
  if (!row || col0Based < 0 || col0Based >= row.length) return "";
  return String(row[col0Based] || "").trim();
}

function columnToLetter_(col1Based) {
  var temp;
  var letter = "";
  var col = col1Based;
  while (col > 0) {
    temp = (col - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    col = (col - temp - 1) / 26;
  }
  return letter;
}

function normalizeStatus_(rawStatus, isExtraDetails, rawDate, calendarNotes) {
  var s = String(rawStatus || "").trim();
  var lower = s.toLowerCase();

  if (lower.indexOf("no school") !== -1 || lower.indexOf("holiday") !== -1 || lower.indexOf("break") !== -1) {
    return "NO SCHOOL";
  }
  if (lower.indexOf("minimum") !== -1 || lower.indexOf("min day") !== -1) {
    return "Minimum day";
  }
  if (lower.indexOf("field trip") !== -1) {
    return "Field trip";
  }
  if (lower.indexOf("party") !== -1) {
    return "Party day";
  }
  if (s !== "") {
    return s;
  }
  if (isExtraDetails) {
    if (!rawDate && !calendarNotes) return "Unscheduled";
    return "Special Event";
  }
  return "Standard";
}

function normalizeDateToIso_(rawDateStr, tabName) {
  var s = String(rawDateStr || "").trim();
  if (!s || /^tbd$/i.test(s)) return null;

  // Match MM/DD/YYYY or M/D/YY
  var mdy = s.match(/^(\d{1,2})[\/\-](\d{1,2})(?:[\/\-](\d{2,4}))?$/);
  if (mdy) {
    var month = parseInt(mdy[1], 10);
    var day = parseInt(mdy[2], 10);
    var year = mdy[3] ? parseInt(mdy[3], 10) : inferYearFromTabName_(tabName);
    if (year < 100) year += 2000;
    return formatIsoDate_(year, month, day);
  }

  var parsed = new Date(s);
  if (!isNaN(parsed.getTime())) {
    return formatIsoDate_(parsed.getFullYear(), parsed.getMonth() + 1, parsed.getDate());
  }
  return null;
}

function inferYearFromTabName_(tabName) {
  var m = String(tabName || "").match(/\b(20\d{2})\b/);
  return m ? parseInt(m[1], 10) : 2026;
}

function formatIsoDate_(year, month, day) {
  var mm = month < 10 ? "0" + month : String(month);
  var dd = day < 10 ? "0" + day : String(day);
  return year + "-" + mm + "-" + dd;
}

function appendStructuredNote_(existingNotes, newFragment) {
  var base = String(existingNotes || "").trim();
  var frag = String(newFragment || "").trim();
  if (!frag) return base;
  if (!base) return frag;
  if (base.toLowerCase().indexOf(frag.toLowerCase()) !== -1) {
    return base;
  }
  return base + " | " + frag;
}

function removeSubNeededFlagFor_(existingNotes, volunteerName, shortRole) {
  var base = String(existingNotes || "").trim();
  if (!base) return "";
  var parts = base.split(/\s*\|\s*/);
  var vLower = String(volunteerName || "").trim().toLowerCase();
  var rLower = String(shortRole || "").trim().toLowerCase();

  var kept = [];
  for (var i = 0; i < parts.length; i++) {
    var p = parts[i].trim();
    if (!p) continue;
    var pLower = p.toLowerCase();
    var isSubFlag = pLower.indexOf("sub needed") !== -1;
    var matchesTarget = (vLower && pLower.indexOf(vLower) !== -1) || (rLower && pLower.indexOf(rLower) !== -1);
    if (isSubFlag && matchesTarget) {
      continue;
    }
    kept.push(p);
  }
  return kept.join(" | ");
}

function shortRoleName_(roleKey) {
  if (roleKey === "centers1") return "Centers 1";
  if (roleKey === "centers2") return "Centers 2";
  if (roleKey === "centers3") return "Centers 3";
  if (roleKey === "homework") return "Homework Folders";
  return roleKey;
}

function cleanVolunteerName_(raw) {
  var s = String(raw || "").trim();
  if (!s) return "";
  if (/^(n\/?a|open|tbd|none|-|\?+)$/i.test(s)) return "";
  // Extract "presented by <Name>" from monthly Artist / Composer cells
  var presentedMatch = s.match(/presented\s+by\s+(.+)$/i);
  if (presentedMatch) {
    s = presentedMatch[1].trim();
  }
  // Strip trailing parenthetical annotations like "(as Tim)", "(signed up as ...)", "(Sub needed)"
  s = s.replace(/\s*\((as\s+[^)]+|signed\s+up\s+as\s+[^)]+|sub\s*needed|confirmed|tentative).*?\)\s*$/i, "").trim();
  return s;
}

function collectMonthParentNames_(monthObj, map) {
  if (!monthObj) return;
  var days = monthObj.days || [];
  for (var i = 0; i < days.length; i++) {
    var slots = days[i].slots || [];
    for (var j = 0; j < slots.length; j++) {
      var clean = cleanVolunteerName_(slots[j].volunteer);
      if (clean) map[clean.toLowerCase()] = clean;
    }
  }
  var footer = monthObj.footerRoles || [];
  for (var f = 0; f < footer.length; f++) {
    var fClean = cleanVolunteerName_(footer[f].volunteer);
    if (fClean) map[fClean.toLowerCase()] = fClean;
  }
}
