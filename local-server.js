#!/usr/bin/env node
/**
 * Local Live Google Sheet Dev Server (`local-server.js`)
 *
 * Serves `index.html` on http://localhost:8765 and proxies `/api` directly to the live
 * Google Sheet (`1W2KJJienqrxDjRQ8OkhPYT8wiML3bLqafSVbZZcB1U8`) via the `gsheets` CLI
 * while running the exact `Code.gs` parser and Compare-And-Swap mutation logic.
 *
 * Zero demo data — reads and writes directly against the live Google Sheet!
 */

const http = require("http");
const fs = require("fs");
const path = require("path");
const vm = require("vm");
const { execFileSync } = require("child_process");

const PORT = 8765;
const SPREADSHEET_ID = "1W2KJJienqrxDjRQ8OkhPYT8wiML3bLqafSVbZZcB1U8";
const ROOT_DIR = __dirname;
const CODE_GS_PATH = path.join(ROOT_DIR, "Code.gs");
const SNAPSHOT_CACHE_PATH = "/tmp/live_parsed_payload.json";

const ALL_TABS = [
  "Original Sign-Ups",
  "Oct 2026",
  "Nov 2026",
  "Dec 2026",
  "Jan 2027",
  "Feb 2027",
  "Mar 2027",
  "Apr 2027",
  "May 2027",
  "Jun 2027"
];

// In-memory cache of raw 2D tab grids so GETs are fast and writes update the exact tab
const rawGridsByTab = {};

function findGsheetsBinary() {
  const candidates = [
    process.env.GSHEETS_BIN,
    "/Users/hanwenyan/.gemini/antigravity/bin/gsheets",
    "gsheets"
  ].filter(Boolean);
  for (const c of candidates) {
    try {
      execFileSync(c, ["--version"], { stdio: "ignore" });
      return c;
    } catch (_) {}
  }
  return "gsheets";
}

const GSHEETS_BIN = findGsheetsBinary();

function padGrid(rawRows) {
  if (!Array.isArray(rawRows) || rawRows.length === 0) return [[""]];
  const maxCols = rawRows.reduce((m, r) => Math.max(m, Array.isArray(r) ? r.length : 0), 1);
  return rawRows.map(r => {
    const row = Array.isArray(r) ? r.map(v => String(v ?? "")) : [];
    while (row.length < maxCols) row.push("");
    return row;
  });
}

function readTabFromLiveSheet(tabName) {
  const range = tabName === "Original Sign-Ups" ? `${tabName}!A1:E30` : `${tabName}!A1:Z30`;
  const out = execFileSync(GSHEETS_BIN, ["readonly", "read", SPREADSHEET_ID, range, "--json"], {
    encoding: "utf8",
    timeout: 20000
  });
  const parsed = JSON.parse(out || "[]");
  const grid = padGrid(parsed);
  rawGridsByTab[tabName] = grid;
  return grid;
}

function writeCellToLiveSheet(tabName, row1Based, col1Based, value) {
  const colLetter = columnToLetter(col1Based);
  const a1 = `${tabName}!${colLetter}${row1Based}`;
  const strVal = String(value ?? "");
  if (strVal === "") {
    execFileSync(GSHEETS_BIN, ["mutate", "clear", SPREADSHEET_ID, a1], {
      encoding: "utf8",
      timeout: 20000
    });
  } else {
    // Escape commas or pipes if gsheets mutate write splits on them, or check how gsheets mutate write handles single cell
    const escaped = strVal.replace(/,/g, "\\,").replace(/\|/g, "\\|");
    execFileSync(GSHEETS_BIN, ["mutate", "write", SPREADSHEET_ID, a1, escaped], {
      encoding: "utf8",
      timeout: 20000
    });
  }

  // Update in-memory grid
  if (rawGridsByTab[tabName]) {
    const g = rawGridsByTab[tabName];
    while (g.length < row1Based) {
      g.push(new Array(g[0] ? g[0].length : col1Based).fill(""));
    }
    while (g[row1Based - 1].length < col1Based) {
      g[row1Based - 1].push("");
    }
    g[row1Based - 1][col1Based - 1] = strVal;
  }
}

function appendRowToAuditLog(rowValues) {
  // Best-effort local audit log tracking; in production Apps Script writes to _Audit_Log
  const logLine = JSON.stringify({ timestamp: new Date().toISOString(), rowValues }) + "\n";
  fs.appendFileSync("/tmp/cape_volunteers_audit.log", logLine);
}

function columnToLetter(col1Based) {
  let temp;
  let letter = "";
  let col = col1Based;
  while (col > 0) {
    temp = (col - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    col = (col - temp - 1) / 26;
  }
  return letter;
}

function createCodeGsRuntime() {
  const codeGs = fs.readFileSync(CODE_GS_PATH, "utf8");

  const makeSheetMock = (tabName) => ({
    getName: () => tabName,
    hideSheet: () => {},
    setFrozenRows: () => {},
    appendRow: (vals) => appendRowToAuditLog(vals),
    getDataRange: () => ({
      getDisplayValues: () => {
        if (!rawGridsByTab[tabName]) {
          readTabFromLiveSheet(tabName);
        }
        return rawGridsByTab[tabName];
      }
    }),
    getRange: (row1, col1) => ({
      getDisplayValue: () => {
        if (!rawGridsByTab[tabName]) {
          readTabFromLiveSheet(tabName);
        }
        const g = rawGridsByTab[tabName];
        return (g[row1 - 1] && g[row1 - 1][col1 - 1]) ? String(g[row1 - 1][col1 - 1]) : "";
      },
      setValue: (val) => {
        writeCellToLiveSheet(tabName, row1, col1, val);
      }
    })
  });

  const ssMock = {
    getName: () => "Mrs_White_Volunteer_Schedule_2026-27",
    getSheetByName: (tabName) => {
      if (tabName === "_Audit_Log") return makeSheetMock("_Audit_Log");
      if (!ALL_TABS.includes(tabName)) return null;
      return makeSheetMock(tabName);
    },
    insertSheet: (tabName) => makeSheetMock(tabName)
  };

  const ctx = {
    console,
    Logger: { log: (msg) => console.log("[Code.gs]", msg) },
    CacheService: { getScriptCache: () => null },
    LockService: {
      getDocumentLock: () => ({
        waitLock: () => {},
        releaseLock: () => {}
      })
    },
    SpreadsheetApp: {
      openById: () => ssMock,
      getActiveSpreadsheet: () => ssMock,
      flush: () => {}
    },
    ContentService: {
      MimeType: { JSON: "application/json" },
      createTextOutput: (str) => ({
        _content: str,
        setMimeType: function() { return this; },
        getContent: function() { return this._content; }
      })
    }
  };

  vm.createContext(ctx);
  vm.runInContext(codeGs, ctx);
  return ctx;
}

// Preload initial grids from /tmp/tab_*.json if present so startup is instant, then refresh on demand
const preloadMap = {
  "Original Sign-Ups": "/tmp/tab_roster.json",
  "Oct 2026": "/tmp/tab_oct.json",
  "Nov 2026": "/tmp/tab_nov.json",
  "Dec 2026": "/tmp/tab_dec.json",
  "Jan 2027": "/tmp/tab_jan.json",
  "Feb 2027": "/tmp/tab_feb.json",
  "Mar 2027": "/tmp/tab_mar.json",
  "Apr 2027": "/tmp/tab_apr.json",
  "May 2027": "/tmp/tab_may.json",
  "Jun 2027": "/tmp/tab_jun.json"
};

for (const [tabName, filePath] of Object.entries(preloadMap)) {
  if (fs.existsSync(filePath)) {
    try {
      rawGridsByTab[tabName] = padGrid(JSON.parse(fs.readFileSync(filePath, "utf8")));
    } catch (_) {}
  }
}

const server = http.createServer((req, res) => {
  const parsedUrl = new URL(req.url, `http://${req.headers.host || "localhost"}`);

  // CORS headers
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");

  if (req.method === "OPTIONS") {
    res.writeHead(204);
    res.end();
    return;
  }

  // API Endpoint (`/api`) executing `Code.gs` against the live Google Sheet
  if (parsedUrl.pathname === "/api") {
    const runtime = createCodeGsRuntime();

    if (req.method === "GET") {
      try {
        const parameter = Object.fromEntries(parsedUrl.searchParams.entries());
        if (parameter.nocache === "1" || parameter.nocache === "true") {
          // Re-fetch active tabs from Google Sheets
          if (parameter.month && ALL_TABS.includes(parameter.month)) {
            readTabFromLiveSheet(parameter.month);
          } else {
            ALL_TABS.forEach(t => readTabFromLiveSheet(t));
          }
        }
        const out = runtime.doGet({ parameter });
        const jsonStr = out.getContent();
        fs.writeFileSync(SNAPSHOT_CACHE_PATH, jsonStr);
        res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
        res.end(jsonStr);
      } catch (err) {
        console.error("GET /api error:", err);
        res.writeHead(500, { "Content-Type": "application/json; charset=utf-8" });
        res.end(JSON.stringify({ ok: false, message: String(err.message || err) }));
      }
      return;
    }

    if (req.method === "POST") {
      let body = "";
      req.on("data", chunk => { body += chunk; });
      req.on("end", () => {
        try {
          const payload = JSON.parse(body || "{}");
          // Always pull fresh tab state from the live sheet right before Compare-And-Swap
          if (payload.month && ALL_TABS.includes(payload.month)) {
            readTabFromLiveSheet(payload.month);
          }
          const out = runtime.doPost({ postData: { contents: body } });
          const jsonStr = out.getContent();
          res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
          res.end(jsonStr);
        } catch (err) {
          console.error("POST /api error:", err);
          res.writeHead(500, { "Content-Type": "application/json; charset=utf-8" });
          res.end(JSON.stringify({ ok: false, message: String(err.message || err) }));
        }
      });
      return;
    }
  }

  // Static files (`/` or `/index.html`)
  let filePath = parsedUrl.pathname === "/" ? "/index.html" : parsedUrl.pathname;
  const safePath = path.normalize(filePath).replace(/^(\.\.[\/\\])+/, "");
  const fullPath = path.join(ROOT_DIR, safePath);

  if (!fs.existsSync(fullPath) || fs.statSync(fullPath).isDirectory()) {
    res.writeHead(404, { "Content-Type": "text/plain" });
    res.end("Not found");
    return;
  }

  const ext = path.extname(fullPath).toLowerCase();
  const mimeTypes = {
    ".html": "text/html; charset=utf-8",
    ".js": "application/javascript; charset=utf-8",
    ".css": "text/css; charset=utf-8",
    ".json": "application/json; charset=utf-8"
  };

  res.writeHead(200, {
    "Content-Type": mimeTypes[ext] || "text/plain; charset=utf-8",
    "Cache-Control": "no-store"
  });
  res.end(fs.readFileSync(fullPath));
});

server.listen(PORT, () => {
  console.log(`Live Google Sheet server listening on http://localhost:${PORT}`);
});
