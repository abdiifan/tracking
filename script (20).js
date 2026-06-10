// =============================================================================
// PharmaTrack v2 — Pharmaceutical Inventory Management System
// =============================================================================
// Fixes applied vs v8:
//  BUG-1  Preview download now exports full filtered dataset (not 500-row slice)
//  BUG-2  Chain/circular reconciliation rules detected and blocked on save
//  BUG-3  pageFilters reset on new file upload so stale plant/MG never persists
//  BUG-4  "Already Expired" KPI now counts only stock-qty > 0 rows (matches table)
//  BUG-5  Target materials blocked from being selected as a new source
//  BUG-6  QC page no longer drops items with QC qty > 0 but zero ETB value
//  BUG-7  rpSetSelected chip close button uses addEventListener, not inline onclick
//  BUG-8  String expiry dates parsed as LOCAL midnight not UTC (timezone fix)
//  BUG-9  CSV tab-character cells now quoted correctly
//  BUG-10 groupBy empty-string bucket renamed to "(Blank)" for chart clarity
//  PERF-1 applyReconciliationToData result memoised; invalidated on file/rule change
//  PERF-2 File size warning before parse (>25 MB)
//  ROBUST localStorage schema validated on load; corrupt entries discarded
//  ROBUST Column header matching is now case-insensitive
//  ROBUST Total Qty removed from QTY_COLS scaling (was scaled then overwritten)
//  ROBUST Conversion factor stored with 9dp rounding to suppress float drift
// =============================================================================

// ── CONSTANTS ──────────────────────────────────────────────────────────────
const REQUIRED_COLUMNS = [
  "Material","Material Description","Plant","Plant Name",
  "Storage Location","Description of Storage Location",
  "Special Stock Type","Special Stock Type Description",
  "Unrestricted Stock","Stock in Quality Inspection","Blocked Stock",
  "Batch","Inventory Valuation Type","Material Group Name",
  "Shelf Life Expiration Date","Stock in Transit",
  "Value of Stock in Quality Inspection","Value of Stock in Transit",
  "Value of Unrestricted Stock",
];

const COLORWAY = ["#58a6ff","#3fb950","#d29922","#f85149","#a371f7","#79c0ff","#56d364","#e3b341","#ff7b72","#d2a8ff","#ffa657","#70d9a0"];

// ── FILTER STUBS (formerly filters.js — integrated, all exclusions disabled) ──
// Every function returns the "keep this row" value so 100 % of raw Excel data
// passes through. Original logic is preserved in comments for easy restoration.

function isNonMedicalCode(code) {   // was: true for NT* prefix / non-[1-4] codes
  return false;
}
function isMedicalCode(code) {      // was: true only for codes starting with 1-4
  return true;
}
function isNonMedicalGroup(groupName) { // was: true for NON TRADE, PROJECT STOCK, etc.
  return false;
}
function isProjectStockDescription(description) { // was: true if description === "PROJECT STOCK"
  return false;
}
function isExcludedStorageLocation(storageLoc) { // was: true for 21 hard-coded location codes
  return false;
}
// ─────────────────────────────────────────────────────────────────────────────
const PLOTLY_LAYOUT = {
  paper_bgcolor: "rgba(0,0,0,0)", plot_bgcolor: "rgba(0,0,0,0)",
  font: { family: "IBM Plex Sans", color: "#8b949e", size: 12 },
  xaxis: { gridcolor: "#21262d", zerolinecolor: "#21262d", tickfont: { color: "#8b949e" } },
  yaxis: { gridcolor: "#21262d", zerolinecolor: "#21262d", tickfont: { color: "#8b949e" } },
  legend: { bgcolor: "rgba(0,0,0,0)", font: { color: "#8b949e" } },
  margin: { l: 20, r: 20, t: 40, b: 40 },
  colorway: COLORWAY,
};
const PLOTLY_CONFIG = { displayModeBar: false, responsive: true };

// ── STATE ──────────────────────────────────────────────────────────────────
let rawDf  = [];
let filtDf = [];
let currentPage = "dashboard";

// Stock-in-Transit separate file state
let stockTransitRaw    = [];   // raw rows from the transit xlsx
let stFilterState      = { purDoc: "", supPlant: "" };  // filter state

// Page-level filter state — now arrays for multi-select support
const pageFilters = {
  dashboard: { plants: [], mgs: [] },
  transit:   { plants: [], mgs: [] },
  expiry:    { plants: [], mgs: [] },
  qc:        { plants: [], mgs: [] },
  branch:    { mgs: [] },
  flow:      { plants: [], mgs: [] },
};

// FIX BUG-3: reset all page filters when a new file is loaded so stale plant/MG
// values from the previous file can never produce a blank result set.
function resetPageFilters() {
  Object.keys(pageFilters).forEach(page => {
    pageFilters[page].plants = [];
    pageFilters[page].mgs    = [];
  });
}

// ── RECONCILIATION STATE ───────────────────────────────────────────────────
// Each entry: { sourceMaterial, sourceDesc, sourceUnit,
//               conversionFactor,
//               targetMaterial, targetDesc, targetUnit,
//               _builtin: true   ← marks rules that ship with the system }
// Qty rule: targetQty = sourceQty × conversionFactor
// Value rule: values are in ETB → summed as-is (no factor applied)
// Expiry rule: earliest (soonest) expiry date is kept (safest for pharma)
let reconcileGroups = [];

// ── BUILT-IN DEFAULT RECONCILIATION RULES ─────────────────────────────────
// These rules consolidate ASA pack-size variants into the canonical 20×10 pack
// (102-ACET-0102-02) for QC inspection and branch comparison reporting.
//
//  Source                 Factor   Target              Rationale
//  102-ACET-0102-01       × 0.5  → 102-ACET-0102-02   Microfined 10×10 → 20×10 equivalent
//  102-ACET-0102-04       × 0.5  → 102-ACET-0102-02   Enteric Coated 10×10 → 20×10 equivalent
//  102-ACET-0102-03       × 1.0  → 102-ACET-0102-02   Enteric Coated 200 → same qty basis
//
// The target code 102-ACET-0102-02 (ASA 81mg E.Coated 20×10) is the canonical form.
const DEFAULT_RECONCILE_RULES = [
  {
    sourceMaterial:   "102-ACET-0102-01",
    sourceDesc:       "Acetylsalicylic Acid - 81mg - Tablet (Microfined) of 10x10",
    sourceUnit:       "PAC",
    conversionFactor: 0.5,
    targetMaterial:   "102-ACET-0102-02",
    targetDesc:       "ASA - 81mg -Tablet (E. Coated) of 20x10",
    targetUnit:       "PAC",
    _builtin:         true,
  },
  {
    sourceMaterial:   "102-ACET-0102-04",
    sourceDesc:       "Acetylsalicylic Acid - 81mg - Tablet( Enteric Coated) of 10x10",
    sourceUnit:       "PAC",
    conversionFactor: 0.5,
    targetMaterial:   "102-ACET-0102-02",
    targetDesc:       "ASA - 81mg -Tablet (E. Coated) of 20x10",
    targetUnit:       "PAC",
    _builtin:         true,
  },
  {
    sourceMaterial:   "102-ACET-0102-03",
    sourceDesc:       "Acetylsalicylic Acid - 81mg - Tablet( Enteric Coated) of 200",
    sourceUnit:       "PAC",
    conversionFactor: 1,
    targetMaterial:   "102-ACET-0102-02",
    targetDesc:       "ASA - 81mg -Tablet (E. Coated) of 20x10",
    targetUnit:       "PAC",
    _builtin:         true,
  },
];

// ── RECONCILIATION CACHE (PERF-1) ─────────────────────────────────────────
// Cache the reconciled base so repeated page renders don't re-run the full
// merge loop. Invalidated when rawDf or reconcileGroups changes.
// Token-based invalidation: a short fingerprint of both arrays so that
// same-length mutations (delete rule A, add rule B) are correctly detected.
let _reconCache       = null;
let _reconCacheToken  = "";

function _makeReconToken() {
  // rawDf length + first/last material codes + reconcileGroups serialised codes
  const rawSig = rawDf.length + "|" +
    (rawDf[0]?.["Material"] || "") + "|" +
    (rawDf[rawDf.length - 1]?.["Material"] || "");
  const grpSig = reconcileGroups.map(g => g.sourceMaterial + ">" + g.targetMaterial + "x" + g.conversionFactor).join(",");
  return rawSig + "||" + grpSig;
}

function invalidateReconCache() {
  _reconCache      = null;
  _reconCacheToken = "";
}

function getReconciledBase() {
  const token = _makeReconToken();
  if (_reconCache !== null && _reconCacheToken === token) {
    return _reconCache;
  }
  _reconCache      = applyReconciliationToData(rawDf);
  _reconCacheToken = token;
  return _reconCache;
}

// ── FORMAT HELPERS ─────────────────────────────────────────────────────────
const fmtETB = v => `ETB ${Number(v || 0).toLocaleString("en-US", { maximumFractionDigits: 0 })}`;
const fmtQty = v => Number(v || 0).toLocaleString("en-US", { maximumFractionDigits: 0 });

// ── HTML ESCAPE (used by buildTable and reconciliation UI) ──────────────────
function escHtml(str) {
  return String(str ?? "").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
}

// ── MATERIAL COLUMN HELPERS ────────────────────────────────────────────────
// SAP sometimes stores the description text in the Material field when no
// numeric/structured code exists. We detect and flag this clearly.

// Returns true if the value looks like free-text description rather than a code.
function looksLikeDescription(val) {
  if (!val) return false;
  const s = String(val).trim();
  if (!s) return false;
  return s.includes(" ") || (s.length > 22 && !/^[\w\-\.\/]+$/.test(s));
}

// Gets the description sibling field from the row, handling both main-data
// rows (Material Description) and transit rows (_st_desc, desc).
function getSiblingDesc(row) {
  if (!row) return "";
  return String(
    row["Material Description"] ?? row["_st_desc"] ?? row["desc"] ?? ""
  ).trim();
}

// Gets the code sibling field — used by desc renderer to detect duplicates.
function getSiblingCode(row) {
  if (!row) return "";
  return String(
    row["Material"] ?? row["_st_material"] ?? row["mat"] ?? ""
  ).trim();
}

// ── renderMatCode(val, row) ────────────────────────────────────────────────
// Renders the "Material Code" cell.
//  • Normal code  → purple monospace
//  • Val looks like a description (has spaces / long) → amber "NAME" badge,
//    styled differently so it's obvious this isn't a structured code
function renderMatCode(val, row) {
  const s = escHtml(String(val ?? "").trim());
  if (!s) return '<span style="color:var(--dim)">—</span>';

  if (looksLikeDescription(val)) {
    // The "code" field actually contains a descriptive name
    return `<span class="mat-name-as-code" title="No structured code — SAP stores the name here">${s}</span>`
         + `<span class="mat-desc-badge" title="Material field contains a name, not a code">NAME</span>`;
  }
  return `<span class="col-mat-code">${s}</span>`;
}

// ── renderMatDesc(val, row) ────────────────────────────────────────────────
// Renders the "Material Description" cell.
//  • If description === code (SAP duplicate) → show italic muted "(same as code)"
//  • Otherwise → normal readable text
function renderMatDesc(val, row) {
  const desc = String(val ?? "").trim();
  const code = getSiblingCode(row);

  if (!desc) return '<span style="color:var(--dim)">—</span>';

  // Description is identical to the code field → don't repeat it
  if (desc === code) {
    return `<span class="mat-desc-same" title="Description is identical to the material code field">— same as code —</span>`;
  }

  return `<span class="col-mat-desc">${escHtml(desc)}</span>`;
}

// ── FIX BUG-8: Timezone-safe expiry date parser ────────────────────────────
// new Date("2024-03-15") is parsed as UTC midnight → in UTC+3 it appears as
// 2024-03-14 after 21:00 local time, causing day-off expiry errors.
// This parser treats yyyy-mm-dd strings as LOCAL midnight to avoid that shift.
function parseExpiryDate(d) {
  if (d instanceof Date) return isNaN(d.getTime()) ? null : d;
  if (!d) return null;
  const s = String(d).trim();
  // yyyy-mm-dd → local date (not UTC)
  const isoMatch = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (isoMatch) {
    const dt = new Date(+isoMatch[1], +isoMatch[2] - 1, +isoMatch[3]);
    return isNaN(dt.getTime()) ? null : dt;
  }
  // Fallback for other string formats
  const p = new Date(d);
  return isNaN(p.getTime()) ? null : p;
}

// ── LOAD & PROCESS EXCEL ───────────────────────────────────────────────────
function loadFile(file) {
  // FIX PERF-2: warn before parsing very large files
  if (file.size > 25 * 1024 * 1024) {
    if (!confirm(`This file is ${(file.size / 1024 / 1024).toFixed(1)} MB. Large files may take a few seconds to parse. Continue?`)) return;
  }

  const statusEl = document.getElementById("fileStatus");
  statusEl.style.display = "block";
  statusEl.innerHTML = `<div class="status-ok">⏳ LOADING…</div><div class="status-name">Parsing ${escHtml(file.name)}</div>`;

  const reader = new FileReader();
  reader.onload = e => {
    setTimeout(() => {
      try {
        const wb   = XLSX.read(new Uint8Array(e.target.result), { type: "array", cellDates: true });
        const ws   = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(ws, { defval: "" });
        if (!data.length) { showError("The uploaded file contains no data."); return; }

        const trimmed = data.map(row => {
          const r = {};
          for (const [k, v] of Object.entries(row)) r[k.trim()] = v;
          return r;
        });

        // FIX ROBUST: case-insensitive column header matching
        const colsLower = Object.keys(trimmed[0]).map(c => c.toLowerCase());
        const missing = REQUIRED_COLUMNS.filter(c => !colsLower.includes(c.toLowerCase()));
        if (missing.length) { showError(`Missing columns: ${missing.join(", ")}`); return; }

        // [ALL EXCLUSIONS DISABLED] — keep every row from the raw Excel file.
        // Original chain filtered out NT* codes, non-[1-4] materials, non-medical
        // groups, project stock descriptions, Special Stock Type "Q", and
        // excluded storage locations. All stubs in filters.js now return false.
        let df = trimmed.slice(); // shallow copy, no rows dropped

        const numCols = [
          "Unrestricted Stock","Stock in Quality Inspection","Blocked Stock","Stock in Transit",
          "Value of Stock in Quality Inspection","Value of Stock in Transit","Value of Unrestricted Stock",
        ];
        df.forEach(row => {
          numCols.forEach(c => { row[c] = parseFloat(row[c]) || 0; });
          // FIX BUG-8: use timezone-safe parser
          row._expiry = parseExpiryDate(row["Shelf Life Expiration Date"]);
          row["Total Value"] = row["Value of Unrestricted Stock"] + row["Value of Stock in Transit"] + row["Value of Stock in Quality Inspection"];
          row["Total Qty"]   = row["Unrestricted Stock"] + row["Stock in Transit"] + row["Stock in Quality Inspection"];
        });

        df = df.filter(r =>
          r["Unrestricted Stock"] > 0 ||
          r["Stock in Transit"] > 0 ||
          r["Stock in Quality Inspection"] > 0 ||
          r["Blocked Stock"] > 0
        );

        rawDf  = df;
        filtDf = df;

        // FIX BUG-3: clear stale page filters from the previous file
        resetPageFilters();
        // Invalidate reconciliation cache for the new dataset
        invalidateReconCache();

        showSuccess(file.name, df.length);
        clearError();
        hideLanding();
        document.getElementById("global-search-bar").style.display = "block";
        populateAllFilters();
        // Re-render home KPIs then switch to dashboard
        renderPage("home");
        renderPage(currentPage === "home" ? "dashboard" : currentPage);
      } catch (err) {
        showError(`Could not read Excel file: ${err.message}`);
      }
    }, 30);
  };
  reader.readAsArrayBuffer(file);
}

// ── MULTI-SELECT DROPDOWN BUILDER ─────────────────────────────────────────
// Creates a searchable checkbox dropdown inside .ms-wrap elements.
// wrapId = id of the .ms-wrap container
// items  = array of string values
// onLabel = optional function(selectedArr) → button label string
function buildMultiSelect(wrapId, ddId, items, placeholder) {
  const wrap = document.getElementById(wrapId);
  const dd   = document.getElementById(ddId);
  if (!wrap || !dd) return;

  const btn  = wrap.querySelector(".ms-btn");

  // Render options
  function renderItems(filter) {
    const filtered = filter ? items.filter(v => v.toLowerCase().includes(filter.toLowerCase())) : items;
    dd.querySelectorAll(".ms-item").forEach(el => el.remove());
    filtered.forEach(val => {
      const label = document.createElement("label");
      label.className = "ms-item";
      const cb = document.createElement("input");
      cb.type  = "checkbox";
      cb.value = val;
      // Restore checked state
      const page = wrap.dataset.page, key = wrap.dataset.key;
      if (page && key && pageFilters[page] && (pageFilters[page][key] || []).includes(val)) {
        cb.checked = true;
      }
      cb.addEventListener("change", updateLabel);
      label.appendChild(cb);
      label.appendChild(document.createTextNode(val));
      dd.appendChild(label);
    });
  }

  function updateLabel() {
    const checked = [...dd.querySelectorAll("input:checked")].map(c => c.value);
    if (checked.length === 0) {
      btn.innerHTML = `${escHtml(placeholder)} <span class="ms-arrow">▾</span>`;
      btn.classList.remove("ms-active");
    } else if (checked.length <= 2) {
      const names = checked.map(v => escHtml(v)).join(", ");
      btn.innerHTML = `<span class="ms-selected-names" title="${escHtml(checked.join(', '))}">${names}</span> <span class="ms-count-badge">${checked.length}</span> <span class="ms-arrow">▾</span>`;
      btn.classList.add("ms-active");
    } else {
      const badgeHtml = `<span class="ms-count-badge">${checked.length}</span>`;
      btn.innerHTML = `${escHtml(placeholder)} ${badgeHtml} <span class="ms-arrow">▾</span>`;
      btn.classList.add("ms-active");
    }
  }

  // Build search box + items
  dd.innerHTML = "";
  const searchInput = document.createElement("input");
  searchInput.className   = "ms-search";
  searchInput.placeholder = "Search…";
  searchInput.type        = "text";
  searchInput.addEventListener("input", e => renderItems(e.target.value));
  dd.appendChild(searchInput);
  renderItems("");

  // Toggle open/close
  btn.addEventListener("click", e => {
    e.stopPropagation();
    // Close all others first
    document.querySelectorAll(".ms-wrap.open").forEach(w => { if (w !== wrap) w.classList.remove("open"); });
    wrap.classList.toggle("open");
    if (wrap.classList.contains("open")) searchInput.focus();
  });

  // Expose refresh function on the wrap element
  wrap._refreshOptions = function(newItems) {
    renderItems(searchInput.value || "");
    updateLabel();
  };
  wrap._getSelected = function() {
    return [...dd.querySelectorAll("input:checked")].map(c => c.value);
  };
  wrap._clearSelected = function() {
    dd.querySelectorAll("input:checked").forEach(cb => { cb.checked = false; });
    updateLabel();
  };

  updateLabel();
}

// Close dropdowns when clicking outside
document.addEventListener("click", () => {
  document.querySelectorAll(".ms-wrap.open").forEach(w => w.classList.remove("open"));
});

// ── POPULATE FILTER DROPDOWNS ──────────────────────────────────────────────
function populateAllFilters() {
  const plants = [...new Set(rawDf.map(r => r["Plant Name"]))].filter(Boolean).sort();
  const mgs    = [...new Set(rawDf.map(r => r["Material Group Name"]))]
    .filter(Boolean)
    // [EXCLUSION DISABLED] — all MG names shown, including non-medical groups
    .sort();

  // Plant multi-selects
  const plantConfigs = [
    { wrapId:"ms-dash-plant",    ddId:"ms-dash-plant-dd",    page:"dashboard", key:"plants" },
    { wrapId:"ms-transit-plant", ddId:"ms-transit-plant-dd", page:"transit",   key:"plants" },
    { wrapId:"ms-expiry-plant",  ddId:"ms-expiry-plant-dd",  page:"expiry",    key:"plants" },
    { wrapId:"ms-qc-plant",      ddId:"ms-qc-plant-dd",      page:"qc",        key:"plants" },
    { wrapId:"ms-flow-plant",    ddId:"ms-flow-plant-dd",    page:"flow",      key:"plants" },
  ];
  plantConfigs.forEach(cfg => {
    const wrap = document.getElementById(cfg.wrapId);
    if (wrap) { wrap.dataset.page = cfg.page; wrap.dataset.key = "plants"; }
    buildMultiSelect(cfg.wrapId, cfg.ddId, plants, "All Plants");
  });

  // MG multi-selects
  const mgConfigs = [
    { wrapId:"ms-dash-mg",    ddId:"ms-dash-mg-dd",    page:"dashboard", key:"mgs" },
    { wrapId:"ms-transit-mg", ddId:"ms-transit-mg-dd", page:"transit",   key:"mgs" },
    { wrapId:"ms-expiry-mg",  ddId:"ms-expiry-mg-dd",  page:"expiry",    key:"mgs" },
    { wrapId:"ms-qc-mg",      ddId:"ms-qc-mg-dd",      page:"qc",        key:"mgs" },
    { wrapId:"ms-branch-mg",  ddId:"ms-branch-mg-dd",  page:"branch",    key:"mgs" },
    { wrapId:"ms-flow-mg",    ddId:"ms-flow-mg-dd",    page:"flow",      key:"mgs" },
  ];
  mgConfigs.forEach(cfg => {
    const wrap = document.getElementById(cfg.wrapId);
    if (wrap) { wrap.dataset.page = cfg.page; wrap.dataset.key = "mgs"; }
    buildMultiSelect(cfg.wrapId, cfg.ddId, mgs, "All Material Groups");
  });

  // Legacy multi-select for Data Preview (kept as-is)
  const plantSelLegacy  = ["filter-plant"];
  const mgSelLegacy     = ["filter-mg"];
  const mgNameSelLegacy = ["filter-mgname"];

  plantSelLegacy.forEach(id => {
    const el = document.getElementById(id); if (!el) return;
    el.innerHTML = plants.map(p => `<option value="${escHtml(p)}">${escHtml(p)}</option>`).join("");
  });
  mgSelLegacy.forEach(id => {
    const el = document.getElementById(id); if (!el) return;
    el.innerHTML = mgs.map(m => `<option value="${escHtml(m)}">${escHtml(m)}</option>`).join("");
  });
  mgNameSelLegacy.forEach(id => {
    const el = document.getElementById(id); if (!el) return;
    el.innerHTML = mgs.map(v => `<option value="${escHtml(v)}">${escHtml(v)}</option>`).join("");
  });
}

// ── APPLY PAGE FILTER ──────────────────────────────────────────────────────
// Uses the memoised reconciled base for performance.
// Also re-enforces base exclusion rules so excluded rows never appear on any page
// even if rawDf somehow contains them (e.g. after reconciliation merges).
function applyPageFilter(page) {
  const f    = pageFilters[page] || {};
  const base = getReconciledBase();
  const plants = f.plants || [];
  const mgs    = f.mgs    || [];
  return base.filter(r =>
    // [ALL EXCLUSIONS DISABLED] — base exclusion rules removed.
    // Only page-level plant / material group UI filters are applied.
    (!plants.length || plants.includes(r["Plant Name"])) &&
    (!mgs.length    || mgs.includes(r["Material Group Name"]))
  );
}

// ── UI HELPERS ─────────────────────────────────────────────────────────────
function showError(msg) {
  const el = document.getElementById("errorBanner");
  el.textContent = `⚠️ ${msg}`;
  el.style.display = "block";
}
function clearError() { document.getElementById("errorBanner").style.display = "none"; }
function showSuccess(name, n) {
  const el = document.getElementById("fileStatus");
  el.style.display = "block";
  el.innerHTML = `<div class="status-ok">✓ FILE LOADED</div><div class="status-name">${escHtml(name)} (${n.toLocaleString()} records)</div>`;
  document.getElementById("uploadBtnText").textContent = "📂 Change File";
}
function hideLanding() { document.getElementById("landingView").style.display = "none"; }

function kpiCard(label, value, sub, color) {
  return `<div class="kpi-card ${color}"><div class="kpi-label">${escHtml(label)}</div><div class="kpi-value">${escHtml(value)}</div><div class="kpi-sub">${escHtml(sub)}</div></div>`;
}
function setKpis(id, cards) {
  document.getElementById(id).innerHTML = cards.map(([l,v,s,c]) => kpiCard(l,v,s,c)).join("");
}

// ── GROUPBY HELPERS ────────────────────────────────────────────────────────
function groupBy(data, key, aggCols) {
  const map = {};
  data.forEach(row => {
    // FIX BUG-10: label blank keys clearly so charts don't show an invisible bar
    const k = row[key] || "(Blank)";
    if (!map[k]) { map[k] = { [key]: k }; aggCols.forEach(([c]) => { map[k][c] = 0; }); }
    aggCols.forEach(([c,src]) => { map[k][c] += row[src] || 0; });
  });
  return Object.values(map);
}
function sortBy(arr, key, asc=false) { return [...arr].sort((a,b) => asc ? a[key]-b[key] : b[key]-a[key]); }

// ── TABLE BUILDER ──────────────────────────────────────────────────────────
// Columns with raw:true may contain trusted HTML (badges etc.) — all others
// are escaped to prevent XSS from Excel data landing in the DOM.
function buildTable(rows, cols, rowClass, extraClass="") {
  if (!rows.length) return `<div class="alert-info">No data to display.</div>`;
  const thead = `<thead><tr>${cols.map(c => `<th>${escHtml(c.label)}</th>`).join("")}</tr></thead>`;
  const tbody = `<tbody>${rows.map(row => {
    const cls = rowClass ? rowClass(row) : "";
    return `<tr class="${cls}">${cols.map(c => {
      // Pass both the cell value AND the full row so fmt functions can cross-check sibling fields
      const raw     = c.fmt ? c.fmt(row[c.key], row) : (row[c.key] ?? "");
      const val     = c.raw ? raw : escHtml(String(raw));
      const cellCls = c.cellClass || "";
      return `<td class="${cellCls}">${val}</td>`;
    }).join("")}</tr>`;
  }).join("")}</tbody>`;
  return `<div class="tbl-wrap"><table class="${extraClass}">${thead}${tbody}</table></div>`;
}

// ── EXCEL DOWNLOAD ─────────────────────────────────────────────────────────
function downloadExcel(data, cols, filename) {
  const header = cols.map(c => c.label);
  const rows   = data.map(row => cols.map(c => {
    const v   = row[c.key];
    const raw = c.rawKey ? (row[c.rawKey] ?? v) : v;
    if (c.fmt) return (typeof raw === "number") ? raw : (raw ?? "");
    return raw ?? "";
  }));
  const wsData = [header, ...rows];
  const ws = XLSX.utils.aoa_to_sheet(wsData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Data");
  XLSX.writeFile(wb, filename);
}

// ── CSV DOWNLOAD ───────────────────────────────────────────────────────────
function downloadCSV(data, cols, filename) {
  const header = cols.map(c => c.label).join(",");
  const rows   = data.map(row => cols.map(c => {
    let v = c.rawKey ? (row[c.rawKey] ?? row[c.key] ?? "") : (row[c.key] ?? "");
    v = String(v ?? "");
    // CSV injection guard — prefix dangerous leading chars
    if (/^[=+\-@\t\r]/.test(v)) v = `'${v}`;
    // FIX BUG-9: also quote cells that contain tab characters
    if (v.includes(",") || v.includes('"') || v.includes("\n") || v.includes("\t")) {
      v = `"${v.replace(/"/g, '""')}"`;
    }
    return v;
  }).join(","));
  const blob = new Blob(["\uFEFF" + header + "\n" + rows.join("\n")], { type: "text/csv;charset=utf-8" });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement("a");
  a.href = url; a.download = filename; a.click();
  URL.revokeObjectURL(url);
}

// ── PLOTLY LAYOUT MERGE ────────────────────────────────────────────────────
function pl(extra={}) {
  return Object.assign({}, PLOTLY_LAYOUT, extra, {
    xaxis:  Object.assign({}, PLOTLY_LAYOUT.xaxis,  extra.xaxis  || {}),
    yaxis:  Object.assign({}, PLOTLY_LAYOUT.yaxis,  extra.yaxis  || {}),
    legend: Object.assign({}, PLOTLY_LAYOUT.legend, extra.legend || {}),
    margin: Object.assign({}, PLOTLY_LAYOUT.margin, extra.margin || {}),
  });
}

// ═══════════════════════════════════════════════════════════════════════════
// DASHBOARD
// ═══════════════════════════════════════════════════════════════════════════
function renderDashboard() {
  const df = applyPageFilter("dashboard");

  const totalVal   = df.reduce((s,r) => s + r["Total Value"], 0);
  const transitVal = df.reduce((s,r) => s + r["Value of Stock in Transit"], 0);
  const qcVal      = df.reduce((s,r) => s + r["Value of Stock in Quality Inspection"], 0);
  const availVal   = df.reduce((s,r) => s + r["Value of Unrestricted Stock"], 0);
  const totalQty   = df.reduce((s,r) => s + r["Total Qty"], 0);

  setKpis("dash-kpis", [
    ["Total Inventory Value",    fmtETB(totalVal),   `${fmtQty(totalQty)} total units`,      "blue"],
    ["Stock in Transit Value",   fmtETB(transitVal), `${fmtQty(df.reduce((s,r) => s+r["Stock in Transit"],0))} units`, "amber"],
    ["Value in QC",              fmtETB(qcVal),      `${fmtQty(df.reduce((s,r) => s+r["Stock in Quality Inspection"],0))} units`, "red"],
    ["Available (Unrestricted)", fmtETB(availVal),   `${fmtQty(df.reduce((s,r) => s+r["Unrestricted Stock"],0))} units`, "green"],
    ["Unique Materials",         new Set(df.map(r=>r["Material"])).size.toLocaleString(), `${new Set(df.map(r=>r["Plant"])).size} plants`, "purple"],
  ]);

  // Plant bar — dual axis qty+value
  const plantAgg = sortBy(groupBy(df, "Plant Name", [["val","Total Value"],["qty","Total Qty"]]), "val");
  Plotly.newPlot("chart-plant-val", [
    { type:"bar", name:"Value (ETB)", x:plantAgg.map(r=>r["Plant Name"]), y:plantAgg.map(r=>r.val), yaxis:"y", marker:{color:"#58a6ff"}, hovertemplate:"<b>%{x}</b><br>ETB %{y:,.0f}<extra></extra>" },
    { type:"scatter", mode:"lines+markers", name:"Quantity", x:plantAgg.map(r=>r["Plant Name"]), y:plantAgg.map(r=>r.qty), yaxis:"y2", marker:{color:"#3fb950",size:8}, line:{color:"#3fb950"}, hovertemplate:"<b>%{x}</b><br>Qty: %{y:,.0f}<extra></extra>" },
  ], pl({ height:280, margin:{l:20,r:60,t:20,b:80}, yaxis2:{overlaying:"y",side:"right",gridcolor:"transparent",tickfont:{color:"#3fb950"},title:{text:"Qty",font:{color:"#3fb950"}}}, barmode:"group" }), PLOTLY_CONFIG);

  // Material Group pie (by value)
  const mgAgg = sortBy(groupBy(df, "Material Group Name", [["val","Total Value"]]), "val").slice(0, 12);
  Plotly.newPlot("chart-cat-pie", [{
    type:"pie", labels:mgAgg.map(r=>r["Material Group Name"]), values:mgAgg.map(r=>r.val),
    hole:0.55, textposition:"outside", textinfo:"percent+label",
    marker:{colors:COLORWAY}, hovertemplate:"<b>%{label}</b><br>ETB %{value:,.0f}<br>%{percent}<extra></extra>",
  }], pl({ showlegend:false, height:280, margin:{l:10,r:10,t:30,b:10} }), PLOTLY_CONFIG);

  // Near-expiry by plant (within 6 months)
  const nearCutoff = new Date(); nearCutoff.setMonth(nearCutoff.getMonth() + 6);
  const nearToday  = new Date();
  const nearExpiry = df.filter(r =>
    r._expiry instanceof Date && !isNaN(r._expiry) &&
    r._expiry >= nearToday && r._expiry <= nearCutoff &&
    (r["Unrestricted Stock"] || 0) > 0
  );
  const nearByPlant = sortBy(
    groupBy(nearExpiry, "Plant Name", [["val","Value of Unrestricted Stock"],["qty","Unrestricted Stock"]]),
    "val"
  );
  if (nearByPlant.length) {
    Plotly.newPlot("chart-mg-bar", [
      { type:"bar", name:"Value at Risk (ETB)", x:nearByPlant.map(r=>r["Plant Name"]), y:nearByPlant.map(r=>r.val), yaxis:"y",  marker:{color:"#d29922"}, hovertemplate:"<b>%{x}</b><br>ETB %{y:,.0f}<extra></extra>" },
      { type:"scatter", mode:"lines+markers", name:"Qty at Risk", x:nearByPlant.map(r=>r["Plant Name"]), y:nearByPlant.map(r=>r.qty), yaxis:"y2", marker:{color:"#f85149",size:8}, line:{color:"#f85149"}, hovertemplate:"<b>%{x}</b><br>Qty: %{y:,.0f}<extra></extra>" },
    ], pl({ height:420, margin:{l:20,r:60,t:20,b:100}, barmode:"group",
      yaxis2:{overlaying:"y",side:"right",gridcolor:"transparent",tickfont:{color:"#f85149"},title:{text:"Qty",font:{color:"#f85149"}}}
    }), PLOTLY_CONFIG);
  } else {
    document.getElementById("chart-mg-bar").innerHTML = `<div class="alert-info" style="margin:1rem 0">✓ No near-expiry stock (within 6 months) with quantity on hand.</div>`;
  }

  // Download
  const dlCols = [
    {key:"Plant Name",         label:"Plant"},
    {key:"Material Group Name",label:"Material Group"},
    {key:"Total Value",        label:"Total Value (ETB)", fmt:fmtETB, rawKey:"Total Value"},
    {key:"Total Qty",          label:"Total Qty",         fmt:fmtQty, rawKey:"Total Qty"},
  ];
  const aggForDl = groupBy(df, "Plant Name", [["Total Value","Total Value"],["Total Qty","Total Qty"]]);
  document.getElementById("btn-dl-dash-xlsx").onclick = () => downloadExcel(aggForDl, dlCols, "dashboard_summary.xlsx");
  document.getElementById("btn-dl-dash-csv").onclick  = () => downloadCSV(aggForDl,   dlCols, "dashboard_summary.csv");
}

// ═══════════════════════════════════════════════════════════════════════════
// STOCK IN TRANSIT FILE LOADER
// Loads the separate stock-in-transit Excel (columns: Material, Material
// Description, Plant, Name 1, Purchasing Document, Item, Supplying Plant,
// Special Stock, Quantity, Base Unit of Measure, …).
// Applies the same isNonMedicalCode / isNonMedicalGroup filters as the
// main inventory file so only medical items appear.
// ═══════════════════════════════════════════════════════════════════════════
function loadTransitFile(file) {
  const statusEl = document.getElementById("transitFileStatus");
  statusEl.style.display = "block";
  statusEl.innerHTML = `<div class="status-ok">⏳ LOADING…</div><div class="status-name">Parsing ${escHtml(file.name)}</div>`;

  const reader = new FileReader();
  reader.onload = e => {
    setTimeout(() => {
      try {
        const wb   = XLSX.read(new Uint8Array(e.target.result), { type: "array" });
        const ws   = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(ws, { defval: "" });
        if (!data.length) {
          statusEl.innerHTML = `<div class="status-ok" style="color:var(--red)">✗ Empty file</div>`;
          return;
        }

        // Trim all column headers
        const trimmed = data.map(row => {
          const r = {};
          for (const [k, v] of Object.entries(row)) r[k.trim()] = v;
          return r;
        });

        // Normalise key column names (case-insensitive lookup)
        const colMap = {};
        if (trimmed.length) {
          Object.keys(trimmed[0]).forEach(k => { colMap[k.toLowerCase()] = k; });
        }
        const getCol = name => colMap[name.toLowerCase()] || name;

        // [EXCLUSION DISABLED] — keep all rows including NT* materials
        let df = trimmed.filter(r => {
          const mat = String(r[getCol("Material")] ?? "").trim();
          return !!mat; // only drop rows with a completely blank Material column
        });

        // Normalise Purchasing Document (may come as scientific notation from Excel)
        df = df.map(r => {
          const raw = String(r[getCol("Purchasing Document")] ?? "").trim();
          let purDoc = raw;
          if (/e/i.test(raw)) purDoc = String(Math.round(Number(raw)));
          return {
            "_st_material":     String(r[getCol("Material")]             ?? "").trim(),
            "_st_desc":         String(r[getCol("Material Description")] ?? "").trim(),
            "_st_plant":        String(r[getCol("Plant")]                ?? "").trim(),
            "_st_plantName":    String(r[getCol("Name 1")]               ?? r[getCol("Plant Name")] ?? "").trim(),
            "_st_purDoc":       purDoc,
            "_st_supPlant":     String(r[getCol("Supplying Plant")]      ?? "").trim(),
            "_st_qty":          parseFloat(r[getCol("Quantity")] ?? r[getCol("Order Quantity")] ?? 0) || 0,
            "_st_uom":          String(r[getCol("Base Unit of Measure")] ?? r[getCol("Order Unit")] ?? "").trim(),
            "_st_item":         String(r[getCol("Item")]                 ?? "").trim(),
            "_st_specialStock": String(r[getCol("Special Stock")]        ?? "").trim(),
          };
        });

        stockTransitRaw = df;
        stFilterState   = { purDoc: "", supPlant: "" };

        // Update status
        statusEl.innerHTML = `<div class="status-ok">✓ TRANSIT FILE LOADED</div><div class="status-name">${escHtml(file.name)} (${df.length.toLocaleString()} records)</div>`;
        document.getElementById("transitUploadBtnText").textContent = "📦 Change Transit File";

        // If currently on transit page, re-render to show the new section
        if (currentPage === "transit") renderStockTransitSection();
      } catch (err) {
        statusEl.innerHTML = `<div class="status-ok" style="color:var(--red)">✗ ${escHtml(err.message)}</div>`;
      }
    }, 30);
  };
  reader.readAsArrayBuffer(file);
}

// ─── Render the Stock in Transit detail section (lower half of Transit page) ─
function renderStockTransitSection() {
  const noFileEl  = document.getElementById("stock-transit-no-file");
  const contentEl = document.getElementById("stock-transit-content");

  if (!noFileEl || !contentEl) return; // elements not in DOM yet

  if (!stockTransitRaw.length) {
    noFileEl.style.display  = "block";
    contentEl.style.display = "none";
    return;
  }

  noFileEl.style.display  = "none";
  contentEl.style.display = "block";

  // Populate Purchasing Document filter dropdown
  const purDocs = [...new Set(stockTransitRaw.map(r => r._st_purDoc).filter(Boolean))].sort();
  const supPlants = [...new Set(stockTransitRaw.map(r => r._st_supPlant).filter(Boolean))].sort();

  const purDocEl   = document.getElementById("st-filter-pur-doc");
  const supPlantEl = document.getElementById("st-filter-sup-plant");

  purDocEl.innerHTML   = `<option value="">All Purchasing Documents</option>` +
    purDocs.map(d => `<option value="${escHtml(d)}"${stFilterState.purDoc === d ? " selected" : ""}>${escHtml(d)}</option>`).join("");
  supPlantEl.innerHTML = `<option value="">All Supplying Plants</option>` +
    supPlants.map(p => `<option value="${escHtml(p)}"${stFilterState.supPlant === p ? " selected" : ""}>${escHtml(p)}</option>`).join("");

  // Apply active filters from stFilterState
  let df = stockTransitRaw.filter(r =>
    (!stFilterState.purDoc   || r._st_purDoc   === stFilterState.purDoc) &&
    (!stFilterState.supPlant || r._st_supPlant === stFilterState.supPlant)
  );

  // KPIs
  const uniqMats    = new Set(df.map(r => r._st_material)).size;
  const uniqPurDocs = new Set(df.map(r => r._st_purDoc).filter(Boolean)).size;
  const uniqSup     = new Set(df.map(r => r._st_supPlant).filter(Boolean)).size;
  const totalQty    = df.reduce((s, r) => s + r._st_qty, 0);
  setKpis("st-kpis", [
    ["Total Records",          df.length.toLocaleString(),    "After filter",           "blue"],
    ["Unique Materials",       uniqMats.toLocaleString(),     "Distinct SKUs",          "green"],
    ["Purchasing Documents",   uniqPurDocs.toLocaleString(),  "Distinct POs/STO docs",  "amber"],
    ["Supplying Plants",       uniqSup.toLocaleString(),      "Source locations",       "purple"],
    ["Total Qty in Transit",   fmtQty(totalQty),              "Units",                  "blue"],
  ]);

  // Table columns
  const stCols = [
    { key: "_st_material",  label: "Material Code", fmt:(val,r)=>renderMatCode(val,r), raw:true, cellClass:"col-mat-code-wrap" },
    { key: "_st_desc",     label: "Material Description", fmt:(val,r)=>renderMatDesc(val,r), raw:true, cellClass:"col-mat-desc-wrap" },
    { key: "_st_plant",     label: "Plant Code" },
    { key: "_st_plantName", label: "Plant Name" },
    { key: "_st_purDoc",    label: "Purchasing Document" },
    { key: "_st_item",      label: "Item" },
    { key: "_st_supPlant",  label: "Supplying Plant" },
    { key: "_st_qty",       label: "Quantity", fmt: fmtQty, rawKey: "_st_qty", cellClass: "col-qty" },
    { key: "_st_uom",       label: "UOM" },
  ];

  document.getElementById("st-table-wrap").innerHTML = buildTable(df, stCols);
  document.getElementById("btn-dl-st-csv").onclick  = () => downloadCSV(df,   stCols, "stock_in_transit_detail.csv");
  document.getElementById("btn-dl-st-xlsx").onclick = () => downloadExcel(df, stCols, "stock_in_transit_detail.xlsx");
}

// ─── Lookup helper: get Purchasing Document(s) and Supplying Plant(s) ─────
// For a given material code + plant code, scans stockTransitRaw and returns
// deduplicated comma-separated values. Falls back to "—" when no transit file
// is loaded or no matching rows exist.
function getTransitInfo(material, plantCode) {
  if (!stockTransitRaw.length) return { purDoc: "—", supPlant: "—" };
  const mat  = String(material  || "").trim();
  const plt  = String(plantCode || "").trim().toUpperCase();
  const hits = stockTransitRaw.filter(r =>
    r._st_material === mat &&
    (plt === "" || r._st_plant.toUpperCase() === plt)
  );
  if (!hits.length) return { purDoc: "—", supPlant: "—" };
  const purDocs  = [...new Set(hits.map(r => r._st_purDoc).filter(Boolean))];
  const supPlants= [...new Set(hits.map(r => r._st_supPlant).filter(Boolean))];
  return {
    purDoc:   purDocs.length   ? purDocs.join(", ")   : "—",
    supPlant: supPlants.length ? supPlants.join(", ") : "—",
  };
}

// ═══════════════════════════════════════════════════════════════════════════
// TRANSIT
// ═══════════════════════════════════════════════════════════════════════════

// Holds the full transit rows (pre-built) so the search filter can re-slice them.
let _transitRowsCache = [];
let _transitColsCache = [];
let _ho01RowsCache    = [];

function renderTransit() {
  // [ALL EXCLUSIONS DISABLED] — show all transit rows regardless of material type
  const df = applyPageFilter("transit").filter(r =>
    r["Stock in Transit"] > 0 &&
    r["Value of Stock in Transit"] > 0
  );

  const totalTV = df.reduce((s,r) => s + r["Value of Stock in Transit"], 0);
  const totalTQ = df.reduce((s,r) => s + r["Stock in Transit"], 0);
  const uniqMat = new Set(df.map(r => r["Material"])).size;
  setKpis("transit-kpis", [
    ["Total Transit Value",        fmtETB(totalTV), "Across all plants",  "amber"],
    ["Total Transit Quantity",     fmtQty(totalTQ), "Units in movement",  "blue"],
    ["Unique Materials in Transit",String(uniqMat), "Distinct SKUs",      "green"],
  ]);

  const transitCols = [
    {key:"Material", label:"Material Code", fmt:(val,r)=>renderMatCode(val,r), raw:true, cellClass:"col-mat-code-wrap"},
    {key:"Material Description", label:"Material Description", fmt:(val,r)=>renderMatDesc(val,r), raw:true, cellClass:"col-mat-desc-wrap"},
    {key:"Material Group Name",       label:"Material Group"},
    {key:"Plant Name",                label:"Plant"},
    {key:"_purDoc",                   label:"Purchasing Document"},
    {key:"_supPlant",                 label:"Supplying Plant"},
    {key:"Stock in Transit",          label:"Transit Qty",       fmt:fmtQty, rawKey:"Stock in Transit",          cellClass:"col-qty"},
    {key:"Value of Stock in Transit", label:"Transit Value (ETB)",fmt:fmtETB, rawKey:"Value of Stock in Transit", cellClass:"col-val"},
    {key:"_status",                   label:"Status", raw:true},
  ];
  const transitRows = sortBy([...df], "Value of Stock in Transit").map(r => {
    const info = getTransitInfo(r["Material"], r["Plant"]);
    return {
      ...r,
      _purDoc:   info.purDoc,
      _supPlant: info.supPlant,
      _status: r["Value of Stock in Transit"] > 100000 ? "<span class='badge badge-red'>Critical</span>"
             : r["Value of Stock in Transit"] > 50000  ? "<span class='badge badge-amber'>High</span>"
             : r["Value of Stock in Transit"] > 10000  ? "<span class='badge badge-amber'>Medium</span>"
             : "<span class='badge badge-green'>Low</span>",
    };
  });

  // Cache rows for search filtering
  _transitRowsCache = transitRows;
  _transitColsCache = transitCols;

  // Wire chart
  if (df.length) {
    const plantAgg = sortBy(groupBy(df, "Plant Name", [["val","Value of Stock in Transit"],["qty","Stock in Transit"]]), "val");
    Plotly.newPlot("chart-transit-plant", [
      {type:"bar",  name:"Value (ETB)", x:plantAgg.map(r=>r["Plant Name"]), y:plantAgg.map(r=>r.val), yaxis:"y",  marker:{color:"#d29922"}, hovertemplate:"<b>%{x}</b><br>ETB %{y:,.0f}<extra></extra>"},
      {type:"scatter", mode:"lines+markers", name:"Qty", x:plantAgg.map(r=>r["Plant Name"]), y:plantAgg.map(r=>r.qty), yaxis:"y2", marker:{color:"#3fb950",size:8}, line:{color:"#3fb950"}, hovertemplate:"<b>%{x}</b><br>Qty: %{y:,.0f}<extra></extra>"},
    ], pl({height:280,margin:{l:20,r:60,t:20,b:80},yaxis2:{overlaying:"y",side:"right",gridcolor:"transparent",tickfont:{color:"#3fb950"}}}), PLOTLY_CONFIG);
  } else {
    document.getElementById("chart-transit-plant").innerHTML = "";
  }

  document.getElementById("btn-dl-transit").onclick      = () => downloadCSV(_transitRowsCache,   transitCols.slice(0,-1), "transit_analysis.csv");
  document.getElementById("btn-dl-transit-xlsx").onclick = () => downloadExcel(_transitRowsCache, transitCols.slice(0,-1), "transit_analysis.xlsx");

  // Show all filtered transit items directly (no search gate)
  document.getElementById("transit-table-wrap").innerHTML = transitRows.length
    ? buildTable(transitRows, transitCols)
    : `<div class="alert-info">No pharmaceutical transit items found.</div>`;
}

// ── Transit material search — filters main transit table ──────────────────
function renderTransitSearch() {
  const query = (document.getElementById("transit-search-input").value || "").trim().toLowerCase();
  const transitCols = _transitColsCache;

  if (!query) {
    document.getElementById("transit-search-results").innerHTML = "";
    document.getElementById("transit-table-wrap").innerHTML = _transitRowsCache.length
      ? buildTable(_transitRowsCache, transitCols)
      : `<div class="alert-info">No pharmaceutical transit items found.</div>`;
    return;
  }

  // Filter transit rows by search query
  const filtered = _transitRowsCache.filter(r => {
    const code = String(r["Material"] || "").toLowerCase();
    const desc = String(r["Material Description"] || "").toLowerCase();
    return code.includes(query) || desc.includes(query);
  });

  document.getElementById("transit-search-results").innerHTML =
    `<div style="font-size:0.78rem;color:var(--muted);margin-bottom:0.4rem">
      Found <b style="color:var(--text)">${filtered.length}</b> transit record(s) matching "<b style="color:var(--text)">${escHtml(query)}</b>"
    </div>`;

  document.getElementById("transit-table-wrap").innerHTML = filtered.length
    ? buildTable(filtered, transitCols)
    : `<div class="alert-info">No transit items match "<b>${escHtml(query)}</b>".</div>`;
}

function clearTransitSearch() {
  document.getElementById("transit-search-input").value = "";
  renderTransitSearch();
}

// ═══════════════════════════════════════════════════════════════════════════
// EXPIRY
// ═══════════════════════════════════════════════════════════════════════════
function renderExpiry() {
  const baseDf  = applyPageFilter("expiry");
  const months  = parseInt(document.querySelector('input[name="expWin"]:checked')?.value || 6);
  const today   = new Date();
  const cutoff  = new Date(today); cutoff.setMonth(cutoff.getMonth() + months);
  const valid   = baseDf.filter(r => r._expiry instanceof Date && !isNaN(r._expiry));

  const expiring     = valid.filter(r => r._expiry >= today && r._expiry <= cutoff && (r["Unrestricted Stock"]||0) > 0 && (r["Value of Unrestricted Stock"]||0) > 0);
  const expired      = valid.filter(r => r._expiry < today);
  // FIX BUG-4: filter zero-qty BEFORE the KPI count so KPI matches the table
  const expiredWithStock = expired.filter(r => (r["Unrestricted Stock"] || 0) > 0);
  const expiredZeroQty   = expired.length - expiredWithStock.length;

  setKpis("expiry-kpis", [
    ["Expiring in Window", String(expiring.length),       `Items within next ${months} months`,             "amber"],
    // FIX BUG-4: use expiredWithStock.length to match what the table shows
    ["Already Expired",   String(expiredWithStock.length),"Items with stock on hand requiring action",      "red"],
    ["At-Risk Value",     fmtETB(expiring.reduce((s,r) => s+r["Value of Unrestricted Stock"],0)),           "Unrestricted stock value","purple"],
    ["At-Risk Quantity",  fmtQty(expiring.reduce((s,r) => s+r["Unrestricted Stock"],0)),                   "Units expiring soon",     "amber"],
  ]);

  if (expiring.length) {
    const monthMap = {}, valMap = {};
    expiring.forEach(r => {
      const key = `${r._expiry.getFullYear()}-${String(r._expiry.getMonth()+1).padStart(2,"0")}`;
      monthMap[key] = (monthMap[key] || 0) + 1;
      valMap[key]   = (valMap[key]   || 0) + r["Value of Unrestricted Stock"];
    });
    const ms = Object.keys(monthMap).sort();
    Plotly.newPlot("chart-expiry-timeline", [
      {type:"bar",   name:"Items Count",   x:ms, y:ms.map(m=>monthMap[m]), marker:{color:"#d29922"}, hovertemplate:"<b>%{x}</b><br>%{y} items<extra></extra>"},
      {type:"scatter",mode:"lines+markers",name:"Value at Risk", x:ms, y:ms.map(m=>valMap[m]), yaxis:"y2", marker:{color:"#f85149",size:8}, line:{color:"#f85149"}, hovertemplate:"<b>%{x}</b><br>ETB %{y:,.0f}<extra></extra>"},
    ], pl({height:260,margin:{l:20,r:60,t:20,b:60},yaxis2:{overlaying:"y",side:"right",gridcolor:"transparent",tickfont:{color:"#f85149"}}}), PLOTLY_CONFIG);

    document.getElementById("chart-expiry-timeline").on("plotly_click", function(data) {
      const pt = data.points[0];
      const monthKey = pt.x;
      const [yr, mo] = monthKey.split("-").map(Number);
      const monthItems = expiring.filter(r => r._expiry.getFullYear() === yr && r._expiry.getMonth() + 1 === mo);
      const drillCols = [
        {key:"Material", label:"Material Code", fmt:(val,r)=>renderMatCode(val,r), raw:true, cellClass:"col-mat-code-wrap"},
        {key:"Material Description", label:"Material Description", fmt:(val,r)=>renderMatDesc(val,r), raw:true, cellClass:"col-mat-desc-wrap"},
        {key:"Material Group Name",         label:"Material Group"},
        {key:"Plant Name",                  label:"Plant"},
        {key:"Description of Storage Location", label:"Storage Location"},
        {key:"_expiryStr",                  label:"Expiry Date"},
        {key:"Unrestricted Stock",          label:"Qty",        fmt:fmtQty, rawKey:"Unrestricted Stock",       cellClass:"col-qty"},
        {key:"Value of Unrestricted Stock", label:"Value (ETB)",fmt:fmtETB, rawKey:"Value of Unrestricted Stock",cellClass:"col-val"},
        {key:"_daysLeft",                   label:"Days Left"},
      ];
      const drillRows = sortBy(
        monthItems.map(r => ({
          ...r,
          _expiryStr: r._expiry ? r._expiry.toISOString().slice(0,10) : "",
          _daysLeft:  r._expiry ? Math.floor((r._expiry - new Date()) / 86400000) : 9999,
        })),
        "_daysLeft", true
      );
      const totalVal   = monthItems.reduce((s,r) => s+r["Value of Unrestricted Stock"], 0);
      const totalQty   = monthItems.reduce((s,r) => s+r["Unrestricted Stock"], 0);
      const monthLabel = new Date(yr, mo-1, 1).toLocaleString("default", {month:"long", year:"numeric"});
      document.getElementById("expiry-drill-title").textContent = "📅 " + monthLabel;
      document.getElementById("expiry-drill-meta").textContent  = `${drillRows.length} items · ${fmtQty(totalQty)} units · ${fmtETB(totalVal)}`;
      document.getElementById("expiry-drill-table").innerHTML   = drillRows.length
        ? buildTable(drillRows, drillCols, r => r._daysLeft <= 30 ? "row-red" : r._daysLeft <= 90 ? "row-amber" : "")
        : '<div class="alert-info">No items for this month.</div>';
      const drillEl = document.getElementById("expiry-drilldown");
      drillEl.style.display = "block";
      drillEl.scrollIntoView({ behavior:"smooth", block:"nearest" });
      document.getElementById("expiry-drill-dl-csv").onclick  = () => downloadCSV(drillRows,  drillCols, `expiry_${monthKey}.csv`);
      document.getElementById("expiry-drill-dl-xlsx").onclick = () => downloadExcel(drillRows, drillCols, `expiry_${monthKey}.xlsx`);
    });
    document.getElementById("expiry-drill-close").onclick = () => {
      document.getElementById("expiry-drilldown").style.display = "none";
    };
  } else {
    document.getElementById("chart-expiry-timeline").innerHTML = "";
    document.getElementById("expiry-drilldown").style.display  = "none";
  }

  document.getElementById("expiry-table-wrap").innerHTML = "";

  if (expiredWithStock.length) {
    document.getElementById("expired-section").style.display = "block";
    const zeroNote = expiredZeroQty
      ? ` <span style="font-size:0.72rem;color:var(--muted);font-weight:400">(${expiredZeroQty} zero-qty records hidden)</span>`
      : "";
    document.getElementById("expired-header").innerHTML = `🔴 Already Expired Items (${expiredWithStock.length})${zeroNote}`;
    const expiredRows = expiredWithStock.map(r => ({...r, _expiryStr: r._expiry ? r._expiry.toISOString().slice(0,10) : ""}));
    document.getElementById("expired-table-wrap").innerHTML = buildTable(expiredRows, [
      {key:"Material", label:"Material Code", fmt:(val,r)=>renderMatCode(val,r), raw:true, cellClass:"col-mat-code-wrap"},
      {key:"Material Description", label:"Material Description", fmt:(val,r)=>renderMatDesc(val,r), raw:true, cellClass:"col-mat-desc-wrap"},
      {key:"Material Group Name",            label:"Material Group"},
      {key:"Plant Name",                     label:"Plant"},
      {key:"Description of Storage Location",label:"Storage Location"},
      {key:"_expiryStr",                     label:"Expiry Date"},
      {key:"Unrestricted Stock",             label:"Qty", fmt:fmtQty, rawKey:"Unrestricted Stock", cellClass:"col-qty"},
    ]);
    document.getElementById("btn-dl-expired").onclick = () => downloadCSV(expiredRows, [
      {key:"Material", label:"Material Code", fmt:(val,r)=>renderMatCode(val,r), raw:true, cellClass:"col-mat-code-wrap"},
      {key:"Material Description", label:"Material Description", fmt:(val,r)=>renderMatDesc(val,r), raw:true, cellClass:"col-mat-desc-wrap"},
      {key:"Plant Name",                     label:"Plant"},
      {key:"Description of Storage Location",label:"Storage Location"},
      {key:"_expiryStr",                     label:"Expiry Date"},
      {key:"Unrestricted Stock",             label:"Qty", rawKey:"Unrestricted Stock"},
    ], "expired_items.csv");
  } else {
    document.getElementById("expired-section").style.display = "none";
  }
}

// ── MATERIAL EXPIRY LOOKUP — filters the main expiry table ───────────────
function renderExpirySearch() {
  const query     = document.getElementById("expiry-search-input").value.trim().toLowerCase();
  const resultsEl = document.getElementById("expiry-search-results");

  if (!query) {
    resultsEl.innerHTML = "";
    // Hide main table and show prompt
    document.getElementById("expiry-table-wrap").innerHTML =
      `<div class="alert-info">🔍 Use the search box above to find and display expiry items.</div>`;
    document.getElementById("expired-section").style.display = "none";
    return;
  }
  if (!rawDf.length) { resultsEl.innerHTML = `<div class="alert-info">No data loaded yet.</div>`; return; }

  const today  = new Date();
  const baseDf = applyPageFilter("expiry");
  const matches = baseDf.filter(r => {
    const code     = String(r["Material"] || "").toLowerCase();
    const desc     = String(r["Material Description"] || "").toLowerCase();
    const hasStock = (r["Unrestricted Stock"] || 0) > 0 && (r["Value of Unrestricted Stock"] || 0) > 0;
    return hasStock && (code.includes(query) || desc.includes(query));
  });

  if (!matches.length) {
    resultsEl.innerHTML = `<div class="alert-info">No materials found matching "<b>${escHtml(query)}</b>".</div>`;
    document.getElementById("expiry-table-wrap").innerHTML = "";
    document.getElementById("expired-section").style.display = "none";
    return;
  }

  const annotated = matches.map(r => {
    const expiryStr = r._expiry ? r._expiry.toISOString().slice(0,10) : "—";
    let daysLeft = null, statusLabel = "No Expiry Date", statusClass = "";
    if (r._expiry instanceof Date && !isNaN(r._expiry)) {
      daysLeft = Math.floor((r._expiry - today) / 86400000);
      if      (daysLeft < 0)   { statusLabel = `Expired ${Math.abs(daysLeft)}d ago`; statusClass = "row-red";   }
      else if (daysLeft <= 30)  { statusLabel = `${daysLeft}d left`;                  statusClass = "row-red";   }
      else if (daysLeft <= 180) { statusLabel = `${daysLeft}d left`;                  statusClass = "row-amber"; }
      else                      { statusLabel = `${daysLeft}d left`;                  statusClass = "";          }
    }
    return { ...r, _expiryStr: expiryStr, _daysLeft: daysLeft ?? 99999, _statusLabel: statusLabel, _statusClass: statusClass };
  });

  const sorted     = annotated.sort((a,b) => a._daysLeft - b._daysLeft);
  const uniqueMats = [...new Set(sorted.map(r => r["Material"]))];
  const summary    = `<div style="font-size:0.78rem;color:var(--muted);margin-bottom:0.5rem">
    Found <b style="color:var(--text)">${sorted.length}</b> batch/location record(s) across
    <b style="color:var(--text)">${uniqueMats.length}</b> material code(s)
  </div>`;

  const cols = [
    {key:"Material", label:"Material Code", fmt:(val,r)=>renderMatCode(val,r), raw:true, cellClass:"col-mat-code-wrap"},
    {key:"Material Description", label:"Material Description", fmt:(val,r)=>renderMatDesc(val,r), raw:true, cellClass:"col-mat-desc-wrap"},
    {key:"Plant Name",                     label:"Plant"},
    {key:"Description of Storage Location",label:"Storage Location"},
    {key:"Batch",                          label:"Batch"},
    {key:"_expiryStr",                     label:"Expiry Date"},
    {key:"_statusLabel",                   label:"Status"},
    {key:"Unrestricted Stock",             label:"Avail Qty",   fmt:fmtQty, rawKey:"Unrestricted Stock",          cellClass:"col-qty"},
    {key:"Value of Unrestricted Stock",    label:"Value (ETB)", fmt:fmtETB, rawKey:"Value of Unrestricted Stock", cellClass:"col-val"},
  ];

  // Show summary in the lookup result area AND render the main table below
  resultsEl.innerHTML = summary;
  document.getElementById("expiry-table-wrap").innerHTML = buildTable(sorted, cols, r => r._statusClass);

  // Also show the expired-items sub-section if any expired results exist
  const expiredRows = sorted.filter(r => r._daysLeft < 0);
  if (expiredRows.length) {
    document.getElementById("expired-section").style.display = "block";
    document.getElementById("expired-header").innerHTML = `🔴 Already Expired Items (${expiredRows.length})`;
    document.getElementById("expired-table-wrap").innerHTML = buildTable(expiredRows, cols, r => r._statusClass);
  } else {
    document.getElementById("expired-section").style.display = "none";
  }
}

function clearExpirySearch() {
  document.getElementById("expiry-search-input").value = "";
  document.getElementById("expiry-search-results").innerHTML = "";
  document.getElementById("expiry-table-wrap").innerHTML =
    `<div class="alert-info">🔍 Use the search box above to find and display expiry items.</div>`;
  document.getElementById("expired-section").style.display = "none";
}

// ── MATERIAL QC LOOKUP ────────────────────────────────────────────────────
function renderQCSearch() {
  const query     = document.getElementById("qc-search-input").value.trim().toLowerCase();
  const resultsEl = document.getElementById("qc-search-results");
  if (!query) { resultsEl.innerHTML = ""; return; }
  if (!rawDf.length) { resultsEl.innerHTML = `<div class="alert-info">No data loaded yet.</div>`; return; }

  // Aggregate by material so source codes are consolidated into their target code
  // before searching — ensures searching "102-ACET-0102-01" returns the consolidated row
  const baseDf = aggregateByMaterial(
    applyPageFilter("qc").filter(r => (r["Stock in Quality Inspection"] || 0) > 0)
  );

  // Also allow searching by any original source code (maps to target description)
  const srcToTarget = {};
  reconcileGroups.forEach(g => { srcToTarget[g.sourceMaterial.toLowerCase()] = g.targetMaterial; });

  const matches = baseDf.filter(r => {
    const code  = String(r["Material"] || "").toLowerCase();
    const desc  = String(r["Material Description"] || "").toLowerCase();
    // Also match if the query matches any source code that points to this target
    const srcMatch = r._sourceBreakdown?.some(s => s.origCode.toLowerCase().includes(query));
    return (code.includes(query) || desc.includes(query) || srcMatch);
  });

  if (!matches.length) {
    resultsEl.innerHTML = `<div class="alert-info">No QC stock found matching "<b>${escHtml(query)}</b>".</div>`;
    return;
  }

  const today = new Date();
  const annotated = matches.map(r => {
    const expiryStr = r._expiry ? r._expiry.toISOString().slice(0,10) : "—";
    let daysLeft = null, statusLabel = "No Expiry Date", statusClass = "";
    if (r._expiry instanceof Date && !isNaN(r._expiry)) {
      daysLeft = Math.floor((r._expiry - today) / 86400000);
      if      (daysLeft < 0)   { statusLabel = `Expired ${Math.abs(daysLeft)}d ago`; statusClass = "row-red";   }
      else if (daysLeft <= 30)  { statusLabel = `${daysLeft}d left`;                  statusClass = "row-red";   }
      else if (daysLeft <= 180) { statusLabel = `${daysLeft}d left`;                  statusClass = "row-amber"; }
      else                      { statusLabel = `${daysLeft}d left`;                  statusClass = "";          }
    }

    // Build "Reconciled From" cell
    let reconSrcHtml = `<span style="color:var(--dim);font-size:0.68rem">—</span>`;
    if (r._isReconciled && r._sourceBreakdown && r._sourceBreakdown.length > 1) {
      reconSrcHtml = r._sourceBreakdown.map(s => {
        const rawQty  = fmtQty(s.qcQty);
        const convQty = fmtQty(s.convertedQCQty);
        const isTarget = s.origCode === r["Material"];
        if (isTarget) {
          return `<span style="font-size:0.68rem;color:var(--green)">` +
                 `<span class="col-mat-code" style="font-size:0.65rem">${escHtml(s.origCode)}</span> = ${convQty}</span>`;
        }
        return `<span style="font-size:0.68rem;color:var(--muted)">` +
               `<span class="col-mat-code" style="font-size:0.65rem">${escHtml(s.origCode)}</span>` +
               ` ${rawQty} ×${s.convFactor} → <b style="color:var(--text)">${convQty}</b></span>`;
      }).join(`<span style="color:var(--dim);margin:0 2px">+</span>`);
      reconSrcHtml = `<div style="display:flex;flex-wrap:wrap;gap:3px;align-items:center">${reconSrcHtml}</div>`;
    }

    return { ...r, _expiryStr: expiryStr, _daysLeft: daysLeft ?? 99999, _statusLabel: statusLabel, _statusClass: statusClass, _reconSrcCol: reconSrcHtml };
  });

  const sorted     = annotated.sort((a,b) => a._daysLeft - b._daysLeft);
  const uniqueMats = [...new Set(sorted.map(r => r["Material"]))];
  const summary    = `<div style="font-size:0.78rem;color:var(--muted);margin-bottom:0.5rem">
    Found <b style="color:var(--text)">${sorted.length}</b> QC material(s)
    (consolidated to <b style="color:var(--text)">${uniqueMats.length}</b> canonical code(s))
  </div>`;

  const cols = [
    {key:"Material", label:"Material Code", fmt:(val,r)=>renderMatCode(val,r), raw:true, cellClass:"col-mat-code-wrap"},
    {key:"Material Description", label:"Material Description", fmt:(val,r)=>renderMatDesc(val,r), raw:true, cellClass:"col-mat-desc-wrap"},
    {key:"_reconSrcCol",                          label:"Reconciled From", raw:true},
    {key:"_expiryStr",                            label:"Shelf Life Expiry"},
    {key:"_statusLabel",                          label:"Expiry Status"},
    {key:"Stock in Quality Inspection",           label:"QC Qty (Consolidated)", fmt:(val,r)=>`${fmtQty(val)}${renderReconBadge(r,"qc")}`, raw:true, rawKey:"Stock in Quality Inspection", cellClass:"col-qty"},
    {key:"Value of Stock in Quality Inspection",  label:"QC Value (ETB)", fmt:fmtETB, rawKey:"Value of Stock in Quality Inspection", cellClass:"col-val"},
  ];
  resultsEl.innerHTML = summary + buildTable(sorted, cols, r => r._statusClass);
}

function clearQCSearch() {
  document.getElementById("qc-search-input").value = "";
  document.getElementById("qc-search-results").innerHTML = "";
}

// ── MATERIAL FLOW LOOKUP ──────────────────────────────────────────────────
function renderFlowSearch() {
  const query     = document.getElementById("flow-search-input").value.trim().toLowerCase();
  const resultsEl = document.getElementById("flow-search-results");
  if (!query) { resultsEl.innerHTML = ""; return; }
  if (!rawDf.length) { resultsEl.innerHTML = `<div class="alert-info">No data loaded yet.</div>`; return; }

  const baseDf = applyPageFilter("flow");
  const matches = baseDf.filter(r => {
    const code = String(r["Material"] || "").toLowerCase();
    const desc = String(r["Material Description"] || "").toLowerCase();
    return code.includes(query) || desc.includes(query);
  });

  if (!matches.length) {
    resultsEl.innerHTML = `<div class="alert-info">No materials found matching "<b>${escHtml(query)}</b>".</div>`;
    return;
  }

  const uniqueMats = [...new Set(matches.map(r => r["Material"]))];
  const summary = `<div style="font-size:0.78rem;color:var(--muted);margin-bottom:0.5rem">
    Found <b style="color:var(--text)">${matches.length}</b> record(s) across
    <b style="color:var(--text)">${uniqueMats.length}</b> material code(s)
  </div>`;

  const cols = [
    {key:"Material", label:"Material Code", fmt:(val,r)=>renderMatCode(val,r), raw:true, cellClass:"col-mat-code-wrap"},
    {key:"Material Description", label:"Material Description", fmt:(val,r)=>renderMatDesc(val,r), raw:true, cellClass:"col-mat-desc-wrap"},
    {key:"Plant Name",                        label:"Plant"},
    {key:"Material Group Name",               label:"Material Group"},
    {key:"Unrestricted Stock",                label:"Avail Qty",      fmt:fmtQty, rawKey:"Unrestricted Stock",          cellClass:"col-qty"},
    {key:"Stock in Transit",                  label:"Transit Qty",    fmt:fmtQty, rawKey:"Stock in Transit",            cellClass:"col-qty"},
    {key:"Stock in Quality Inspection",       label:"QC Qty",         fmt:fmtQty, rawKey:"Stock in Quality Inspection", cellClass:"col-qty"},
    {key:"Value of Unrestricted Stock",       label:"Avail Value (ETB)", fmt:fmtETB, rawKey:"Value of Unrestricted Stock", cellClass:"col-val"},
    {key:"Value of Stock in Transit",         label:"Transit Value (ETB)", fmt:fmtETB, rawKey:"Value of Stock in Transit", cellClass:"col-val"},
  ];
  resultsEl.innerHTML = summary + buildTable(matches, cols);
}

function clearFlowSearch() {
  document.getElementById("flow-search-input").value = "";
  document.getElementById("flow-search-results").innerHTML = "";
}

// ═══════════════════════════════════════════════════════════════════════════
// QC
// ═══════════════════════════════════════════════════════════════════════════
function renderQC() {
  // FIX BUG-6: removed "&& r["Value of Stock in Quality Inspection"] > 0"
  // SAP sometimes records QC qty > 0 with zero ETB value (non-valuated batches,
  // consignment stock) — these must still appear for physical count audits.
  // RECONCILIATION: aggregate all source codes into their target canonical code
  // so each material appears exactly once (e.g. three ASA variants → one total).
  const rawFiltered = applyPageFilter("qc").filter(r => r["Stock in Quality Inspection"] > 0);
  const df          = aggregateByMaterial(rawFiltered).filter(r => r["Stock in Quality Inspection"] > 0);

  const totalQCVal = df.reduce((s,r) => s + r["Value of Stock in Quality Inspection"], 0);
  const totalQCQty = df.reduce((s,r) => s + r["Stock in Quality Inspection"], 0);
  setKpis("qc-kpis", [
    ["Total Value in QC", fmtETB(totalQCVal), "Across all plants",      "red"],
    ["Total QC Quantity", fmtQty(totalQCQty), "Units under inspection", "amber"],
    ["Unique Materials",  String(new Set(df.map(r=>r["Material"])).size),"Distinct SKUs","blue"],
  ]);

  if (!df.length) { document.getElementById("qc-table-wrap").innerHTML = `<div class="alert-info">✓ No items in quality inspection.</div>`; return; }

  const plantQC = sortBy(groupBy(rawFiltered, "Plant Name", [["val","Value of Stock in Quality Inspection"],["qty","Stock in Quality Inspection"]]), "val");
  Plotly.newPlot("chart-qc-plant", [
    {type:"bar",     name:"Value (ETB)", x:plantQC.map(r=>r["Plant Name"]), y:plantQC.map(r=>r.val), yaxis:"y",  marker:{color:"#f85149"}, hovertemplate:"<b>%{x}</b><br>ETB %{y:,.0f}<extra></extra>"},
    {type:"scatter", mode:"lines+markers", name:"Qty", x:plantQC.map(r=>r["Plant Name"]), y:plantQC.map(r=>r.qty), yaxis:"y2", marker:{color:"#3fb950",size:8}, line:{color:"#3fb950"}, hovertemplate:"<b>%{x}</b><br>Qty: %{y:,.0f}<extra></extra>"},
  ], pl({height:280,margin:{l:20,r:60,t:20,b:80},yaxis2:{overlaying:"y",side:"right",gridcolor:"transparent",tickfont:{color:"#3fb950"}}}), PLOTLY_CONFIG);

  const qcCols = [
    {key:"Material", label:"Material Code", fmt:(val,r)=>renderMatCode(val,r), raw:true, cellClass:"col-mat-code-wrap"},
    {key:"Material Description", label:"Material Description", fmt:(val,r)=>renderMatDesc(val,r), raw:true, cellClass:"col-mat-desc-wrap"},
    {key:"Material Group Name",                  label:"Material Group"},
    {key:"_reconSrcCol",                         label:"Reconciled From", raw:true},
    {key:"_expiryStr",                           label:"Shelf Life Expiry"},
    {key:"Stock in Quality Inspection",          label:"QC Qty",        fmt:(val,r)=>`${fmtQty(val)}${renderReconBadge(r,"qc")}`, raw:true, rawKey:"Stock in Quality Inspection", cellClass:"col-qty"},
    {key:"Value of Stock in Quality Inspection", label:"QC Value (ETB)",fmt:fmtETB, rawKey:"Value of Stock in Quality Inspection", cellClass:"col-val"},
  ];

  const qcRows = sortBy(
    [...df].map(r => {
      let reconSrcHtml = "";
      if (r._isReconciled && r._sourceBreakdown && r._sourceBreakdown.length > 1) {
        reconSrcHtml = r._sourceBreakdown.map(s => {
          const rawQty  = fmtQty(s.qcQty);
          const convQty = fmtQty(s.convertedQCQty);
          const isTarget = s.origCode === r["Material"];
          if (isTarget) {
            return `<span style="font-size:0.7rem;color:var(--green)">` +
                   `<span class="col-mat-code" style="font-size:0.65rem">${escHtml(s.origCode)}</span>` +
                   ` = ${convQty}</span>`;
          }
          return `<span style="font-size:0.7rem;color:var(--muted)">` +
                 `<span class="col-mat-code" style="font-size:0.65rem">${escHtml(s.origCode)}</span>` +
                 ` ${rawQty} ×${s.convFactor} → <b style="color:var(--text)">${convQty}</b></span>`;
        }).join(`<span style="color:var(--muted);margin:0 2px">+</span>`);
        reconSrcHtml = `<div style="display:flex;flex-wrap:wrap;gap:3px;align-items:center;font-size:0.68rem">${reconSrcHtml}</div>`;
      } else {
        reconSrcHtml = `<span style="color:var(--dim);font-size:0.68rem">—</span>`;
      }
      return {
        ...r,
        _expiryStr:   r._expiry ? r._expiry.toISOString().slice(0,10) : "",
        _reconSrcCol: reconSrcHtml,
      };
    }),
    "Value of Stock in Quality Inspection"
  );
  document.getElementById("qc-table-wrap").innerHTML = buildTable(qcRows, qcCols, r => r["Value of Stock in Quality Inspection"] > 10000 ? "row-red" : "");
  document.getElementById("btn-dl-qc").onclick      = () => downloadCSV(qcRows,   qcCols, "qc_inspection.csv");
  document.getElementById("btn-dl-qc-xlsx").onclick = () => downloadExcel(qcRows, qcCols, "qc_inspection.xlsx");
}

// ═══════════════════════════════════════════════════════════════════════════
// BRANCH COMPARISON
// ═══════════════════════════════════════════════════════════════════════════
function renderBranch() {
  const baseDf = applyPageFilter("branch");
  // Aggregate by material so source codes consolidate into canonical target codes
  // across all plants before the branch-level breakdown is built.
  const df = aggregateByMaterial(baseDf);

  const plants = [...new Set(df.map(r => String(r["Plant"]).toUpperCase()))];
  let centralCode, centralName;
  if (plants.includes("HO01")) {
    centralCode = "HO01";
    centralName = df.find(r => String(r["Plant"]).toUpperCase() === "HO01")?.["Plant Name"] || "HO01";
    document.getElementById("branch-central-info").style.display = "none";
  } else {
    const totals = {};
    df.forEach(r => { const p = r["Plant Name"]; totals[p] = (totals[p] || 0) + r["Total Value"]; });
    centralName = Object.entries(totals).sort((a,b) => b[1]-a[1])[0]?.[0] || "";
    document.getElementById("branch-central-info").style.display = "block";
    document.getElementById("branch-central-info").innerHTML = `ℹ️ HO01 not found — using <b>${escHtml(centralName)}</b> as central branch (highest inventory value).`;
  }

  const aggMap = {};
  df.forEach(r => {
    const k = r["Plant Name"];
    if (!aggMap[k]) aggMap[k] = {PlantName:k,Plant:r["Plant"],TotalValue:0,Unrestricted:0,Transit:0,QC:0,UnrestrictedQty:0,TransitQty:0,QCQty:0,Items:0};
    aggMap[k].TotalValue     += r["Total Value"];
    aggMap[k].Unrestricted   += r["Value of Unrestricted Stock"];
    aggMap[k].Transit        += r["Value of Stock in Transit"];
    aggMap[k].QC             += r["Value of Stock in Quality Inspection"];
    aggMap[k].UnrestrictedQty += r["Unrestricted Stock"];
    aggMap[k].TransitQty     += r["Stock in Transit"];
    aggMap[k].QCQty          += r["Stock in Quality Inspection"];
    aggMap[k].Items++;
  });
  const branchAgg = Object.values(aggMap);
  const others    = branchAgg.map(r => r.PlantName).filter(b => b !== centralName);

  const matPlantMap = {};
  df.forEach(r => {
    const mat = r["Material"], pln = r["Plant Name"];
    if (!matPlantMap[mat]) {
      matPlantMap[mat] = {
        desc:             r["Material Description"],
        group:            r["Material Group Name"],
        _isReconciled:    r._isReconciled    || false,
        _sourceBreakdown: r._sourceBreakdown || [],
      };
    }
    // Propagate reconciliation metadata from any row that carries it
    if (r._isReconciled && !matPlantMap[mat]._isReconciled) {
      matPlantMap[mat]._isReconciled    = true;
      matPlantMap[mat]._sourceBreakdown = r._sourceBreakdown || [];
    }
    if (!matPlantMap[mat][pln]) matPlantMap[mat][pln] = {Unrestricted:0,Transit:0,QC:0,TotalValue:0,TotalQty:0,UnrestrictedQty:0,TransitQty:0,QCQty:0};
    matPlantMap[mat][pln].Unrestricted    += r["Value of Unrestricted Stock"];
    matPlantMap[mat][pln].Transit         += r["Value of Stock in Transit"];
    matPlantMap[mat][pln].QC             += r["Value of Stock in Quality Inspection"];
    matPlantMap[mat][pln].TotalValue      += r["Total Value"];
    matPlantMap[mat][pln].TotalQty        += r["Total Qty"];
    matPlantMap[mat][pln].UnrestrictedQty += r["Unrestricted Stock"];
    matPlantMap[mat][pln].TransitQty      += r["Stock in Transit"];
    matPlantMap[mat][pln].QCQty           += r["Stock in Quality Inspection"];
  });

  const tabsHtml = `
    <div class="branch-tabs" id="branch-tabs">
      <button class="branch-tab active" data-tab="value">📊 Total Value &amp; Quantity</button>
      <button class="branch-tab" data-tab="material">🔬 Line-Item (Material Across Branches)</button>
    </div>
    <div id="branch-tab-value"></div>
    <div id="branch-tab-material" style="display:none"></div>`;
  document.getElementById("branch-tabs-wrap").innerHTML = tabsHtml;

  document.querySelectorAll(".branch-tab").forEach(btn => {
    btn.addEventListener("click", () => {
      document.querySelectorAll(".branch-tab").forEach(b => b.classList.remove("active"));
      btn.classList.add("active");
      const tab = btn.dataset.tab;
      document.getElementById("branch-tab-value").style.display    = tab === "value"    ? "block" : "none";
      document.getElementById("branch-tab-material").style.display = tab === "material" ? "block" : "none";
      if (tab === "material") renderMaterialTab();
    });
  });

  const sel = document.getElementById("branch-select");
  sel.innerHTML = "";
  others.forEach(b => {
    const opt = document.createElement("option");
    opt.value = b; opt.textContent = b; opt.selected = true;
    sel.appendChild(opt);
  });

  // ── TAB 1: Total Value ──
  function updateBranchCharts() {
    const selected = [...sel.selectedOptions].map(o => o.value);
    const wrap     = document.getElementById("branch-tab-value");
    if (!selected.length) { wrap.innerHTML = `<div class="alert-warning">⚠️ Select at least one branch.</div>`; return; }
    const compareNames = [centralName, ...selected];
    const compareDf    = branchAgg.filter(r => compareNames.includes(r.PlantName));

    const bCols = [
      {key:"PlantName",       label:"Plant Name"},
      {key:"TotalValue",      label:"Total Value (ETB)",    fmt:fmtETB, rawKey:"TotalValue"},
      {key:"Unrestricted",    label:"Unrestricted (ETB)",   fmt:fmtETB, rawKey:"Unrestricted"},
      {key:"UnrestrictedQty", label:"Avail Qty",            fmt:fmtQty, rawKey:"UnrestrictedQty", cellClass:"col-qty"},
      {key:"Transit",         label:"Transit (ETB)",        fmt:fmtETB, rawKey:"Transit"},
      {key:"TransitQty",      label:"Transit Qty",          fmt:fmtQty, rawKey:"TransitQty",      cellClass:"col-qty"},
      {key:"QC",              label:"QC (ETB)",             fmt:fmtETB, rawKey:"QC"},
      {key:"QCQty",           label:"QC Qty",               fmt:fmtQty, rawKey:"QCQty",           cellClass:"col-qty"},
      {key:"Items",           label:"# Line Items"},
    ];
    wrap.innerHTML = `<div id="branch-table-wrap-inner" style="margin-bottom:1rem">${buildTable(compareDf, bCols, r => r.PlantName === centralName ? "row-blue" : "")}</div>`;
    document.getElementById("btn-dl-branch-csv").onclick  = () => downloadCSV(compareDf,   bCols, "branch_comparison.csv");
    document.getElementById("btn-dl-branch-xlsx").onclick = () => downloadExcel(compareDf, bCols, "branch_comparison.xlsx");
  }

  // ── TAB 2: Material Across Branches ──
  let matTabInitialized = false;
  function renderMaterialTab() {
    const wrap         = document.getElementById("branch-tab-material");
    const allPlantNames = [...new Set(df.map(r => r["Plant Name"]))].sort((a,b) => {
      if (a === centralName) return -1; if (b === centralName) return 1; return a.localeCompare(b);
    });

    if (!matTabInitialized) {
      matTabInitialized = true;
      const mgNamesForFilter = [...new Set(df.map(r => r["Material Group Name"]))].filter(Boolean).sort(); // [EXCLUSION DISABLED]
      wrap.innerHTML = `
        <div style="display:flex;gap:0.8rem;flex-wrap:wrap;align-items:flex-end;margin-bottom:1rem;background:var(--surface);border:1px solid var(--border);border-radius:8px;padding:0.8rem">
          <div>
            <div class="nav-label" style="font-size:0.65rem;margin-bottom:3px">Search Material</div>
            <input id="mat-search" type="text" placeholder="code or description…" style="background:var(--surface2);border:1px solid var(--border2);color:var(--text);padding:6px 10px;border-radius:6px;width:220px;font-size:13px">
          </div>
          <div>
            <div class="nav-label" style="font-size:0.65rem;margin-bottom:3px">Metric</div>
            <select id="mat-metric" style="background:var(--surface2);border:1px solid var(--border2);color:var(--text);padding:6px 10px;border-radius:6px;font-size:13px">
              <option value="TotalValue">Total Value (ETB)</option>
              <option value="Unrestricted">Unrestricted Value (ETB)</option>
              <option value="Transit">Transit Value (ETB)</option>
              <option value="QC">QC Value (ETB)</option>
              <option value="TotalQty">Total Quantity</option>
              <option value="UnrestrictedQty">Available Quantity</option>
              <option value="TransitQty">Transit Quantity</option>
              <option value="QCQty">QC Quantity</option>
            </select>
          </div>
          <div>
            <div class="nav-label" style="font-size:0.65rem;margin-bottom:3px">Material Group</div>
            <select id="mat-mgfilter" style="background:var(--surface2);border:1px solid var(--border2);color:var(--text);padding:6px 10px;border-radius:6px;font-size:13px">
              <option value="">All Material Groups</option>
              ${mgNamesForFilter.map(m => `<option value="${escHtml(m)}">${escHtml(m)}</option>`).join("")}
            </select>
          </div>
          <div>
            <div class="nav-label" style="font-size:0.65rem;margin-bottom:3px">Sort By</div>
            <select id="mat-sort" style="background:var(--surface2);border:1px solid var(--border2);color:var(--text);padding:6px 10px;border-radius:6px;font-size:13px">
              <option value="total_desc">Highest Total ↓</option>
              <option value="total_asc">Lowest Total ↑</option>
              <option value="desc_asc">Description A–Z</option>
              <option value="spread_desc">Most Branches ↓</option>
            </select>
          </div>
          <button id="mat-apply" class="apply-btn">Apply</button>
          <button id="mat-dl-csv" class="dl-btn">⬇ CSV</button>
          <button id="mat-dl-xlsx" class="dl-btn">⬇ Excel</button>
        </div>
        <div id="mat-chart-wrap" style="margin-bottom:1rem"></div>
        <div id="mat-table-wrap"></div>`;
      document.getElementById("mat-apply").addEventListener("click", refreshMaterialView);
      document.getElementById("mat-search").addEventListener("keydown", e => { if (e.key === "Enter") refreshMaterialView(); });
    }
    refreshMaterialView();

    function refreshMaterialView() {
      const searchVal = (document.getElementById("mat-search").value || "").toLowerCase().trim();
      const metric    = document.getElementById("mat-metric").value;
      const sortMode  = document.getElementById("mat-sort").value;
      const mgFilter  = document.getElementById("mat-mgfilter").value;
      const isQty     = metric.includes("Qty");
      const fmtFn     = isQty ? fmtQty : fmtETB;

      let materials = Object.entries(matPlantMap)
        .filter(([mat, info]) => {
          if (mgFilter && info.group !== mgFilter) return false;
          if (searchVal) {
            // Also match by any original source code that maps to this target
            const srcMatch = info._sourceBreakdown?.some(s => s.origCode.toLowerCase().includes(searchVal));
            return mat.toLowerCase().includes(searchVal) || info.desc.toLowerCase().includes(searchVal) || !!srcMatch;
          }
          return true;
        })
        .map(([mat, info]) => {
          const plantData = {};
          let grandTotal = 0, branchCount = 0;
          allPlantNames.forEach(pn => {
            const v = info[pn] ? info[pn][metric] : 0;
            plantData[pn] = v || 0;
            grandTotal   += plantData[pn];
            if ((info[pn]?.TotalValue || 0) > 0) branchCount++;
          });
          return {mat, desc:info.desc, group:info.group, plantData, grandTotal, branchCount,
                  _isReconciled: info._isReconciled, _sourceBreakdown: info._sourceBreakdown};
        });

      if (sortMode === "total_desc") materials.sort((a,b) => b.grandTotal - a.grandTotal);
      if (sortMode === "total_asc")  materials.sort((a,b) => a.grandTotal - b.grandTotal);
      if (sortMode === "desc_asc")   materials.sort((a,b) => a.desc.localeCompare(b.desc));
      if (sortMode === "spread_desc")materials.sort((a,b) => b.branchCount - a.branchCount);

      const top      = materials.slice(0, 30);
      const chartWrap = document.getElementById("mat-chart-wrap");
      if (!top.length) {
        chartWrap.innerHTML = "";
        document.getElementById("mat-table-wrap").innerHTML = `<div class="alert-info">No materials found.</div>`;
        return;
      }
      chartWrap.innerHTML = "";

      const colDefs = [
        {key:"mat",  label:"Material Code", fmt:(val,r)=>renderMatCode(val,r), raw:true, cellClass:"col-mat-code-wrap"},
        {key:"desc", label:"Material Description", fmt:(val,r)=>renderMatDesc(val,r), raw:true, cellClass:"col-mat-desc-wrap"},
        {key:"group",label:"Material Group"},
        {key:"_reconCol", label:"Reconciled From", raw:true},
        ...allPlantNames.map(pn => ({key:`__p__${pn}`, label:pn, fmt:fmtFn, rawKey:`__r__${pn}`, cellClass:isQty?"col-qty":"col-val"})),
        {key:"grandTotal",  label:"Grand Total", fmt:fmtFn, rawKey:"grandTotal", cellClass:isQty?"col-qty":"col-val"},
        {key:"branchCount", label:"# Branches"},
      ];
      const tableRows = materials.slice(0, 200).map(m => {
        const row = {mat:m.mat, desc:m.desc, group:m.group, grandTotal:m.grandTotal, branchCount:m.branchCount};
        allPlantNames.forEach(pn => { row[`__p__${pn}`] = m.plantData[pn] || 0; row[`__r__${pn}`] = m.plantData[pn] || 0; });
        row["__r__grandTotal"] = m.grandTotal;

        // Build reconciliation summary for this material
        if (m._isReconciled && m._sourceBreakdown && m._sourceBreakdown.length > 1) {
          const qtyKey  = isQty ? (metric === "QCQty" ? "convertedQCQty" : metric === "TransitQty" ? "convertedTransitQty" : "convertedUnrestQty") : null;
          const rawKey  = isQty ? (metric === "QCQty" ? "qcQty"          : metric === "TransitQty" ? "transitQty"          : "unrestQty")          : null;
          row._reconCol = m._sourceBreakdown.map(s => {
            const isTarget = s.origCode === m.mat;
            if (isTarget) {
              return `<span style="font-size:0.65rem;color:var(--green)"><span class="col-mat-code" style="font-size:0.6rem">${escHtml(s.origCode)}</span>${isQty && qtyKey ? ` = ${fmtQty(s[qtyKey])}` : ""}</span>`;
            }
            const rawQty  = isQty && rawKey  ? ` ${fmtQty(s[rawKey])}`       : "";
            const convQty = isQty && qtyKey  ? ` → <b style="color:var(--text)">${fmtQty(s[qtyKey])}</b>` : "";
            return `<span style="font-size:0.65rem;color:var(--muted)"><span class="col-mat-code" style="font-size:0.6rem">${escHtml(s.origCode)}</span>${rawQty} ×${s.convFactor}${convQty}</span>`;
          }).join(`<span style="color:var(--dim);margin:0 1px">+</span>`);
          row._reconCol = `<div style="display:flex;flex-wrap:wrap;gap:2px;align-items:center">${row._reconCol}</div>`;
        } else {
          row._reconCol = `<span style="color:var(--dim);font-size:0.65rem">—</span>`;
        }
        return row;
      });

      const centralKey = `__p__${centralName}`;
      const thead = `<thead><tr>${colDefs.map(c =>
        `<th${c.key === centralKey ? ' style="color:#58a6ff;background:#0d2035"' : ""}>${escHtml(c.label)}</th>`
      ).join("")}</tr></thead>`;
      const tbody = tableRows.map(r => {
        const cells = colDefs.map(c => {
          const v       = r[c.key];
          const raw     = c.raw ? v : null;           // raw HTML — don't escape
          const display = raw != null ? (raw ?? "")
                        : c.fmt ? c.fmt(v)
                        : (v == null ? "" : escHtml(String(v)));
          const isZero  = typeof v === "number" && v === 0;
          const style   = c.key === centralKey ? 'style="color:#58a6ff;background:#0d2035"' : isZero ? 'style="color:#484f58"' : "";
          const cls     = c.cellClass || "";
          return `<td class="${cls}" ${style}>${display}</td>`;
        }).join("");
        return `<tr>${cells}</tr>`;
      }).join("");
      document.getElementById("mat-table-wrap").innerHTML = `
        <div style="color:var(--muted);font-size:12px;margin-bottom:6px">Showing ${tableRows.length} of ${materials.length} materials · Blue = Central (${escHtml(centralName)})</div>
        <div class="tbl-wrap"><table>${thead}<tbody>${tbody}</tbody></table></div>
        ${materials.length > 200 ? `<div class="alert-info">Showing first 200 of ${materials.length}. Refine search.</div>` : ""}`;

      const flatCols = [
        {key:"mat", label:"Material Code", fmt:(val,r)=>renderMatCode(val,r), raw:true, cellClass:"col-mat-code-wrap"}, {key:"desc", label:"Material Description", fmt:(val,r)=>renderMatDesc(val,r), raw:true, cellClass:"col-mat-desc-wrap"}, {key:"group",label:"Material Group"},
        ...allPlantNames.map(pn => ({key:`__p__${pn}`, label:pn, rawKey:`__r__${pn}`})),
        {key:"grandTotal",label:"Grand Total"},
      ];
      const btnCsv  = document.getElementById("mat-dl-csv");
      const btnXlsx = document.getElementById("mat-dl-xlsx");
      if (btnCsv)  btnCsv.onclick  = () => downloadCSV(tableRows,   flatCols, "materials_by_branch.csv");
      if (btnXlsx) btnXlsx.onclick = () => downloadExcel(tableRows, flatCols, "materials_by_branch.xlsx");
    }
  }

  sel.addEventListener("change", updateBranchCharts);
  updateBranchCharts();
}

// ═══════════════════════════════════════════════════════════════════════════
// INVENTORY FLOW
// ═══════════════════════════════════════════════════════════════════════════
function renderFlow() {
  const df = applyPageFilter("flow");

  const totalVal   = df.reduce((s,r) => s + r["Total Value"], 0);
  const transitVal = df.reduce((s,r) => s + r["Value of Stock in Transit"], 0);
  const qcVal      = df.reduce((s,r) => s + r["Value of Stock in Quality Inspection"], 0);
  const availVal   = df.reduce((s,r) => s + r["Value of Unrestricted Stock"], 0);
  const totalQty   = df.reduce((s,r) => s + r["Total Qty"], 0);
  const availQty   = df.reduce((s,r) => s + r["Unrestricted Stock"], 0);

  const reorderItems = df.filter(r => r["Unrestricted Stock"] === 0 && (r["Stock in Transit"] > 0 || r["Stock in Quality Inspection"] > 0));

  setKpis("flow-kpis", [
    ["Total Inventory",      fmtETB(totalVal),   `${fmtQty(totalQty)} units`,               "blue"],
    ["Available Stock",      fmtETB(availVal),   `${fmtQty(availQty)} units unrestricted`,   "green"],
    ["In Transit (Inbound)", fmtETB(transitVal), `${fmtQty(df.reduce((s,r) => s+r["Stock in Transit"],0))} units`, "amber"],
    ["In QC",                fmtETB(qcVal),      `${fmtQty(df.reduce((s,r) => s+r["Stock in Quality Inspection"],0))} units`, "red"],
    ["Reorder Alerts",       String(reorderItems.length), "Zero unrestricted stock", "red"],
  ]);

  const reorderCols = [
    {key:"Material", label:"Material Code", fmt:(val,r)=>renderMatCode(val,r), raw:true, cellClass:"col-mat-code-wrap"},
    {key:"Material Description", label:"Material Description", fmt:(val,r)=>renderMatDesc(val,r), raw:true, cellClass:"col-mat-desc-wrap"},
    {key:"Material Group Name",       label:"Material Group"},
    {key:"Plant Name",                label:"Plant"},
    {key:"Unrestricted Stock",        label:"Avail Qty",        fmt:fmtQty, rawKey:"Unrestricted Stock",        cellClass:"col-qty"},
    {key:"Stock in Transit",          label:"In Transit",        fmt:fmtQty, rawKey:"Stock in Transit",          cellClass:"col-qty"},
    {key:"Stock in Quality Inspection",label:"In QC",           fmt:fmtQty, rawKey:"Stock in Quality Inspection",cellClass:"col-qty"},
    {key:"Value of Stock in Transit", label:"Transit Value (ETB)",fmt:fmtETB,rawKey:"Value of Stock in Transit", cellClass:"col-val"},
    {key:"_purDoc",                   label:"Purchasing Document"},
    {key:"_supPlant",                 label:"Supplying Plant"},
    {key:"_alert",                    label:"Alert", raw:true},
  ];
  const reorderRows = reorderItems.map(r => {
    const info = getTransitInfo(r["Material"], r["Plant"]);
    return {
      ...r,
      _purDoc:   info.purDoc,
      _supPlant: info.supPlant,
      _alert: r["Stock in Transit"] > 0 && r["Stock in Quality Inspection"] > 0
        ? "<span class='badge badge-red'>Transit+QC</span>"
        : r["Stock in Transit"] > 0
        ? "<span class='badge badge-amber'>Awaiting Transit</span>"
        : "<span class='badge badge-amber'>Awaiting QC Release</span>",
    };
  });
  document.getElementById("reorder-table-wrap").innerHTML = reorderRows.length
    ? buildTable(reorderRows, reorderCols, () => "row-amber")
    : `<div class="alert-info">✓ No reorder alerts — all materials have available unrestricted stock.</div>`;

  // Stock levels chart
  const plantAgg = sortBy(
    groupBy(df, "Plant Name", [["avail","Unrestricted Stock"],["transit","Stock in Transit"],["qc","Stock in Quality Inspection"]]),
    "avail"
  );
  Plotly.newPlot("chart-stock-levels", [
    {type:"bar", name:"Available",  x:plantAgg.map(r=>r["Plant Name"]), y:plantAgg.map(r=>r.avail),  marker:{color:"#3fb950"}},
    {type:"bar", name:"In Transit", x:plantAgg.map(r=>r["Plant Name"]), y:plantAgg.map(r=>r.transit), marker:{color:"#d29922"}},
    {type:"bar", name:"In QC",      x:plantAgg.map(r=>r["Plant Name"]), y:plantAgg.map(r=>r.qc),     marker:{color:"#f85149"}},
  ], pl({height:300,barmode:"stack",margin:{l:20,r:20,t:20,b:80}}), PLOTLY_CONFIG);

  // Inter-location transfers
  const transferCols = [
    {key:"Material", label:"Material Code", fmt:(val,r)=>renderMatCode(val,r), raw:true, cellClass:"col-mat-code-wrap"},
    {key:"Material Description", label:"Material Description", fmt:(val,r)=>renderMatDesc(val,r), raw:true, cellClass:"col-mat-desc-wrap"},
    {key:"Plant Name",                label:"Receiving Plant"},
    {key:"_purDoc",                   label:"Purchasing Document"},
    {key:"_supPlant",                 label:"Supplying Plant"},
    {key:"Stock in Transit",          label:"Transit Qty",        fmt:fmtQty, rawKey:"Stock in Transit",          cellClass:"col-qty"},
    {key:"Value of Stock in Transit", label:"Transit Value (ETB)",fmt:fmtETB, rawKey:"Value of Stock in Transit", cellClass:"col-val"},
  ];
  const transferRows = sortBy(df.filter(r => r["Stock in Transit"] > 0), "Value of Stock in Transit").map(r => {
    const info = getTransitInfo(r["Material"], r["Plant"]);
    return { ...r, _purDoc: info.purDoc, _supPlant: info.supPlant };
  });
  document.getElementById("transfer-table-wrap").innerHTML = transferRows.length
    ? buildTable(transferRows, transferCols)
    : `<div class="alert-info">No inter-location transfers currently in progress.</div>`;

  // Inbound vs available chart
  const inboundAgg = sortBy(
    groupBy(df.filter(r => r["Stock in Transit"] > 0), "Plant Name", [["avail","Unrestricted Stock"],["inbound","Stock in Transit"]]),
    "inbound"
  );
  if (inboundAgg.length) {
    Plotly.newPlot("chart-inbound-outbound", [
      {type:"bar", name:"Available Stock", x:inboundAgg.map(r=>r["Plant Name"]), y:inboundAgg.map(r=>r.avail),   marker:{color:"#3fb950"}},
      {type:"bar", name:"Inbound Transit", x:inboundAgg.map(r=>r["Plant Name"]), y:inboundAgg.map(r=>r.inbound), marker:{color:"#d29922"}},
    ], pl({height:280,barmode:"group",margin:{l:20,r:20,t:20,b:80}}), PLOTLY_CONFIG);
  } else {
    document.getElementById("chart-inbound-outbound").innerHTML = `<div class="alert-info">No transit data to chart.</div>`;
  }

  const flowDlCols = [
    {key:"Material", label:"Material Code", fmt:(val,r)=>renderMatCode(val,r), raw:true, cellClass:"col-mat-code-wrap"},
    {key:"Material Description", label:"Material Description", fmt:(val,r)=>renderMatDesc(val,r), raw:true, cellClass:"col-mat-desc-wrap"},
    {key:"Plant Name",                        label:"Plant"},
    {key:"Material Group Name",               label:"Material Group"},
    {key:"Unrestricted Stock",                label:"Available Qty",      rawKey:"Unrestricted Stock"},
    {key:"Stock in Transit",                  label:"Transit Qty",        rawKey:"Stock in Transit"},
    {key:"Stock in Quality Inspection",       label:"QC Qty",             rawKey:"Stock in Quality Inspection"},
    {key:"Value of Unrestricted Stock",       label:"Available Value (ETB)",rawKey:"Value of Unrestricted Stock"},
    {key:"Value of Stock in Transit",         label:"Transit Value (ETB)",rawKey:"Value of Stock in Transit"},
  ];
  document.getElementById("btn-dl-flow-csv").onclick  = () => downloadCSV(df,   flowDlCols, "inventory_flow.csv");
  document.getElementById("btn-dl-flow-xlsx").onclick = () => downloadExcel(df, flowDlCols, "inventory_flow.xlsx");
}

// ═══════════════════════════════════════════════════════════════════════════
// DATA PREVIEW
// ═══════════════════════════════════════════════════════════════════════════
function renderPreview() {
  filtDf = getReconciledBase();
  populatePreviewFilters();
  renderPreviewTable();
}

function populatePreviewFilters() {
  function fill(id, key, excludeFn) {
    const sel = document.getElementById(id); if (!sel) return;
    const vals = [...new Set(rawDf.map(r => r[key]))]
      .filter(Boolean)
      .filter(v => !excludeFn || !excludeFn(v))
      .sort();
    sel.innerHTML = vals.map(v => `<option value="${escHtml(v)}">${escHtml(v)}</option>`).join("");
  }
  fill("filter-plant",  "Plant Name",         null);
  fill("filter-mg",     "Material Group Name", null); // [EXCLUSION DISABLED]
  fill("filter-mgname", "Material Group Name", null); // [EXCLUSION DISABLED]
}

function applyPreviewFilters() {
  const baseDf     = getReconciledBase();
  const getSelected = id => [...document.querySelectorAll(`#${id} option:checked`)].map(o => o.value);
  const plants      = getSelected("filter-plant");
  const mgs         = getSelected("filter-mg");
  const mgnames     = getSelected("filter-mgname");
  filtDf = baseDf.filter(r =>
    (!plants.length  || plants.includes(r["Plant Name"])) &&
    (!mgs.length     || mgs.includes(r["Material Group Name"])) &&
    (!mgnames.length || mgnames.includes(r["Material Group Name"]))
  );
  renderPreviewTable();
}

function renderPreviewTable() {
  const df = filtDf;
  setKpis("preview-kpis", [
    ["Total Records",    df.length.toLocaleString(),                            "After filtering",           "blue"],
    ["Unique Materials", new Set(df.map(r=>r["Material"])).size.toLocaleString(),"Distinct SKUs",            "green"],
    ["Total Plants",     new Set(df.map(r=>r["Plant"])).size.toLocaleString(),   "Stocking locations",       "amber"],
    ["Material Groups",  new Set(df.map(r=>r["Material Group Name"])).size.toLocaleString(),"Therapeutic categories","purple"],
  ]);
  document.getElementById("preview-count").innerHTML = `Showing <b>${df.length.toLocaleString()}</b> of <b>${rawDf.length.toLocaleString()}</b> records`;

  const cols = [
    {key:"Material", label:"Material Code", fmt:(val,r)=>renderMatCode(val,r), raw:true, cellClass:"col-mat-code-wrap"},
    {key:"Material Description", label:"Material Description", fmt:(val,r)=>renderMatDesc(val,r), raw:true, cellClass:"col-mat-desc-wrap"},
    {key:"Plant Name",                        label:"Plant"},
    {key:"Material Group Name",               label:"Material Group"},
    {key:"Unrestricted Stock",                label:"Avail Qty",        fmt:fmtQty, rawKey:"Unrestricted Stock",          cellClass:"col-qty"},
    {key:"Stock in Transit",                  label:"Transit Qty",      fmt:fmtQty, rawKey:"Stock in Transit",            cellClass:"col-qty"},
    {key:"Stock in Quality Inspection",       label:"QC Qty",           fmt:fmtQty, rawKey:"Stock in Quality Inspection", cellClass:"col-qty"},
    {key:"Value of Unrestricted Stock",       label:"Avail Value (ETB)",fmt:fmtETB, rawKey:"Value of Unrestricted Stock", cellClass:"col-val"},
    {key:"Value of Stock in Transit",         label:"Transit Value (ETB)",fmt:fmtETB,rawKey:"Value of Stock in Transit",  cellClass:"col-val"},
    {key:"Value of Stock in Quality Inspection",label:"QC Value (ETB)", fmt:fmtETB, rawKey:"Value of Stock in Quality Inspection",cellClass:"col-val"},
    {key:"Total Value",                       label:"Total Value (ETB)",fmt:fmtETB, rawKey:"Total Value",                 cellClass:"col-val"},
    {key:"_expiryStr",                        label:"Expiry Date"},
  ];

  // FIX BUG-1: display rows are sliced to 500 for the table, but download rows
  // use the FULL filtered dataset so the export is never silently truncated.
  const displayRows  = df.slice(0, 500).map(r => ({...r, _expiryStr: r._expiry ? r._expiry.toISOString().slice(0,10) : ""}));
  const downloadRows = df.map(r => ({...r, _expiryStr: r._expiry ? r._expiry.toISOString().slice(0,10) : ""}));

  document.getElementById("preview-table-wrap").innerHTML =
    buildTable(displayRows, cols) +
    (df.length > 500
      ? `<div class="alert-warning">⚠️ Showing first 500 of ${df.length.toLocaleString()} records. Downloads include all ${df.length.toLocaleString()} rows.</div>`
      : "");

  // FIX BUG-1: wire download buttons to downloadRows (full set)
  document.getElementById("btn-dl-preview").onclick      = () => downloadCSV(downloadRows,   cols, "pharma_inventory_filtered.csv");
  document.getElementById("btn-dl-preview-xlsx").onclick = () => downloadExcel(downloadRows, cols, "pharma_inventory_filtered.xlsx");
}

// ═══════════════════════════════════════════════════════════════════════════
// MATERIAL UOM CONVERSION & RECONCILIATION
// ═══════════════════════════════════════════════════════════════════════════
//
// Data model: reconcileGroups = [{
//   sourceMaterial, sourceDesc, sourceUnit,
//   conversionFactor,
//   targetMaterial,  targetDesc,  targetUnit
// }]
//
// Qty rule  : targetQty = sourceQty × conversionFactor
// Value rule: values are ETB monetary → summed as-is (no factor)
// Expiry    : earliest (soonest) expiry kept — most conservative for pharma
// QC / Transit quantities: also × conversionFactor (physical units)

// Returns conversion rule for a source code, or null.
function getConversion(code) {
  return reconcileGroups.find(g => g.sourceMaterial === code) || null;
}

// Legacy helper still used by a few callers.
function getCanonicalCode(code) {
  const conv = getConversion(code);
  return conv ? conv.targetMaterial : code;
}

// Applies all UOM conversions to a data frame.
// Source rows are converted (qty × factor) and merged into target rows.
// Values (ETB) are summed as-is. Expiry = earliest date across merged rows.
function applyReconciliationToData(df) {
  if (!reconcileGroups.length) return df;

  // FIX ROBUST: Total Qty removed from QTY_COLS — it was scaled by factor then
  // immediately overwritten by the correct sum below, which was confusing and
  // left an incorrect intermediate value briefly in the merged array.
  const QTY_COLS = [
    "Unrestricted Stock",
    "Stock in Quality Inspection",
    "Blocked Stock",
    "Stock in Transit",
  ];
  const VAL_COLS = [
    "Value of Unrestricted Stock",
    "Value of Stock in Quality Inspection",
    "Value of Stock in Transit",
  ];

  const merged  = [];
  const mergeKey = r => `${r["Material"]}||${r["Plant"]}||${r["Storage Location"]}||${r["Batch"] || ""}`;
  const keyMap  = {};

  df.forEach(row => {
    const conv = getConversion(row["Material"]);

    if (!conv) {
      const k = mergeKey(row);
      if (keyMap[k] !== undefined) {
        const target = merged[keyMap[k]];
        QTY_COLS.forEach(c => { target[c] = (target[c] || 0) + (row[c] || 0); });
        VAL_COLS.forEach(c => { target[c] = (target[c] || 0) + (row[c] || 0); });
        _mergeExpiry(target, row);
      } else {
        keyMap[k] = merged.length;
        merged.push({...row});
      }
      return;
    }

    const convertedRow = {...row};
    convertedRow["Material"]             = conv.targetMaterial;
    convertedRow["Material Description"] = conv.targetDesc || row["Material Description"];

    // FIX ROBUST: round to 9dp to suppress IEEE 754 float drift (e.g. 0.1 × 3 = 0.30000000000000004)
    QTY_COLS.forEach(c => {
      convertedRow[c] = Math.round((row[c] || 0) * conv.conversionFactor * 1e9) / 1e9;
    });
    VAL_COLS.forEach(c => { convertedRow[c] = row[c] || 0; });

    const k = mergeKey(convertedRow);
    if (keyMap[k] !== undefined) {
      const target = merged[keyMap[k]];
      QTY_COLS.forEach(c => { target[c] = (target[c] || 0) + (convertedRow[c] || 0); });
      VAL_COLS.forEach(c => { target[c] = (target[c] || 0) + (convertedRow[c] || 0); });
      _mergeExpiry(target, convertedRow);
      if (!target["Batch"] && convertedRow["Batch"]) target["Batch"] = convertedRow["Batch"];
      if (!target["Description of Storage Location"] && convertedRow["Description of Storage Location"])
        target["Description of Storage Location"] = convertedRow["Description of Storage Location"];
    } else {
      keyMap[k] = merged.length;
      merged.push(convertedRow);
    }
  });

  // Recompute derived totals after merge
  merged.forEach(row => {
    row["Total Qty"] = (row["Unrestricted Stock"] || 0)
                     + (row["Stock in Transit"] || 0)
                     + (row["Stock in Quality Inspection"] || 0);
    row["Total Value"] = (row["Value of Unrestricted Stock"] || 0)
                       + (row["Value of Stock in Transit"] || 0)
                       + (row["Value of Stock in Quality Inspection"] || 0);
  });

  return merged;
}

// ═══════════════════════════════════════════════════════════════════════════
// AGGREGATE BY MATERIAL (used by QC and Branch Comparison)
// ═══════════════════════════════════════════════════════════════════════════
// After reconciliation renames source codes → target codes, rows that share
// the same target material code but differ in Plant / Storage Location / Batch
// still appear as separate rows.  This function collapses them into ONE row
// per material code, summing all qty and value columns and keeping the
// earliest expiry.  It also attaches a human-readable "_reconSummary" string
// that lists every original source contribution for audit purposes.
//
// Example:
//   102-ACET-0102-01 (QC=10 × 0.5) → 102-ACET-0102-02 (QC=5)
//   102-ACET-0102-04 (QC=20 × 0.5) → 102-ACET-0102-02 (QC=10)
//   102-ACET-0102-03 (QC=10 × 1.0) → 102-ACET-0102-02 (QC=10)
//   Already in data: 102-ACET-0102-02 (QC=90)
//   ──────────────────────────────────────────────────────
//   Consolidated:    102-ACET-0102-02 (QC=115)
//
// The "_sourceBreakdown" array carries per-source-code objects for tooltip/table use.
// The "_isReconciled" flag is true when at least two different original codes merged.

function aggregateByMaterial(df) {
  const QTY_COLS = [
    "Unrestricted Stock", "Stock in Quality Inspection",
    "Blocked Stock",      "Stock in Transit",
    "Total Qty",
  ];
  const VAL_COLS = [
    "Value of Unrestricted Stock", "Value of Stock in Quality Inspection",
    "Value of Stock in Transit",   "Value of Stock in Quality Inspection",
    "Total Value",
  ];

  // Build a lookup: original source → target + factor from reconcileGroups
  const srcToRule = {};
  reconcileGroups.forEach(g => { srcToRule[g.sourceMaterial] = g; });

  // We need to know what original (pre-reconciliation) material each row came from.
  // applyReconciliationToData renames Material to the target code but does NOT
  // preserve the original code.  However, we can look it up via _originalMaterial
  // if present (we will start stamping it below), or infer it from reconcileGroups.

  // First pass: re-apply raw source → row mapping to attach _originalMaterial
  // We do this by re-scanning rawDf and reconcileGroups together.
  const targetToSources = {}; // targetCode → Set of source codes
  reconcileGroups.forEach(g => {
    if (!targetToSources[g.targetMaterial]) targetToSources[g.targetMaterial] = new Set();
    targetToSources[g.targetMaterial].add(g.sourceMaterial);
  });

  // Group all rows by Material code
  const matMap = {}; // materialCode → aggregated row

  df.forEach(row => {
    const mat = row["Material"];
    if (!mat) return;

    if (!matMap[mat]) {
      // First row for this material — use it as the base
      matMap[mat] = {
        ...row,
        _sourceBreakdown: [],   // [{origCode, convFactor, qcQty, unrestQty, transitQty}]
        _isReconciled:    false,
      };
    } else {
      // Subsequent row for same material code — aggregate
      const target = matMap[mat];
      QTY_COLS.forEach(c => { target[c] = (target[c] || 0) + (row[c] || 0); });
      // De-duplicate VAL_COLS (Total Value included twice above for safety — fix)
      [
        "Value of Unrestricted Stock",
        "Value of Stock in Quality Inspection",
        "Value of Stock in Transit",
        "Total Value",
      ].forEach(c => { target[c] = (target[c] || 0) + (row[c] || 0); });
      _mergeExpiry(target, row);
      // Merge non-blank string fields for display
      if (!target["Plant Name"] && row["Plant Name"]) target["Plant Name"] = row["Plant Name"];
      if (!target["Material Group Name"] && row["Material Group Name"]) target["Material Group Name"] = row["Material Group Name"];
      target._isReconciled = true;
    }
  });

  // Third pass: for each target material that has reconcile rules pointing to it,
  // build the breakdown summary from rawDf (pre-conversion) for audit display.
  const rawByMat = {};
  rawDf.forEach(r => {
    const c = r["Material"];
    if (!rawByMat[c]) rawByMat[c] = [];
    rawByMat[c].push(r);
  });

  Object.keys(matMap).forEach(tgtCode => {
    const row    = matMap[tgtCode];
    const srcSet = targetToSources[tgtCode]; // source codes that map → tgtCode
    if (!srcSet || !srcSet.size) return;

    const breakdown = [];

    // Include the target code itself if it exists in raw data
    const selfRaws = rawByMat[tgtCode] || [];
    if (selfRaws.length) {
      const selfQCQty  = selfRaws.reduce((s,r) => s + (r["Stock in Quality Inspection"] || 0), 0);
      const selfUnrQty = selfRaws.reduce((s,r) => s + (r["Unrestricted Stock"] || 0), 0);
      const selfTrQty  = selfRaws.reduce((s,r) => s + (r["Stock in Transit"]   || 0), 0);
      if (selfQCQty + selfUnrQty + selfTrQty > 0) {
        breakdown.push({
          origCode: tgtCode, convFactor: 1,
          qcQty: selfQCQty, unrestQty: selfUnrQty, transitQty: selfTrQty,
          convertedQCQty: selfQCQty, convertedUnrestQty: selfUnrQty, convertedTransitQty: selfTrQty,
        });
      }
    }

    // Include each source code
    srcSet.forEach(srcCode => {
      const rule    = srcToRule[srcCode];
      const srcRaws = rawByMat[srcCode] || [];
      if (!srcRaws.length || !rule) return;
      const rawQCQty  = srcRaws.reduce((s,r) => s + (r["Stock in Quality Inspection"] || 0), 0);
      const rawUnrQty = srcRaws.reduce((s,r) => s + (r["Unrestricted Stock"] || 0), 0);
      const rawTrQty  = srcRaws.reduce((s,r) => s + (r["Stock in Transit"]   || 0), 0);
      if (rawQCQty + rawUnrQty + rawTrQty === 0) return;
      breakdown.push({
        origCode: srcCode, convFactor: rule.conversionFactor,
        qcQty:      rawQCQty,
        unrestQty:  rawUnrQty,
        transitQty: rawTrQty,
        convertedQCQty:      Math.round(rawQCQty  * rule.conversionFactor * 1e9) / 1e9,
        convertedUnrestQty:  Math.round(rawUnrQty * rule.conversionFactor * 1e9) / 1e9,
        convertedTransitQty: Math.round(rawTrQty  * rule.conversionFactor * 1e9) / 1e9,
      });
    });

    if (breakdown.length > 1) {
      row._sourceBreakdown = breakdown;
      row._isReconciled    = true;
    }
  });

  return Object.values(matMap);
}

// Render a compact reconciliation tooltip/badge showing how a consolidated
// material was built from its sources.  Used in QC and Branch tables.
function renderReconBadge(row, qtyField) {
  if (!row._isReconciled || !row._sourceBreakdown || !row._sourceBreakdown.length) return "";

  const qtyKey = qtyField === "qc"       ? "qcQty"
               : qtyField === "unrest"   ? "unrestQty"
               : qtyField === "transit"  ? "transitQty"
               : "qcQty"; // default

  const convKey = qtyField === "qc"      ? "convertedQCQty"
                : qtyField === "unrest"  ? "convertedUnrestQty"
                : qtyField === "transit" ? "convertedTransitQty"
                : "convertedQCQty";

  const lines = row._sourceBreakdown.map(s => {
    const rawQty   = fmtQty(s[qtyKey]);
    const convQty  = fmtQty(s[convKey]);
    const factor   = s.convFactor === 1 ? "×1.0" : `×${s.convFactor}`;
    const arrow    = s.convFactor === 1 && s.origCode === row["Material"] ? "" : ` ${factor} → ${convQty}`;
    return `${escHtml(s.origCode)}: ${rawQty}${arrow}`;
  }).join("&#10;"); // newline in tooltip

  const total = fmtQty(row._sourceBreakdown.reduce((s,b) => s + (b[convKey] || 0), 0));

  return `<span class="recon-badge" title="Consolidated from ${row._sourceBreakdown.length} source(s):&#10;${lines}&#10;Total = ${total}" style="cursor:help;margin-left:4px;font-size:0.6rem;background:#1f3558;color:#58a6ff;border:1px solid #1f6feb;border-radius:3px;padding:1px 4px;vertical-align:middle;white-space:nowrap">⟳ ${row._sourceBreakdown.length} src</span>`;
}

// Keep earliest (soonest) expiry — safest approach for pharma stock management.
function _mergeExpiry(target, src) {
  const te = target["_expiry"], se = src["_expiry"];
  if (se instanceof Date && !isNaN(se)) {
    if (!(te instanceof Date) || isNaN(te) || se < te) {
      target["_expiry"] = se;
    }
  }
}

// ── FIX BUG-2 + BUG-5: Reconciliation rule validation ────────────────────
// Validates a new source→target pair against existing rules.
// Returns an array of error strings (empty = valid).
function validateNewConversionRule(srcCode, tgtCode) {
  const errors = [];
  if (srcCode === tgtCode) {
    errors.push("Source and target must be different materials.");
  }
  // Duplicate source
  if (reconcileGroups.some(g => g.sourceMaterial === srcCode)) {
    errors.push(`"${srcCode}" already has a conversion rule. Delete it first.`);
  }
  // FIX BUG-5: prevent a target material from becoming a new source
  if (reconcileGroups.some(g => g.targetMaterial === srcCode)) {
    errors.push(`"${srcCode}" is already a canonical target in another rule. Creating a rule from it would cause conflicts.`);
  }
  // FIX BUG-2: detect chain rule (source → target where target is already someone's source)
  if (reconcileGroups.some(g => g.sourceMaterial === tgtCode)) {
    errors.push(`"${tgtCode}" is already a source in another rule. Chained conversions (A→B→C) are not supported — map directly to the final canonical code.`);
  }
  // Circular: tgtCode already maps back to srcCode
  if (reconcileGroups.some(g => g.sourceMaterial === tgtCode && g.targetMaterial === srcCode)) {
    errors.push(`This would create a circular conversion (${tgtCode} → ${srcCode} already exists).`);
  }
  return errors;
}

// ── Panel open / close ────────────────────────────────────────────────────
function openReconcilePanel() {
  document.getElementById("reconcile-panel").style.display  = "flex";
  document.getElementById("reconcile-overlay").style.display = "block";
  refreshReconcileGroupsList();
}
function closeReconcilePanel() {
  document.getElementById("reconcile-panel").style.display  = "none";
  document.getElementById("reconcile-overlay").style.display = "none";
}

// ── Upload mapping Excel ──────────────────────────────────────────────────
function handleReconcileFileUpload(file) {
  const statusEl = document.getElementById("rp-upload-status");
  statusEl.style.color   = "var(--muted)";
  statusEl.textContent   = "⏳ Reading mapping file…";
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb   = XLSX.read(e.target.result, {type:"array"});
      const ws   = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:""});
      let added = 0, skipped = 0, errors = [];
      for (let i = 1; i < rows.length; i++) {
        const r       = rows[i];
        const srcMat  = String(r[0] || "").trim();
        const srcDesc = String(r[1] || "").trim();
        const factor  = parseFloat(r[2]);
        const tgtMat  = String(r[3] || "").trim();
        const tgtDesc = String(r[4] || "").trim();
        const srcUnit = "";
        const tgtUnit = "";

        if (!srcMat || !tgtMat || isNaN(factor) || factor <= 0) { skipped++; continue; }

        // Skip self-mapping rows (source === target) — these are canonical identity
        // rows in the mapping file and must NOT be registered as conversion rules
        // or they block real conversion rules from being added.
        if (srcMat === tgtMat) { skipped++; continue; }

        // FIX BUG-2 + BUG-5: validate via shared rule validator
        const ruleErrors = validateNewConversionRule(srcMat, tgtMat);
        if (ruleErrors.length) { skipped++; errors.push(`Row ${i+1}: ${ruleErrors[0]}`); continue; }

        // FIX ROBUST: store factor with 9dp precision cap
        reconcileGroups.push({
          sourceMaterial: srcMat, sourceDesc: srcDesc, sourceUnit: srcUnit,
          conversionFactor: Math.round(factor * 1e9) / 1e9,
          targetMaterial: tgtMat, targetDesc: tgtDesc, targetUnit: tgtUnit,
        });
        added++;
      }
      invalidateReconCache();
      saveReconcileGroups();
      refreshReconcileGroupsList();
      statusEl.style.color = "var(--green)";
      let msg = `✓ Added ${added} conversion(s)`;
      if (skipped)       msg += ` · ${skipped} skipped`;
      if (errors.length) msg += ` · First issue: ${errors[0]}`;
      statusEl.textContent = msg;
      if (rawDf.length) renderPage(currentPage);
    } catch(err) {
      statusEl.style.color = "var(--red)";
      statusEl.textContent = "✗ Failed to read file: " + err.message;
    }
  };
  reader.readAsArrayBuffer(file);
}

// ── Manual search & select ────────────────────────────────────────────────
let rpSrcSelected = null;
let rpTgtSelected = null;

function rpSearch(query, resultsElId, onSelect) {
  const el = document.getElementById(resultsElId);
  if (!query.trim() || !rawDf.length) { el.innerHTML = ""; return; }
  const q          = query.toLowerCase();
  const seen       = new Set();
  // Search the raw (unreconciled) data so both source and target codes are findable.
  const searchBase = rawDf;
  const matches    = searchBase.filter(r => {
    const c = String(r["Material"] || ""), d = String(r["Material Description"] || "");
    if (seen.has(c)) return false;
    if (c.toLowerCase().includes(q) || d.toLowerCase().includes(q)) { seen.add(c); return true; }
    return false;
  }).slice(0, 12);

  if (!matches.length) {
    el.innerHTML = `<div style="padding:6px;font-size:0.72rem;color:var(--muted)">No materials found</div>`;
    return;
  }
  el.innerHTML = matches.map((m, i) =>
    `<div class="rp-result-item" data-idx="${i}" style="cursor:pointer">
      <span class="rp-result-code">${escHtml(m["Material"])}</span>
      <span class="rp-result-desc">${escHtml(m["Material Description"])}</span>
    </div>`
  ).join("");
  el.querySelectorAll(".rp-result-item").forEach(item => {
    item.addEventListener("click", () => {
      const m = matches[parseInt(item.dataset.idx)];
      onSelect({ code: m["Material"], desc: m["Material Description"] });
      el.innerHTML = "";
    });
  });
}

function rpSetSelected(side, data) {
  const chipId   = side === "src" ? "rp-src-selected" : "rp-tgt-selected";
  const searchId = side === "src" ? "rp-src-search"   : "rp-tgt-search";
  if (side === "src") rpSrcSelected = data; else rpTgtSelected = data;
  const chip = document.getElementById(chipId);
  if (data) {
    chip.style.display = "block";
    // FIX BUG-7: use createElement + addEventListener instead of inline onclick
    chip.innerHTML = "";
    const codeSpan = document.createElement("span");
    codeSpan.className = "rp-result-code";
    codeSpan.textContent = data.code;
    const descSpan = document.createElement("span");
    descSpan.className = "rp-result-desc";
    descSpan.textContent = data.desc;
    const closeBtn = document.createElement("button");
    closeBtn.textContent = "✕";
    closeBtn.style.cssText = "background:none;border:none;color:var(--muted);cursor:pointer;font-size:11px;margin-left:4px";
    closeBtn.addEventListener("click", () => rpClearSelected(side));
    chip.appendChild(codeSpan);
    chip.appendChild(descSpan);
    chip.appendChild(closeBtn);
    document.getElementById(searchId).value = "";
  } else {
    chip.style.display = "none";
    chip.innerHTML = "";
  }
}

function rpClearSelected(side) { rpSetSelected(side, null); }

function addManualConversion() {
  if (!rpSrcSelected) { alert("Select a source material first.");  return; }
  if (!rpTgtSelected) { alert("Select a target material first."); return; }
  const factor = parseFloat(document.getElementById("rp-conv-factor").value);
  if (isNaN(factor) || factor <= 0) { alert("Enter a valid conversion factor > 0."); return; }

  // FIX BUG-2 + BUG-5: validate before pushing
  const errors = validateNewConversionRule(rpSrcSelected.code, rpTgtSelected.code);
  if (errors.length) { alert(errors.join("\n")); return; }

  reconcileGroups.push({
    sourceMaterial:   rpSrcSelected.code,
    sourceDesc:       rpSrcSelected.desc,
    sourceUnit:       "",
    // FIX ROBUST: cap precision to 9dp to avoid float drift in downstream math
    conversionFactor: Math.round(factor * 1e9) / 1e9,
    targetMaterial:   rpTgtSelected.code,
    targetDesc:       rpTgtSelected.desc,
    targetUnit:       "",
  });
  invalidateReconCache();
  saveReconcileGroups();
  refreshReconcileGroupsList();
  rpClearSelected("src");
  rpClearSelected("tgt");
  document.getElementById("rp-conv-factor").value = "1";
  if (rawDf.length) renderPage(currentPage);
}

// ── Render existing conversion list ──────────────────────────────────────
function refreshReconcileGroupsList() {
  const el      = document.getElementById("rp-groups-list");
  const countEl = document.getElementById("rp-group-count");
  countEl.textContent = reconcileGroups.length ? `(${reconcileGroups.length})` : "";
  if (!reconcileGroups.length) {
    el.innerHTML = `<div style="color:var(--muted);font-size:0.75rem;padding:0.5rem 0">No conversions defined yet.</div>`;
    return;
  }
  el.innerHTML = reconcileGroups.map((g, i) => {
    const isBuiltin = !!g._builtin;
    const builtinBadge = isBuiltin
      ? `<span style="font-size:0.6rem;background:#0d2035;color:#58a6ff;border:1px solid #58a6ff;border-radius:4px;padding:1px 5px;vertical-align:middle;margin-left:4px">BUILT-IN</span>`
      : "";
    const deleteBtn = isBuiltin
      ? `<span style="font-size:0.68rem;color:var(--muted);padding:2px 8px;margin-top:2px;cursor:default" title="Built-in rules cannot be deleted">🔒</span>`
      : `<button class="rp-group-del" data-group-idx="${i}" style="flex-shrink:0;font-size:0.68rem;padding:2px 8px;margin-top:2px">Delete</button>`;
    return `
    <div class="rp-group-card" style="margin-bottom:0.6rem${isBuiltin ? ";border-color:#1f3558;background:#0a1929" : ""}">
      <div style="display:flex;align-items:flex-start;justify-content:space-between;gap:0.5rem">
        <div style="flex:1;min-width:0">
          <div style="display:flex;align-items:center;gap:0.4rem;flex-wrap:wrap;margin-bottom:4px">
            <span class="rp-code-tag">${escHtml(g.sourceMaterial)}</span>
            <span style="color:var(--muted);font-size:0.7rem;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:180px">${escHtml(g.sourceDesc || "")}</span>
            ${builtinBadge}
          </div>
          <div style="display:flex;align-items:center;gap:0.4rem;margin-bottom:4px">
            <span style="color:var(--amber);font-size:0.85rem">⟶</span>
            <span class="rp-code-tag" style="border-color:var(--green);color:var(--green)">${escHtml(g.targetMaterial)}</span>
            <span style="color:var(--muted);font-size:0.7rem;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:180px">${escHtml(g.targetDesc || "")}</span>
          </div>
          <div style="font-size:0.7rem;color:var(--amber)">
            × ${g.conversionFactor}
            <span style="color:var(--muted)">&nbsp;qty × factor → target qty &nbsp;|&nbsp; value summed as-is</span>
          </div>
        </div>
        ${deleteBtn}
      </div>
    </div>`;
  }).join("");

  // Use a single delegated listener on the stable container element
  // (el is replaced each call so no listener accumulation)
  el.addEventListener("click", e => {
    const btn = e.target.closest(".rp-group-del");
    if (!btn) return;
    const idx = parseInt(btn.dataset.groupIdx);
    if (!isNaN(idx)) {
      // Never allow deletion of built-in rules
      if (reconcileGroups[idx]?._builtin) {
        alert("Built-in conversion rules cannot be deleted. They are required for ASA pack-size reconciliation.");
        return;
      }
      reconcileGroups.splice(idx, 1);
      invalidateReconCache();
      saveReconcileGroups();
      refreshReconcileGroupsList();
      if (rawDf.length) renderPage(currentPage);
    }
  });
}

// ── Persistence ──────────────────────────────────────────────────────────
const RECONCILE_STORE_KEY = "pharmatrack_reconcile_v3";

function saveReconcileGroups() {
  try { localStorage.setItem(RECONCILE_STORE_KEY, JSON.stringify(reconcileGroups)); } catch(e) {}
}

// FIX ROBUST: validate schema of each persisted entry before accepting it.
// Corrupt or manually edited localStorage entries could inject bad shapes that
// cause NaN propagation in the reconciliation engine.
function isValidReconcileGroup(g) {
  return (
    g !== null && typeof g === "object" &&
    typeof g.sourceMaterial   === "string" && g.sourceMaterial.trim().length > 0 &&
    typeof g.targetMaterial   === "string" && g.targetMaterial.trim().length > 0 &&
    typeof g.conversionFactor === "number" && isFinite(g.conversionFactor) && g.conversionFactor > 0
  );
}

function loadReconcileGroups() {
  try {
    const saved = localStorage.getItem(RECONCILE_STORE_KEY);
    let userRules = [];
    if (saved) {
      const parsed = JSON.parse(saved);
      if (Array.isArray(parsed)) {
        const valid = parsed.filter(isValidReconcileGroup);
        if (valid.length !== parsed.length) {
          console.warn(`PharmaTrack: ${parsed.length - valid.length} invalid reconcile rule(s) discarded on load.`);
        }
        // Strip any previously-persisted built-in rules so we always
        // use the fresh DEFAULT_RECONCILE_RULES definition.
        userRules = valid.filter(g => !g._builtin);
      }
    }

    // Merge: built-in rules first, then user-defined rules.
    // Skip any user rule whose sourceMaterial would conflict with a default.
    const builtinSources = new Set(DEFAULT_RECONCILE_RULES.map(g => g.sourceMaterial));
    const filteredUser   = userRules.filter(g => !builtinSources.has(g.sourceMaterial));
    reconcileGroups = [...DEFAULT_RECONCILE_RULES, ...filteredUser];

  } catch(e) {
    console.error("PharmaTrack: failed to load reconcile groups from localStorage:", e);
    reconcileGroups = [...DEFAULT_RECONCILE_RULES];
  }
}

// Re-persist after loading so built-in rules are always in the saved state.
// (Called once at startup, after loadReconcileGroups runs.)

// ═══════════════════════════════════════════════════════════════════════════
// HOME PAGE
// ═══════════════════════════════════════════════════════════════════════════
function renderHome() {
  // Module cards navigate to their respective page (only if data loaded)
  document.querySelectorAll(".home-module-card[data-page]").forEach(card => {
    card.onclick = () => {
      const target = card.dataset.page;
      if (!rawDf.length) {
        // No file yet — briefly highlight the upload button
        const uploadBtn = document.querySelector(".upload-btn");
        if (uploadBtn) {
          uploadBtn.style.borderColor = "var(--amber)";
          uploadBtn.style.boxShadow   = "0 0 0 3px rgba(210,153,34,0.25)";
          setTimeout(() => {
            uploadBtn.style.borderColor = "";
            uploadBtn.style.boxShadow   = "";
          }, 1600);
        }
        return;
      }
      renderPage(target);
    };
  });

  // Show summary KPIs if data is already loaded
  const kpiRow = document.getElementById("home-kpis");
  if (!rawDf.length) {
    kpiRow.style.display = "none";
    return;
  }
  kpiRow.style.display = "";

  const base       = getReconciledBase();
  const totalVal   = base.reduce((s,r) => s + r["Total Value"], 0);
  const totalQty   = base.reduce((s,r) => s + r["Total Qty"],   0);
  const transitVal = base.reduce((s,r) => s + r["Value of Stock in Transit"], 0);
  const qcVal      = base.reduce((s,r) => s + r["Value of Stock in Quality Inspection"], 0);

  const today      = new Date();
  const in90       = new Date(); in90.setDate(in90.getDate() + 90);
  const expiryCount = base.filter(r =>
    r._expiry instanceof Date && !isNaN(r._expiry) &&
    r._expiry >= today && r._expiry <= in90 &&
    (r["Unrestricted Stock"] || 0) > 0
  ).length;

  setKpis("home-kpis", [
    ["Total Inventory Value",    fmtETB(totalVal),   `${fmtQty(totalQty)} units across all plants`,      "blue"],
    ["Stock in Transit",         fmtETB(transitVal), `Moving between locations`,                          "amber"],
    ["In Quality Inspection",    fmtETB(qcVal),      `Pending QC release`,                               "red"],
    ["Expiring within 90 Days",  expiryCount.toLocaleString() + " items", `Requiring urgent action`,     "purple"],
    ["Unique Materials",         new Set(base.map(r=>r["Material"])).size.toLocaleString(), `Across ${new Set(base.map(r=>r["Plant"])).size} plants`, "green"],
  ]);
}

// ═══════════════════════════════════════════════════════════════════════════
// PAGE SWITCHING
// ═══════════════════════════════════════════════════════════════════════════
const PAGE_RENDERERS = {
  home:      renderHome,
  dashboard: renderDashboard,
  transit:   renderTransit,
  expiry:    renderExpiry,
  qc:        renderQC,
  branch:    renderBranch,
  flow:      renderFlow,
  preview:   renderPreview,
};

function renderPage(id) {
  // Home page works even before a file is loaded
  if (id !== "home" && !rawDf.length) return;
  currentPage = id;
  // Hide the pre-data landing splash whenever any page is shown
  document.getElementById("landingView").style.display = "none";
  document.querySelectorAll(".page").forEach(el => { el.style.display = "none"; });
  const pg = document.getElementById(`page-${id}`);
  if (pg) pg.style.display = "block";
  document.querySelectorAll(".nav-btn").forEach(btn => btn.classList.toggle("active", btn.dataset.page === id));
  try {
    PAGE_RENDERERS[id]?.();
  } catch(e) {
    console.error(`Error rendering ${id}:`, e);
    // Show a friendly in-page error rather than a blank page
    if (pg) pg.innerHTML = `<div class="alert-danger" style="margin-top:2rem">
      ⚠️ An error occurred while rendering this page: <b>${escHtml(e.message)}</b>
      <br><small style="opacity:0.7">Check the browser console for details.</small>
    </div>`;
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// EVENT LISTENERS
// ═══════════════════════════════════════════════════════════════════════════
document.addEventListener("DOMContentLoaded", () => {
  // Load persisted reconcile groups (always merges DEFAULT_RECONCILE_RULES)
  loadReconcileGroups();
  // Immediately re-persist so the merged state (including built-ins) is saved
  saveReconcileGroups();

  // Show Home page immediately (works without data)
  renderPage("home");

  // Nav
  document.querySelectorAll(".nav-btn[data-page]").forEach(btn => {
    btn.addEventListener("click", () => renderPage(btn.dataset.page));
  });

  // File upload
  document.getElementById("fileInput").addEventListener("change", e => {
    const f = e.target.files[0]; if (f) loadFile(f);
  });

  // Stock in Transit file upload
  document.getElementById("transitFileInput").addEventListener("change", e => {
    const f = e.target.files[0]; if (f) loadTransitFile(f);
    e.target.value = "";
  });

  // Stock in Transit section filter wiring
  document.getElementById("st-filter-apply").addEventListener("click", () => {
    stFilterState.purDoc   = (document.getElementById("st-filter-pur-doc")   || {}).value || "";
    stFilterState.supPlant = (document.getElementById("st-filter-sup-plant") || {}).value || "";
    renderStockTransitSection();
  });
  document.getElementById("st-filter-clear").addEventListener("click", () => {
    stFilterState = { purDoc: "", supPlant: "" };
    const purDocEl   = document.getElementById("st-filter-pur-doc");
    const supPlantEl = document.getElementById("st-filter-sup-plant");
    if (purDocEl)   purDocEl.value   = "";
    if (supPlantEl) supPlantEl.value = "";
    renderStockTransitSection();
  });

  // Material transit lookup
  document.getElementById("transit-search-btn").addEventListener("click", renderTransitSearch);
  document.getElementById("transit-search-clear").addEventListener("click", clearTransitSearch);
  document.getElementById("transit-search-input").addEventListener("keydown", e => {
    if (e.key === "Enter") renderTransitSearch();
  });

  // Expiry window radio
  document.getElementById("expiry-window-group").addEventListener("change", () => {
    if (rawDf.length && currentPage === "expiry") renderExpiry();
  });

  // Material expiry lookup
  document.getElementById("expiry-search-btn").addEventListener("click", renderExpirySearch);
  document.getElementById("expiry-search-clear").addEventListener("click", clearExpirySearch);
  document.getElementById("expiry-search-input").addEventListener("keydown", e => {
    if (e.key === "Enter") renderExpirySearch();
  });

  // Material QC lookup
  document.getElementById("qc-search-btn").addEventListener("click", renderQCSearch);
  document.getElementById("qc-search-clear").addEventListener("click", clearQCSearch);
  document.getElementById("qc-search-input").addEventListener("keydown", e => {
    if (e.key === "Enter") renderQCSearch();
  });

  // Material Flow lookup
  document.getElementById("flow-search-btn").addEventListener("click", renderFlowSearch);
  document.getElementById("flow-search-clear").addEventListener("click", clearFlowSearch);
  document.getElementById("flow-search-input").addEventListener("keydown", e => {
    if (e.key === "Enter") renderFlowSearch();
  });

  // Preview filters
  document.getElementById("btn-apply-filter").addEventListener("click", applyPreviewFilters);
  document.getElementById("btn-clear-filter").addEventListener("click", () => {
    document.querySelectorAll("#filter-plant option,#filter-mg option,#filter-mgname option").forEach(o => { o.selected = false; });
    filtDf = getReconciledBase();
    renderPreviewTable();
  });

  // ── Page filter wiring (event delegation) ──────────────────────────────
  // Uses document-level delegation so listeners survive any DOM rebuild
  // (e.g. the renderPage error path replaces pg.innerHTML entirely).
  // Each Apply/Clear button is identified by its stable ID.

  const PAGE_FILTER_MAP = {
    "dash-filter-apply":    { page:"dashboard", plantWrap:"ms-dash-plant",    mgWrap:"ms-dash-mg",    action:"apply" },
    "dash-filter-clear":    { page:"dashboard", plantWrap:"ms-dash-plant",    mgWrap:"ms-dash-mg",    action:"clear" },
    "transit-filter-apply": { page:"transit",   plantWrap:"ms-transit-plant", mgWrap:"ms-transit-mg", action:"apply" },
    "transit-filter-clear": { page:"transit",   plantWrap:"ms-transit-plant", mgWrap:"ms-transit-mg", action:"clear" },
    "expiry-filter-apply":  { page:"expiry",    plantWrap:"ms-expiry-plant",  mgWrap:"ms-expiry-mg",  action:"apply" },
    "expiry-filter-clear":  { page:"expiry",    plantWrap:"ms-expiry-plant",  mgWrap:"ms-expiry-mg",  action:"clear" },
    "qc-filter-apply":      { page:"qc",        plantWrap:"ms-qc-plant",      mgWrap:"ms-qc-mg",      action:"apply" },
    "qc-filter-clear":      { page:"qc",        plantWrap:"ms-qc-plant",      mgWrap:"ms-qc-mg",      action:"clear" },
    "branch-filter-apply":  { page:"branch",    plantWrap:null,               mgWrap:"ms-branch-mg",  action:"apply" },
    "branch-filter-clear":  { page:"branch",    plantWrap:null,               mgWrap:"ms-branch-mg",  action:"clear" },
    "flow-filter-apply":    { page:"flow",      plantWrap:"ms-flow-plant",    mgWrap:"ms-flow-mg",    action:"apply" },
    "flow-filter-clear":    { page:"flow",      plantWrap:"ms-flow-plant",    mgWrap:"ms-flow-mg",    action:"clear" },
  };

  document.body.addEventListener("click", (e) => {
    const btn = e.target.closest("button[id]");
    if (!btn) return;
    const cfg = PAGE_FILTER_MAP[btn.id];
    if (!cfg) return;
    if (!rawDf.length) return;

    e.stopPropagation();
    // Close any open dropdowns first
    document.querySelectorAll(".ms-wrap.open").forEach(w => w.classList.remove("open"));

    if (cfg.action === "apply") {
      if (cfg.plantWrap) {
        const wrap = document.getElementById(cfg.plantWrap);
        pageFilters[cfg.page].plants = (wrap && wrap._getSelected) ? wrap._getSelected() : [];
      }
      if (cfg.mgWrap) {
        const wrap = document.getElementById(cfg.mgWrap);
        pageFilters[cfg.page].mgs = (wrap && wrap._getSelected) ? wrap._getSelected() : [];
      }
    } else {
      if (cfg.plantWrap) {
        pageFilters[cfg.page].plants = [];
        const wrap = document.getElementById(cfg.plantWrap);
        if (wrap && wrap._clearSelected) wrap._clearSelected();
      }
      if (cfg.mgWrap) {
        pageFilters[cfg.page].mgs = [];
        const wrap = document.getElementById(cfg.mgWrap);
        if (wrap && wrap._clearSelected) wrap._clearSelected();
      }
    }
    renderPage(cfg.page);
  });

  // ── Reconciliation panel ──
  document.getElementById("open-reconcile-btn").addEventListener("click", openReconcilePanel);
  document.getElementById("reconcile-close").addEventListener("click", closeReconcilePanel);
  document.getElementById("reconcile-overlay").addEventListener("click", closeReconcilePanel);

  // Mapping file upload
  document.getElementById("rp-file-input").addEventListener("change", e => {
    const f = e.target.files[0]; if (f) handleReconcileFileUpload(f);
    e.target.value = ""; // allow re-upload of same file
  });

  // Manual source search
  document.getElementById("rp-src-search").addEventListener("input", e => {
    rpSearch(e.target.value, "rp-src-results", data => rpSetSelected("src", data));
  });

  // Manual target search
  document.getElementById("rp-tgt-search").addEventListener("input", e => {
    rpSearch(e.target.value, "rp-tgt-results", data => rpSetSelected("tgt", data));
  });

  // Add manual conversion
  document.getElementById("rp-add-manual").addEventListener("click", addManualConversion);

  // Clear all reconciliation rules
  document.getElementById("rp-clear-all").addEventListener("click", () => {
    const userRules = reconcileGroups.filter(g => !g._builtin);
    if (!userRules.length) {
      alert("No user-defined conversions to clear. Built-in rules are always retained.");
      return;
    }
    if (confirm(`Delete all ${userRules.length} user-defined conversion(s)? Built-in rules will be kept.`)) {
      // Keep only built-in rules
      reconcileGroups = reconcileGroups.filter(g => g._builtin);
      invalidateReconCache();
      saveReconcileGroups();
      refreshReconcileGroupsList();
      if (rawDf.length) renderPage(currentPage);
    }
  });
});

// ── GLOBAL MATERIAL SEARCH ─────────────────────────────────────────────────
(function () {
  function fmt(n) {
    if (n == null || isNaN(+n)) return "—";
    return (+n).toLocaleString(undefined, { maximumFractionDigits: 2 });
  }

  // FIX BUG: argument order was (cols, rows) — reversed vs the outer buildTable(rows, cols).
  // Also added fmt/raw/cellClass support so formatted cells render correctly.
  function buildTable(rows, cols) {
    if (!rows.length) return '<p class="gsr-no-data">No matching records found.</p>';
    let html = '<div class="tbl-wrap"><table><thead><tr>';
    cols.forEach(c => { html += `<th>${escHtml(c.label)}</th>`; });
    html += "</tr></thead><tbody>";
    rows.slice(0, 200).forEach(r => {
      html += "<tr>";
      cols.forEach(c => {
        const rawVal = r[c.key] ?? "";
        const display = c.fmt ? c.fmt(rawVal, r) : rawVal;
        const val = c.raw ? display : escHtml(String(display ?? ""));
        const cls = (c.cellClass || c.cls) ? ` class="${c.cellClass || c.cls}"` : "";
        html += `<td${cls}>${val}</td>`;
      });
      html += "</tr>";
    });
    html += "</tbody></table></div>";
    if (rows.length > 200) {
      html += `<p class="gsr-no-data" style="margin-top:0.4rem">Showing first 200 of ${rows.length} rows. Narrow your search for more precision.</p>`;
    }
    return html;
  }

  function showResultsPanel() {
    document.getElementById("global-search-results-panel").style.display = "block";
  }
  function hideResultsPanel() {
    document.getElementById("global-search-results-panel").style.display = "none";
  }

  function runSearch() {
    const q = (document.getElementById("global-search-input").value || "").trim().toLowerCase();
    const out = document.getElementById("global-search-results");
    if (!q) { out.innerHTML = ""; hideResultsPanel(); return; }

    // ── In-Stock results ──
    const base = rawDf;
    const stockRows = base.filter(r => {
      const code = String(r["Material"] || "").toLowerCase();
      const desc = String(r["Material Description"] || "").toLowerCase();
      return code.includes(q) || desc.includes(q);
    });

    const stockCols = [
      { key: "Material", label: "Material Code", fmt:(val,r)=>renderMatCode(val,r), raw:true, cellClass:"col-mat-code-wrap" },
      { key: "Material Description", label: "Material Description", fmt:(val,r)=>renderMatDesc(val,r), raw:true, cellClass:"col-mat-desc-wrap" },
      { key: "Plant",                label: "Plant" },
      { key: "Plant Name",           label: "Plant Name" },
      { key: "Material Group Name",  label: "Material Group" },
      { key: "Unrestricted Stock",   label: "Unrestricted Qty",  cls: "col-qty" },
      { key: "Value of Unrestricted Stock", label: "Value (ETB)", cls: "col-val" },
      { key: "Shelf Life Expiration Date",  label: "Expiry" },
    ];

    // ── Transit results (from separate transit file) ──
    const transitCols = [
      { key: "_st_material", label: "Material Code", fmt:(val,r)=>renderMatCode(val,r), raw:true, cellClass:"col-mat-code-wrap" },
      { key: "_st_desc",     label: "Material Description", fmt:(val,r)=>renderMatDesc(val,r), raw:true, cellClass:"col-mat-desc-wrap" },
      { key: "_st_purDoc",   label: "Purch. Doc." },
      { key: "_st_supPlant", label: "Supplying Plant" },
      { key: "_st_qty",      label: "Qty", cls: "col-qty" },
      { key: "_st_uom",      label: "UoM" },
    ];
    const transitRows = stockTransitRaw.filter(r => {
      const code = String(r["_st_material"] || "").toLowerCase();
      const desc = String(r["_st_desc"]     || "").toLowerCase();
      return code.includes(q) || desc.includes(q);
    });

    // ── Also search "Stock in Transit" column in main data ──
    const inTransitMain = base.filter(r => {
      const code = String(r["Material"] || "").toLowerCase();
      const desc = String(r["Material Description"] || "").toLowerCase();
      const hasTransit = parseFloat(r["Stock in Transit"] || 0) > 0;
      return hasTransit && (code.includes(q) || desc.includes(q));
    });

    let html = "";

    // In-Stock section
    html += `<div class="gsr-section-title">
      <span class="gsr-badge gsr-badge-stock">In Stock</span>
      ${stockRows.length} record${stockRows.length !== 1 ? "s" : ""} found
    </div>`;
    html += buildTable(stockRows, stockCols);

    // Transit from separate file (if uploaded)
    if (stockTransitRaw.length > 0) {
      html += `<div class="gsr-section-title" style="margin-top:1.2rem">
        <span class="gsr-badge gsr-badge-transit">In Transit (Transit File)</span>
        ${transitRows.length} record${transitRows.length !== 1 ? "s" : ""} found
      </div>`;
      html += buildTable(transitRows, transitCols);
    } else if (inTransitMain.length > 0) {
      // Fallback: show in-transit column from main data
      html += `<div class="gsr-section-title" style="margin-top:1.2rem">
        <span class="gsr-badge gsr-badge-transit">In Transit (from inventory data)</span>
        ${inTransitMain.length} record${inTransitMain.length !== 1 ? "s" : ""} found
      </div>`;
      const tCols = [
        { key: "Material", label: "Material Code", fmt:(val,r)=>renderMatCode(val,r), raw:true, cellClass:"col-mat-code-wrap" },
        { key: "Material Description", label: "Material Description", fmt:(val,r)=>renderMatDesc(val,r), raw:true, cellClass:"col-mat-desc-wrap" },
        { key: "Plant",                label: "Plant" },
        { key: "Stock in Transit",     label: "Transit Qty", cls: "col-qty" },
        { key: "Value of Stock in Transit", label: "Transit Value (ETB)", cls: "col-val" },
      ];
      html += buildTable(inTransitMain, tCols);
    }

    out.innerHTML = html;
    showResultsPanel();
  }

  function clearSearch() {
    document.getElementById("global-search-input").value = "";
    document.getElementById("global-search-results").innerHTML = "";
    hideResultsPanel();
  }

  document.addEventListener("DOMContentLoaded", () => {
    document.getElementById("global-search-btn").addEventListener("click", runSearch);
    document.getElementById("global-search-clear").addEventListener("click", clearSearch);
    document.getElementById("global-search-input").addEventListener("keydown", e => {
      if (e.key === "Enter") runSearch();
    });
    document.getElementById("global-search-results-close").addEventListener("click", hideResultsPanel);
  });
})();
