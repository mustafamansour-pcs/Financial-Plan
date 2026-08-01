// ============================================================
// HOUSEHOLD LEDGER — APP LOGIC
// Reads "Cash Flow" for the Top Sheet, writes new rows to "Actuals",
// and reads a "Users" tab (Name | PIN | Google Email | Allowed
// Categories) to gate the PIN screen and filter the Top Sheet.
// ============================================================

const SHEETS_API = "https://sheets.googleapis.com/v4/spreadsheets";

let state = {
  pin: "",
  currentUser: null,      // { name, pin, email, allowed: [...] }
  users: [],              // all rows from Users tab
  cashFlow: null,         // parsed Cash Flow data
  googleToken: null,      // OAuth access token, once signed in
  isAdmin: false,
};

// ---------- Helpers ----------

function fmt(n) {
  if (n === null || n === undefined || isNaN(n)) return "—";
  return Number(n).toLocaleString("en-US", {
    notation: "standard",
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

// Cells arrive as formatted strings like "  112,555.48 " or "  -   ".
// Strip everything except digits, minus sign, and decimal point before
// parsing, so commas and padding spaces don't corrupt the number.
function parseMoney(raw) {
  if (raw === null || raw === undefined) return 0;
  const cleaned = String(raw).replace(/[^0-9.-]/g, "");
  const n = parseFloat(cleaned);
  return isNaN(n) ? 0 : n;
}

// Sheet cells become HTML in the drill-through, and item/description
// text is free-typed by whoever added the entry. Escape it so a stray
// "&" or "<" in a shop name can't break the markup.
function esc(v) {
  return String(v === null || v === undefined ? "" : v)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

// Budget line names and typed-in item names never agree perfectly:
// "Beef " vs "beef", "Electricity Bill" vs "Electricity bill",
// Arabic-Indic digits, stray punctuation. Flatten all of that before
// comparing so a real match isn't reported as unbudgeted.
function normKey(v) {
  return String(v === null || v === undefined ? "" : v)
    .toLowerCase()
    .replace(/[\u0660-\u0669]/g, d => String(d.charCodeAt(0) - 0x0660))
    .replace(/[\u064B-\u0652]/g, "")
    .replace(/[^\p{L}\p{N}]+/gu, " ")
    .trim();
}

// Shows the cell's own formatting exactly as the sheet displays it —
// just trims the padding spaces Sheets adds around currency values,
// without touching decimals, commas, or the "-" it uses for zero.
function rawMoney(raw) {
  if (raw === null || raw === undefined || raw === "") return "—";
  const trimmed = String(raw).trim();
  return trimmed || "—";
}

const MONTH_NAMES = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

// Google Sheets returns date cells in a few different shapes depending on
// the cell's number format: a serial number, an ISO-ish string, or — as
// this sheet uses — "DD-Mon-YY" (e.g. "31-Aug-26"). The generic Date
// constructor mishandles that last format's two-digit year unreliably
// (it can land in the wrong century), so we parse it explicitly first.
function toDate(raw) {
  if (raw === null || raw === undefined || raw === "") return null;

  if (typeof raw === "number") {
    // Sheets/Excel serial date: days since 1899-12-30.
    return new Date(Date.UTC(1899, 11, 30) + raw * 86400000);
  }

  const str = String(raw).trim();

  // "31-Aug-26" / "1-Aug-26" style — DD-Mon-YY, two-digit year.
  const ddMonYY = str.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{2})$/);
  if (ddMonYY) {
    const day = parseInt(ddMonYY[1], 10);
    const monIdx = MONTH_NAMES.indexOf(ddMonYY[2]);
    const yy = parseInt(ddMonYY[3], 10);
    // Two-digit years in this sheet are always 20xx, not 19xx.
    const year = 2000 + yy;
    if (monIdx >= 0) return new Date(Date.UTC(year, monIdx, day));
  }

  // "Aug-26" style — Mon-YY, no day component (used in the month header
  // row). Default to the 1st of that month; callers only compare
  // year/month, so the exact day doesn't affect month resolution.
  const monYY = str.match(/^([A-Za-z]{3})-(\d{2})$/);
  if (monYY) {
    const monIdx = MONTH_NAMES.indexOf(monYY[1]);
    const yy = parseInt(monYY[2], 10);
    const year = 2000 + yy;
    if (monIdx >= 0) return new Date(Date.UTC(year, monIdx, 1));
  }

  const d = new Date(str);
  return isNaN(d.getTime()) ? null : d;
}

// Finds which column (row-4 header) covers "today." Prefers the real,
// always-current calendar date (device clock) so the Top Sheet advances
// automatically every month with no manual upkeep; falls back to the
// sheet's own Data Date cell only if the calendar date can't be read for
// some reason, and finally to the first month column if nothing resolves.
function findCurrentMonthColumn(calendarDate, headerRow, fallbackDataDateRaw) {
  const dataDate = calendarDate || toDate(fallbackDataDateRaw) || new Date();
  const dataYM = dataDate.getUTCFullYear() * 12 + dataDate.getUTCMonth();

  for (let c = 3; c < headerRow.length; c++) {
    const headerDate = toDate(headerRow[c]);
    if (!headerDate) continue;
    const headerYM = headerDate.getUTCFullYear() * 12 + headerDate.getUTCMonth();
    if (headerYM === dataYM) return c;
  }
  return 3; // fallback: first month column
}

async function sheetsGet(range) {
  if (!state.googleToken) throw new Error("Not signed in");
  const url = `${SHEETS_API}/${CONFIG.SHEET_ID}/values/${encodeURIComponent(range)}`;
  const res = await fetch(url, {
    headers: { "Authorization": `Bearer ${state.googleToken}` },
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.error?.message || `Sheets read failed (${res.status})`);
  }
  const json = await res.json();
  return json.values || [];
}

async function sheetsAppend(range, rowValues) {
  if (!state.googleToken) throw new Error("Not signed in");
  const url = `${SHEETS_API}/${CONFIG.SHEET_ID}/values/${encodeURIComponent(range)}:append?valueInputOption=USER_ENTERED`;
  const res = await fetch(url, {
    method: "POST",
    headers: {
      "Authorization": `Bearer ${state.googleToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ values: [rowValues] }),
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.error?.message || `Sheets write failed (${res.status})`);
  }
  return res.json();
}

// Writes to an exact cell range rather than letting Sheets guess where
// the table "ends" — needed because Actuals has pre-filled formulas
// (Ser. numbering, Total Cost) running far down the sheet, which makes
// the plain append API think the table extends much further than the
// real data does.
async function sheetsUpdate(range, rowValues) {
  if (!state.googleToken) throw new Error("Not signed in");
  const url = `${SHEETS_API}/${CONFIG.SHEET_ID}/values/${encodeURIComponent(range)}?valueInputOption=USER_ENTERED`;
  const res = await fetch(url, {
    method: "PUT",
    headers: {
      "Authorization": `Bearer ${state.googleToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ range, values: [rowValues] }),
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.error?.message || `Sheets write failed (${res.status})`);
  }
  return res.json();
}

// Finds the first genuinely empty data row by checking column B
// (Category) rather than trusting "last row with any content" — column
// A and I have formulas pre-filled hundreds of rows down, which would
// otherwise make every row look occupied.
async function findNextEmptyRow() {
  const rows = await sheetsGet(`${CONFIG.TABS.ACTUALS}!B4:B400`);
  let offset = rows.length;
  // rows[] may include trailing entries the API omits for fully blank
  // cells, so walk from the top to find the first truly empty one too.
  for (let i = 0; i < rows.length; i++) {
    if (!rows[i] || !rows[i][0]) { offset = i; break; }
  }
  return 4 + offset; // data starts at row 4
}

// ---------- PIN screen ----------

function initPinScreen() {
  const form = document.getElementById("pin-form");
  const input = document.getElementById("pin-input");
  const errEl = document.getElementById("pin-error");

  form.addEventListener("submit", async (e) => {
    e.preventDefault();
    errEl.textContent = "";
    state.pin = input.value.trim();
    if (!state.pin) return;
    await tryUnlock(errEl, input);
  });
}

async function tryUnlock(errEl, input) {
  try {
    if (!state.users.length) {
      state.users = await loadUsers();
    }
    // Case-insensitive match, since PINs mix letters and numbers.
    const match = state.users.find(
      u => u.pin.toUpperCase() === state.pin.toUpperCase()
    );
    if (!match) {
      errEl.textContent = "That PIN doesn't match anyone on the ledger.";
      if (input) { input.value = ""; input.focus(); }
      return;
    }
    state.currentUser = match;
    state.isAdmin = match.allowed.includes("ALL");
    enterApp();
  } catch (e) {
    errEl.textContent = "Couldn't reach the sheet. Check your connection.";
  }
}

async function loadUsers() {
  const rows = await sheetsGet(`${CONFIG.TABS.USERS}!A2:D`);
  return rows
    .filter(r => r[0] && r[1])
    .map(r => ({
      name: r[0],
      pin: String(r[1]).trim(),
      email: (r[2] || "").trim(),
      allowed: (r[3] || "").split(",").map(s => s.trim()).filter(Boolean),
    }));
}

// ---------- App shell ----------

function enterApp() {
  document.getElementById("pin-screen").classList.add("hidden");
  document.getElementById("app").classList.remove("hidden");
  document.getElementById("current-user").textContent = state.currentUser.name;
  document.getElementById("admin-tab").classList.toggle("hidden", !state.isAdmin);

  // If no family-photo.jpg has been uploaded yet, hide the broken image
  // and let the banner's ink background + gradient stand alone.
  const photo = document.getElementById("family-photo");
  photo.addEventListener("error", () => photo.classList.add("broken"));

  document.getElementById("f-date").valueAsDate = new Date();
  populateDropdowns();
  loadTopSheet();
}

function populateDropdowns() {
  const catSel = document.getElementById("f-category");
  CATEGORIES.forEach(c => {
    const opt = document.createElement("option");
    opt.value = c; opt.textContent = c;
    catSel.appendChild(opt);
  });

  const subSel = document.getElementById("f-subcategory");
  SUBCATEGORIES.forEach(s => {
    const opt = document.createElement("option");
    opt.value = s; opt.textContent = s;
    subSel.appendChild(opt);
  });
}

// ---------- Budget vs Actual by subcategory ----------

// Only "Home Expenses" has real subcategory-level budgets in the sheet
// (Budget!D8:E24) — every other category is tracked as a single line
// item with no sub-breakdown. We're honest about that in the UI rather
// than fabricating numbers for categories that don't have them.
const SUBCATEGORY_BUDGET_SOURCE = {
  "Home Expenses": {
    sheet: "Budget",
    range: "D8:E24",           // subcategory name | subcategory budget

    // Line-level budget for the overrun drill-through. Column D of the
    // Budget Breakdown tab holds the item name, and column D of the
    // Actuals tab holds the same item name — that shared column is what
    // the two sides are joined on.
    detail: {
      sheet: "Budget Breakdown",  // auto-corrected if the tab is named slightly differently
      range: "A2:P400",           // read wide; unused columns are ignored
      keyCol: 3,                  // column D within A2:P400 = index 3 (the item name)

      // Leave these null to let the app work them out from the data.
      // amountCol is chosen by testing each numeric column against the
      // subcategory totals in `range` above and keeping whichever one
      // reconciles; subcatCol falls back to the item→subcategory map
      // built from the Actuals tab. Set either to a zero-based index to
      // override the detection.
      amountCol: null,
      subcatCol: null,
    },
  },
};

let subcatState = { selectedCategory: null };

async function loadSubcategoryView() {
  const allowed = state.isAdmin ? CATEGORIES : state.currentUser.allowed;
  const chipsEl = document.getElementById("subcat-category-chips");

  if (!allowed.length) {
    chipsEl.innerHTML = "";
    document.getElementById("subcat-list").innerHTML =
      `<div class="empty-state">No categories assigned to you yet. Ask the admin to grant access.</div>`;
    return;
  }

  if (!subcatState.selectedCategory || !allowed.includes(subcatState.selectedCategory)) {
    subcatState.selectedCategory = allowed[0];
  }

  chipsEl.innerHTML = allowed.map(c => `
    <button type="button" class="filter-chip ${c === subcatState.selectedCategory ? 'active' : ''}" data-cat="${c}">${c}</button>
  `).join("");

  chipsEl.querySelectorAll(".filter-chip").forEach(chip => {
    chip.addEventListener("click", () => {
      subcatState.selectedCategory = chip.dataset.cat;
      loadSubcategoryView();
    });
  });

  await renderSubcategoryBreakdown(subcatState.selectedCategory);
}

async function renderSubcategoryBreakdown(category) {
  const listEl = document.getElementById("subcat-list");
  listEl.innerHTML = `<div class="empty-state">Loading subcategories…</div>`;

  const source = SUBCATEGORY_BUDGET_SOURCE[category];
  if (!source) {
    listEl.innerHTML = `<div class="empty-state">${category} is tracked as a single line item — there's no subcategory breakdown for it in the sheet.</div>`;
    document.getElementById("subcat-period-label").textContent = "";
    return;
  }

  try {
    const [budgetRows, actualsRows] = await Promise.all([
      sheetsGet(`${source.sheet}!${source.range}`),
      sheetsGet(`${CONFIG.TABS.ACTUALS}!B4:J400`),
    ]);

    const today = new Date();
    document.getElementById("subcat-period-label").textContent =
      today.toLocaleDateString("en-US", { month: "long", year: "numeric" });

    // Sum this month's actuals per subcategory, restricted to the
    // selected category and the current calendar month.
    const actualBySubcat = {};
    for (const row of actualsRows) {
      const rowCategory = (row[0] || "").trim();
      const rowSubcat = (row[1] || "").trim();
      const rowTotal = parseMoney(row[7]); // column I (Total Cost) = index 7 within B4:J
      const rowDate = toDate(row[8]);       // column J (Date) = index 8 within B4:J
      if (rowCategory !== category || !rowDate) continue;
      if (rowDate.getUTCFullYear() !== today.getFullYear() || rowDate.getUTCMonth() !== today.getMonth()) continue;
      actualBySubcat[rowSubcat] = (actualBySubcat[rowSubcat] || 0) + rowTotal;
    }

    const subcats = budgetRows
      .filter(r => r[0])
      .map(r => ({
        name: String(r[0]).trim(),
        budget: parseMoney(r[1]),
        actual: actualBySubcat[String(r[0]).trim()] || 0,
      }));

    if (!subcats.length) {
      listEl.innerHTML = `<div class="empty-state">No subcategories found for ${category}.</div>`;
      return;
    }

    listEl.innerHTML = subcats.map(s => {
      const pct = s.budget > 0 ? (s.actual / s.budget) * 100 : 0;
      const over = pct > 100;
      const maxVal = Math.max(s.budget, s.actual, 1);
      const budgetWidth = (s.budget / maxVal) * 100;
      const actualWidth = (s.actual / maxVal) * 100;
      return `
        <div class="cat-card">
          <div class="cat-card-head">
            <span class="cat-name">${s.name}</span>
            <span class="cat-pct ${over ? 'over' : 'under'}">${pct.toFixed(0)}%</span>
          </div>
          <div class="bars">
            <div class="bar-row">
              <span class="bar-label">Budget</span>
              <div class="bar-track"><div class="bar-fill budget" style="width:${budgetWidth}%"></div></div>
            </div>
            <div class="bar-row">
              <span class="bar-label">Actual</span>
              <div class="bar-track"><div class="bar-fill actual" style="width:${actualWidth}%"></div></div>
            </div>
          </div>
          <div class="cat-figures">
            <span>Budget <strong>${fmt(s.budget)}</strong></span>
            <span>Actual <strong>${fmt(s.actual)}</strong></span>
          </div>
        </div>`;
    }).join("");
  } catch (e) {
    listEl.innerHTML = `<div class="empty-state">Couldn't load subcategory data. Try refreshing.</div>`;
  }
}

// ---------- Top 10 overruns ----------

// Everything the drill-through needs is collected during the ranking
// pass and parked here, so opening a card is instant rather than a
// second round-trip to Sheets.
let overrunState = {
  details: {},     // "Category||Subcat" -> { budget, actual, over, budgetLines, actualLines }
  expanded: null,  // key of the open card, or null — accordion, one at a time
};

function overrunKey(category, name) {
  return `${category}||${name}`;
}

async function loadOverrunsView() {
  const listEl = document.getElementById("overruns-list");
  listEl.innerHTML = `<div class="empty-state">Loading overruns…</div>`;

  const allowed = state.isAdmin ? CATEGORIES : state.currentUser.allowed;
  // Only categories with real subcategory budget data can produce a
  // meaningful overrun ranking — same honest limitation as the "By
  // subcategory" tab (only Home Expenses has this in the sheet today).
  const eligibleCategories = allowed.filter(c => SUBCATEGORY_BUDGET_SOURCE[c]);

  if (!eligibleCategories.length) {
    listEl.innerHTML = `<div class="empty-state">None of your permitted categories have subcategory-level budgets in the sheet, so an overrun ranking isn't available yet.</div>`;
    document.getElementById("overruns-period-label").textContent = "";
    return;
  }

  try {
    const today = new Date();
    document.getElementById("overruns-period-label").textContent =
      today.toLocaleDateString("en-US", { month: "long", year: "numeric" });

    const actualsRows = await sheetsGet(`${CONFIG.TABS.ACTUALS}!B4:J400`);

    let allSubcats = [];
    for (const category of eligibleCategories) {
      const source = SUBCATEGORY_BUDGET_SOURCE[category];
      const budgetRows = await sheetsGet(`${source.sheet}!${source.range}`);

      // Everything the detail loader needs to check its own work: the
      // subcategory each item gets booked under (learned from Actuals,
      // joined on the item name in column D) and the subcategory budget
      // totals to reconcile against.
      const itemToSubcat = {};
      for (const row of actualsRows) {
        if ((row[0] || "").trim() !== category) continue;
        const item = normKey(row[2]);
        const sub = (row[1] || "").trim();
        if (item && sub && !itemToSubcat[item]) itemToSubcat[item] = sub;
      }
      const knownBudgets = {};
      for (const r of budgetRows) {
        if (!r[0]) continue;
        knownBudgets[String(r[0]).trim()] = parseMoney(r[1]);
      }

      const detail = await loadBudgetDetail(source, { itemToSubcat, knownBudgets });
      const budgetDetail = detail.lines;
      const detailMeta = detail.meta;

      // Keep the individual rows this month, not just the sum — the
      // drill-through is the whole point of holding on to them.
      const actualBySubcat = {};
      const linesBySubcat = {};
      for (const row of actualsRows) {
        const rowCategory = (row[0] || "").trim();
        const rowSubcat = (row[1] || "").trim();
        const rowTotal = parseMoney(row[7]);
        const rowDate = toDate(row[8]);
        if (rowCategory !== category || !rowDate) continue;
        if (rowDate.getUTCFullYear() !== today.getFullYear() || rowDate.getUTCMonth() !== today.getMonth()) continue;

        actualBySubcat[rowSubcat] = (actualBySubcat[rowSubcat] || 0) + rowTotal;
        (linesBySubcat[rowSubcat] = linesBySubcat[rowSubcat] || []).push({
          date: rowDate,
          item: row[2] || "",
          description: row[3] || "",
          uom: row[4] || "",
          qty: row[5] || "",
          unitCost: row[6] || "",
          total: rowTotal,
        });
      }

      const subcats = budgetRows
        .filter(r => r[0])
        .map(r => {
          const name = String(r[0]).trim();
          const budget = parseMoney(r[1]);
          const actual = actualBySubcat[name] || 0;
          return {
            category, name, budget, actual,
            over: actual - budget,
            budgetLines: budgetDetail[name] || [],
            detailMeta,
            actualLines: (linesBySubcat[name] || []).sort((a, b) => b.total - a.total),
          };
        });

      allSubcats = allSubcats.concat(subcats);
    }

    // Rank by how far over budget, descending — only genuine overruns.
    const overruns = allSubcats
      .filter(s => s.over > 0)
      .sort((a, b) => b.over - a.over)
      .slice(0, 10);

    if (!overruns.length) {
      listEl.innerHTML = `<div class="empty-state">No subcategories are over budget this month — nice.</div>`;
      overrunState = { details: {}, expanded: null };
      return;
    }

    // Rebuild the detail cache, but remember which card was open so a
    // refresh doesn't collapse what you were reading.
    const wasExpanded = overrunState.expanded;
    overrunState.details = {};
    for (const s of overruns) overrunState.details[overrunKey(s.category, s.name)] = s;
    overrunState.expanded = overrunState.details[wasExpanded] ? wasExpanded : null;

    listEl.innerHTML = overruns.map((s, i) => {
      const key = overrunKey(s.category, s.name);
      const open = key === overrunState.expanded;
      const pct = s.budget > 0 ? (s.actual / s.budget) * 100 : 0;
      const maxVal = Math.max(s.budget, s.actual, 1);
      const budgetWidth = (s.budget / maxVal) * 100;
      const actualWidth = (s.actual / maxVal) * 100;
      return `
        <div class="cat-card">
          <button type="button" class="drill-toggle" data-key="${esc(key)}" aria-expanded="${open}">
            <span class="cat-card-head">
              <span class="cat-name">${i + 1}. ${esc(s.name)} <span class="overrun-parent">(${esc(s.category)})</span></span>
              <span class="cat-pct over">+${fmt(s.over)}</span>
            </span>
            <svg class="drill-chevron ${open ? 'open' : ''}" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M6 9l6 6 6-6"/></svg>
          </button>
          <div class="bars">
            <div class="bar-row">
              <span class="bar-label">Budget</span>
              <div class="bar-track"><div class="bar-fill budget" style="width:${budgetWidth}%"></div></div>
            </div>
            <div class="bar-row">
              <span class="bar-label">Actual</span>
              <div class="bar-track"><div class="bar-fill actual" style="width:${actualWidth}%"></div></div>
            </div>
          </div>
          <div class="cat-figures">
            <span>Budget <strong>${fmt(s.budget)}</strong></span>
            <span>Actual <strong>${fmt(s.actual)}</strong></span>
            <span>${pct.toFixed(0)}% used</span>
          </div>
          <div class="drill-panel ${open ? '' : 'hidden'}" data-key="${esc(key)}">${open ? renderOverrunDetail(key) : ''}</div>
        </div>`;
    }).join("");

    attachOverrunDrill(listEl);
  } catch (e) {
    listEl.innerHTML = `<div class="empty-state">Couldn't load overruns. Try refreshing.</div>`;
  }
}

// Pulls line-level budget rows and groups them under their subcategory.
// Detail is a bonus, never a dependency — if the range is missing or
// misconfigured, the ranking still renders and the panel just falls
// back to the single budget figure.
// Reads the Budget Breakdown tab and returns budget lines grouped by
// subcategory, keyed on the item name in column D. Which column holds
// the money and which holds the subcategory is worked out from the
// data rather than assumed, because getting either wrong silently
// turns every item into a phantom "unmatched" row.
async function loadBudgetDetail(source, ctx) {
  const cfg = source.detail;
  if (!cfg) return { lines: {}, meta: null };

  let rows, sheetName = cfg.sheet;
  try {
    rows = await sheetsGet(`${sheetName}!${cfg.range}`);
  } catch (e) {
    // Tab name didn't resolve — find the closest match and retry once,
    // so a stray capital or double space isn't fatal.
    sheetName = await resolveTabName(cfg.sheet);
    if (!sheetName) return { lines: {}, meta: { error: `Couldn't find a tab named "${cfg.sheet}".` } };
    try {
      rows = await sheetsGet(`${sheetName}!${cfg.range}`);
    } catch (e2) {
      return { lines: {}, meta: { error: `Couldn't read ${sheetName}!${cfg.range}.` } };
    }
  }

  const keyCol = cfg.keyCol ?? 3;
  const entries = rows
    .map(r => ({ row: r, item: String(r[keyCol] || "").trim() }))
    .filter(e => e.item);

  if (!entries.length) {
    return { lines: {}, meta: { error: `No item names found in ${sheetName}!${cfg.range} at column index ${keyCol}.` } };
  }

  // Which subcategory does each line belong to? Prefer an explicit
  // column; otherwise infer it from where that item actually gets
  // booked in the Actuals tab.
  const subcatOf = (e) => {
    if (cfg.subcatCol !== null && cfg.subcatCol !== undefined) {
      return String(e.row[cfg.subcatCol] || "").trim();
    }
    return ctx.itemToSubcat[normKey(e.item)] || "";
  };

  // Score every plausible money column by how closely its per-subcategory
  // sums reproduce the subcategory budgets already read from the Budget
  // tab. The real budget column reconciles; a quantity or unit-cost
  // column doesn't.
  let amountCol = cfg.amountCol;
  let confidence = null;
  if (amountCol === null || amountCol === undefined) {
    const width = Math.max(...entries.map(e => e.row.length));
    let best = null;
    for (let j = 0; j < width; j++) {
      if (j === keyCol) continue;
      const numeric = entries.filter(e => parseMoney(e.row[j]) > 0).length;
      if (numeric < 2) continue;

      const sums = {};
      for (const e of entries) {
        const sub = subcatOf(e);
        if (!sub) continue;
        sums[sub] = (sums[sub] || 0) + parseMoney(e.row[j]);
      }

      let diff = 0, base = 0, compared = 0;
      for (const sub of Object.keys(ctx.knownBudgets)) {
        if (!(sub in sums)) continue;
        diff += Math.abs(sums[sub] - ctx.knownBudgets[sub]);
        base += ctx.knownBudgets[sub];
        compared++;
      }
      if (!compared || !base) continue;

      const score = diff / base;
      if (!best || score < best.score) best = { col: j, score, compared };
    }
    if (best) {
      amountCol = best.col;
      confidence = best.score;
    }
  }

  if (amountCol === null || amountCol === undefined) {
    return { lines: {}, meta: { sheet: sheetName, error: "Couldn't identify which column holds the budgeted amount. Set detail.amountCol to a zero-based index." } };
  }

  const grouped = {};
  for (const e of entries) {
    const sub = subcatOf(e);
    if (!sub) continue;
    const amount = parseMoney(e.row[amountCol]);
    if (!amount) continue;
    (grouped[sub] = grouped[sub] || []).push({ label: e.item, amount });
  }

  return {
    lines: grouped,
    meta: {
      sheet: sheetName,
      keyCol,
      amountCol,
      confidence,
      detected: cfg.amountCol === null || cfg.amountCol === undefined,
      lineCount: entries.length,
    },
  };
}

// Finds the real tab name when the configured one doesn't resolve —
// case, spacing and punctuation insensitive.
async function resolveTabName(wanted) {
  try {
    const url = `${SHEETS_API}/${CONFIG.SHEET_ID}?fields=sheets.properties.title`;
    const res = await fetch(url, { headers: { "Authorization": `Bearer ${state.googleToken}` } });
    if (!res.ok) return null;
    const json = await res.json();
    const titles = (json.sheets || []).map(s => s.properties.title);
    const target = normKey(wanted);
    return titles.find(t => normKey(t) === target)
      || titles.find(t => normKey(t).includes(target) || target.includes(normKey(t)))
      || null;
  } catch (e) {
    return null;
  }
}

function renderOverrunDetail(key) {
  const d = overrunState.details[key];
  if (!d) return "";

  // Roll the month's entries up to one row per item, so an item bought
  // four times reads as a single line worth comparing to a budget line
  // rather than four fragments of one.
  const byItem = {};
  for (const r of d.actualLines) {
    const name = String(r.item || "").trim() || "(unnamed)";
    const k = name.toLowerCase();
    if (!byItem[k]) byItem[k] = { name, total: 0, count: 0, first: r.date, last: r.date };
    byItem[k].total += r.total;
    byItem[k].count += 1;
    if (r.date < byItem[k].first) byItem[k].first = r.date;
    if (r.date > byItem[k].last) byItem[k].last = r.date;
  }

  const body = d.budgetLines.length
    ? renderLineComparison(d, byItem)
    : renderCumulativeBurn(d, byItem);

  const pct = d.budget > 0 ? (d.actual / d.budget) * 100 : 0;

  return `
    ${body}
    <div class="drill-verdict">
      <span>Budget <strong>${fmt(d.budget)}</strong></span>
      <span>Actual <strong>${fmt(d.actual)}</strong></span>
      <span class="drill-verdict-over">Over by <strong>${fmt(d.over)}</strong>${d.budget > 0 ? ` (${pct.toFixed(0)}%)` : ""}</span>
    </div>`;
}

// --- Mode A: real budget lines exist, so compare like for like -------
// Every budgeted line gets its actual alongside it, and anything spent
// with no budget line behind it is called out as unbudgeted — that's
// usually where an overrun is actually hiding.
function renderLineComparison(d, byItem) {
  // Index the month's items by normalised name once, then match each
  // budget line against it: exact first, then containment either way
  // ("Beef" budget line vs "Beef mince 1kg" entry). One actual item
  // can only be claimed by one budget line.
  const index = Object.values(byItem).map(it => ({ it, key: normKey(it.name) }));
  const claimed = new Set();

  function findMatch(label) {
    const target = normKey(label);
    if (!target) return null;
    let hit = index.find(e => !claimed.has(e.it.name) && e.key === target);
    if (!hit) {
      hit = index.find(e => !claimed.has(e.it.name) && e.key &&
        (e.key.includes(target) || target.includes(e.key)));
    }
    if (hit) claimed.add(hit.it.name);
    return hit ? hit.it : null;
  }

  const rows = d.budgetLines.map(l => {
    const hit = findMatch(l.label);
    return {
      label: l.label,
      budget: l.amount,
      actual: hit ? hit.total : 0,
      count: hit ? hit.count : 0,
      unbudgeted: false,
    };
  });

  for (const e of index) {
    if (claimed.has(e.it.name)) continue;
    rows.push({
      label: e.it.name,
      budget: 0,
      actual: e.it.total,
      count: e.it.count,
      unbudgeted: true,
    });
  }

  // Worst variance first — the reason you opened the card.
  rows.sort((a, b) => (b.actual - b.budget) - (a.actual - a.budget));

  const totBudget = rows.reduce((s, r) => s + r.budget, 0);
  const totActual = rows.reduce((s, r) => s + r.actual, 0);

  return `
    <div class="drill-section">
      <div class="drill-head">Line by line vs budget</div>
      <table class="drill-table drill-compare">
        <thead>
          <tr>
            <th>Item</th>
            <th class="num">Budget</th>
            <th class="num">Actual</th>
            <th class="num">Variance</th>
          </tr>
        </thead>
        <tbody>
          ${rows.map(r => {
            const v = r.actual - r.budget;
            const cls = v > 0.005 ? "over" : (v < -0.005 ? "under" : "level");
            return `
              <tr class="${r.unbudgeted ? 'unbudgeted' : ''}">
                <td>
                  <span class="drill-item">${esc(r.label)}</span>
                  <span class="drill-meta">${r.unbudgeted ? `no budget line matched · ${r.count} ${r.count === 1 ? "entry" : "entries"}` : (r.count ? `${r.count} ${r.count === 1 ? "entry" : "entries"}` : "nothing spent")}</span>
                </td>
                <td class="num">${r.budget ? fmt(r.budget) : "—"}</td>
                <td class="num">${r.actual ? fmt(r.actual) : "—"}</td>
                <td class="num drill-var ${cls}">${v > 0 ? "+" : ""}${fmt(v)}</td>
              </tr>`;
          }).join("")}
        </tbody>
        <tfoot>
          <tr>
            <td>Total</td>
            <td class="num">${fmt(totBudget)}</td>
            <td class="num">${fmt(totActual)}</td>
            <td class="num drill-var ${totActual - totBudget > 0 ? "over" : "under"}">${totActual - totBudget > 0 ? "+" : ""}${fmt(totActual - totBudget)}</td>
          </tr>
        </tfoot>
      </table>
      ${renderDetailSource(d)}
      ${rows.some(r => r.unbudgeted) ? `<p class="drill-foot">"No budget line matched" means the item name in Actuals didn't line up with any line in the configured budget range — not that it's missing from your budget. Run <code>LedgerDebug.peek()</code> in the console to see exactly which budget rows the app is reading.</p>` : ""}
      ${Math.abs(totBudget - d.budget) > 0.01 ? `<p class="drill-foot">These lines sum to ${fmt(totBudget)}, but the subcategory total on the Budget sheet is ${fmt(d.budget)}. The ranking uses the sheet total.</p>` : ""}
    </div>`;
}

// --- Mode B: one budget figure, no lines behind it -------------------
// Without line-level budget there's nothing to compare each item to,
// so compare the running total instead: spend biggest-first and mark
// the point where the budget ran out. Everything below that line is,
// quite literally, the overrun.
function renderCumulativeBurn(d, byItem) {
  const items = Object.values(byItem).sort((a, b) => b.total - a.total);

  if (!items.length) {
    return `
      <div class="drill-section">
        <div class="drill-head">vs budget</div>
        <p class="drill-empty">The sheet shows ${fmt(d.actual)} spent against a budget of ${fmt(d.budget)}, but no individual entries match this subcategory this month — the sub category name in the Actuals tab probably doesn't match the Budget tab exactly.</p>
      </div>`;
  }

  let running = 0;
  let breached = false;
  const rows = items.map(it => {
    const before = running;
    running += it.total;
    const crossesHere = !breached && running > d.budget && d.budget > 0;
    if (crossesHere) breached = true;
    const remaining = d.budget - running;
    return `
      <tr class="${before >= d.budget && d.budget > 0 ? 'past-budget' : ''}">
        <td>
          <span class="drill-item">${esc(it.name)}</span>
          <span class="drill-meta">${it.count} ${it.count === 1 ? "entry" : "entries"} · ${it.first.toLocaleDateString("en-US", { month: "short", day: "numeric" })}${it.count > 1 ? `–${it.last.toLocaleDateString("en-US", { month: "short", day: "numeric" })}` : ""}</span>
        </td>
        <td class="num">${fmt(it.total)}</td>
        <td class="num drill-var ${remaining < 0 ? "over" : "under"}">${remaining < 0 ? "+" + fmt(-remaining) : fmt(remaining)}</td>
      </tr>
      ${crossesHere ? `<tr class="drill-breach"><td colspan="3">Budget of ${fmt(d.budget)} exhausted here — everything below is over</td></tr>` : ""}`;
  }).join("");

  return `
    <div class="drill-section">
      <div class="drill-head">vs budget — biggest first</div>
      <table class="drill-table drill-compare">
        <thead>
          <tr>
            <th>Item</th>
            <th class="num">Spent</th>
            <th class="num">Left of budget</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
        <tfoot>
          <tr>
            <td>${items.length} ${items.length === 1 ? "item" : "items"}</td>
            <td class="num">${fmt(d.actual)}</td>
            <td class="num drill-var over">+${fmt(d.over)}</td>
          </tr>
        </tfoot>
      </table>
      <p class="drill-foot">${d.detailMeta && d.detailMeta.error
        ? esc(d.detailMeta.error) + " Falling back to the single subcategory figure."
        : `No budget lines on the Budget Breakdown tab roll up to this subcategory, so there's no per-item budget to compare against here.`}</p>
    </div>`;
}

// Says out loud which tab and column the budget figures came from —
// column detection is a guess until it's confirmed, and a wrong guess
// should be visible rather than quietly skewing every comparison.
function renderDetailSource(d) {
  const m = d.detailMeta;
  if (!m || m.error) return "";
  const col = i => String.fromCharCode(65 + i);
  const conf = m.detected && m.confidence !== null
    ? (m.confidence < 0.02
        ? " — reconciles with the subcategory total"
        : ` — off the subcategory total by ${(m.confidence * 100).toFixed(0)}%, so check it`)
    : "";
  return `<p class="drill-foot">Budget lines from <strong>${esc(m.sheet)}</strong>, amounts in column ${col(m.amountCol)}${m.detected ? " (auto-detected" + conf + ")" : ""}, matched on item name in column ${col(m.keyCol)}.</p>`;
}

function attachOverrunDrill(listEl) {
  listEl.querySelectorAll(".drill-toggle").forEach(btn => {
    btn.addEventListener("click", () => {
      const key = btn.dataset.key;
      overrunState.expanded = overrunState.expanded === key ? null : key;

      listEl.querySelectorAll(".drill-panel").forEach(panel => {
        const open = panel.dataset.key === overrunState.expanded;
        // Render on first open only — ten panels of line items is a lot
        // of DOM to build for cards nobody taps.
        if (open && !panel.dataset.rendered) {
          panel.innerHTML = renderOverrunDetail(panel.dataset.key);
          panel.dataset.rendered = "1";
        }
        panel.classList.toggle("hidden", !open);
      });

      listEl.querySelectorAll(".drill-toggle").forEach(b => {
        const open = b.dataset.key === overrunState.expanded;
        b.setAttribute("aria-expanded", String(open));
        b.querySelector(".drill-chevron").classList.toggle("open", open);
      });
    });
  });
}

// ---------- Console inspector (config troubleshooting) ----------
// Nothing calls these; they exist so the budget range can be pointed
// at the right tab without guessing. Open the browser console after
// signing in and run LedgerDebug.tabs().

const LedgerDebug = {
  // Lists every tab name in the spreadsheet, exactly as the API sees
  // them — spelling and spacing included, which is what the range
  // string has to match.
  async tabs() {
    const url = `${SHEETS_API}/${CONFIG.SHEET_ID}?fields=sheets.properties(title,gridProperties)`;
    const res = await fetch(url, { headers: { "Authorization": `Bearer ${state.googleToken}` } });
    const json = await res.json();
    const rows = (json.sheets || []).map(s => ({
      tab: s.properties.title,
      rows: s.properties.gridProperties?.rowCount,
      cols: s.properties.gridProperties?.columnCount,
    }));
    console.table(rows);
    return rows;
  },

  // Dumps a range as a grid with the column letter AND the zero-based
  // index that detailCols wants, so there's no counting by hand.
  // e.g. LedgerDebug.peek("Budget Breakdown!A1:L40")
  async peek(range) {
    const source = SUBCATEGORY_BUDGET_SOURCE["Home Expenses"];
    const target = range || `${source.sheet}!${source.detailRange || source.range}`;
    const rows = await sheetsGet(target);
    console.log(`%c${target} — ${rows.length} rows`, "font-weight:bold");
    const startCol = (target.match(/!([A-Z]+)/) || [, "A"])[1];
    const base = startCol.split("").reduce((n, ch) => n * 26 + (ch.charCodeAt(0) - 64), 0) - 1;
    console.table(rows.map((r, i) => {
      const o = { "#": i };
      r.forEach((cell, j) => {
        const letter = String.fromCharCode(65 + base + j);
        o[`${letter} (idx ${j})`] = cell;
      });
      return o;
    }));
    return rows;
  },
};

window.LedgerDebug = LedgerDebug;

// ---------- Budget breakdown ----------

// Same table treatment as Actual records, minus the date filters —
// budget lines aren't dated. Columns aren't hard-coded: row 1 of the
// tab supplies the headers, so adding a column to the sheet shows up
// here without a code change.
const BREAKDOWN_SOURCE = { sheet: "Budget Breakdown", range: "A1:P400" };

let breakdownState = { headers: [], rows: [], numericCols: [], catCol: -1 };

async function loadBreakdownView() {
  const tbody = document.getElementById("breakdown-tbody");
  const searchEl = document.getElementById("breakdown-search");
  tbody.innerHTML = `<tr><td class="empty-state">Loading budget breakdown…</td></tr>`;

  try {
    let sheetName = BREAKDOWN_SOURCE.sheet;
    let raw;
    try {
      raw = await sheetsGet(`${sheetName}!${BREAKDOWN_SOURCE.range}`);
    } catch (e) {
      sheetName = await resolveTabName(BREAKDOWN_SOURCE.sheet);
      if (!sheetName) throw new Error(`Couldn't find a tab named "${BREAKDOWN_SOURCE.sheet}".`);
      raw = await sheetsGet(`${sheetName}!${BREAKDOWN_SOURCE.range}`);
    }

    if (!raw.length) {
      tbody.innerHTML = `<tr><td class="empty-state">That tab is empty.</td></tr>`;
      return;
    }

    const width = Math.max(...raw.map(r => r.length));
    const headerRow = raw[0] || [];
    let rows = raw.slice(1).filter(r => r.some(c => String(c || "").trim()));

    // Drop columns that are blank top to bottom — sheets are usually
    // read wider than they are.
    const keep = [];
    for (let j = 0; j < width; j++) {
      const hasHeader = String(headerRow[j] || "").trim();
      const hasData = rows.some(r => String(r[j] || "").trim());
      if (hasHeader || hasData) keep.push(j);
    }

    const headers = keep.map(j => String(headerRow[j] || "").trim() || colLetter(j));

    // A column whose values are mostly real category names is the
    // category column — used to respect each person's permissions.
    let catCol = -1;
    if (typeof CATEGORIES !== "undefined") {
      const catKeys = new Set(CATEGORIES.map(normKey));
      keep.forEach((j, idx) => {
        const vals = rows.map(r => String(r[j] || "").trim()).filter(Boolean);
        if (!vals.length) return;
        const hits = vals.filter(v => catKeys.has(normKey(v))).length;
        if (catCol === -1 && hits / vals.length > 0.5) catCol = idx;
      });
    }

    const allowed = state.isAdmin ? null : state.currentUser.allowed;
    let table = rows.map(r => keep.map(j => String(r[j] === undefined || r[j] === null ? "" : r[j])));
    if (allowed && catCol >= 0) {
      table = table.filter(r => allowed.includes(r[catCol].trim()));
    }

    // Right-align and total the columns that are predominantly numbers.
    const numericCols = headers.map((_, idx) => {
      const vals = table.map(r => r[idx]).filter(v => String(v).trim());
      if (vals.length < 2) return false;
      const nums = vals.filter(v => /\d/.test(v) && parseMoney(v) !== 0).length;
      return nums / vals.length > 0.6;
    });

    breakdownState = { headers, rows: table, numericCols, catCol, sheetName };

    document.getElementById("breakdown-source").innerHTML =
      `Read live from <strong>${esc(sheetName)}</strong>${allowed && catCol >= 0 ? " — showing only the categories you have access to." : "."}`;

    const apply = () => renderBreakdownTable(searchEl.value);
    apply();
    searchEl.oninput = apply;
  } catch (e) {
    tbody.innerHTML = `<tr><td class="empty-state">${esc(e.message || "Couldn't load the budget breakdown.")}</td></tr>`;
    document.getElementById("breakdown-thead").innerHTML = "";
    document.getElementById("breakdown-tfoot").innerHTML = "";
    document.getElementById("breakdown-count").textContent = "";
  }
}

function colLetter(j) {
  let s = "", n = j;
  do { s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26) - 1; } while (n >= 0);
  return s;
}

function renderBreakdownTable(query) {
  const { headers, rows, numericCols } = breakdownState;
  const thead = document.getElementById("breakdown-thead");
  const tbody = document.getElementById("breakdown-tbody");
  const tfoot = document.getElementById("breakdown-tfoot");
  const q = normKey(query || "");

  const filtered = q ? rows.filter(r => r.some(c => normKey(c).includes(q))) : rows;

  thead.innerHTML = `<tr>${headers.map((h, i) =>
    `<th class="${numericCols[i] ? "num" : ""}">${esc(h)}</th>`).join("")}</tr>`;

  document.getElementById("breakdown-count").textContent =
    `${filtered.length} of ${rows.length} lines`;

  if (!filtered.length) {
    tbody.innerHTML = `<tr><td colspan="${headers.length}" class="empty-state">Nothing matches that search.</td></tr>`;
    tfoot.innerHTML = "";
    return;
  }

  tbody.innerHTML = filtered.map(r => `
    <tr>${r.map((c, i) => `<td class="${numericCols[i] ? "num" : ""}">${esc(c)}</td>`).join("")}</tr>
  `).join("");

  // Totals for the numeric columns, reflecting the current search.
  const anyNumeric = numericCols.some(Boolean);
  tfoot.innerHTML = anyNumeric
    ? `<tr>${headers.map((_, i) => {
        if (!numericCols[i]) return `<td>${i === 0 ? "Total" : ""}</td>`;
        const sum = filtered.reduce((s, r) => s + parseMoney(r[i]), 0);
        return `<td class="num">${fmt(sum)}</td>`;
      }).join("")}</tr>`
    : "";
}

// ---------- Actual records table ----------

async function loadRecordsView() {
  const tbody = document.getElementById("records-tbody");
  const filterSel = document.getElementById("records-subcat-filter");
  const dateFrom = document.getElementById("records-date-from");
  const dateTo = document.getElementById("records-date-to");
  tbody.innerHTML = `<tr><td colspan="9" class="empty-state">Loading records…</td></tr>`;

  try {
    const rows = await sheetsGet(`${CONFIG.TABS.ACTUALS}!B4:J400`);
    const allowed = state.isAdmin ? null : state.currentUser.allowed;

    const records = rows
      .filter(r => r[0]) // has a category
      .filter(r => !allowed || allowed.includes(String(r[0]).trim()))
      .map(r => ({
        category: r[0] || "",
        subcategory: r[1] || "",
        item: r[2] || "",
        description: r[3] || "",
        uom: r[4] || "",
        qty: r[5] || "",
        unitCost: r[6] || "",
        total: r[7] || "",
        date: toDate(r[8]),
      }))
      .sort((a, b) => (b.date?.getTime() || 0) - (a.date?.getTime() || 0));

    // Populate the sub category filter from whatever's actually present
    // in the user's visible records, not the full master list — so the
    // dropdown never offers a sub category they have no entries for.
    const subcatsPresent = [...new Set(records.map(r => r.subcategory).filter(Boolean))].sort();
    const currentFilter = filterSel.value;
    filterSel.innerHTML = `<option value="">All sub categories</option>` +
      subcatsPresent.map(s => `<option value="${s}">${s}</option>`).join("");
    filterSel.value = subcatsPresent.includes(currentFilter) ? currentFilter : "";

    const applyFilters = () => renderRecordsTable(records, filterSel.value, dateFrom.value, dateTo.value);
    applyFilters();

    filterSel.onchange = applyFilters;
    dateFrom.onchange = applyFilters;
    dateTo.onchange = applyFilters;
  } catch (e) {
    tbody.innerHTML = `<tr><td colspan="9" class="empty-state">Couldn't load records. Try refreshing.</td></tr>`;
  }
}

function renderRecordsTable(records, subcatFilter, dateFromStr, dateToStr) {
  const tbody = document.getElementById("records-tbody");
  let filtered = subcatFilter ? records.filter(r => r.subcategory === subcatFilter) : records;

  if (dateFromStr) {
    const from = new Date(dateFromStr + "T00:00:00Z");
    filtered = filtered.filter(r => r.date && r.date.getTime() >= from.getTime());
  }
  if (dateToStr) {
    const to = new Date(dateToStr + "T23:59:59Z");
    filtered = filtered.filter(r => r.date && r.date.getTime() <= to.getTime());
  }

  if (!filtered.length) {
    tbody.innerHTML = `<tr><td colspan="9" class="empty-state">No records found.</td></tr>`;
    return;
  }

  tbody.innerHTML = filtered.map(r => `
    <tr>
      <td>${r.date ? r.date.toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" }) : "—"}</td>
      <td>${r.category}</td>
      <td>${r.subcategory}</td>
      <td>${r.item}</td>
      <td>${r.description}</td>
      <td>${r.uom}</td>
      <td class="num">${r.qty}</td>
      <td class="num">${r.unitCost}</td>
      <td class="num">${r.total}</td>
    </tr>
  `).join("");
}

// ---------- Tab navigation ----------

document.querySelectorAll(".tab-btn").forEach(btn => {
  btn.addEventListener("click", () => {
    document.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("active"));
    btn.classList.add("active");
    document.querySelectorAll(".view").forEach(v => v.classList.add("hidden"));
    document.getElementById(`view-${btn.dataset.view}`).classList.remove("hidden");
    if (btn.dataset.view === "admin") renderAdmin();
    if (btn.dataset.view === "breakdown") loadBreakdownView();
    if (btn.dataset.view === "subcat") loadSubcategoryView();
    if (btn.dataset.view === "records") loadRecordsView();
    if (btn.dataset.view === "overruns") loadOverrunsView();
  });
});

document.getElementById("refresh-btn").addEventListener("click", () => {
  loadTopSheet();
});

// ---------- Top Sheet ----------

async function loadTopSheet() {
  const listEl = document.getElementById("cat-list");
  listEl.innerHTML = `<div class="empty-state">Refreshing your numbers…</div>`;

  try {
    // Cash Flow layout: col A = category label, col B = "Budget"/"Actual Cost",
    // col C = 12-month TOTAL, columns D onward = one column per month
    // (row 4 holds each column's month-end date). We find "today's" column
    // by matching the real calendar date against row 4's headers, so the
    // Top Sheet always reflects the current month without manual updates.
    // Rows ~21-24 hold the overall Total Budget / Total Actual Cost / ETC /
    // Savings-Overruns figures shown at the bottom of the view.
    const rows = await sheetsGet(`${CONFIG.TABS.CASH_FLOW}!A1:P30`);

    const dataDateRaw = (rows[1] || [])[1]; // B2, kept only as a fallback
    const headerRow = rows[3] || [];        // row 4: month-end dates start at col D (index 3)
    // Use the real, always-current calendar date rather than trusting the
    // sheet's own Data Date cell to be kept up to date manually — this way
    // the correct month is picked every day with zero maintenance.
    const today = new Date();
    const monthCol = findCurrentMonthColumn(today, headerRow, dataDateRaw);

    const categories = [];
    let totals = { budget: null, actual: null, etc: null, savings: null };

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const label = (row[0] || "").trim();
      const kind = (row[1] || "").trim();

      if (kind === "Budget" && label && label !== "Total") {
        const budget = parseMoney(row[monthCol]);
        const budgetRaw = rawMoney(row[monthCol]);
        const nextRow = rows[i + 1] || [];
        const actual = parseMoney(nextRow[monthCol]);
        const actualRaw = rawMoney(nextRow[monthCol]);
        categories.push({ name: label, budget, budgetRaw, actual, actualRaw });
      }

      if (kind === "Total Budget") totals.budget = rawMoney(row[monthCol]);
      if (kind.replace(/\s+/g, " ") === "Total Actual Cost") totals.actual = rawMoney(row[monthCol]);
      if (label.startsWith("ETC")) totals.etc = rawMoney(row[monthCol]);
      if (label.startsWith("Cuurent Budget") || label.startsWith("Current Budget")) totals.savings = rawMoney(row[monthCol]);
    }

    state.cashFlow = { categories, totals };
    const resolvedMonthDate = toDate(headerRow[monthCol]);
    document.getElementById("period-label").textContent = resolvedMonthDate
      ? resolvedMonthDate.toLocaleDateString("en-US", { month: "long", year: "numeric" })
      : "This month";
    renderTopSheet();
  } catch (e) {
    listEl.innerHTML = `<div class="empty-state">Couldn't load the sheet. Pull down or tap refresh to try again.</div>`;
  }
}

function renderTopSheet() {
  const { categories, totals } = state.cashFlow;
  const allowed = state.isAdmin ? null : state.currentUser.allowed;
  const visible = allowed ? categories.filter(c => allowed.includes(c.name)) : categories;

  const totalBudget = visible.reduce((s, c) => s + c.budget, 0);
  const totalActual = visible.reduce((s, c) => s + c.actual, 0);

  // Admins see the sheet's own Total Budget / Total Actual Cost rows —
  // exact, cent-for-cent match to the household total. Restricted users
  // instead see the sum of just their permitted categories, since the
  // sheet has no single cell representing "this person's subset" —
  // computed here from the same category cards shown below.
  if (state.isAdmin) {
    document.getElementById("sum-budget").textContent = totals.budget || "—";
    document.getElementById("sum-actual").textContent = totals.actual || "—";
    const remaining = parseMoney(totals.budget) - parseMoney(totals.actual);
    document.getElementById("sum-remaining").textContent = fmt(remaining);
  } else {
    document.getElementById("sum-budget").textContent = fmt(totalBudget);
    document.getElementById("sum-actual").textContent = fmt(totalActual);
    document.getElementById("sum-remaining").textContent = fmt(totalBudget - totalActual);
  }

  // Bottom totals card only makes sense as whole-household figures — hide
  // it for members restricted to a single category, since ETC/Savings are
  // household-wide numbers from the sheet, not per-category subsets.
  const totalsCard = document.getElementById("totals-card");
  if (state.isAdmin) {
    totalsCard.classList.remove("hidden");
    document.getElementById("etc-total-budget").textContent = totals.budget || "—";
    document.getElementById("etc-total-actual").textContent = totals.actual || "—";
    document.getElementById("etc-remaining").textContent = totals.etc || "—";
    document.getElementById("etc-savings").textContent = totals.savings || "—";
  } else {
    totalsCard.classList.add("hidden");
  }

  const listEl = document.getElementById("cat-list");
  if (!visible.length) {
    listEl.innerHTML = `<div class="empty-state">No categories assigned to you yet. Ask the admin to grant access.</div>`;
    return;
  }

  listEl.innerHTML = visible.map(c => {
    const pct = c.budget > 0 ? (c.actual / c.budget) * 100 : 0;
    const over = pct > 100;
    // Each bar is scaled against the larger of budget/actual for that
    // category, so the two bars stay visually comparable to one another
    // rather than each maxing out independently.
    const maxVal = Math.max(c.budget, c.actual, 1);
    const budgetWidth = (c.budget / maxVal) * 100;
    const actualWidth = (c.actual / maxVal) * 100;
    return `
      <div class="cat-card">
        <div class="cat-card-head">
          <span class="cat-name">${c.name}</span>
          <span class="cat-pct ${over ? 'over' : 'under'}">${pct.toFixed(0)}%</span>
        </div>
        <div class="bars">
          <div class="bar-row">
            <span class="bar-label">Budget</span>
            <div class="bar-track"><div class="bar-fill budget" style="width:${budgetWidth}%"></div></div>
          </div>
          <div class="bar-row">
            <span class="bar-label">Actual</span>
            <div class="bar-track"><div class="bar-fill actual" style="width:${actualWidth}%"></div></div>
          </div>
        </div>
        <div class="cat-figures">
          <span>Budget <strong>${c.budgetRaw}</strong></span>
          <span>Actual <strong>${c.actualRaw}</strong></span>
        </div>
      </div>`;
  }).join("");
}

// ---------- Add Entry ----------

function recalcTotal() {
  const qty = parseFloat(document.getElementById("f-qty").value) || 0;
  const unit = parseFloat(document.getElementById("f-unitcost").value) || 0;
  document.getElementById("f-total").textContent = (qty * unit).toFixed(2);
}
document.getElementById("f-qty").addEventListener("input", recalcTotal);
document.getElementById("f-unitcost").addEventListener("input", recalcTotal);

document.getElementById("entry-form").addEventListener("submit", async (e) => {
  e.preventDefault();
  const statusEl = document.getElementById("form-status");
  const btn = document.getElementById("submit-btn");
  statusEl.textContent = "";
  statusEl.className = "form-status";
  btn.disabled = true;

  try {
    const date = document.getElementById("f-date").value;
    const category = document.getElementById("f-category").value;
    const subcategory = document.getElementById("f-subcategory").value;
    const item = document.getElementById("f-item").value;
    const description = document.getElementById("f-description").value;
    const uom = document.getElementById("f-uom").value;
    const qty = parseFloat(document.getElementById("f-qty").value) || 0;
    const unitCost = parseFloat(document.getElementById("f-unitcost").value) || 0;
    const total = qty * unitCost;

    // Guard against the two ways a bad entry can slip through: no
    // category selected, or a total that computes to zero (usually a
    // missing qty or unit cost). Checked explicitly here rather than
    // relying only on the HTML required attribute, so the sheet never
    // receives a row that can't be attributed to a budget category or
    // that would silently show as a zero-cost entry.
    if (!category) {
      statusEl.textContent = "Choose a category before saving — entries can't be recorded without one.";
      statusEl.classList.add("error");
      btn.disabled = false;
      document.getElementById("f-category").focus();
      return;
    }
    if (total <= 0) {
      statusEl.textContent = "Total cost is zero — check the quantity and unit cost before saving.";
      statusEl.classList.add("error");
      btn.disabled = false;
      document.getElementById("f-unitcost").focus();
      return;
    }

    // Find the true next empty row (column A/I have formulas pre-filled
    // hundreds of rows down, so a plain append would land far past your
    // real data — see findNextEmptyRow). We write only columns B-H and
    // J, leaving column A (Ser. number) and column I (Total Cost) as the
    // sheet's own formulas, so they keep auto-calculating exactly as
    // they do when you type directly into the sheet.
    const row = await findNextEmptyRow();
    await sheetsUpdate(`${CONFIG.TABS.ACTUALS}!B${row}:H${row}`, [
      category, subcategory, item, description, uom, qty, unitCost,
    ]);
    await sheetsUpdate(`${CONFIG.TABS.ACTUALS}!J${row}`, [date]);

    statusEl.textContent = `Added — ${item} (${fmt(total)}) logged by ${state.currentUser.name}.`;
    statusEl.classList.add("success");
    e.target.reset();
    document.getElementById("f-date").valueAsDate = new Date();
    document.getElementById("f-qty").value = 1;
    recalcTotal();
    loadTopSheet();
  } catch (err) {
    statusEl.textContent = err.message || "Couldn't save that entry. Try again.";
    statusEl.classList.add("error");
  } finally {
    btn.disabled = false;
  }
});

// ---------- Admin ----------

function renderAdmin() {
  const listEl = document.getElementById("admin-list");
  listEl.innerHTML = state.users.map(u => `
    <div class="admin-person" data-name="${u.name}">
      <div class="admin-person-name">${u.name}</div>
      <div class="admin-chips">
        ${CATEGORIES.map(c => `
          <button type="button" class="admin-chip ${u.allowed.includes(c) || u.allowed.includes('ALL') ? 'active' : ''}" data-cat="${c}">
            ${c}
          </button>`).join("")}
      </div>
    </div>
  `).join("");

  listEl.querySelectorAll(".admin-chip").forEach(chip => {
    chip.addEventListener("click", () => {
      if (chip.closest(".admin-person").dataset.name === state.currentUser.name && state.isAdmin) return;
      chip.classList.toggle("active");
      // Persisting this back to the Users tab requires write access to that
      // tab too — wire this up the same way as the Actuals append once
      // you're ready (see SETUP.md, "Optional: editable permissions").
      document.getElementById("admin-status").textContent =
        "Change staged. Update the Users tab in Google Sheets to save it permanently.";
    });
  });
}

// ---------- Daily note (PIN screen) ----------

async function fetchSignedInEmail(token) {
  try {
    const res = await fetch("https://www.googleapis.com/oauth2/v2/userinfo", {
      headers: { "Authorization": `Bearer ${token}` },
    });
    if (!res.ok) return null;
    const json = await res.json();
    return json.email || null;
  } catch (e) {
    return null;
  }
}

// Shows one line from DAILY_NOTE_CONFIG.MESSAGES, rotating by day of
// year, but only when the signed-in account matches HER_EMAIL. Defined
// in daily-note.js, kept separate so the message list is easy to edit
// without touching app logic.
function showDailyNoteIfHers() {
  const noteEl = document.getElementById("daily-note");
  const sigEl = document.getElementById("daily-signature");
  if (!noteEl || typeof DAILY_NOTE_CONFIG === "undefined") return;

  const email = (state.signedInEmail || "").trim().toLowerCase();
  const herEmail = (DAILY_NOTE_CONFIG.HER_EMAIL || "").trim().toLowerCase();
  if (!email || !herEmail || email !== herEmail) {
    noteEl.classList.add("hidden");
    if (sigEl) sigEl.classList.add("hidden");
    return;
  }

  const messages = DAILY_NOTE_CONFIG.MESSAGES || [];
  if (!messages.length) return;

  const dayOfYear = Math.floor(
    (Date.now() - new Date(new Date().getFullYear(), 0, 0)) / 86400000
  );
  const message = messages[dayOfYear % messages.length];

  noteEl.textContent = message;
  noteEl.classList.remove("hidden");
  if (sigEl) sigEl.classList.remove("hidden");
}

// ---------- Google Sign-In (now the first gate, before PIN) ----------

window.onload = () => {
  // Graceful fallback if unlock-photo.jpg hasn't been uploaded yet —
  // hide the broken image and let the plain ink background stand alone.
  ["signin-photo", "unlock-photo"].forEach(id => {
    const img = document.getElementById(id);
    if (img) img.addEventListener("error", () => img.classList.add("broken"));
  });

  const signinError = document.getElementById("signin-error");

  if (!window.google || !CONFIG.OAUTH_CLIENT_ID || CONFIG.OAUTH_CLIENT_ID.startsWith("YOUR_")) {
    signinError.textContent = "Sign-in isn't configured yet — check OAUTH_CLIENT_ID in config.js.";
    return;
  }

  const tokenClient = google.accounts.oauth2.initTokenClient({
    client_id: CONFIG.OAUTH_CLIENT_ID,
    scope: "https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/userinfo.email",
    callback: async (resp) => {
      if (resp.access_token) {
        state.googleToken = resp.access_token;
        state.signedInEmail = await fetchSignedInEmail(resp.access_token);
        // Reads now require this same token, so the PIN screen (which
        // fetches the Users tab) only appears once sign-in succeeds.
        document.getElementById("signin-screen").classList.add("hidden");
        document.getElementById("pin-screen").classList.remove("hidden");
        showDailyNoteIfHers();
        initPinScreen();
      } else {
        signinError.textContent = "Sign-in didn't complete. Try again.";
      }
    },
  });

  const btn = document.createElement("button");
  btn.textContent = "Sign in with Google";
  btn.className = "submit-btn";
  btn.style.width = "auto";
  btn.style.padding = "12px 24px";
  btn.onclick = () => tokenClient.requestAccessToken();
  document.getElementById("google-signin-gate-btn").appendChild(btn);
};
