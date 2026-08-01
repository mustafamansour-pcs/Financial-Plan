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
    range: "D8:E24",

    // OPTIONAL — line-level budget detail for the overrun drill-through.
    // Without this, the drill-through still works but the budget side
    // shows only the single subcategory figure from `range` above.
    //
    // detailRange : the block on the Budget sheet holding one row per
    //               budgeted line (not per subcategory).
    // detailCols  : zero-based column positions WITHIN detailRange —
    //               subcat = which subcategory the line rolls up to,
    //               label  = the line's name, amount = its budget.
    //
    // The values below are a guess at your layout. Point them at the
    // real columns, or delete the two keys to turn budget detail off.
    detailRange: "D8:H24",
    detailCols: { subcat: 0, label: 2, amount: 4 },
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
      const budgetDetail = await loadBudgetDetail(source);

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
async function loadBudgetDetail(source) {
  if (!source.detailRange || !source.detailCols) return {};
  try {
    const rows = await sheetsGet(`${source.sheet}!${source.detailRange}`);
    const c = source.detailCols;
    const grouped = {};
    for (const row of rows) {
      const sub = String(row[c.subcat] || "").trim();
      if (!sub) continue;
      const label = String(row[c.label] || "").trim();
      const amount = parseMoney(row[c.amount]);
      if (!label && !amount) continue;
      (grouped[sub] = grouped[sub] || []).push({ label: label || sub, amount });
    }
    return grouped;
  } catch (e) {
    return {};
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
  const matched = new Set();
  const rows = d.budgetLines.map(l => {
    const k = String(l.label).trim().toLowerCase();
    const hit = byItem[k];
    if (hit) matched.add(k);
    return {
      label: l.label,
      budget: l.amount,
      actual: hit ? hit.total : 0,
      count: hit ? hit.count : 0,
      unbudgeted: false,
    };
  });

  for (const k of Object.keys(byItem)) {
    if (matched.has(k)) continue;
    rows.push({
      label: byItem[k].name,
      budget: 0,
      actual: byItem[k].total,
      count: byItem[k].count,
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
                  <span class="drill-meta">${r.unbudgeted ? "not in budget" : (r.count ? `${r.count} ${r.count === 1 ? "entry" : "entries"}` : "nothing spent")}</span>
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
      <p class="drill-foot">The Budget sheet holds a single figure for this subcategory, so there's no per-item budget to compare against. Set <code>detailRange</code> in <code>SUBCATEGORY_BUDGET_SOURCE</code> to get a true line-by-line comparison.</p>
    </div>`;
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
