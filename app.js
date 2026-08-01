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
  return Number(n).toLocaleString("en-EG", { maximumFractionDigits: 0 });
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

// Shows the cell's own formatting exactly as the sheet displays it —
// just trims the padding spaces Sheets adds around currency values,
// without touching decimals, commas, or the "-" it uses for zero.
function rawMoney(raw) {
  if (raw === null || raw === undefined || raw === "") return "—";
  const trimmed = String(raw).trim();
  return trimmed || "—";
}

// Google Sheets returns date cells as either an ISO-ish string
// ("2026-08-31") or, less often, a serial number — handle both.
function toDate(raw) {
  if (raw === null || raw === undefined || raw === "") return null;
  if (typeof raw === "number") {
    // Sheets/Excel serial date: days since 1899-12-30.
    return new Date(Date.UTC(1899, 11, 30) + raw * 86400000);
  }
  const d = new Date(raw);
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
  const url = `${SHEETS_API}/${CONFIG.SHEET_ID}/values/${encodeURIComponent(range)}?key=${CONFIG.API_KEY}`;
  const res = await fetch(url);
  if (!res.ok) throw new Error(`Sheets read failed (${res.status})`);
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

// ---------- Tab navigation ----------

document.querySelectorAll(".tab-btn").forEach(btn => {
  btn.addEventListener("click", () => {
    document.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("active"));
    btn.classList.add("active");
    document.querySelectorAll(".view").forEach(v => v.classList.add("hidden"));
    document.getElementById(`view-${btn.dataset.view}`).classList.remove("hidden");
    if (btn.dataset.view === "admin") renderAdmin();
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

  document.getElementById("sum-budget").textContent = fmt(totalBudget);
  document.getElementById("sum-actual").textContent = fmt(totalActual);
  document.getElementById("sum-remaining").textContent = fmt(totalBudget - totalActual);

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

    // Ser. left blank — sheet formula or next-row logic can number it;
    // if you want auto-numbering here instead, tell me and I'll add it.
    await sheetsAppend(`${CONFIG.TABS.ACTUALS}!A:J`, [
      "", category, subcategory, item, description, uom, qty, unitCost, total, date,
    ]);

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

// ---------- Google Sign-In ----------

window.onload = () => {
  initPinScreen();

  if (window.google && CONFIG.OAUTH_CLIENT_ID && !CONFIG.OAUTH_CLIENT_ID.startsWith("YOUR_")) {
    const tokenClient = google.accounts.oauth2.initTokenClient({
      client_id: CONFIG.OAUTH_CLIENT_ID,
      scope: "https://www.googleapis.com/auth/spreadsheets",
      callback: (resp) => {
        if (resp.access_token) {
          state.googleToken = resp.access_token;
          document.getElementById("signin-gate").classList.add("hidden");
          document.getElementById("entry-form").classList.remove("hidden");
        }
      },
    });

    const btn = document.createElement("button");
    btn.textContent = "Sign in with Google";
    btn.className = "submit-btn";
    btn.style.width = "auto";
    btn.style.padding = "12px 24px";
    btn.onclick = () => tokenClient.requestAccessToken();
    document.getElementById("google-signin-btn").appendChild(btn);
  }
};
