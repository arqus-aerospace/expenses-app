import { CONFIG, isConfigured, isApprover } from "./config.js";
import { initAuth, signIn, signOut, currentUser } from "./auth.js";
import * as graph from "./graph.js";
import { computeStats, cumulativeSeries, lastMonths, monthLabel, fmtMoney } from "./stats.js";
import { columnChart, lineChart, hbarChart, sparkline } from "./charts.js";
import { demoExpenses } from "./demo.js";

const $ = (id) => document.getElementById(id);
const show = (el, on = true) => { el.hidden = !on; };

// ---------------------------------------------------------------- state ----

let demo = false;
let user = null;
let founder = false;   // founders (approver emails) see Dashboard + Approvals
let expenses = null;          // cached list
let demoStore = null;
let range = "12m";
let pendingFile = null;

// Data source: Microsoft Graph in real mode, in-memory store in demo mode.
const api = {
  async list() {
    if (demo) return [...demoStore];
    return graph.listExpenses();
  },
  async submit(exp, file) {
    if (demo) {
      demoStore.push({ ...exp, rowIndex: demoStore.length, submitted: new Date(),
        receipt: file ? file.name : "", receiptUrl: "", status: "Pending",
        decidedBy: "", decidedOn: null });
      return;
    }
    let receiptName = "", receiptUrl = "";
    if (file) {
      setStatus("Uploading receipt to SharePoint…");
      const up = await graph.uploadReceipt(file, exp);
      receiptName = up.name;
      receiptUrl = up.webUrl;
    }
    setStatus("Writing to the Excel tracker…");
    await graph.addExpense({ ...exp, receiptName, receiptUrl });
  },
  async decide(item, status) {
    if (demo) {
      const row = demoStore.find((e) => e.id === item.id);
      row.status = status;
      row.decidedBy = user.name;
      row.decidedOn = new Date();
      return;
    }
    await graph.decideExpense(item.rowIndex, status, user.name);
  },
};

// ------------------------------------------------------------ pin  gate ----

async function sha256(text) {
  const buf = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(text));
  return [...new Uint8Array(buf)].map((b) => b.toString(16).padStart(2, "0")).join("");
}

async function tryUnlock() {
  const code = $("gate-input").value.trim();
  if ((await sha256(code)) === CONFIG.accessCodeHash) {
    sessionStorage.setItem("gateOk", "1");
    show($("screen-gate"), false);
    startAuth();
  } else {
    show($("gate-error"));
    $("gate-input").value = "";
    $("gate-input").focus();
  }
}

// ----------------------------------------------------------------- auth ----

async function startAuth() {
  // ?demo=1 keeps demo mode reachable even once the real connection is
  // configured (handy for previews and automated tests).
  const wantDemo = new URLSearchParams(location.search).has("demo");
  if (!isConfigured() || wantDemo) {
    show($("screen-connect"));
    show($("connect-unconfigured"));
    show($("unconfigured-note"), !isConfigured());
    if (!isConfigured()) return;
  }
  show($("connect-ready"));
  try {
    const account = await initAuth();
    if (account) {
      user = currentUser();
      enterApp();
    } else {
      show($("screen-connect"));
    }
  } catch (e) {
    show($("screen-connect"));
    const err = $("connect-error");
    err.textContent = `Sign-in failed: ${e.message}`;
    show(err);
  }
}

function enterApp() {
  show($("screen-connect"), false);
  show($("screen-app"));
  $("user-chip").textContent = user.name;
  show($("demo-badge"), demo);
  $("signout-btn").textContent = demo ? "Exit demo" : "Sign out";
  // Founders get the company-wide dashboard and approvals; everyone else
  // gets a personal record of their own filings instead.
  founder = isApprover(user.email);
  show($("tab-dashboard"), founder);
  show($("tab-approvals"), founder);
  show($("tab-mine"), !founder);
  if (!demo) {
    graph.workbookWebUrl()
      .then((url) => { $("workbook-link").href = url; show($("workbook-link")); })
      .catch(() => {});
    // warm the connection so first submit is fast, and prime the approvals badge
    refreshData().then(updateApprovalsBadge).catch(() => {});
  } else {
    refreshData().then(updateApprovalsBadge);
  }
  switchView("submit");
}

// ----------------------------------------------------------------- tabs ----

function switchView(name) {
  // non-founders can never land on the company views
  if (!founder && (name === "dashboard" || name === "approvals")) name = "mine";
  document.querySelectorAll(".tab").forEach((t) =>
    t.classList.toggle("active", t.dataset.view === name));
  for (const v of ["submit", "dashboard", "mine", "approvals"]) show($(`view-${v}`), v === name);
  if (name === "dashboard") renderDashboard();
  if (name === "mine") renderMine();
  if (name === "approvals") renderApprovals();
}

// --------------------------------------------------------------- submit ----

function setStatus(msg) {
  // progress lines while sending (shown on the review card, where send happens)
  const el = $("review-status");
  el.textContent = msg;
  show(el, Boolean(msg));
}

function acceptFile(file) {
  if (!file) return;
  if (file.size > CONFIG.maxFileMB * 1024 * 1024) {
    toast(`File is too large (max ${CONFIG.maxFileMB} MB).`);
    return;
  }
  pendingFile = file;
  show($("drop-idle"), false);
  show($("drop-preview"));
  const img = $("preview-img");
  const chip = $("preview-file");
  if (file.type.startsWith("image/")) {
    img.src = URL.createObjectURL(file);
    show(img); show(chip, false);
  } else {
    chip.textContent = `📎 ${file.name} (${(file.size / 1024 / 1024).toFixed(1)} MB)`;
    show(chip); show(img, false);
  }
}

function clearFile() {
  pendingFile = null;
  if ($("preview-img").src) URL.revokeObjectURL($("preview-img").src);
  $("preview-img").src = "";
  show($("drop-preview"), false);
  show($("drop-idle"));
  $("file-input").value = "";
  $("camera-input").value = "";
}

function wireSubmit() {
  const dz = $("dropzone");
  dz.addEventListener("click", (e) => {
    if (e.target.id === "camera-btn" || e.target.id === "remove-file") return;
    $("file-input").click();
  });
  dz.addEventListener("keydown", (e) => { if (e.key === "Enter") $("file-input").click(); });
  $("camera-btn").addEventListener("click", () => $("camera-input").click());
  $("remove-file").addEventListener("click", clearFile);
  $("file-input").addEventListener("change", (e) => acceptFile(e.target.files[0]));
  $("camera-input").addEventListener("change", (e) => acceptFile(e.target.files[0]));

  ["dragover", "dragenter"].forEach((ev) =>
    dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.add("drag"); }));
  ["dragleave", "drop"].forEach((ev) =>
    dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.remove("drag"); }));
  dz.addEventListener("drop", (e) => acceptFile(e.dataTransfer.files[0]));
  // paste a screenshot anywhere on the submit view
  document.addEventListener("paste", (e) => {
    if ($("view-submit").hidden) return;
    const f = [...(e.clipboardData?.files || [])][0];
    if (f) acceptFile(f);
  });

  $("f-date").value = new Date().toISOString().slice(0, 10);
  fillSelect($("f-currency"), CONFIG.currencies, CONFIG.defaultCurrency);
  fillSelect($("f-category"), CONFIG.categories);
  fillSelect($("f-payment"), CONFIG.paymentMethods);
  suggestVat();

  // suggest the usual VAT rate for the category; live gross → VAT/net preview
  $("f-category").addEventListener("change", suggestVat);
  ["input", "change"].forEach((ev) => {
    $("f-amount").addEventListener(ev, updateVatHint);
    $("f-vat").addEventListener(ev, updateVatHint);
  });

  $("expense-form").addEventListener("submit", onSubmit);
  $("review-cancel").addEventListener("click", cancelReview);
  $("review-send").addEventListener("click", () => finishReview(true));
  $("another-btn").addEventListener("click", () => {
    show($("submit-done"), false);
    show(document.querySelector("#view-submit .form-card"));
  });
}

function suggestVat() {
  fillVatSelect(CONFIG.vatByCategory[$("f-category").value] ?? CONFIG.defaultVatRate);
}

function fillVatSelect(selected) {
  const sel = $("f-vat");
  sel.innerHTML = "";
  for (const r of CONFIG.vatRates) {
    const opt = document.createElement("option");
    opt.value = String(r);
    opt.textContent = `${Math.round(r * 100)}%`;
    if (r === selected) opt.selected = true;
    sel.appendChild(opt);
  }
  updateVatHint();
}

function vatBreakdown() {
  const gross = Math.round(parseFloat($("f-amount").value || "0") * 100) / 100;
  const rate = Number($("f-vat").value) || 0;
  const vat = Math.round((gross - gross / (1 + rate)) * 100) / 100;
  return { gross, rate, vat, net: Math.round((gross - vat) * 100) / 100 };
}

function updateVatHint() {
  const { gross, rate, vat, net } = vatBreakdown();
  $("vat-hint").textContent = gross
    ? `Net ${fmtMoney(net, $("f-currency").value)} + ${Math.round(rate * 100)}% VAT ${fmtMoney(vat, $("f-currency").value)}`
    : "";
}

function fillSelect(sel, options, selected) {
  sel.innerHTML = "";
  for (const o of options) {
    const opt = document.createElement("option");
    opt.value = opt.textContent = o;
    if (o === selected) opt.selected = true;
    sel.appendChild(opt);
  }
}

// Submit is a two-step flow: the form builds the expense, then a review card
// shows a summary and counts down before actually sending, so a mistake can
// be caught and undone.
let reviewExp = null;
let reviewTimer = null;

function onSubmit(e) {
  e.preventDefault();
  if (!pendingFile && !confirm("No receipt attached — submit without one?")) return;
  const { gross, rate, vat, net } = vatBreakdown();
  reviewExp = {
    id: `EXP-${Date.now().toString(36).toUpperCase()}${Math.random().toString(36).slice(2, 5).toUpperCase()}`,
    dateISO: $("f-date").value,
    employee: user.name,
    email: user.email,
    vendor: $("f-vendor").value.trim(),
    category: $("f-category").value,
    description: $("f-desc").value.trim(),
    amount: gross,
    vatRate: rate,
    vat,
    net,
    currency: $("f-currency").value,
    payment: $("f-payment").value,
  };
  showReview(reviewExp);
}

function showReview(exp) {
  const cur = exp.currency;
  const rows = [
    ["Amount", `${fmtMoney(exp.amount, cur)} gross — ${fmtMoney(exp.net, cur)} net + ${Math.round(exp.vatRate * 100)}% VAT (${fmtMoney(exp.vat, cur)})`],
    ["Vendor", exp.vendor],
    ["Category", exp.category],
    ["Date", exp.dateISO],
    ["Paid with", exp.payment],
    ["Description", exp.description],
    ["Receipt", pendingFile ? pendingFile.name : "none attached"],
  ];
  $("review-list").innerHTML = rows
    .map(([k, v]) => `<div><dt>${k}</dt><dd>${esc(v)}</dd></div>`)
    .join("");
  show(document.querySelector("#view-submit .form-card"), false);
  show($("review-card"));
  $("review-send").disabled = false;
  $("review-cancel").disabled = false;
  show($("review-status"), false);

  let left = CONFIG.reviewSeconds;
  const fill = $("countbar-fill");
  fill.style.transition = "none";
  fill.style.width = "100%";
  // force reflow so the transition restarts cleanly on repeat submissions
  void fill.offsetWidth;
  fill.style.transition = `width ${left}s linear`;
  fill.style.width = "0%";
  const tick = () => {
    $("review-timer").textContent = `Sending automatically in ${left}s — press “Go back” to change something.`;
    if (left-- <= 0) { finishReview(false); return; }
    reviewTimer = setTimeout(tick, 1000);
  };
  tick();
}

function cancelReview() {
  clearTimeout(reviewTimer);
  reviewTimer = null;
  reviewExp = null;
  show($("review-card"), false);
  show(document.querySelector("#view-submit .form-card"));
}

async function finishReview(manual) {
  if (!reviewExp) return;
  clearTimeout(reviewTimer);
  reviewTimer = null;
  const exp = reviewExp;
  reviewExp = null;
  $("review-send").disabled = true;
  $("review-cancel").disabled = true;
  const rs = $("review-status");
  rs.textContent = manual ? "Sending…" : "Time’s up — sending…";
  show(rs);
  try {
    await api.submit(exp, pendingFile);
    setStatus("");
    expenses = null; // dashboard cache is stale now
    $("done-summary").textContent =
      `${fmtMoney(exp.amount, exp.currency)} · ${esc(exp.vendor)} · ${exp.category} · ${exp.dateISO}`;
    show($("review-card"), false);
    show($("submit-done"));
    $("expense-form").reset();
    $("f-date").value = new Date().toISOString().slice(0, 10);
    fillSelect($("f-currency"), CONFIG.currencies, CONFIG.defaultCurrency);
    suggestVat();
    clearFile();
    updateApprovalsBadge();
  } catch (err) {
    setStatus("");
    // keep the filled form so nothing is lost — user can try again
    reviewExp = null;
    show($("review-card"), false);
    show(document.querySelector("#view-submit .form-card"));
    toast(`Could not submit: ${err.message}`);
  }
}

// ------------------------------------------------------------ dashboard ----

async function refreshData() {
  expenses = await api.list();
  return expenses;
}

function rangeWindow(today = new Date()) {
  const end = today;
  if (range === "ytd") return { start: new Date(today.getFullYear(), 0, 1), end, months: today.getMonth() + 1 };
  const n = { "3m": 3, "6m": 6, "12m": 12 }[range] || 12;
  return { start: new Date(today.getFullYear(), today.getMonth() - (n - 1), 1), end, months: n };
}

async function renderDashboard() {
  show($("dash-loading"));
  show($("dash-content"), false);
  try {
    if (!expenses) await refreshData();
  } catch (e) {
    $("dash-loading").textContent = `Could not load expenses: ${e.message}`;
    return;
  }
  show($("dash-loading"), false);
  show($("dash-content"));

  const today = new Date();
  const stats = computeStats(expenses, today);
  const { start, end, months } = rangeWindow(today);
  const inRange = expenses.filter((x) => {
    const d = new Date(x.dateISO);
    return x.dateISO && d >= start && d <= end;
  });
  const rangeStats = computeStats(inRange, today);
  const cur = CONFIG.defaultCurrency;
  const money = (v) => fmtMoney(v, cur);
  const compact = (v) =>
    v >= 1000 ? `${Math.round(v / 100) / 10}k` : String(Math.round(v));

  // KPI tiles (fixed definitions, independent of the range filter)
  $("kpi-month").textContent = money(stats.thisMonth);
  const delta = stats.lastMonth
    ? Math.round(((stats.thisMonth - stats.lastMonth) / stats.lastMonth) * 100)
    : null;
  const de = $("kpi-month-delta");
  if (delta === null) { de.textContent = "no data last month"; de.className = "kpi-delta muted"; }
  else {
    de.textContent = `${delta >= 0 ? "▲" : "▼"} ${Math.abs(delta)}% vs last month`;
    de.className = `kpi-delta ${delta > 0 ? "up" : "down"}`;
  }
  sparkline($("kpi-spark"), stats.monthly);
  $("kpi-avg").textContent = money(stats.avgMonth);
  $("kpi-avg-sub").textContent = "active months, last 12";
  $("kpi-pending").textContent = String(stats.pendingCount);
  $("kpi-pending-sub").textContent = stats.pendingCount ? `${money(stats.pendingAmount)} waiting` : "nothing waiting";
  $("kpi-ytd").textContent = money(stats.ytd);
  $("kpi-ytd-sub").textContent = `${today.getFullYear()}, excl. rejected`;

  // Monthly columns over the selected range
  const monthKeys = lastMonths(months, today);
  const monthVals = monthKeys.map((k) =>
    inRange.filter((x) => x.status !== "Rejected" && x.dateISO.startsWith(k))
      .reduce((a, x) => a + x.amount, 0));
  $("title-monthly").textContent = `Expenses per month · last ${months} months`;
  columnChart($("chart-monthly"), {
    labels: monthKeys.map(monthLabel),
    values: monthVals,
    fmt: money,
    fmtAxis: compact,
  });

  // Cumulative line over the selected range
  $("title-cumulative").textContent = "Spend over time · cumulative";
  const series = cumulativeSeries(expenses, start, end);
  const compactMoney = new Intl.NumberFormat(undefined, {
    style: "currency", currency: cur, notation: "compact", maximumFractionDigits: 1,
  });
  lineChart($("chart-cumulative"), {
    labels: series.labels,
    values: series.values,
    fmt: money,
    fmtEnd: (v) => compactMoney.format(v),
    fmtAxis: compact,
    fmtLabel: (iso) => new Date(iso).toLocaleDateString(undefined, { day: "numeric", month: "short" }),
  });

  // Category breakdown — fixed categorical hue per category (never re-cycled)
  const catColor = (label) => {
    const i = CONFIG.categories.indexOf(label);
    return `var(--series-${(i >= 0 ? i : 9) % 8 + 1})`;
  };
  hbarChart($("chart-category"), {
    items: rangeStats.byCategory.slice(0, 8).map(([label, value]) => ({ label, value })),
    fmt: money,
    colorFor: catColor,
  });

  // Employee breakdown — one measure, single hue
  hbarChart($("chart-employee"), {
    items: rangeStats.byEmployee.slice(0, 8).map(([emp, value]) => ({
      label: emp || "(imported)",
      value,
      extra: `${inRange.filter((x) => x.employee === emp && x.status !== "Rejected").length} expenses`,
    })),
    fmt: money,
  });

  // Recent table
  const tbody = $("recent-table").querySelector("tbody");
  tbody.innerHTML = "";
  [...inRange]
    .sort((a, b) => b.dateISO.localeCompare(a.dateISO))
    .slice(0, 15)
    .forEach((x) => {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${x.dateISO}</td>
        <td>${esc(x.vendor || x.employee || "—")}</td>
        <td>${esc(x.category)}</td>
        <td class="desc-cell">${esc(x.description)}</td>
        <td class="num">${fmtMoney(x.amount, x.currency)}</td>
        <td class="num">${x.vat ? fmtMoney(x.vat, x.currency) : "—"}</td>
        <td><span class="status-pill ${x.status.toLowerCase()}">${x.status}</span></td>`;
      tbody.appendChild(tr);
    });
}

// -------------------------------------------------- my expenses (staff) ----

async function renderMine() {
  show($("mine-loading"));
  show($("mine-content"), false);
  try {
    if (!expenses) await refreshData();
  } catch (e) {
    $("mine-loading").textContent = `Could not load: ${e.message}`;
    return;
  }
  show($("mine-loading"), false);
  show($("mine-content"));

  const mine = expenses
    .filter((x) => x.email === user.email)
    .sort((a, b) => b.dateISO.localeCompare(a.dateISO));

  const today = new Date();
  const mKey = today.toISOString().slice(0, 7);
  const counted = (x) => x.status !== "Rejected";
  const sum = (list) => list.reduce((a, x) => a + x.amount, 0);
  const cur = CONFIG.defaultCurrency;

  $("mine-month").textContent =
    fmtMoney(sum(mine.filter((x) => counted(x) && x.dateISO.startsWith(mKey))), cur);
  const yearMine = mine.filter(
    (x) => counted(x) && x.dateISO.startsWith(String(today.getFullYear())));
  $("mine-year").textContent = fmtMoney(sum(yearMine), cur);
  $("mine-year-sub").textContent = `${yearMine.length} expense${yearMine.length === 1 ? "" : "s"}`;
  $("mine-pending").textContent = String(mine.filter((x) => x.status === "Pending").length);

  const tbody = $("mine-table").querySelector("tbody");
  tbody.innerHTML = "";
  for (const x of mine) {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${x.dateISO}</td>
      <td>${esc(x.vendor || "—")}</td>
      <td>${esc(x.category)}</td>
      <td class="desc-cell">${esc(x.description)}</td>
      <td class="num">${fmtMoney(x.amount, x.currency)}</td>
      <td class="num">${x.vat ? fmtMoney(x.vat, x.currency) : "—"}</td>
      <td><span class="status-pill ${x.status.toLowerCase()}">${x.status}</span></td>
      <td>${x.receiptUrl
        ? `<a href="${esc(x.receiptUrl)}" target="_blank" rel="noopener">view ↗</a>`
        : esc(x.receipt || "—")}</td>`;
    tbody.appendChild(tr);
  }
  show($("mine-empty"), mine.length === 0);
}

const esc = (s) => String(s).replace(/[&<>"']/g, (c) =>
  ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

// ------------------------------------------------------------ approvals ----

async function updateApprovalsBadge() {
  if ($("tab-approvals").hidden) return;
  try {
    if (!expenses) await refreshData();
    const n = expenses.filter((e) => e.status === "Pending").length;
    $("approvals-count").textContent = String(n);
    show($("approvals-count"), n > 0);
  } catch { /* badge is best-effort */ }
}

async function renderApprovals() {
  show($("approvals-loading"));
  show($("approvals-empty"), false);
  $("approvals-list").innerHTML = "";
  try {
    await refreshData();
  } catch (e) {
    $("approvals-loading").textContent = `Could not load: ${e.message}`;
    return;
  }
  show($("approvals-loading"), false);
  const pending = expenses
    .filter((e) => e.status === "Pending")
    .sort((a, b) => a.dateISO.localeCompare(b.dateISO));
  updateApprovalsBadge();
  if (!pending.length) { show($("approvals-empty")); return; }

  for (const item of pending) {
    const card = document.createElement("div");
    card.className = "card approval-card";
    card.innerHTML = `
      <div class="appr-main">
        <div class="appr-amount">${fmtMoney(item.amount, item.currency)}</div>
        <div class="appr-meta">
          <b>${esc(item.vendor || item.employee || "—")}</b> · ${esc(item.category)} · ${item.dateISO}<br>
          <span class="sub">${esc(item.description)}${item.employee ? ` — by ${esc(item.employee)}` : ""}</span>
          ${item.receiptUrl ? `<br><a href="${esc(item.receiptUrl)}" target="_blank" rel="noopener">View receipt ↗</a>`
                            : item.receipt ? `<br><span class="sub">Receipt: ${esc(item.receipt)}</span>` : ""}
        </div>
      </div>
      <div class="appr-actions">
        <button class="btn approve">Approve</button>
        <button class="btn reject">Reject</button>
      </div>`;
    const act = async (status) => {
      card.classList.add("busy");
      try {
        await api.decide(item, status);
        card.classList.add("decided");
        card.querySelector(".appr-actions").innerHTML =
          `<span class="decided-label ${status.toLowerCase()}">${status} by ${esc(user.name)}</span>`;
        expenses = null;
        updateApprovalsBadge();
      } catch (e) {
        card.classList.remove("busy");
        toast(`Could not update: ${e.message}`);
      }
    };
    card.querySelector(".approve").addEventListener("click", () => act("Approved"));
    card.querySelector(".reject").addEventListener("click", () => act("Rejected"));
    $("approvals-list").appendChild(card);
  }
}

// ----------------------------------------------------------------- misc ----

let toastTimer;
function toast(msg) {
  const t = $("toast");
  t.textContent = msg;
  show(t);
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => show(t, false), 5000);
}

function wireChrome() {
  document.querySelectorAll(".tab").forEach((t) =>
    t.addEventListener("click", () => switchView(t.dataset.view)));
  $("signin-btn").addEventListener("click", () =>
    signIn().catch((e) => toast(e.message)));
  $("signout-btn").addEventListener("click", () => (demo ? location.reload() : signOut()));
  $("demo-btn").addEventListener("click", () => {
    demo = true;
    demoStore = demoExpenses();
    // ?demo=employee previews the restricted (non-founder) experience
    const asEmployee = new URLSearchParams(location.search).get("demo") === "employee";
    user = asEmployee
      ? { name: "Elena (demo)", email: "elena@arqusaerospace.com" }
      : { name: "Marnix (demo)", email: CONFIG.approvers[0] };
    enterApp();
  });
  $("gate-btn").addEventListener("click", tryUnlock);
  $("gate-input").addEventListener("keydown", (e) => { if (e.key === "Enter") tryUnlock(); });
  $("refresh-btn").addEventListener("click", async () => {
    expenses = null;
    renderDashboard();
  });
  document.querySelectorAll("#range-chips .chip").forEach((c) =>
    c.addEventListener("click", () => {
      range = c.dataset.range;
      document.querySelectorAll("#range-chips .chip").forEach((x) =>
        x.classList.toggle("active", x === c));
      renderDashboard();
    }));
  let rt;
  window.addEventListener("resize", () => {
    clearTimeout(rt);
    rt = setTimeout(() => { if (!$("view-dashboard").hidden && expenses) renderDashboard(); }, 200);
  });
}

// ----------------------------------------------------------------- boot ----

wireChrome();
wireSubmit();
if (sessionStorage.getItem("gateOk") === "1") {
  startAuth();
} else {
  show($("screen-gate"));
  $("gate-input").focus();
}
