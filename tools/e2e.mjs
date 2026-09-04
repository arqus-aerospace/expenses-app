// End-to-end regression suite. Drives the real UI in demo mode: access-code
// gate, submit, the review/undo countdown, credits (incl. that they require
// approval), the dashboard charts, approvals, the restricted employee role,
// and who may file/approve at all (company accounts only; approvals limited
// to the three founders). Screenshots are written for eyeballing layout.
//
//   python3 -m http.server 8123 &      # serve the app
//   npm i playwright && node tools/e2e.mjs
//
// Exits non-zero on any failed assertion or console error.
import fs from "node:fs";
import { chromium } from "playwright";
import { CONFIG, isApprover, isCompanyAccount } from "../js/config.js";
import { NO_ACCESS_TEXT, isAccessDenied, explainError, errorDetail } from "../js/errors.js";

// Screenshots land in $SHOTS (default ./.screenshots, gitignored).
// CHROMIUM_PATH pins a browser when Playwright's own download is unavailable.
const OUT = process.env.SHOTS || ".screenshots";
const PORT = process.env.PORT || 8123;
const URL = `http://localhost:${PORT}/?demo=1`;

fs.mkdirSync(OUT, { recursive: true });
const browser = await chromium.launch(
  process.env.CHROMIUM_PATH ? { executablePath: process.env.CHROMIUM_PATH } : {});
const errors = [];

async function run(scheme, mobile) {
  const ctx = await browser.newContext({
    colorScheme: scheme,
    viewport: mobile ? { width: 390, height: 844 } : { width: 1280, height: 900 },
    deviceScaleFactor: 2,
  });
  const page = await ctx.newPage();
  page.on("console", (m) => { if (m.type() === "error") errors.push(`[${scheme}] ${m.text()}`); });
  page.on("pageerror", (e) => errors.push(`[${scheme}] pageerror: ${e.message}`));

  await page.goto(URL);

  // 1. wrong code is rejected
  await page.fill("#gate-input", "0000");
  await page.click("#gate-btn");
  await page.waitForSelector("#gate-error:not([hidden])");

  // 2. correct code unlocks
  await page.fill("#gate-input", "1876");
  await page.click("#gate-btn");
  await page.waitForSelector("#demo-btn", { state: "visible" });
  await page.click("#demo-btn");
  await page.waitForSelector("#view-submit:not([hidden])");
  if (await page.locator("#tab-dashboard").isHidden()) errors.push(`[${scheme}] founder missing dashboard tab`);
  if (await page.locator("#access-warning").isVisible()) errors.push(`[${scheme}] access banner shown with access`);
  if (await page.locator("#tab-mine").isVisible()) errors.push(`[${scheme}] founder should not see My-expenses tab`);
  await page.screenshot({ path: `${OUT}/shot-submit-${scheme}${mobile ? "-mobile" : ""}.png` });

  // 3. submit an expense (no receipt -> confirm dialog), then the review step
  page.once("dialog", (d) => d.accept());
  await page.fill("#f-amount", "42.50");
  await page.fill("#f-vendor", "Taxi Dresden");
  await page.fill("#f-desc", "Taxi to airport");
  const hint = await page.textContent("#vat-hint");
  if (!/VAT/.test(hint)) errors.push(`[${scheme}] VAT hint missing: "${hint}"`);
  await page.click("#submit-btn");
  await page.waitForSelector("#review-card:not([hidden])");
  await page.screenshot({ path: `${OUT}/shot-review-${scheme}${mobile ? "-mobile" : ""}.png` });

  // 3a. undo works: back to the form with values intact
  await page.click("#review-cancel");
  await page.waitForSelector("#view-submit .form-card:not([hidden])");
  const kept = await page.inputValue("#f-vendor");
  if (kept !== "Taxi Dresden") errors.push(`[${scheme}] undo lost form values`);

  // 3b. submit again and send immediately
  page.once("dialog", (d) => d.accept());
  await page.click("#submit-btn");
  await page.waitForSelector("#review-card:not([hidden])");
  const timerText = await page.textContent("#review-timer");
  if (!/Sending automatically/.test(timerText)) errors.push(`[${scheme}] countdown text missing`);
  await page.click("#review-send");
  await page.waitForSelector("#submit-done:not([hidden])");

  // 3c. (light desktop only) countdown expiry auto-sends
  if (scheme === "light" && !mobile) {
    await page.click("#another-btn");
    page.once("dialog", (d) => d.accept());
    await page.fill("#f-amount", "10.00");
    await page.fill("#f-vendor", "Auto Send GmbH");
    await page.fill("#f-desc", "auto-send test");
    await page.click("#submit-btn");
    await page.waitForSelector("#review-card:not([hidden])");
    await page.waitForSelector("#submit-done:not([hidden])", { timeout: 15000 });
    console.log("light: auto-send after countdown works");
  }

  // 3d. credit / refund: typing a minus flips the toggle, row is negative
  if (scheme === "light" && !mobile) {
    await page.click("#another-btn");
    page.once("dialog", (d) => d.accept());
    await page.fill("#f-amount", "-150");
    await page.fill("#f-vendor", "Max Emanuel Brauerei");
    await page.fill("#f-desc", "Voucher refund");
    const normalized = await page.inputValue("#f-amount");
    if (normalized !== "150") errors.push(`[credit] amount not normalized: ${normalized}`);
    if (!(await page.locator('.sign-btn[data-sign="-1"].active').count()))
      errors.push("[credit] typing minus did not select Credit");
    const chint = await page.textContent("#vat-hint");
    if (!/Credit:/.test(chint)) errors.push(`[credit] hint wrong: ${chint}`);
    await page.click("#submit-btn");
    await page.waitForSelector("#review-card:not([hidden])");
    const review = (await page.textContent("#review-list")).replace(/−/g, "-");
    if (!/Credit \/ refund/.test(review)) errors.push("[credit] review missing credit type row");
    if (!/-\D*150/.test(review)) errors.push(`[credit] review amount not negative: ${review.slice(0, 140)}`);
    await page.screenshot({ path: `${OUT}/shot-credit-review.png` });
    await page.click("#review-send");
    await page.waitForSelector("#submit-done:not([hidden])");
    await page.click("#another-btn");
    if (!(await page.locator('.sign-btn[data-sign="1"].active').count()))
      errors.push("[credit] toggle did not reset to Expense");

    // a credit must go through approval like any other entry
    await page.click('.tab[data-view="approvals"]');
    await page.waitForSelector("#approvals-list .approval-card");
    const creditCard = page.locator(".approval-card", { hasText: "Voucher refund" });
    if (!(await creditCard.count())) {
      errors.push("[credit] credit did not appear in Approvals");
    } else {
      const shown = await creditCard.locator(".appr-amount").textContent();
      if (!/-\D*150/.test(shown.replace(/−/g, "-")))
        errors.push(`[credit] approval card amount not negative: ${shown}`);
      if (!(await creditCard.locator(".appr-amount.credit").count()))
        errors.push("[credit] approval amount missing credit styling");
      await page.screenshot({ path: `${OUT}/shot-credit-approval.png` });
      await creditCard.locator(".approve").click();
      await page.waitForSelector(".approval-card.decided");
      console.log("light: credit requires approval and approves cleanly");
    }
    await page.click('.tab[data-view="submit"]');
    console.log("light: credit entry works (minus flips toggle, negative row filed)");
  }

  // 4. dashboard
  await page.click('.tab[data-view="dashboard"]');
  await page.waitForSelector("#dash-content:not([hidden])");
  await page.waitForTimeout(300);
  await page.screenshot({ path: `${OUT}/shot-dash-${scheme}${mobile ? "-mobile" : ""}.png`, fullPage: true });

  // every rendered SVG path must be finite (negative values used to break these)
  const badPaths = await page.evaluate(() =>
    [...document.querySelectorAll("#view-dashboard svg path")]
      .map((p) => p.getAttribute("d") || "")
      .filter((d) => /NaN|Infinity|undefined/.test(d)).length);
  if (badPaths) errors.push(`[${scheme}] ${badPaths} chart path(s) with NaN/Infinity`);

  if (!mobile) {
    const bar = page.locator("#chart-monthly svg rect").nth(5);
    await bar.hover();
    await page.waitForTimeout(150);
    if (!(await page.locator(".chart-tip").isVisible()))
      errors.push(`[${scheme}] tooltip did not appear on bar hover`);
  }

  // 5. approvals: approve the expense we just filed
  await page.click('.tab[data-view="approvals"]');
  await page.waitForSelector("#approvals-list .approval-card");
  const before = await page.locator(".approval-card").count();
  await page.screenshot({ path: `${OUT}/shot-approvals-${scheme}${mobile ? "-mobile" : ""}.png` });
  await page.locator(".approval-card .approve").first().click();
  await page.waitForSelector(".approval-card.decided");
  console.log(`${scheme}${mobile ? "/mobile" : ""}: ok — ${before} pending card(s), approve worked`);
  await ctx.close();
}

// Who may file and who may approve — the rules straight from config.js.
function checkRules() {
  const T = CONFIG.tenantId;
  const MSA = "9188040d-6c67-4c5b-b112-36a304b66dad";
  const cases = [
    // [email, tenantId, mayFile, mayApprove, label]
    ["marnix@arqusaerospace.com", T, true, true, "founder Marnix"],
    ["stijn@arqusaerospace.com", T, true, true, "founder Stijn"],
    ["anton@arqusaerospace.com", T, true, true, "founder Anton"],
    ["MARNIX@ArqusAerospace.com", T, true, true, "founder, mixed case"],
    ["elena@arqusaerospace.com", T, true, false, "company employee"],
    ["tom@arqusaerospace.com", T, true, false, "company employee"],
    ["visitor@example.com", "00000000-0000-0000-0000-000000000000", false, false, "other tenant"],
    ["guest@gmail.com", T, false, false, "guest invited into the tenant"],
    ["marnix@arqusaerospace.com", MSA, false, false, "personal Microsoft account"],
    ["marnix@arqusaerospace.com", "", false, false, "no tenant claim"],
    ["", T, false, false, "no address"],
  ];
  for (const [email, tenantId, mayFile, mayApprove, label] of cases) {
    const u = { email, tenantId };
    if (isCompanyAccount(u) !== mayFile) {
      errors.push(`[rules] ${label}: may file should be ${mayFile}`);
    }
    if (isApprover(u) !== mayApprove) {
      errors.push(`[rules] ${label}: may approve should be ${mayApprove}`);
    }
  }
  if (CONFIG.approvers.length !== 3) {
    errors.push(`[rules] expected 3 approvers, found ${CONFIG.approvers.length}`);
  }
  console.log(`rules: ok — ${cases.length} file/approve cases`);
}

// A refused SharePoint write must come out as something the employee can act
// on, not as Graph's bare "Access denied".
function checkErrors() {
  const denied = { message: "Access denied", status: 403, code: "accessDenied" };
  const denied403 = { message: "Access denied", status: 403 };
  const expired = { message: "token expired", status: 401 };
  const other = { message: "Item not found", status: 404, code: "itemNotFound" };

  for (const [err, want, label] of [
    [denied, true, "403 + accessDenied"],
    [denied403, true, "bare 403"],
    [{ message: "x", code: "accessDenied" }, true, "code only"],
    [expired, false, "401"],
    [other, false, "404"],
    [undefined, false, "no error"],
  ]) {
    if (isAccessDenied(err) !== want) errors.push(`[errors] ${label}: isAccessDenied should be ${want}`);
  }
  if (explainError(denied) !== NO_ACCESS_TEXT) errors.push("[errors] 403 not explained");
  if (!/Marnix, Stijn or Anton/.test(explainError(denied))) errors.push("[errors] 403 names nobody to ask");
  if (!/sign out/i.test(explainError(expired))) errors.push("[errors] 401 not explained");
  if (explainError(other) !== "Item not found") errors.push("[errors] other error text lost");
  if (explainError(undefined) !== "Something went wrong.") errors.push("[errors] missing error not handled");
  if (errorDetail(denied) !== "Microsoft Graph said: Access denied (accessDenied, HTTP 403)") {
    errors.push(`[errors] detail line wrong: "${errorDetail(denied)}"`);
  }
  if (errorDetail({}) !== "Microsoft Graph said: Access denied") {
    errors.push(`[errors] empty detail line wrong: "${errorDetail({})}"`);
  }
  console.log("errors: ok — access-denied messages are actionable");
}

// A non-company account gets turned away: no app, nothing to file with.
async function runOutsider() {
  const ctx = await browser.newContext({ viewport: { width: 1280, height: 900 }, deviceScaleFactor: 2 });
  const page = await ctx.newPage();
  page.on("console", (m) => { if (m.type() === "error") errors.push(`[outsider] ${m.text()}`); });
  page.on("pageerror", (e) => errors.push(`[outsider] pageerror: ${e.message}`));
  await page.goto(URL.replace("demo=1", "demo=outsider"));
  await page.fill("#gate-input", "1876");
  await page.click("#gate-btn");
  await page.waitForSelector("#demo-btn", { state: "visible" });
  await page.click("#demo-btn");

  await page.waitForSelector("#connect-blocked:not([hidden])");
  if (await page.locator("#screen-app").isVisible()) errors.push("[outsider] reached the app");
  if (await page.locator("#connect-ready").isVisible()) errors.push("[outsider] sign-in card still shown");
  const who = await page.textContent("#blocked-account");
  if (!who.includes("@")) errors.push(`[outsider] blocked card names no account: "${who}"`);
  await page.screenshot({ path: `${OUT}/shot-blocked-outsider.png` });
  console.log(`outsider: ok — refused (${who}), app never opened`);
  await ctx.close();
}

// Non-founder (employee) role: no dashboard/approvals, personal list instead.
async function runEmployee() {
  const ctx = await browser.newContext({ viewport: { width: 1280, height: 900 }, deviceScaleFactor: 2 });
  const page = await ctx.newPage();
  page.on("console", (m) => { if (m.type() === "error") errors.push(`[employee] ${m.text()}`); });
  page.on("pageerror", (e) => errors.push(`[employee] pageerror: ${e.message}`));
  await page.goto(URL.replace("demo=1", "demo=employee"));
  await page.fill("#gate-input", "1876");
  await page.click("#gate-btn");
  await page.waitForSelector("#demo-btn", { state: "visible" });
  await page.click("#demo-btn");
  await page.waitForSelector("#view-submit:not([hidden])");

  if (await page.locator("#tab-dashboard").isVisible()) errors.push("[employee] dashboard tab visible");
  if (await page.locator("#tab-approvals").isVisible()) errors.push("[employee] approvals tab visible");
  if (await page.locator("#tab-mine").isHidden()) errors.push("[employee] My-expenses tab missing");

  await page.click('.tab[data-view="mine"]');
  await page.waitForSelector("#mine-content:not([hidden])");
  const rows = await page.locator("#mine-table tbody tr").count();
  const month = await page.textContent("#mine-month");
  if (!rows) errors.push("[employee] personal list is empty");
  if (!month.trim()) errors.push("[employee] personal KPI empty");
  await page.screenshot({ path: `${OUT}/shot-mine-employee.png`, fullPage: true });

  // forcing the hidden approvals tab must not open a queue to decide on
  await page.evaluate(() => document.getElementById("tab-approvals").click());
  if (await page.locator("#view-approvals").isVisible()) errors.push("[employee] approvals view opened");
  if (await page.locator(".approval-card").count()) errors.push("[employee] approval cards rendered");

  // an employee submission shows up in their own list
  page.once("dialog", (d) => d.accept());
  await page.click('.tab[data-view="submit"]');
  await page.fill("#f-amount", "19.99");
  await page.fill("#f-vendor", "Obi");
  await page.fill("#f-desc", "Sandpaper");
  await page.click("#submit-btn");
  await page.waitForSelector("#review-card:not([hidden])");
  await page.click("#review-send");
  await page.waitForSelector("#submit-done:not([hidden])");
  await page.click('.tab[data-view="mine"]');
  await page.waitForSelector("#mine-content:not([hidden])");
  const after = await page.locator("#mine-table tbody tr").count();
  if (after !== rows + 1) errors.push(`[employee] new filing not in list (${rows} -> ${after})`);
  console.log(`employee: ok — ${rows} own rows, +1 after submitting, no company tabs`);
  await ctx.close();
}

checkRules();
checkErrors();
await run("light", false);
await run("dark", false);
await run("light", true);
await runEmployee();
await runOutsider();

if (errors.length) {
  console.log("\nERRORS:");
  errors.forEach((e) => console.log(" -", e));
  process.exit(1);
}
console.log("\nAll flows passed with no console errors.");
await browser.close();
