// Microsoft Graph layer: SharePoint uploads + the maintained Excel workbook.
//
// Data flow (see README for the full plan):
//   receipt file  ->  Documents/Expenses/Receipts/<year>/<year-month>/…
//   expense row   ->  Documents/Expenses/expense-tracker.xlsx, table "Expenses"
// The workbook (with its live Dashboard sheet) is provisioned automatically
// on first run from the template embedded in js/xlsx-template.js.

import { CONFIG, COLUMNS } from "./config.js";
import { getToken } from "./auth.js";
import { XLSX_TEMPLATE_B64 } from "./xlsx-template.js";

const GRAPH = "https://graph.microsoft.com/v1.0";

// ---------------------------------------------------------------- fetch ----

async function gfetch(path, opts = {}, attempt = 0) {
  const token = await getToken();
  const headers = {
    Authorization: `Bearer ${token}`,
    ...(opts.body && !(opts.body instanceof Uint8Array) && !(opts.body instanceof Blob)
      ? { "Content-Type": "application/json" }
      : {}),
    ...(opts.headers || {}),
  };
  const res = await fetch(`${GRAPH}${path}`, { ...opts, headers });

  // Graph throttling: back off and retry a few times.
  if ((res.status === 429 || res.status === 503 || res.status === 504) && attempt < 3) {
    const wait = Number(res.headers.get("Retry-After")) || 2 * (attempt + 1);
    await new Promise((r) => setTimeout(r, wait * 1000));
    return gfetch(path, opts, attempt + 1);
  }
  if (res.status === 204) return null;
  const data = await res.json().catch(() => null);
  if (!res.ok) {
    const err = new Error(data?.error?.message || `Graph error ${res.status}`);
    err.status = res.status;
    err.code = data?.error?.code;
    throw err;
  }
  return data;
}

// ------------------------------------------------------- site and drive ----

let driveId = null;

export async function resolveDrive() {
  if (driveId) return driveId;
  const cached = sessionStorage.getItem("driveId");
  if (cached) return (driveId = cached);
  const site = await gfetch(
    `/sites/${CONFIG.siteHostname}:${CONFIG.sitePath}?$select=id`
  );
  const drive = await gfetch(`/sites/${site.id}/drive?$select=id`);
  driveId = drive.id;
  sessionStorage.setItem("driveId", driveId);
  return driveId;
}

const itemPath = (p) => `/drives/${driveId}/root:/${encodeURI(p)}`;

// Create a folder path segment by segment; existing folders are fine.
async function ensureFolder(path) {
  const parts = path.split("/").filter(Boolean);
  let parent = "";
  for (const name of parts) {
    try {
      await gfetch(
        parent ? `${itemPath(parent)}:/children` : `/drives/${driveId}/root/children`,
        {
          method: "POST",
          body: JSON.stringify({
            name,
            folder: {},
            "@microsoft.graph.conflictBehavior": "fail",
          }),
        }
      );
    } catch (e) {
      if (e.code !== "nameAlreadyExists") throw e;
    }
    parent = parent ? `${parent}/${name}` : name;
  }
}

// ------------------------------------------------------------- receipts ----

const CHUNK = 10 * 320 * 1024; // upload-session chunks must be 320 KiB aligned

export async function uploadReceipt(file, { id, dateISO }) {
  await resolveDrive();
  const [year, month] = dateISO.split("-");
  const folder = `${CONFIG.rootFolder}/Receipts/${year}/${year}-${month}`;
  await ensureFolder(folder);

  const ext = (file.name.match(/\.[A-Za-z0-9]+$/) || [".bin"])[0].toLowerCase();
  const name = `${dateISO}_${id}${ext}`;
  const path = `${folder}/${name}`;

  let item;
  if (file.size <= 4 * 1024 * 1024) {
    item = await gfetch(`${itemPath(path)}:/content?@microsoft.graph.conflictBehavior=rename`, {
      method: "PUT",
      body: file,
      headers: { "Content-Type": file.type || "application/octet-stream" },
    });
  } else {
    const session = await gfetch(`${itemPath(path)}:/createUploadSession`, {
      method: "POST",
      body: JSON.stringify({
        item: { "@microsoft.graph.conflictBehavior": "rename", name },
      }),
    });
    const buf = await file.arrayBuffer();
    for (let start = 0; start < buf.byteLength; start += CHUNK) {
      const end = Math.min(start + CHUNK, buf.byteLength);
      const res = await fetch(session.uploadUrl, {
        method: "PUT",
        headers: {
          "Content-Range": `bytes ${start}-${end - 1}/${buf.byteLength}`,
        },
        body: buf.slice(start, end),
      });
      if (!res.ok) throw new Error(`Receipt upload failed (${res.status})`);
      if (end === buf.byteLength) item = await res.json();
    }
  }
  return { name: item.name, webUrl: item.webUrl };
}

// ------------------------------------------------------------- workbook ----

let workbookId = null;

const wbPath = () =>
  `${itemPath(`${CONFIG.rootFolder}/${CONFIG.workbookName}`)}`;
const wbApi = (suffix) => `/drives/${driveId}/items/${workbookId}/workbook${suffix}`;

async function withSession(fn) {
  const s = await gfetch(wbApi("/createSession"), {
    method: "POST",
    body: JSON.stringify({ persistChanges: true }),
  });
  const headers = { "workbook-session-id": s.id };
  try {
    return await fn(headers);
  } finally {
    gfetch(wbApi("/closeSession"), { method: "POST", headers }).catch(() => {});
  }
}

// Find the workbook; on first ever run upload the embedded template and
// create the "Expenses" table over its header row.
export async function ensureWorkbook() {
  if (workbookId) return workbookId;
  await resolveDrive();
  await ensureFolder(CONFIG.rootFolder);

  try {
    const item = await gfetch(`${wbPath()}?$select=id`);
    workbookId = item.id;
  } catch (e) {
    if (e.status !== 404) throw e;
    const bytes = Uint8Array.from(atob(XLSX_TEMPLATE_B64), (c) => c.charCodeAt(0));
    try {
      const item = await gfetch(`${wbPath()}:/content?@microsoft.graph.conflictBehavior=fail`, {
        method: "PUT",
        body: bytes,
        headers: {
          "Content-Type":
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        },
      });
      workbookId = item.id;
    } catch (conflict) {
      // Another device provisioned it between our check and upload.
      const item = await gfetch(`${wbPath()}?$select=id`);
      workbookId = item.id;
    }
  }

  // Make sure the table exists (idempotent).
  try {
    await gfetch(wbApi(`/tables('${CONFIG.tableName}')?$select=name`));
  } catch (e) {
    if (e.status !== 404) throw e;
    await withSession(async (headers) => {
      const lastCol = String.fromCharCode(64 + COLUMNS.length); // N
      const table = await gfetch(wbApi("/tables/add"), {
        method: "POST",
        headers,
        body: JSON.stringify({ address: `Data!A1:${lastCol}1`, hasHeaders: true }),
      });
      await gfetch(wbApi(`/tables('${table.name}')`), {
        method: "PATCH",
        headers,
        body: JSON.stringify({ name: CONFIG.tableName, showTotals: false }),
      });
    });
  }
  return workbookId;
}

// ------------------------------------------------------ date <-> serial ----
// Dates are stored as native Excel date serials so the workbook's Dashboard
// formulas (SUMIFS on the date column) work. Serial 1 = 1900-01-01.

const EPOCH = Date.UTC(1899, 11, 30);
const DAY = 86400000;

export const toSerial = (iso) => {
  const [y, m, d] = iso.split("-").map(Number);
  return (Date.UTC(y, m - 1, d) - EPOCH) / DAY;
};
export const nowSerial = () => (Date.now() - new Date().getTimezoneOffset() * 60000 - EPOCH) / DAY;
export const fromSerial = (n) => {
  if (typeof n !== "number" || !isFinite(n) || n <= 0) return null;
  return new Date(EPOCH + n * DAY);
};
export const serialToISO = (n) => {
  const d = fromSerial(n);
  return d ? d.toISOString().slice(0, 10) : "";
};

// ------------------------------------------------------------- expenses ----

export async function addExpense(exp) {
  await ensureWorkbook();
  const row = [
    exp.id,
    nowSerial(),
    toSerial(exp.dateISO),
    exp.employee,
    exp.email,
    exp.vendor,
    exp.category,
    exp.description,
    exp.amount,       // gross
    exp.vatRate,      // fraction, e.g. 0.19
    exp.vat,
    exp.net,
    exp.currency,
    exp.payment,
    exp.receiptUrl ? `=HYPERLINK("${exp.receiptUrl}","${exp.receiptName}")` : "",
    "Pending",
    "",
    "",
  ];
  await withSession((headers) =>
    gfetch(wbApi(`/tables('${CONFIG.tableName}')/rows/add`), {
      method: "POST",
      headers,
      body: JSON.stringify({ values: [row] }),
    })
  );
}

export async function listExpenses() {
  await ensureWorkbook();
  // Read the whole table range once; `formulas` recovers the receipt URL out
  // of the HYPERLINK() cell (row values only carry the display text).
  const range = await gfetch(
    wbApi(`/tables('${CONFIG.tableName}')/range?$select=values,formulas`)
  );
  const [, ...dataRows] = range.values; // drop the header row
  const formulas = range.formulas.slice(1);
  return dataRows
    .map((v, i) => {
      const link = String(formulas[i]?.[14] ?? "").match(/HYPERLINK\("([^"]+)"/);
      return {
        rowIndex: i,
        id: String(v[0] ?? ""),
        submitted: fromSerial(v[1]),
        dateISO: serialToISO(v[2]),
        employee: String(v[3] ?? ""),
        email: String(v[4] ?? "").toLowerCase(),
        vendor: String(v[5] ?? ""),
        category: String(v[6] ?? ""),
        description: String(v[7] ?? ""),
        amount: Number(v[8]) || 0,     // gross
        vatRate: Number(v[9]) || 0,
        vat: Number(v[10]) || 0,
        net: Number(v[11]) || 0,
        currency: String(v[12] ?? CONFIG.defaultCurrency),
        payment: String(v[13] ?? ""),
        receipt: String(v[14] ?? ""),
        receiptUrl: link ? link[1] : "",
        status: String(v[15] ?? "Pending"),
        decidedBy: String(v[16] ?? ""),
        decidedOn: fromSerial(v[17]),
      };
    })
    .filter((e) => e.id);
}

// Approve or reject: patch the Status / DecidedBy / DecidedOn cells of the row.
export async function decideExpense(rowIndex, status, decidedBy) {
  await ensureWorkbook();
  await withSession(async (headers) => {
    const range = await gfetch(
      wbApi(`/tables('${CONFIG.tableName}')/rows/itemAt(index=${rowIndex})/range?$select=address`),
      { headers }
    );
    // address looks like "Data!A5:R5" -> patch the Status/DecidedBy/DecidedOn
    // cells (P:R) on that row
    const m = range.address.match(/!([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
    if (!m) throw new Error(`Unexpected row address: ${range.address}`);
    const rowNum = m[2];
    await gfetch(
      wbApi(`/worksheets('Data')/range(address='P${rowNum}:R${rowNum}')`),
      {
        method: "PATCH",
        headers,
        body: JSON.stringify({ values: [[status, decidedBy, nowSerial()]] }),
      }
    );
  });
}

// Link shown in the header so people can jump to the live workbook/folder.
export async function workbookWebUrl() {
  await ensureWorkbook();
  const item = await gfetch(`/drives/${driveId}/items/${workbookId}?$select=webUrl`);
  return item.webUrl;
}
