// Aggregations for the dashboard. "Spend" counts Pending + Approved expenses
// (everything that is not Rejected), matching the workbook's Dashboard sheet.

export const monthKey = (d) =>
  `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;

export const monthLabel = (key) => {
  const [y, m] = key.split("-").map(Number);
  return new Date(y, m - 1, 1).toLocaleDateString(undefined, {
    month: "short",
    year: "2-digit",
  });
};

// Last n month keys ending with the current month, oldest first.
export function lastMonths(n, today = new Date()) {
  const out = [];
  for (let i = n - 1; i >= 0; i--) {
    out.push(monthKey(new Date(today.getFullYear(), today.getMonth() - i, 1)));
  }
  return out;
}

const counted = (e) => e.status !== "Rejected" && e.dateISO;

export function computeStats(expenses, today = new Date()) {
  const spend = expenses.filter(counted);
  const thisKey = monthKey(today);
  const lastKey = monthKey(new Date(today.getFullYear(), today.getMonth() - 1, 1));

  const byMonth = new Map();
  const byCategory = new Map();
  const byEmployee = new Map();
  let ytd = 0;

  for (const e of spend) {
    const key = e.dateISO.slice(0, 7);
    byMonth.set(key, (byMonth.get(key) || 0) + e.amount);
    byCategory.set(e.category, (byCategory.get(e.category) || 0) + e.amount);
    byEmployee.set(e.employee, (byEmployee.get(e.employee) || 0) + e.amount);
    if (key.slice(0, 4) === String(today.getFullYear())) ytd += e.amount;
  }

  const months12 = lastMonths(12, today);
  const monthly = months12.map((k) => byMonth.get(k) || 0);
  const activeMonths = monthly.filter((v) => v > 0).length || 1;

  const pending = expenses.filter((e) => e.status === "Pending");

  return {
    thisMonth: byMonth.get(thisKey) || 0,
    lastMonth: byMonth.get(lastKey) || 0,
    avgMonth: monthly.reduce((a, b) => a + b, 0) / activeMonths,
    ytd,
    pendingCount: pending.length,
    pendingAmount: pending.reduce((a, e) => a + e.amount, 0),
    months12,
    monthly,
    byCategory: [...byCategory.entries()].sort((a, b) => b[1] - a[1]),
    byEmployee: [...byEmployee.entries()].sort((a, b) => b[1] - a[1]),
  };
}

// Cumulative spend day-by-day inside a window (for the over-time line).
export function cumulativeSeries(expenses, fromDate, toDate) {
  const spend = expenses
    .filter(counted)
    .filter((e) => {
      const d = new Date(e.dateISO);
      return d >= fromDate && d <= toDate;
    })
    .sort((a, b) => a.dateISO.localeCompare(b.dateISO));

  const perDay = new Map();
  for (const e of spend) perDay.set(e.dateISO, (perDay.get(e.dateISO) || 0) + e.amount);

  const labels = [];
  const values = [];
  let running = 0;
  for (let d = new Date(fromDate); d <= toDate; d.setDate(d.getDate() + 1)) {
    const iso = d.toISOString().slice(0, 10);
    running += perDay.get(iso) || 0;
    labels.push(iso);
    values.push(running);
  }
  return { labels, values };
}

export const fmtMoney = (v, currency = "EUR") =>
  new Intl.NumberFormat(undefined, {
    style: "currency",
    currency,
    maximumFractionDigits: v >= 1000 ? 0 : 2,
  }).format(v);
