// Demo mode: seeded fake data so the app is fully explorable (and the
// dashboard reviewable) before the Microsoft 365 connection is configured.

import { CONFIG } from "./config.js";

// Deterministic PRNG so the demo looks the same on every load.
function mulberry32(seed) {
  return () => {
    seed |= 0; seed = (seed + 0x6d2b79f5) | 0;
    let t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

const PEOPLE = [
  ["Marnix", "marnix@arqusaerospace.com"],
  ["Stijn", "stijn@arqusaerospace.com"],
  ["Anton", "anton@arqusaerospace.com"],
  ["Elena", "elena@arqusaerospace.com"],
  ["Tom", "tom@arqusaerospace.com"],
];

const DESCRIPTIONS = {
  Travel: ["Train to Schiphol", "Flight BRU–TLS", "Taxi to supplier", "Parking airport"],
  Meals: ["Lunch with client", "Team dinner", "Coffee meeting"],
  Accommodation: ["Hotel Toulouse 2 nights", "Hotel Bremen"],
  "Office supplies": ["Printer paper + toner", "Whiteboard markers"],
  "Software & subscriptions": ["Figma seats", "CAD license renewal", "GitHub Team"],
  Equipment: ["USB-C dock", "Test bench PSU", "Torque wrench set"],
  "Fuel & mileage": ["Fuel company van", "Mileage claim March"],
  Training: ["EASA Part-21 course", "Welding cert renewal"],
  "Client entertainment": ["Airshow tickets", "Client dinner Paris"],
  Other: ["Customs handling fee", "Notary copy"],
};

export function demoExpenses(today = new Date()) {
  const rand = mulberry32(1876);
  const out = [];
  let n = 0;
  // ~14 months of history with a gentle upward drift + seasonal bumps
  for (let m = 13; m >= 0; m--) {
    const base = new Date(today.getFullYear(), today.getMonth() - m, 1);
    const count = 6 + Math.floor(rand() * 7) + (m < 2 ? 2 : 0);
    for (let i = 0; i < count; i++) {
      const cat = CONFIG.categories[Math.floor(rand() * CONFIG.categories.length)];
      const [employee, email] = PEOPLE[Math.floor(rand() * PEOPLE.length)];
      const day = 1 + Math.floor(rand() * 27);
      const date = new Date(base.getFullYear(), base.getMonth(), day);
      if (date > today) continue;
      const big = cat === "Accommodation" || cat === "Travel" || cat === "Equipment";
      const amount = Math.round((big ? 120 + rand() * 700 : 8 + rand() * 190) * 100) / 100;
      const recent = m === 0 && day > today.getDate() - 12;
      const status = recent && rand() < 0.55 ? "Pending" : rand() < 0.06 ? "Rejected" : "Approved";
      const descs = DESCRIPTIONS[cat];
      out.push({
        rowIndex: n,
        id: `EXP-DEMO${String(++n).padStart(3, "0")}`,
        submitted: date,
        dateISO: date.toISOString().slice(0, 10),
        employee,
        email,
        category: cat,
        description: descs[Math.floor(rand() * descs.length)],
        amount,
        currency: "EUR",
        payment: CONFIG.paymentMethods[Math.floor(rand() * CONFIG.paymentMethods.length)],
        receipt: "receipt.jpg",
        receiptUrl: "",
        status,
        decidedBy: status === "Pending" ? "" : "Marnix",
        decidedOn: status === "Pending" ? null : date,
      });
    }
  }
  return out.sort((a, b) => a.dateISO.localeCompare(b.dateISO));
}
