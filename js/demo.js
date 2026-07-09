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

// [vendor, description, bigTicket] per category
const SAMPLES = {
  "Travel Expenses": [["Deutsche Bahn", "Train to trade fair", true], ["Booking.com", "Hotel 2 nights", true], ["Taxi Dresden", "Taxi to supplier", false]],
  Hardware: [["Bambulab", "PLA & nozzles", false], ["Vevor", "Workshop tooling", true], ["Elegoo", "Filament restock", false]],
  "Software/SaaS": [["Anthropic", "Claude subscription", false], ["GitHub", "Team seats", false], ["Autodesk", "CAD license", true]],
  Infrastructure: [["Amazon", "Workbench & storage", true], ["Hornbach", "Workshop shelving", false]],
  "Office & Team": [["Lidl", "Team dinner groceries", false], ["HIT", "Drinks company dinner", false]],
  "Marketing/Sales": [["Copy Planet Dresden", "Business cards", false], ["Gerstäcker", "Print materials", false]],
  "Legal & Notary": [["Engel und Hain GbR", "Legal advisory", true]],
  Miscellaneous: [["bavAIRia e.V.", "Membership fee", true]],
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
      const [vendor, desc, big] = SAMPLES[cat][Math.floor(rand() * SAMPLES[cat].length)];
      const amount = Math.round((big ? 120 + rand() * 700 : 8 + rand() * 190) * 100) / 100;
      const vatRate = CONFIG.vatByCategory[cat] ?? (cat === "Miscellaneous" ? 0 : 0.19);
      const vat = Math.round((amount - amount / (1 + vatRate)) * 100) / 100;
      const recent = m === 0 && day > today.getDate() - 12;
      const status = recent && rand() < 0.55 ? "Pending" : rand() < 0.06 ? "Rejected" : "Approved";
      out.push({
        rowIndex: n,
        id: `EXP-DEMO${String(++n).padStart(3, "0")}`,
        submitted: date,
        dateISO: date.toISOString().slice(0, 10),
        employee,
        email,
        vendor,
        category: cat,
        description: desc,
        amount,
        vatRate,
        vat,
        net: Math.round((amount - vat) * 100) / 100,
        currency: "EUR",
        payment: CONFIG.paymentMethods[Math.floor(rand() * 2)],
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
