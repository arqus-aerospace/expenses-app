// ---------------------------------------------------------------------------
// Arqus Expenses — configuration
//
// Everything the app needs to talk to YOUR Microsoft 365 tenant lives here.
// Follow README.md → "Setup" to fill in the two placeholder values below
// (clientId + sitePath). Until then the app runs in Demo mode.
// ---------------------------------------------------------------------------

export const CONFIG = {
  // -- Access gate ----------------------------------------------------------
  // SHA-256 of the shared access code. This is a courtesy lock for the UI;
  // real security is the Microsoft 365 sign-in (see README → Security).
  accessCodeHash:
    "1421ff611c93756cbc675b827ea48e8f3ef11c922b0046364c3953d936ef9394",

  // -- Microsoft Entra ID (Azure AD) app registration ------------------------
  // Create a free "Single-page application" registration in Entra ID and put
  // its Application (client) ID here. See README → Setup, step 1.
  clientId: "ca05b475-986e-4b68-be8b-f388e0070a89",                                // e.g. "6f1b2c3d-...."
  tenant: "arqusaerospace.com",                // your M365 tenant domain

  // -- Who may use the app at all --------------------------------------------
  // Filing expenses is open to everyone with a company Microsoft 365 account —
  // and to nobody else. Two checks run after sign-in (see isCompanyAccount):
  //   tenantId            the directory the account must live in ("tid" claim).
  //                       This is what keeps invited guests and personal
  //                       Microsoft accounts out; leave "" to skip the check.
  //   allowedEmailDomains the domains a sign-in address may end in.
  tenantId: "234c9a25-2a21-4de0-97d2-e5136f5c9b5f",   // Arqus Aerospace directory
  allowedEmailDomains: ["arqusaerospace.com"],

  // -- SharePoint destination -------------------------------------------------
  // The site that holds the expense archive + workbook. The app stores
  // everything in the site's default "Documents" library.
  siteHostname: "arqusaerospace.sharepoint.com",
  sitePath: "/sites/Finance",                  // SharePoint site to use

  // Folder layout inside the Documents library (created automatically):
  //   Expenses/expense-tracker.xlsx            <- the maintained Excel file
  //   Expenses/Receipts/<year>/<year-month>/   <- every receipt photo/document
  rootFolder: "Expenses",
  workbookName: "expense-tracker.xlsx",
  tableName: "Expenses",                       // Excel table the app maintains

  // -- Approval --------------------------------------------------------------
  // Only these people see the Approvals tab and can approve/reject; every
  // other company account can file expenses and see its own filings only.
  approvers: [
    "marnix@arqusaerospace.com",
    "stijn@arqusaerospace.com",
    "anton@arqusaerospace.com",
  ],

  // -- Form options -----------------------------------------------------------
  currencies: ["EUR", "USD", "GBP", "CHF"],
  defaultCurrency: "EUR",
  categories: [
    "Travel Expenses",
    "Hardware",
    "Software/SaaS",
    "Infrastructure",
    "Office & Team",
    "Marketing/Sales",
    "Legal & Notary",
    "Miscellaneous",
  ],
  paymentMethods: [
    "Company Credit Card",
    "Bank Transfer",
    "Personal (reimburse)",
    "Cash (reimburse)",
  ],

  // VAT rates offered in the form; the gross amount is entered and the app
  // computes VAT + net. Default rate per category (German rates).
  vatRates: [0.19, 0.07, 0],
  defaultVatRate: 0.19,
  vatByCategory: { "Travel Expenses": 0.07 },

  // Seconds the post-submit review screen waits before auto-sending.
  reviewSeconds: 10,

  // Max receipt size (MB). Larger files upload in chunks automatically;
  // this is just a hard cap to keep the library sane.
  maxFileMB: 25,
};

// Columns of the Excel table, in order. Must match tools/make_template.py.
export const COLUMNS = [
  "ID", "Submitted", "Date", "Employee", "Email", "Vendor", "Category",
  "Description", "Gross", "VAT %", "VAT", "Net", "Currency", "Payment",
  "ReceiptFile", "Status", "DecidedBy", "DecidedOn",
];

export const isConfigured = () => Boolean(CONFIG.clientId);

const norm = (s) => String(s ?? "").trim().toLowerCase();

// Every consumer ("personal") Microsoft account reports this well-known
// tenant, so it is refused even when tenantId above is left empty.
const PERSONAL_ACCOUNT_TENANT = "9188040d-6c67-4c5b-b112-36a304b66dad";

// Is this a Microsoft 365 account of the company itself?
// Takes the object returned by auth.currentUser(): { email, tenantId }.
// Guests invited into the tenant fail the domain check, accounts from another
// tenant (and personal accounts) fail the tenant check.
export function isCompanyAccount(user) {
  const email = norm(user?.email);
  const tenantId = norm(user?.tenantId);
  if (!email.includes("@")) return false;
  if (tenantId === PERSONAL_ACCOUNT_TENANT) return false;
  if (CONFIG.tenantId && tenantId !== norm(CONFIG.tenantId)) return false;
  const domain = email.slice(email.lastIndexOf("@") + 1);
  return CONFIG.allowedEmailDomains.some((d) => norm(d) === domain);
}

// May this user approve/reject? Approver rights are never granted to an
// account that isn't a company account in the first place.
export function isApprover(user) {
  if (!isCompanyAccount(user)) return false;
  const email = norm(user?.email);
  return CONFIG.approvers.some((a) => norm(a) === email);
}
