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
  // Only these people see the Approvals tab and can approve/reject.
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

export const isApprover = (email) =>
  Boolean(email) &&
  CONFIG.approvers.some((a) => a.toLowerCase() === email.toLowerCase());
