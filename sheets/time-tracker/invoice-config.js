// sheets/time-tracker/invoice-config.js
/**
 * Configuration for Invoice Generation
 * Contains template IDs, folder paths, and column mappings for the Invoices sheet
 */

// Google Doc template ID for invoice
const INVOICE_TEMPLATE_ID = "1atVZFnHGOOYfB9RPcDEYuA2C-HOMryfb39huNASb8Lc";

// Folder IDs for client management
const CLIENTS_PARENT_FOLDER_ID = "1afPjpeFNzd96nv01X4XicRL52NXruQEh";
const CLIENT_FOLDER_TEMPLATE_ID = "1IXXQdHNUyq_qFqAGKQAIRDm9iOy2cfvy";

// Name of the Invoices sheet
const INVOICES_SHEET_NAME = "💼 Invoices";

// Payment terms
const PAYMENT_TERMS = "Net 30";
const DUE_DATE_DAYS = 30;

// Line item defaults
const LINE_ITEM_TITLE = "Bundle of Hours";
const LINE_ITEM_DESCRIPTION = "Prepaid block of service hours at a fixed rate.";

// Processing fee configuration
const PROCESSING_FEE_PERCENTAGE = 0.035; // 3.5%
const INVOICE_NUMBER_FEE_SUFFIX = "-PF";

// Payment links (update these with your actual payment URLs)
const STRIPE_PAYMENT_URL = "https://buy.stripe.com/6oUfZhfzYe6pbz72Bkds400";
const PAYPAL_PAYMENT_URL = "https://paypal.me/your-paypal-link"; // TODO: Update with actual PayPal link

// Card payment disclaimer for standard invoices (without processing fee)
const CARD_PAYMENT_DISCLAIMER =
  "To pay by card, please contact us for an updated invoice that includes this processing fee.";

/**
 * Column indices for the Invoices sheet
 * Columns: Year, Client, Project Title, Active, Source, Month Name, Start, End, Recur,
 *          Status, Hours Total, Hours Adj, Hours Used, Hours Rem, Perc Done, Progress,
 *          PIF, Rate, Gross, Tax, Net, Report, Notes
 *
 * IMPORTANT: Customer pays the GROSS amount. NET is your after-tax take home (internal tracking only).
 */
const InvoiceColIndex = {
  YEAR: 0, // A - Year (e.g., 2025)
  CLIENT: 1, // B - Client name
  PROJECT_TITLE: 2, // C - Project title
  ACTIVE: 3, // D - Active status
  SOURCE: 4, // E - Source
  MONTH_NAME: 5, // F - Month name
  START: 6, // G - Start date
  END: 7, // H - End date
  RECUR: 8, // I - Recur
  STATUS: 9, // J - Status
  HOURS_TOTAL: 10, // K - Hours total
  HOURS_ADJ: 11, // L - Hours adjusted
  HOURS_USED: 12, // M - Hours used
  HOURS_REM: 13, // N - Hours remaining
  PERC_DONE: 14, // O - Percentage done
  PROGRESS: 15, // P - Progress
  PIF: 16, // Q - PIF
  RATE: 17, // R - Hourly rate
  GROSS: 18, // S - Gross amount (what customer pays)
  TAX: 19, // T - Tax amount
  NET: 20, // U - Net amount (your after-tax take home - NOT used on invoice)
  REPORT: 21, // V - Report
  EMAIL_TO: 22, // W - Notes
  NOTES: 23, // X - Notes
};
