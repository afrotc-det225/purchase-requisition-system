// =============================================================================
// Purchase Requisition System — Google Apps Script (TypeScript)
// Bound to a Google Sheet that receives Google Form submissions.
// =============================================================================

// ---------------------------------------------------------------------------
// CONFIGURATION
// ---------------------------------------------------------------------------

/**
 * Script Property key for the approver email address.
 * Set via: Apps Script Editor → Project Settings → Script Properties
 *   Key:   APPROVER_EMAIL
 *   Value:  e.g. det225-finance@nd.edu
 */
const APPROVER_EMAIL_PROPERTY_KEY = "APPROVER_EMAIL";

/** Name of the sheet tab where processed requests are stored. */
const REQUESTS_SHEET_NAME = "Purchase Requests";

// ---------------------------------------------------------------------------
// getApproverEmail()
// Reads the approver email from Script Properties. Throws if not configured.
// ---------------------------------------------------------------------------

function getApproverEmail(): string {
  const props = PropertiesService.getScriptProperties();
  const email = props.getProperty(APPROVER_EMAIL_PROPERTY_KEY);
  if (!email) {
    throw new Error(
      `Script Property "${APPROVER_EMAIL_PROPERTY_KEY}" is not set. ` +
        "Go to Apps Script Editor → Project Settings → Script Properties and add it."
    );
  }
  return email;
}

// ---------------------------------------------------------------------------
// COLUMN INDEX MAP (0-based)
// Keep in sync with the header row written by setupSheet().
// ---------------------------------------------------------------------------

const COL = {
  STATUS: 0,
  TIMESTAMP: 1,
  EMAIL: 2,
  MECHANISM: 3,
  NAME: 4,
  PHONE: 5,
  VENDOR_NAME: 6,
  ADDRESS: 7,
  OTHER_VENDORS: 8,
  VENDOR_REASON: 9,
  PURCHASE_DATE: 10,
  ITEMIZED_TABLE: 11,
  ADDITIONAL_COMMENTS: 12,
  SIGNATURE: 13,
  DECISION_NOTES: 14,
  DECISION_TIMESTAMP: 15,
} as const;

/** Headers written to the first row of the "Purchase Requests" sheet. */
const HEADERS: string[] = [
  "Status",
  "Timestamp",
  "Email Address",
  "Mechanism of Purchase",
  "Requisitioner Name",
  "Phone Number",
  "Vendor Name",
  "Address",
  "Other Vendors",
  "Vendor Reason",
  "Purchase Date",
  "Itemized Table",
  "Additional Comments",
  "Confirmation Signature",
  "Decision Notes",
  "Decision Timestamp",
];

/** Valid status values for the dropdown. */
const STATUS_VALUES = ["Pending", "Approved", "Denied"] as const;

// ---------------------------------------------------------------------------
// FORM FIELD INDEX MAP (0-based, matches Google Form question order)
// Index 0 = Timestamp, Index 1 = Email Address, then form questions follow.
// Adjust these indices if the form question order changes.
// ---------------------------------------------------------------------------

const FORM = {
  TIMESTAMP: 0,
  EMAIL: 1,
  MECHANISM: 2,
  NAME: 3,
  PHONE: 4,
  VENDOR_NAME: 5,
  ADDRESS: 6,
  OTHER_VENDORS: 7,
  VENDOR_REASON: 8,
  PURCHASE_DATE: 9,
  ITEMIZED_TABLE: 10,
  ADDITIONAL_COMMENTS: 11,
  SIGNATURE: 12,
} as const;

// ---------------------------------------------------------------------------
// setupSheet()
// Creates the "Purchase Requests" sheet if it doesn't exist, writes headers,
// and applies Status dropdown validation to the entire Status column.
// Run this once manually or from the Script Editor.
// ---------------------------------------------------------------------------

function setupSheet(): void {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(REQUESTS_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(REQUESTS_SHEET_NAME);
  }

  // Write headers if row 1 is empty or doesn't match.
  const existingHeaders = sheet
    .getRange(1, 1, 1, HEADERS.length)
    .getValues()[0];
  const needsHeaders = existingHeaders.every(
    (cell: string | number | boolean | Date) => cell === ""
  );

  if (needsHeaders) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }

  // Apply dropdown validation to the entire Status column (excluding header).
  const lastRow = Math.max(sheet.getMaxRows(), 100);
  const statusRange = sheet.getRange(2, COL.STATUS + 1, lastRow - 1, 1);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(STATUS_VALUES as unknown as string[], true)
    .setAllowInvalid(false)
    .build();
  statusRange.setDataValidation(rule);
}

// ---------------------------------------------------------------------------
// onFormSubmit(e)
// Installable trigger handler — fires when the linked Google Form submits.
// ---------------------------------------------------------------------------

function onFormSubmit(e: GoogleAppsScript.Events.SheetsOnFormSubmit): void {
  if (!e || !e.values) {
    Logger.log("onFormSubmit: No event data received.");
    return;
  }

  const values = e.values;

  // Ensure the target sheet exists.
  setupSheet();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(REQUESTS_SHEET_NAME);
  if (!sheet) {
    Logger.log("onFormSubmit: Could not find sheet: " + REQUESTS_SHEET_NAME);
    return;
  }

  // Extract form fields (with safe fallbacks).
  const timestamp = values[FORM.TIMESTAMP] ?? "";
  const email = values[FORM.EMAIL] ?? "";
  const mechanism = values[FORM.MECHANISM] ?? "";
  const name = values[FORM.NAME] ?? "";
  const phone = values[FORM.PHONE] ?? "";
  const vendorName = values[FORM.VENDOR_NAME] ?? "";
  const address = values[FORM.ADDRESS] ?? "";
  const otherVendors = values[FORM.OTHER_VENDORS] ?? "";
  const vendorReason = values[FORM.VENDOR_REASON] ?? "";
  const purchaseDate = values[FORM.PURCHASE_DATE] ?? "";
  const itemizedTableRaw = values[FORM.ITEMIZED_TABLE] ?? "";
  const additionalComments = values[FORM.ADDITIONAL_COMMENTS] ?? "";
  const signature = values[FORM.SIGNATURE] ?? "";

  // Convert uploaded file ID(s) to full Drive link(s).
  const itemizedTableLink = buildDriveLinks(itemizedTableRaw);

  // Build the row for "Purchase Requests".
  const newRow: (string | Date)[] = [
    "Pending", // Status
    timestamp,
    email,
    mechanism,
    name,
    phone,
    vendorName,
    address,
    otherVendors,
    vendorReason,
    purchaseDate,
    itemizedTableLink,
    additionalComments,
    signature,
    "", // Decision Notes
    "", // Decision Timestamp
  ];

  sheet.appendRow(newRow);

  // Re-apply validation to the newly added row's Status cell.
  const lastRow = sheet.getLastRow();
  const statusCell = sheet.getRange(lastRow, COL.STATUS + 1);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(STATUS_VALUES as unknown as string[], true)
    .setAllowInvalid(false)
    .build();
  statusCell.setDataValidation(rule);

  // Send notification email to the approver.
  sendSubmissionEmail({
    timestamp,
    email,
    mechanism,
    name,
    phone,
    vendorName,
    address,
    otherVendors,
    vendorReason,
    purchaseDate,
    itemizedTableLink,
    additionalComments,
    signature,
    spreadsheetUrl: ss.getUrl(),
  });
}

// ---------------------------------------------------------------------------
// onEdit(e)
// Installable trigger handler — fires on manual cell edits.
// Watches only the Status column of "Purchase Requests".
// ---------------------------------------------------------------------------

function onEdit(e: GoogleAppsScript.Events.SheetsOnEdit): void {
  if (!e || !e.range) {
    return;
  }

  const sheet = e.range.getSheet();

  // Only act on edits within the "Purchase Requests" sheet.
  if (sheet.getName() !== REQUESTS_SHEET_NAME) {
    return;
  }

  const editedCol = e.range.getColumn(); // 1-based
  const editedRow = e.range.getRow();

  // Ignore header row.
  if (editedRow <= 1) {
    return;
  }

  // Only react to changes in the Status column.
  if (editedCol !== COL.STATUS + 1) {
    return;
  }

  const newValue = e.range.getValue() as string;
  const oldValue = (e.oldValue ?? "") as string;

  // Only trigger when transitioning FROM "Pending" to "Approved" or "Denied".
  if (oldValue !== "Pending") {
    return;
  }
  if (newValue !== "Approved" && newValue !== "Denied") {
    return;
  }

  // Prevent duplicate emails: check if Decision Timestamp is already set.
  const decisionTimestampCell = sheet.getRange(
    editedRow,
    COL.DECISION_TIMESTAMP + 1
  );
  const existingDecisionTimestamp = decisionTimestampCell.getValue();
  if (existingDecisionTimestamp !== "" && existingDecisionTimestamp != null) {
    return;
  }

  // Read row data.
  const rowData = sheet
    .getRange(editedRow, 1, 1, HEADERS.length)
    .getValues()[0];

  const submitterEmail = (rowData[COL.EMAIL] ?? "") as string;
  const vendorName = (rowData[COL.VENDOR_NAME] ?? "") as string;
  const purchaseDate = (rowData[COL.PURCHASE_DATE] ?? "") as string;
  const itemizedTableLink = (rowData[COL.ITEMIZED_TABLE] ?? "") as string;
  const decisionNotes = (rowData[COL.DECISION_NOTES] ?? "") as string;

  if (!submitterEmail) {
    Logger.log("onEdit: No submitter email found in row " + editedRow);
    return;
  }

  // Stamp the decision timestamp.
  const now = new Date();
  decisionTimestampCell.setValue(now);

  // Send decision email to the original submitter.
  sendDecisionEmail({
    recipientEmail: submitterEmail,
    status: newValue as "Approved" | "Denied",
    vendorName,
    purchaseDate,
    decisionNotes,
    itemizedTableLink,
  });
}

// ---------------------------------------------------------------------------
// sendSubmissionEmail()
// Notifies the approver about a new purchase request.
// ---------------------------------------------------------------------------

interface SubmissionEmailParams {
  timestamp: string;
  email: string;
  mechanism: string;
  name: string;
  phone: string;
  vendorName: string;
  address: string;
  otherVendors: string;
  vendorReason: string;
  purchaseDate: string;
  itemizedTableLink: string;
  additionalComments: string;
  signature: string;
  spreadsheetUrl: string;
}

function sendSubmissionEmail(params: SubmissionEmailParams): void {
  const approverEmail = getApproverEmail();
  const subject = `New Purchase Request - ${params.name} - ${params.vendorName}`;

  const body = [
    "A new purchase request has been submitted.",
    "",
    `Requisitioner: ${params.name} (${params.email})`,
    `Phone: ${params.phone}`,
    `Timestamp: ${params.timestamp}`,
    `Mechanism of Purchase: ${params.mechanism}`,
    "",
    "--- Vendor Information ---",
    `Vendor Name: ${params.vendorName}`,
    `Address: ${params.address}`,
    `Other Vendors Considered: ${params.otherVendors}`,
    `Vendor Reason: ${params.vendorReason}`,
    "",
    "--- Purchase Details ---",
    `Purchase Date: ${params.purchaseDate}`,
    `Itemized Table: ${params.itemizedTableLink || "N/A"}`,
    `Additional Comments: ${params.additionalComments || "N/A"}`,
    `Confirmation Signature: ${params.signature || "N/A"}`,
    "",
    "Review this request in the spreadsheet:",
    params.spreadsheetUrl,
  ].join("\n");

  try {
    MailApp.sendEmail(approverEmail, subject, body);
  } catch (err) {
    Logger.log("sendSubmissionEmail error: " + err);
  }
}

// ---------------------------------------------------------------------------
// sendDecisionEmail()
// Notifies the original submitter about an approval / denial.
// ---------------------------------------------------------------------------

interface DecisionEmailParams {
  recipientEmail: string;
  status: "Approved" | "Denied";
  vendorName: string;
  purchaseDate: string;
  decisionNotes: string;
  itemizedTableLink: string;
}

function sendDecisionEmail(params: DecisionEmailParams): void {
  const subject = `Purchase Request ${params.status} - ${params.vendorName}`;

  const lines = [
    `Your purchase request has been ${params.status.toLowerCase()}.`,
    "",
    `Vendor: ${params.vendorName}`,
    `Purchase Date: ${params.purchaseDate}`,
    `Status: ${params.status}`,
  ];

  if (params.decisionNotes) {
    lines.push(`Decision Notes: ${params.decisionNotes}`);
  }

  if (params.itemizedTableLink) {
    lines.push("", `Itemized Table: ${params.itemizedTableLink}`);
  }

  const body = lines.join("\n");

  try {
    MailApp.sendEmail(params.recipientEmail, subject, body);
  } catch (err) {
    Logger.log("sendDecisionEmail error: " + err);
  }
}

// ---------------------------------------------------------------------------
// Utility: buildDriveLinks()
// Converts a raw file-upload value (comma-separated Drive file IDs or URLs)
// into clickable Drive links.
// ---------------------------------------------------------------------------

function buildDriveLinks(raw: string): string {
  if (!raw || raw.trim() === "") {
    return "";
  }

  // Google Forms file uploads store the value as a comma-separated list of
  // Drive file IDs or full URLs. Handle both cases.
  return raw
    .split(",")
    .map((entry) => entry.trim())
    .filter((entry) => entry.length > 0)
    .map((entry) => {
      // Already a full URL — return as-is.
      if (entry.startsWith("http")) {
        return entry;
      }
      // Treat as a bare file ID.
      return `https://drive.google.com/file/d/${entry}/view`;
    })
    .join(", ");
}
