// =============================================================================
// Purchase Requisition System — Google Apps Script (TypeScript)
// Bound to a Google Sheet that receives Google Form submissions.
//
// Two tabs:
//   "Purchase Requests"  — full archive of every form submission
//   "Approval Action"    — slim approver workspace with Decision dropdown
// =============================================================================

// ---------------------------------------------------------------------------
// CONFIGURATION
// ---------------------------------------------------------------------------

/**
 * Script Property key for the approver email address(es).
 * Supports a comma-separated list of emails.
 * Set via: Apps Script Editor → Project Settings → Script Properties
 *   Key:   APPROVER_EMAIL
 *   Value:  e.g. det225-finance@nd.edu,another@nd.edu
 */
const APPROVER_EMAIL_PROPERTY_KEY = "APPROVER_EMAIL";

/** Name of the full-archive sheet tab. */
const REQUESTS_SHEET_NAME = "Purchase Requests";

/** Name of the slim approver-facing sheet tab. */
const APPROVAL_SHEET_NAME = "Approval Action";

// ---------------------------------------------------------------------------
// getApproverEmails()
// Reads the approver email(s) from Script Properties. Throws if not configured.
// Returns a comma-separated string suitable for MailApp to/cc fields.
// ---------------------------------------------------------------------------

function getApproverEmails(): string {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(APPROVER_EMAIL_PROPERTY_KEY);
  if (!raw || raw.trim() === "") {
    throw new Error(
      `Script Property "${APPROVER_EMAIL_PROPERTY_KEY}" is not set. ` +
        "Go to Apps Script Editor → Project Settings → Script Properties and add it."
    );
  }
  // Normalize: trim each address, filter empties, rejoin.
  return raw
    .split(",")
    .map((e) => e.trim())
    .filter((e) => e.length > 0)
    .join(",");
}

// ---------------------------------------------------------------------------
// PURCHASE REQUESTS — COLUMN INDEX MAP (0-based)
// Full archive. Keep in sync with PR_HEADERS.
// ---------------------------------------------------------------------------

const PR_COL = {
  TIMESTAMP: 0,
  EMAIL: 1,
  MECHANISM: 2,
  NAME: 3,
  PHONE: 4,
  VENDOR_NAME: 5,
  VENDOR_ADDRESS: 6,
  OTHER_VENDORS: 7,
  VENDOR_REASON: 8,
  PURCHASE_DATE: 9,
  ITEMIZED_TABLE: 10,
  ADDITIONAL_COMMENTS: 11,
  SIGNATURE: 12,
} as const;

const PR_HEADERS: string[] = [
  "Timestamp",
  "Email Address",
  "Mechanism of Purchase",
  "Requisitioner Name",
  "Phone Number",
  "Vendor Name",
  "Vendor Address",
  "Other Vendors",
  "Vendor Reason",
  "Purchase Date",
  "Itemized Table",
  "Additional Comments",
  "Confirmation Signature",
];

// ---------------------------------------------------------------------------
// APPROVAL ACTION — COLUMN INDEX MAP (0-based)
// Slim approver workspace. Keep in sync with AA_HEADERS.
// ---------------------------------------------------------------------------

const AA_COL = {
  DATE_OF_REQUEST: 0,
  REQUESTING_CADET: 1,
  FUNDS_USE: 2,
  VENDOR_NAME: 3,
  PURCHASE_DATE: 4,
  ITEMIZED_TABLE: 5,
  DECISION: 6,
  ADDITIONAL_NOTES: 7,
  DECISION_TIMESTAMP: 8,
  SUBMITTER_EMAIL: 9,
} as const;

const AA_HEADERS: string[] = [
  "Date of Request",
  "Requesting Cadet",
  "Funds Use",
  "Vendor Name",
  "Purchase Date",
  "Itemized Table",
  "Decision",
  "Additional Notes",
  "Decision Timestamp",
  "Submitter Email",
];

/** Valid decision values for the dropdown. */
const DECISION_VALUES = ["Pending", "Approved", "Denied"] as const;

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
  VENDOR_ADDRESS: 6,
  OTHER_VENDORS: 7,
  VENDOR_REASON: 8,
  PURCHASE_DATE: 9,
  ITEMIZED_TABLE: 10,
  ADDITIONAL_COMMENTS: 11,
  SIGNATURE: 12,
} as const;

// ---------------------------------------------------------------------------
// onOpen()
// Simple trigger — adds a custom menu to the Google Sheets UI.
// ---------------------------------------------------------------------------

function onOpen(): void {
  SpreadsheetApp.getUi()
    .createMenu("Purchase Requests")
    .addItem("Setup Sheets", "setupSheet")
    .addItem("Install Triggers", "installTriggers")
    .addSeparator()
    .addItem("Refresh Decision Dropdowns", "refreshDecisionValidation")
    .addSeparator()
    .addItem("Show Approver Email", "showApproverEmail")
    .addToUi();
}

// ---------------------------------------------------------------------------
// showApproverEmail()
// Displays the currently configured approver email in an alert dialog.
// ---------------------------------------------------------------------------

function showApproverEmail(): void {
  const ui = SpreadsheetApp.getUi();
  try {
    const email = getApproverEmails();
    ui.alert("Approver Email", `Currently set to:\n${email}`, ui.ButtonSet.OK);
  } catch {
    ui.alert(
      "Approver Email Not Set",
      'Go to Apps Script Editor → Project Settings → Script Properties and add "APPROVER_EMAIL".',
      ui.ButtonSet.OK
    );
  }
}

// ---------------------------------------------------------------------------
// refreshDecisionValidation()
// Re-applies the Decision dropdown validation to the entire Approval Action
// Decision column. Useful if rows were added manually or validation was lost.
// ---------------------------------------------------------------------------

function refreshDecisionValidation(): void {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aaSheet = ss.getSheetByName(APPROVAL_SHEET_NAME);
  if (!aaSheet) {
    SpreadsheetApp.getUi().alert(
      "Sheet Not Found",
      `"${APPROVAL_SHEET_NAME}" tab does not exist. Run "Setup Sheets" first.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  const lastRow = Math.max(aaSheet.getMaxRows(), 100);
  const decisionRange = aaSheet.getRange(
    2,
    AA_COL.DECISION + 1,
    lastRow - 1,
    1
  );
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(DECISION_VALUES as unknown as string[], true)
    .setAllowInvalid(false)
    .build();
  decisionRange.setDataValidation(rule);

  SpreadsheetApp.getUi().alert("Done", "Decision dropdowns refreshed.", SpreadsheetApp.getUi().ButtonSet.OK);
}

// ---------------------------------------------------------------------------
// setupSheet()
// Creates both tabs if missing, writes headers, and applies Decision dropdown
// validation to the Approval Action tab.
// Run this once manually or from the Script Editor.
// ---------------------------------------------------------------------------

function setupSheet(): void {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- Purchase Requests (full archive) ---
  let prSheet = ss.getSheetByName(REQUESTS_SHEET_NAME);
  if (!prSheet) {
    prSheet = ss.insertSheet(REQUESTS_SHEET_NAME);
  }

  const prExisting = prSheet
    .getRange(1, 1, 1, PR_HEADERS.length)
    .getValues()[0];
  const prNeedsHeaders = prExisting.every(
    (cell: string | number | boolean | Date) => cell === ""
  );
  if (prNeedsHeaders) {
    prSheet.getRange(1, 1, 1, PR_HEADERS.length).setValues([PR_HEADERS]);
    prSheet.getRange(1, 1, 1, PR_HEADERS.length).setFontWeight("bold");
    prSheet.setFrozenRows(1);
  }

  // --- Approval Action (slim approver workspace) ---
  let aaSheet = ss.getSheetByName(APPROVAL_SHEET_NAME);
  if (!aaSheet) {
    aaSheet = ss.insertSheet(APPROVAL_SHEET_NAME);
  }

  const aaExisting = aaSheet
    .getRange(1, 1, 1, AA_HEADERS.length)
    .getValues()[0];
  const aaNeedsHeaders = aaExisting.every(
    (cell: string | number | boolean | Date) => cell === ""
  );
  if (aaNeedsHeaders) {
    aaSheet.getRange(1, 1, 1, AA_HEADERS.length).setValues([AA_HEADERS]);
    aaSheet.getRange(1, 1, 1, AA_HEADERS.length).setFontWeight("bold");
    aaSheet.setFrozenRows(1);
  }

  // Apply Decision dropdown validation to the entire Decision column.
  const aaLastRow = Math.max(aaSheet.getMaxRows(), 100);
  const decisionRange = aaSheet.getRange(
    2,
    AA_COL.DECISION + 1,
    aaLastRow - 1,
    1
  );
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(DECISION_VALUES as unknown as string[], true)
    .setAllowInvalid(false)
    .build();
  decisionRange.setDataValidation(rule);
}

// ---------------------------------------------------------------------------
// installTriggers()
// Programmatically creates the onFormSubmit and onEdit installable triggers.
// Run this once manually from the Script Editor. Safe to re-run — it clears
// existing triggers for these functions before recreating them.
// ---------------------------------------------------------------------------

function installTriggers(): void {
  const ss = SpreadsheetApp.getActive();

  // Remove existing triggers for our functions to avoid duplicates.
  const existing = ScriptApp.getProjectTriggers();
  for (const trigger of existing) {
    const handlerName = trigger.getHandlerFunction();
    if (handlerName === "onFormSubmit" || handlerName === "onEdit") {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  // Create onFormSubmit trigger.
  ScriptApp.newTrigger("onFormSubmit")
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();

  // Create onEdit trigger (installable, so it can use MailApp/PropertiesService).
  ScriptApp.newTrigger("onEdit").forSpreadsheet(ss).onEdit().create();

  Logger.log("Triggers installed: onFormSubmit, onEdit");
}

// ---------------------------------------------------------------------------
// onFormSubmit(e)
// Installable trigger handler — fires when the linked Google Form submits.
// Writes to both "Purchase Requests" (full archive) and "Approval Action"
// (slim approver workspace).
// ---------------------------------------------------------------------------

function onFormSubmit(e: GoogleAppsScript.Events.SheetsOnFormSubmit): void {
  if (!e || !e.values) {
    Logger.log("onFormSubmit: No event data received.");
    return;
  }

  const values = e.values;

  // Ensure both tabs exist.
  setupSheet();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const prSheet = ss.getSheetByName(REQUESTS_SHEET_NAME);
  const aaSheet = ss.getSheetByName(APPROVAL_SHEET_NAME);

  if (!prSheet || !aaSheet) {
    Logger.log("onFormSubmit: Could not find required sheet tabs.");
    return;
  }

  // Extract form fields (with safe fallbacks).
  const timestamp = values[FORM.TIMESTAMP] ?? "";
  const email = values[FORM.EMAIL] ?? "";
  const mechanism = values[FORM.MECHANISM] ?? "";
  const name = values[FORM.NAME] ?? "";
  const phone = values[FORM.PHONE] ?? "";
  const vendorName = values[FORM.VENDOR_NAME] ?? "";
  const vendorAddress = values[FORM.VENDOR_ADDRESS] ?? "";
  const otherVendors = values[FORM.OTHER_VENDORS] ?? "";
  const vendorReason = values[FORM.VENDOR_REASON] ?? "";
  const purchaseDate = values[FORM.PURCHASE_DATE] ?? "";
  const itemizedTableRaw = values[FORM.ITEMIZED_TABLE] ?? "";
  const additionalComments = values[FORM.ADDITIONAL_COMMENTS] ?? "";
  const signature = values[FORM.SIGNATURE] ?? "";

  // Convert uploaded file ID(s) to full Drive link(s).
  const itemizedTableLink = buildDriveLinks(itemizedTableRaw);

  // --- Write to Purchase Requests (full archive) ---
  const prRow: string[] = [
    timestamp,
    email,
    mechanism,
    name,
    phone,
    vendorName,
    vendorAddress,
    otherVendors,
    vendorReason,
    purchaseDate,
    itemizedTableLink,
    additionalComments,
    signature,
  ];
  prSheet.appendRow(prRow);

  // --- Write to Approval Action (slim approver workspace) ---
  const aaRow: string[] = [
    timestamp, // Date of Request
    name, // Requesting Cadet
    mechanism, // Funds Use
    vendorName, // Vendor Name
    purchaseDate, // Purchase Date
    itemizedTableLink, // Itemized Table
    "Pending", // Decision
    "", // Additional Notes
    "", // Decision Timestamp
    email, // Submitter Email
  ];
  aaSheet.appendRow(aaRow);

  // Apply Decision dropdown validation to the newly added row.
  const aaLastRow = aaSheet.getLastRow();
  const decisionCell = aaSheet.getRange(aaLastRow, AA_COL.DECISION + 1);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(DECISION_VALUES as unknown as string[], true)
    .setAllowInvalid(false)
    .build();
  decisionCell.setDataValidation(rule);

  // Send notification email to the approver.
  sendSubmissionEmail({
    timestamp,
    email,
    mechanism,
    name,
    phone,
    vendorName,
    vendorAddress,
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
// Watches only the Decision column of the "Approval Action" sheet.
// ---------------------------------------------------------------------------

function onEdit(e: GoogleAppsScript.Events.SheetsOnEdit): void {
  Logger.log("onEdit: trigger fired");

  if (!e || !e.range) {
    Logger.log("onEdit: no event or range — exiting");
    return;
  }

  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  Logger.log(`onEdit: sheet="${sheetName}", row=${e.range.getRow()}, col=${e.range.getColumn()}`);

  // Only act on edits within the "Approval Action" sheet.
  if (sheetName !== APPROVAL_SHEET_NAME) {
    Logger.log(`onEdit: wrong sheet (expected "${APPROVAL_SHEET_NAME}") — exiting`);
    return;
  }

  const editedCol = e.range.getColumn(); // 1-based
  const editedRow = e.range.getRow();

  // Ignore header row.
  if (editedRow <= 1) {
    Logger.log("onEdit: header row — exiting");
    return;
  }

  // Only react to changes in the Decision column.
  if (editedCol !== AA_COL.DECISION + 1) {
    Logger.log(`onEdit: wrong column (edited=${editedCol}, expected=${AA_COL.DECISION + 1}) — exiting`);
    return;
  }

  const newValue = e.range.getValue() as string;
  Logger.log(`onEdit: Decision value="${newValue}"`);

  // Only act on Approved or Denied selections.
  if (newValue !== "Approved" && newValue !== "Denied") {
    Logger.log(`onEdit: value is not Approved/Denied — exiting`);
    return;
  }

  // Prevent duplicate emails: check if Decision Timestamp is already set.
  const decisionTimestampCell = sheet.getRange(
    editedRow,
    AA_COL.DECISION_TIMESTAMP + 1
  );
  const existingDecisionTimestamp = decisionTimestampCell.getValue();
  Logger.log(`onEdit: existing Decision Timestamp="${existingDecisionTimestamp}"`);
  if (existingDecisionTimestamp !== "" && existingDecisionTimestamp != null) {
    Logger.log("onEdit: Decision Timestamp already set — exiting (duplicate prevention)");
    return;
  }

  // Read row data.
  const rowData = sheet
    .getRange(editedRow, 1, 1, AA_HEADERS.length)
    .getValues()[0];

  const submitterEmail = (rowData[AA_COL.SUBMITTER_EMAIL] ?? "") as string;
  const vendorName = (rowData[AA_COL.VENDOR_NAME] ?? "") as string;
  const purchaseDate = (rowData[AA_COL.PURCHASE_DATE] ?? "") as string;
  const itemizedTableLink = (rowData[AA_COL.ITEMIZED_TABLE] ?? "") as string;
  const additionalNotes = (rowData[AA_COL.ADDITIONAL_NOTES] ?? "") as string;

  Logger.log(`onEdit: submitterEmail="${submitterEmail}", vendor="${vendorName}", date="${purchaseDate}"`);

  if (!submitterEmail) {
    Logger.log("onEdit: No submitter email found in row " + editedRow + " — exiting");
    return;
  }

  // Stamp the decision timestamp.
  const now = new Date();
  decisionTimestampCell.setValue(now);
  Logger.log(`onEdit: Decision Timestamp stamped: ${now.toISOString()}`);

  // Send decision email to the original submitter.
  Logger.log(`onEdit: sending decision email (${newValue}) to ${submitterEmail}`);
  sendDecisionEmail({
    recipientEmail: submitterEmail,
    status: newValue as "Approved" | "Denied",
    vendorName,
    purchaseDate,
    decisionNotes: additionalNotes,
    itemizedTableLink,
  });
  Logger.log("onEdit: sendDecisionEmail returned successfully");
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
  vendorAddress: string;
  otherVendors: string;
  vendorReason: string;
  purchaseDate: string;
  itemizedTableLink: string;
  additionalComments: string;
  signature: string;
  spreadsheetUrl: string;
}

function sendSubmissionEmail(params: SubmissionEmailParams): void {
  const approverEmail = getApproverEmails();
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
    `Vendor Address: ${params.vendorAddress}`,
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
    GmailApp.sendEmail(approverEmail, subject, body);
  } catch (err) {
    Logger.log("sendSubmissionEmail error: " + err);
  }
}

// ---------------------------------------------------------------------------
// sendDecisionEmail()
// Notifies the original submitter about an approval / denial.
// CCs the approver so they have a record of the outgoing notification.
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
  const approverEmail = getApproverEmails();
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
    GmailApp.sendEmail(params.recipientEmail, subject, body, {
      cc: approverEmail,
    });
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
