# Purchase Requisition System

A Google Apps Script project (TypeScript + clasp) bound to a Google Sheet that receives Google Form submissions for purchase requests. Automates intake, email notifications, and an approval/denial workflow.

## Architecture

```
Google Form  →  Google Sheet ("Form Responses 1")  →  Apps Script
                                                          ↓
                                          ┌───────────────┼───────────────┐
                                          ↓                               ↓
                                  "Purchase Requests"            "Approval Action"
                                   (full archive)              (approver workspace)
                                                                      ↓
                                                              Email notifications
```

Everything runs within the Google Workspace ecosystem — no external services or dependencies.

## Sheet Tabs

### Form Responses 1

Raw Google Form submissions. Managed automatically by Google Forms — do not edit.

### Purchase Requests (full archive)

Every form submission is copied here with all 13 fields intact. This is the permanent record.

| Col | Header                  | Source |
|-----|-------------------------|--------|
| A   | Timestamp               | Form   |
| B   | Email Address           | Form   |
| C   | Mechanism of Purchase   | Form   |
| D   | Requisitioner Name      | Form   |
| E   | Phone Number            | Form   |
| F   | Vendor Name             | Form   |
| G   | Vendor Address          | Form   |
| H   | Other Vendors           | Form   |
| I   | Vendor Reason           | Form   |
| J   | Purchase Date           | Form   |
| K   | Itemized Table          | Form (file upload ID converted to Drive link) |
| L   | Additional Comments     | Form   |
| M   | Confirmation Signature  | Form   |

### Approval Action (approver workspace)

A slim view with only the fields an approver needs. The **Decision** column drives the workflow.

| Col | Header              | Source                                           |
|-----|---------------------|--------------------------------------------------|
| A   | Date of Request     | Form Timestamp                                   |
| B   | Requesting Cadet    | Requisitioner Name                               |
| C   | Funds Use           | Mechanism of Purchase                            |
| D   | Vendor Name         | Form                                             |
| E   | Purchase Date       | Form                                             |
| F   | Itemized Table      | Drive link from form upload                      |
| G   | **Decision**        | Dropdown: Pending / Approved / Denied            |
| H   | Additional Notes    | Approver fills in                                |
| I   | Decision Timestamp  | System (auto-stamped on decision)                |
| J   | Submitter Email     | Form (used for decision notification, far right) |

## Data Flow

### 1. Submission (automated)

1. User fills out the Google Form
2. Form writes the response to "Form Responses 1" tab
3. `onFormSubmit(e)` trigger fires
4. Script creates both tabs if missing (`setupSheet()`)
5. Extracts all 13 form fields from `e.values`
6. Converts file upload IDs to full Google Drive URLs
7. Appends a full row to **Purchase Requests**
8. Appends a slim row to **Approval Action** with Decision = "Pending"
9. Applies dropdown validation to the Decision cell
10. Emails the approver (address read from Script Properties)

### 2. Approval / Denial (manual)

1. Approver opens the **Approval Action** tab
2. Changes the Decision dropdown from **Pending** to **Approved** or **Denied**
3. Optionally fills in the "Additional Notes" column
4. `onEdit(e)` trigger fires
5. Script validates: correct sheet, correct column, Pending → Approved/Denied
6. Checks that Decision Timestamp is empty (prevents duplicate emails)
7. Stamps **Decision Timestamp** with the current time
8. Emails the original submitter with the result

## Emails

**On submission → approver:**
```
Subject: New Purchase Request - <Name> - <Vendor Name>
Body:    All form fields, Drive link, link to spreadsheet
```

**On approval → submitter:**
```
Subject: Purchase Request Approved - <Vendor Name>
Body:    Status, vendor, purchase date, decision notes, Drive link
```

**On denial → submitter:**
```
Subject: Purchase Request Denied - <Vendor Name>
Body:    Status, vendor, purchase date, decision notes, Drive link
```

## Safety Guards

- **Duplicate prevention** — Decision Timestamp is checked before sending; once stamped, subsequent edits to Decision are ignored
- **Sheet-scoped** — `onEdit` only fires on the "Approval Action" tab
- **Column-scoped** — only Decision column edits trigger logic
- **Transition-scoped** — only `Pending → Approved/Denied` transitions fire
- **Dropdown enforcement** — `DataValidation` with `setAllowInvalid(false)` prevents free-text in Decision

## Project Structure

```
purchase-requisition-system/
├── src/
│   └── index.ts           # All application logic
├── dist/                   # Compiled JS output (gitignored)
├── appsscript.json         # Apps Script manifest and OAuth scopes
├── tsconfig.json           # TypeScript config (strict, ES2020, outDir=dist)
├── package.json            # npm scripts for build/deploy via clasp
├── .clasp.json             # Script ID binding (gitignored)
└── .gitignore
```

## Functions

| Function               | Purpose                                                                |
|------------------------|------------------------------------------------------------------------|
| `setupSheet()`         | Creates both tabs, writes headers, sets Decision dropdown validation   |
| `installTriggers()`    | Programmatically installs onFormSubmit and onEdit installable triggers  |
| `onFormSubmit(e)`      | Parses form submission, writes to both tabs, emails approver           |
| `onEdit(e)`            | Watches Decision changes on Approval Action, emails submitter, stamps timestamp |
| `sendSubmissionEmail()`| Builds and sends the new-request notification                          |
| `sendDecisionEmail()`  | Builds and sends the approval/denial notification                      |
| `getApproverEmail()`   | Reads approver email from Script Properties                            |
| `buildDriveLinks()`    | Converts file upload IDs to full Drive URLs                            |

## Setup

### Prerequisites

- Node.js and npm
- A Google account with access to the target Google Sheet
- `clasp` CLI authenticated (`npx clasp login`)

### Install

```bash
npm install
```

### Configure clasp

**Option A — Create a new bound Sheet:**
```bash
npm run create
```

**Option B — Use an existing Sheet:**
1. Open the Sheet → Extensions → Apps Script → copy the Script ID from the URL
2. Create `.clasp.json`:
```json
{
  "scriptId": "YOUR_SCRIPT_ID_HERE",
  "rootDir": "dist"
}
```

### Deploy

```bash
npm run push
```

### Set the Approver Email (Script Property)

1. Run `npm run open` to open the Apps Script editor
2. Go to **Project Settings** (gear icon) → **Script Properties**
3. Add a property:
   - **Key:** `APPROVER_EMAIL`
   - **Value:** the approver's email address (e.g. `det225-finance@nd.edu`)

### Install Triggers

In the Apps Script editor, select `installTriggers` from the function dropdown and click **Run**. This programmatically creates both installable triggers:

- **onFormSubmit** — From spreadsheet, on form submit
- **onEdit** — From spreadsheet, on edit

Safe to re-run — it removes existing triggers for these functions before recreating them.

### Run Initial Setup

In the Apps Script editor, select `setupSheet` from the function dropdown and click **Run**. This creates the "Purchase Requests" and "Approval Action" tabs with headers and dropdown validation.

### OAuth Scopes

Defined in `appsscript.json`:

| Scope                              | Reason                          |
|------------------------------------|---------------------------------|
| `spreadsheets`                     | Read/write the Sheet            |
| `script.send_mail`                 | Send notification emails        |
| `drive.readonly`                   | Access uploaded file metadata   |
| `script.scriptapp`                 | Manage installable triggers     |

## npm Scripts

| Script          | Command                                            |
|-----------------|----------------------------------------------------|
| `npm run build` | Compile TypeScript                                 |
| `npm run push`  | Build + copy manifest + push to Apps Script        |
| `npm run pull`  | Pull from Apps Script + copy manifest back         |
| `npm run create`| Create a new bound Apps Script project             |
| `npm run status`| Show clasp push/pull status                        |
| `npm run logs`  | Tail Apps Script execution logs                    |
| `npm run open`  | Open the script in the browser                     |
| `npm run clean` | Delete the `dist/` directory                       |
| `npm run upgrade`| Check for dependency updates                      |

## License

MIT
