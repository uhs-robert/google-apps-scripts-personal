# ⏱️ Time Tracker for Google Sheets

Automatic time tracking system with PDF export for Google Sheets.

## ✨ Features

- Automatic time tracking when status changes to "In Progress"
- PDF report generation from Google Docs templates
- Company/project filtering
- HTML description support
- Auto-organization in Drive folders

## ⚙️ Setup

**Named Ranges:**

- `assignmentStatus` - Status column
- `companyNames` - List of companies
- `companyProjects` - Company to project mapping

**Sheet Structure:**

- Create sheet named "🕑 Time" with columns: Pay Rate, Company, Project, Title, Description (HTML supported), Links, Status, Start Time, Minutes

**Template:**

- Create Google Docs template with placeholders: `{{companyName}}`, `{{totalHours}}`, `{{formattedStart}}`, `{{formattedEnd}}`

**Configuration:**

- Update `templateDocId` and `destFolderId` in `document-config.js`

## 📁 Project Structure

- `index.js` - Menu creation and triggers
- `time-tracking.js` - Time calculation logic
- `sheet-data.js` - Spreadsheet data operations
- `data.js` - Data filtering
- `export-to-pdf.js` - PDF export orchestration
- `document-generator.js` - Google Docs formatting
- `document-config.js` - Configuration constants
- `dialog.js` - UI dialogs
- `ExportForm.html` - Export form interface
- `utils.js` - Utility functions

## 🚀 Usage

**Time Tracking:**

1. Change task status to "In Progress" to start tracking
2. Mark task complete to calculate elapsed time

**Generating Reports:**

1. Export > Export Time Entries to PDF
2. Select company and project
3. Click Export

## 📦 Deployment

1. Extensions > Apps Script
2. Copy all `.js` and `.html` files
3. Update `templateDocId` and `destFolderId` in `document-config.js`
4. Save and authorize when prompted
