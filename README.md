# 📜 Google Apps Script Projects

Personal Google Apps Script projects for document automation and time tracking.

## 📝 Google Docs Add-on

**Location**: `add-ons/google-docs/`

Menu-based add-on with document utilities:

- Meeting notes formatter with auto-dating
- Empty paragraph remover
- Title case converter for headings
- Milestone cost calculator

See `add-ons/google-docs/README.md` for setup instructions.

## ⏱️ Time Tracker

**Location**: `sheets/time-tracker/`

Time tracking system for Google Sheets with PDF export:

- Automatic time calculation based on task status
- PDF reports from customizable templates
- Company/project filtering
- HTML formatting support

**Setup Requirements:**

- Named ranges: `assignmentStatus`, `companyNames`, `companyProjects`
- Sheet named "🕑 Time" for time data
- Google Docs template for reports
- Google Drive folder for output
