# Time Tracker for Google Sheets

A comprehensive time tracking system for Google Sheets with professional PDF export capabilities. Automatically calculates time spent on tasks and generates formatted reports.

## Features

### Automatic Time Tracking
- Monitors task status changes in your Google Sheet
- Automatically starts tracking when status changes to "In Progress"
- Calculates elapsed time when task is marked complete
- Stores timestamps and duration in the spreadsheet

### PDF Report Generation
- Creates professional time reports using customizable Google Docs templates
- Filters time entries by company and project
- Processes HTML descriptions with rich formatting (links, lists, headers)
- Automatically organizes reports in company-specific Drive folders
- Generates both PDF and editable Google Docs versions

### Smart Data Organization
- Company and project-based filtering
- Named range integration for dynamic data management
- Flexible data structure supporting multiple companies and projects

## Setup Requirements

### Google Sheet Configuration

Your spreadsheet must include the following named ranges:

1. **`assignmentStatus`**: Column containing task status values
   - Must include statuses like "In Progress", "Complete", etc.

2. **`companyNames`**: List of company names for filtering

3. **`companyProjects`**: Mapping of companies to their projects

### Sheet Structure

Create a sheet named **"🕑 Time"** with the following columns:
- Pay Rate
- Company
- Project
- Title
- Description (supports HTML formatting)
- Links
- Status
- Start Time (auto-populated)
- Minutes/Duration (auto-calculated)

### Google Docs Template

Create a Google Docs template with these placeholders:
- `{{companyName}}` - Company name
- `{{totalHours}}` - Total hours for the period
- `{{formattedStart}}` - Report period start date
- `{{formattedEnd}}` - Report period end date
- Time entries will be automatically inserted

### Configuration

Update the following constants in `document-config.js`:
```javascript
const templateDocId = 'YOUR_TEMPLATE_DOC_ID';  // Template document ID
const destFolderId = 'YOUR_FOLDER_ID';         // Destination folder ID
```

## Architecture

The project is organized into modular components for maintainability:

### Core Modules

- **`index.js`**: Main entry point
  - Creates "Export" menu in spreadsheet UI
  - Handles `onOpen()` and `onEdit()` triggers
  - Coordinates time tracking on status changes

- **`time-tracking.js`**: Time calculation logic
  - Calculates elapsed time between start and completion
  - Updates spreadsheet with timestamps and durations

- **`sheet-data.js`**: Spreadsheet data operations
  - Reads data from named ranges
  - Extracts time tracking entries from sheets

- **`data.js`**: Data retrieval and filtering
  - Filters entries by company and project
  - Processes and structures data for export

### Export Pipeline

- **`export-to-pdf.js`**: Export orchestration
  - Coordinates the full PDF generation workflow
  - Manages folder structure and file creation
  - Provides user feedback dialogs

- **`document-generator.js`**: Document formatting
  - Populates Google Docs templates
  - Converts HTML to formatted Google Docs content
  - Handles lists, headers, links, and text styling
  - Generates final PDF from Google Docs

- **`document-config.js`**: Configuration constants
  - Template and folder IDs
  - Font sizes and styling preferences
  - Column indices for data mapping

### UI Components

- **`dialog.js`**: Dialog utilities
  - Shows success/error messages to users
  - Displays HTML-formatted notifications

- **`ExportForm.html`**: Export interface
  - Company and project selection dropdowns
  - Form validation and submission
  - Loading states and user feedback

- **`utils.js`**: Shared utilities
  - Common helper functions
  - Date formatting
  - File and folder management

## Usage

### Setting Up Time Tracking

1. Open your Google Sheet
2. The script will automatically add an "Export" menu on first open
3. When you change a task status to "In Progress", tracking begins automatically
4. When you mark a task complete, elapsed time is calculated and recorded

### Generating Reports

1. Click **Export** > **Export Time Entries to PDF**
2. Select the company from the dropdown
3. Select the project from the dropdown
4. Click "Export"
5. Wait for the success dialog with links to your report

The system will:
- Filter time entries for the selected company/project
- Create a new document from your template
- Populate it with formatted time entry data
- Convert it to PDF
- Save both versions in the appropriate company folder

## Customization

### Modifying the Template

Edit your Google Docs template to change:
- Report header and footer
- Company branding
- Table formatting
- Additional placeholders

### Adjusting Data Structure

If your sheet has a different column layout:
1. Update the column indices in `document-config.js`
2. Modify data extraction logic in `sheet-data.js` if needed

### Styling and Formatting

Customize document appearance in `document-config.js`:
- Font sizes
- Text styles
- Spacing and margins

## Deployment

1. Open your Google Sheet
2. Go to **Extensions** > **Apps Script**
3. Copy all `.js` and `.html` files from this directory
4. Update configuration constants with your template and folder IDs
5. Save and reload your spreadsheet
6. Authorize the script when prompted

## Dependencies

This project uses Google Apps Script built-in services:
- `SpreadsheetApp` - Spreadsheet operations
- `DocumentApp` - Google Docs creation and formatting
- `DriveApp` - File and folder management
- `HtmlService` - UI dialogs and forms
- `Utilities` - Date formatting and general utilities

## Troubleshooting

**Time not calculating:**
- Ensure the status column is part of the `assignmentStatus` named range
- Check that status values match exactly (case-sensitive)

**Export fails:**
- Verify template document ID and destination folder ID are correct
- Ensure you have edit access to both the template and destination folder
- Check that named ranges exist and contain data

**HTML formatting not working:**
- Verify descriptions use valid HTML tags
- Check that the HTML is properly formatted in your spreadsheet

## License

For personal use. Modify as needed for your requirements.
