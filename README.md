# google-apps-scripts-personal
My personal google apps scripts projects, maybe you can take some inspiration from these. But they're intended for my own personal usage.

## Project Structure

This repository is organized by Apps Script project type:
- `sheets/` - Google Sheets add-ons and automation scripts
- `add-ons/` - Google Workspace Editor add-ons (Docs, Sheets, etc.)

## Projects

### Google Docs Add-on (`add-ons/google-docs/`)
A menu-based Google Docs add-on with utilities for document formatting and automation.

**Features:**
- **Meeting Notes Formatter**: Auto-dates headers and applies consistent formatting to meeting notes
- **Empty Paragraph Remover**: Cleans up document formatting by removing empty paragraphs
- **Title Case Converter**: Converts headings to proper title case
- **Milestone Cost Calculator**: Calculates project costs with dynamic hourly rates

**Setup:**
- See `add-ons/google-docs/README.md` for complete installation and deployment instructions
- Includes guidance for publishing as a private organization-wide add-on
- Configured for internal use with admin-managed installation

**Key Files:**
- `Code.js`: Main add-on logic with custom menu and utilities
- `appsscript.json`: Manifest with OAuth scopes and runtime configuration
- `README.md`: Comprehensive setup guide for Google Workspace deployment

### Time Tracker (`sheets/time-tracker/`)
A comprehensive time tracking system for Google Sheets with professional PDF export capabilities.

**Features:**
- **Automatic Time Calculation**: Tracks elapsed time when task status changes from "In Progress" to completion states
- **PDF Report Generation**: Creates professional time reports using customizable Google Docs templates
- **Company/Project Filtering**: Organizes time entries by company and project for targeted reporting
- **Rich Text Support**: Processes HTML descriptions with formatting, links, lists, and headers
- **Drive Integration**: Automatically organizes reports in company-specific folders

**Setup Requirements:**
- Google Sheet with named ranges: `assignmentStatus`, `companyNames`, `companyProjects`
- Time tracking data in sheet named "🕑 Time"
- Google Docs template with placeholders (`{{companyName}}`, `{{totalHours}}`, etc.)
- Destination folder in Google Drive for report storage

**Architecture:**
The project is modularized into focused components:
- `index.js`: Main entry point with menu integration
- `time-tracking.js`: Core time tracking business logic
- `sheet-data.js`: Google Sheets data operations
- `data.js`: Data retrieval and filtering
- `export-to-pdf.js`: PDF export orchestration
- `document-generator.js`: Google Docs formatting and content generation
- `document-config.js`: Configuration and styling constants
- `dialog.js`: UI dialog handling
- `utils.js`: Shared utility functions
- `ExportForm.html`: User interface for export configuration
