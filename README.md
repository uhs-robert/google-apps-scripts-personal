# google-apps-scripts-personal
My personal google apps scripts projects, maybe you can take some inspiration from these. But they're intended for my own personal usage.

## Projects

### time-tracker
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

**Key Components:**
- `index.js`: Core time tracking logic and menu integration
- `ExportFunction.js`: PDF generation pipeline with document processing
- `ExportForm.html`: User interface for export configuration
