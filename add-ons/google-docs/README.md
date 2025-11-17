# 📝 Google Docs Add-on Deployment

Guide for deploying this add-on organization-wide as a private Google Workspace add-on.

This is not a fun process but the benefits are that the script will be pre-authorized resulting in an easier user experience.

## ✅ Prerequisites

- Google Workspace Super Admin access
- GCP project linked to organization

## 🔧 Setup Steps

### 1. GCP Project Setup

1. Create GCP project at [console.cloud.google.com](https://console.cloud.google.com) associated with your organization
2. Link Apps Script to GCP:
   - Copy Project Number from GCP dashboard
   - Apps Script > Project Settings > Change project > Paste Project Number
3. Configure OAuth consent screen (APIs & Services > OAuth consent screen):
   - Set user type to **Internal**
4. Enable required APIs (APIs & Services > Library):
   - Google Docs API
   - Google Workspace Marketplace SDK

### 2. Marketplace SDK Configuration

In GCP, navigate to Google Workspace Marketplace SDK:

**App Configuration:**

- App Visibility: **Private**
- Installation Settings: **Admin Only Install**
- App Integration:
  - UNCHECK "Google Workspace add-on"
  - CHECK "Docs add-on"
- Script Details:
  - Script ID: From Apps Script > Project Settings
  - Script version: From Apps Script > Deploy > Manage deployments

**Store Listing:**

- Fill in descriptions, icons, and OAuth scopes matching `appsscript.json`

**Publish:**

- Click Publish (no Google review required for private apps)

### 3. Organization-Wide Installation

1. Go to [admin.google.com](https://admin.google.com)
2. Navigate to Apps > Google Workspace Marketplace apps > App list
3. Install app > Internal Apps > Select add-on
4. Admin install for entire organization
5. Review permissions and Allow

## 🔄 Updating the Add-on

1. Make code changes
2. Apps Script > Deploy > Manage deployments > Edit deployment > New version > Deploy
3. Update version in Marketplace SDK > Save Draft
4. Store Listing tab > Publish
