# README: Google Docs Add-on Setup for Internal Installation (UpHill Solutions)

This guide provides the complete step-by-step process for creating, publishing, and force-installing a private Google Docs add-on for an entire Google Workspace organization.

The goal is to have the add-on pre-approved by an admin and automatically available to all users inside their Google Docs Extensions menu, with no authorization prompts for the end-user.

## Prerequisites

- **Google Workspace Super Admin:** You must have Super Admin privileges for your organization (e.g., uphillsolutions.tech).
- **Google Cloud Platform (GCP):** You need access to GCP to create a project that will be linked to your organization.
- **Apps Script Project:** The script you want to deploy.

## Step 1: Prepare the Apps Script Project

This involves two files: the script itself (Code.gs) and its manifest (appsscript.json).

### Include the Manifest (appsscript.json)

This file defines what your add-on is, what permissions it needs, and how it runs. Your manifest must include the addOns block to be recognized as an Editor Add-on.

### The Code (Code.gs)

Your script must contain the onOpen(e) function specified in the manifest. This function is responsible for building the menu that appears in Google Docs.

**Important:**

- The function **must** accept the event object (e).
- The onOpen function should _only_ build the menu. Do not call any services that require authorization (like PropertiesService) here, as it will cause the function to fail silently.

## Step 2: Create and Configure the GCP Project

To publish, your Apps Script project must be linked to a standard Google Cloud Platform (GCP) project.

1. **Create a New GCP Project:**
   - Go to the [GCP Console](https://console.cloud.google.com).
   - Create a new project (e.g., "UpHill Docs Add-on").
   - **Crucially:** When creating it, ensure it is associated with your **Organization** (e.g., "uphillsolutions.tech").
2. **Link Apps Script to GCP:**
   - In the GCP Console, find and copy the **Project Number** from the new project's dashboard.
   - In your Apps Script editor, go to **Project Settings (⚙️)**.
   - Scroll to "Google Cloud Platform (GCP) Project" and click **Change project**.
   - Paste the **Project Number** and click **Set project**.
3. **Configure the OAuth Consent Screen (Critical Step):**
   - In your new GCP project, go to **"APIs & Services" > "OAuth consent screen"**.
   - Select **Internal** for the "User Type."
   - Fill in the required fields (App name, support email).
   - Click **Save**. _This step is required to unlock the "Private" visibility setting._
4. **Enable Required APIs:**
   - In your new GCP project, go to **"APIs & Services" > "Library"**.
   - Search for and **Enable** these three APIs:
     1. Google Docs API
     2. Google Drive API
     3. Google Workspace Marketplace SDK

## Step 3: Publish with the Marketplace SDK

This is where you configure the add-on's listing and point it to your script.

1. **Go to the Marketplace SDK:**
   - In your GCP project, search for and navigate to the **"Google Workspace Marketplace SDK"**.
2. **Fill out "App Configuration":**
   - **App Visibility:** Select **Private**. (If this is disabled, your OAuth screen is not set to "Internal". You may need to create a new GCP project if this is permanently locked).
   - **Installation Settings:** Select **Admin Only Install**. This prevents users from installing it themselves and ensures only admins can force-install it.
   - **App Integration:** This is the most important step.
     - **UNCHECK** "Google Workspace add-on." (This creates a global sidebar add-on and will not run onOpen).
     - **CHECK** "Docs add-on."
   - **Enter Script Details:**
     - **Docs add-on Project Script ID:** Get this from your Apps Script **Project Settings (⚙️) > Script ID**.
     - **Docs add-on script version:** Get this from Apps Script by going to **Deploy > Manage deployments**. Create a new "Add-on" deployment if one doesn't exist, and use the **Version number** (e.g., 1 or 2).
3. **Fill out "Store Listing":**
   - Fill in all required fields, including descriptions, developer info, and all required icon and screenshot graphics.
   - **OAuth Scopes:** Manually add the same scopes from your appsscript.json (e.g., <https://www.googleapis.com/auth/documents>).
4. **Publish:**
   - Once all tabs are complete, click **Publish**. Since it's a private app, it will be published almost immediately without a Google review.

## Step 4: Install the Add-on from the Admin Console

This is the final step, where you force-install the add-on for your entire organization.

1. Go to the [Google Admin Console](https://admin.google.com).
2. Navigate to **Apps > Google Workspace Marketplace apps > App list**.
3. Click **"Install app"**.
4. A new window will open. Find the **"Internal Apps"** category (this is your organization's private store).
5. Find your new add-on and click **"Admin install"**.
6. Choose to install it for your **"Entire organization"**.
7. Review the permissions and click **"Allow"**. This is the **pre-approval step** that grants consent on behalf of all users.

## How to Update Your Add-on

1. Make your code changes in Code.gs.
2. In Apps Script, go to **Deploy > Manage deployments**.
3. Click the **Edit (✏️)** icon on your active add-on deployment.
4. From the **"Version"** dropdown, select **"New version"**.
5. Click **Deploy**.
6. Be sure to update the version in Marketplace SDK.

**Update Wait:** Just like the initial install, updates for domain-installed add-ons are **not instant**. It can take 24-48 hours for the new version to roll out to all users.

## Useful Links

- [Google Workspace Add-on Overview](https://developers.google.com/workspace/add-ons/overview)
- [Editor Add-on Authorization Lifecycle](https://developers.google.com/workspace/add-ons/concepts/editor-auth-lifecycle)
- [Editor Add-on Manifest Structure](https://developers.google.com/apps-script/manifest/editor-addons)
