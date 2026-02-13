# Publishing Guide for Jira Sync Add-on

This guide covers all distribution options for the Jira Sync add-on.

---

## 📦 Distribution Options (Without Marketplace)

If you don't want to publish to the Google Workspace Marketplace, you have several alternatives:

### Option 1: Direct Install Link (Simplest)

After deploying your Apps Script project, share the install URL directly:

1. **Deploy the Add-on:**
   ```bash
   npm run push
   ```

2. **Create a Deployment:**
   - Open Apps Script Editor (`npm run open`)
   - Go to **Deploy → New deployment**
   - Select type: **Editor Add-on**
   - Add a description and click **Deploy**
   - Copy the Deployment ID

3. **Share the Install Link:**
   ```
   https://script.google.com/d/SCRIPT_ID/edit?usp=sharing
   ```
   
   Users can install by:
   - Opening the shared link
   - Going to **Deploy → Test deployments**
   - Clicking **Install**

### Option 2: Domain-Wide Installation (For Organizations)

For Google Workspace organizations, the admin can deploy the add-on to all users:

1. **Get the Deployment ID** (from Option 1)

2. **Admin Console Installation:**
   - Go to [Google Admin Console](https://admin.google.com)
   - Navigate to **Apps → Google Workspace Marketplace apps**
   - Click **Add app → Add internal app**
   - Enter your Script ID and configure access

3. **Configure via Admin SDK:**
   ```
   The add-on will be automatically available for all users in the domain.
   ```

### Option 3: Template Sheet (Best for End Users)

Create a template Google Sheet that includes the add-on:

1. **Create a master spreadsheet** with the add-on bound to it
2. **Configure the template** with any default settings
3. **Share as "View Only"** or via Google Drive template gallery
4. Users **make a copy** to get the add-on automatically

**Pros:** No installation required, instant setup  
**Cons:** Users must copy the template to use it

### Option 4: npm Package for Developers

Package the source code for other developers to use:

1. **Create a distributable package:**
   - Include all `.gs` and `.html` files
   - Document the setup process
   - Provide environment variable templates

2. **Share via npm or GitHub:**
   ```bash
   npm pack
   # or
   git clone https://github.com/your-repo/jira-sheets-addon
   ```

3. **Recipient installs:**
   ```bash
   npm install
   npm run login
   npm run push
   ```

### Comparison Table

| Method | Ease of Use | Maintenance | Best For |
|--------|-------------|-------------|----------|
| Direct Link | ⭐⭐⭐ | Updates deploy instantly | Small teams, beta testing |
| Domain-Wide | ⭐⭐⭐⭐⭐ | Admin-controlled | Enterprise/Organizations |
| Template Sheet | ⭐⭐⭐⭐ | Users have standalone copy | Non-technical users |
| npm Package | ⭐⭐ | Recipients manage updates | Developers, customization |
| Marketplace | ⭐⭐⭐⭐⭐ | Automatic updates | Public distribution |

---

## 🏪 Marketplace Publishing (Full Guide)

## Prerequisites

1. **Google Cloud Project** - Create one at [console.cloud.google.com](https://console.cloud.google.com)
2. **Google Workspace Developer Account** - $5 one-time fee at [developers.google.com/workspace/marketplace](https://developers.google.com/workspace/marketplace)
3. **Node.js** - Version 18+ installed locally
4. **clasp** - Google's CLI for Apps Script

## Initial Setup

### 1. Install Dependencies

```bash
npm install
```

### 2. Login to clasp

```bash
npm run login
```

This opens a browser for Google OAuth authentication.

### 3. Create Apps Script Project (First Time Only)

```bash
npm run create
```

This creates a new Apps Script project linked to Google Sheets.

### 4. Update .clasp.json

After creating the project, update `.clasp.json` with your Script ID:

```json
{
  "scriptId": "YOUR_ACTUAL_SCRIPT_ID",
  "rootDir": "."
}
```

Find your Script ID in the Apps Script editor URL: `https://script.google.com/home/projects/SCRIPT_ID/edit`

## Development Workflow

### Push Changes
```bash
npm run push
```

### Pull Remote Changes
```bash
npm run pull
```

### Open in Browser
```bash
npm run open
```

### View Logs
```bash
npm run logs
```

## 📚 Hosting Documentation

The documentation site in `docs/` must be hosted publicly to provide the required Privacy Policy and Terms of Service URLs.

### 1. Build the Site
```bash
cd docs
npm run build
```

### 2. Deploy
Deploy the `docs/build` directory to a static hosting provider.

**Netlify / Vercel:**
- Connect your repository.
- Set Build Command: `npm run build`
- Set Publish Directory: `docs/build` (or just `build` if valid)

**GitHub Pages:**
- Enable GitHub Pages in repo settings.
- Push the build artifacts or use a workflow to deploy `docs/build`.

### 3. Update URLs
Once deployed, update the following:
1. **Sidebar.html**: Update the `https://your-domain.com` links in the footer.
2. **docusaurus.config.js**: Update `url` to your actual domain.
3. **Marketplace SDK**: Update Privacy and Terms URLs.

## Google Cloud Configuration

### 1. Link GCP Project to Apps Script

1. Open Apps Script Editor
2. Go to **Project Settings** (gear icon)
3. Under "Google Cloud Platform (GCP) Project", click **Change project**
4. Enter your GCP Project Number

### 2. Configure OAuth Consent Screen

1. Go to [GCP Console → APIs & Services → OAuth consent screen](https://console.cloud.google.com/apis/credentials/consent)
2. Choose **External** for public access
3. Fill in required fields:
   - App name: `Jira Sync for Sheets`
   - User support email: Your email
   - Developer contact: Your email
4. Add scopes:
   - `https://www.googleapis.com/auth/spreadsheets.currentonly`
   - `https://www.googleapis.com/auth/script.external_request`
5. Save and continue

### 3. Enable Marketplace SDK

1. Go to [GCP Console → APIs & Services → Library](https://console.cloud.google.com/apis/library)
2. Search for "Google Workspace Marketplace SDK"
3. Click **Enable**

### 4. Configure Marketplace SDK

1. Go to [Marketplace SDK Configuration](https://console.cloud.google.com/apis/api/appsmarket-component.googleapis.com/googleapps_sdk)
2. Fill in:
1.  **Store Listing:**
    - **Language:** English
    - **App Name:** Jira Sync for Sheets
    - **App Description:** Sync Jira issues directly to Google Sheets with bi-directional updates.
    - **Detailed Description:** Paste your `README.md` content here.
    - **Short Description:** Connect Jira & Sheets.

2.  **Graphics:**
    - **App Icon (32x32):** Resize `assets/icon-128x128.png` to 32x32.
    - **App Icon (128x128):** Upload `assets/icon-128x128.png`.
    - **Card Banner (440x280):** Upload `assets/card-banner-440x280.png`.
      - `docs/static/img/screenshot-config-1280x800.png`
      - `docs/static/img/screenshot-issues-1280x800.png`
      - `docs/static/img/screenshot-roadmap-1280x800.png`
      - `docs/static/img/screenshot-dashboard.png` (New)
      - `docs/static/img/screenshot-capacity.png` (New)

      *Note: Use the images from `docs/static/img/` as they are the most up-to-date.*

    - **Support URLs:**
      - **Privacy Policy URL:** `https://your-site.com/legal/privacy` (Point to your hosted docs)
      - **Terms of Service URL:** `https://your-site.com/legal/terms` (Point to your hosted docs)

3.  **App Configuration:**
    - **Application Integration:** Sheets Add-on
    - **Script ID:** Paste your Script ID from `.clasp.json`.
    - **Scopes:** Add the scopes listed in `appsscript.json`.
    - **Installation:** Enable "Domain Install" and "Individual Install".
    - **Visibility:** Public.

4.  **Save changes.**

## GitHub Actions Setup (CI/CD)

### 1. Get clasp Credentials

After running `clasp login`, find your credentials:

```bash
cat ~/.clasprc.json
```

### 2. Add GitHub Secret

1. Go to your GitHub repo → **Settings** → **Secrets and variables** → **Actions**
2. Create new secret named `CLASP_CREDENTIALS`
3. Paste the entire contents of `~/.clasprc.json`

### 3. Workflow Usage

- **Push on main**: Automatically pushes code to Apps Script
- **Manual deploy**: Go to Actions → Deploy → Run workflow → Select "deploy"

## Publishing to Marketplace

### 1. Create a Version

```bash
clasp version "Version 1.0.0 - Initial Release"
```

### 2. Deploy

```bash
clasp deploy --description "Production v1.0.0"
```

### 3. Submit for Review

1. Go to [Google Workspace Marketplace](https://console.cloud.google.com/apis/api/appsmarket-component.googleapis.com/googleapps_sdk)
2. Complete all required fields
3. Click **Publish**

### 4. Wait for Approval

Google reviews all public add-ons. This typically takes **1-4 weeks**.

## Post-Publishing Checklist

- [ ] Test installation from Marketplace
- [ ] Verify all features work for new users
- [ ] Set up user support channel
- [ ] Monitor error logs in GCP Console
- [ ] Create user documentation

## Troubleshooting

### "Script ID not found"
- Ensure `.clasp.json` has the correct `scriptId`
- Run `clasp login` to refresh credentials

### "Quota exceeded"
- Check GCP quotas at [console.cloud.google.com/iam-admin/quotas](https://console.cloud.google.com/iam-admin/quotas)

### "OAuth error"
- Regenerate credentials: `clasp login --creds`
- Ensure OAuth consent screen is configured

## Resources

- [clasp Documentation](https://github.com/google/clasp)
- [Apps Script Best Practices](https://developers.google.com/apps-script/guides/support/best-practices)
- [Marketplace Publishing Guide](https://developers.google.com/workspace/marketplace/how-to-publish)
