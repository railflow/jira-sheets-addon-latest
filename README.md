# Jira Sync for Google Sheets

[![Deploy to Apps Script](https://github.com/YOUR_USERNAME/jira-sheets-addon/actions/workflows/deploy.yml/badge.svg)](https://github.com/YOUR_USERNAME/jira-sheets-addon/actions/workflows/deploy.yml)

A powerful Google Sheets Add-on for bi-directional sync with Jira. Pull issues, update fields, and create new issues directly from your spreadsheet.

## ✨ Features

- **Bi-directional Sync** - Pull issues from Jira and push updates back
- **Smart JQL Editor** - Live validation with issue count preview
- **Bulk Operations** - Create and update multiple issues at once
- **Custom Functions** - Use `=JIRA_COUNT()` formulas in cells
- **Auto-Refresh** - Scheduled polling at configurable intervals
- **Change Highlighting** - Visual indicators for recently updated issues
- **Modern UI** - React-based sidebar with clean, intuitive design

## 📁 Project Structure

```
jira-sheets-addon/
├── Code.js              # Main Apps Script logic
├── Sidebar.html         # React UI for the sidebar
├── Logs.html            # Refresh logs viewer
├── appsscript.json      # Apps Script manifest
├── .clasp.json          # Clasp configuration
├── .claspignore         # Files to exclude from deployment
├── package.json         # NPM scripts for development
├── PUBLISHING.md        # Detailed publishing guide
├── .github/
│   └── workflows/
│       └── deploy.yml   # GitHub Actions CI/CD
└── assets/              # Icons and screenshots for Marketplace
```

## 🚀 Quick Start

### Manual Installation (Development)

1. Open a Google Sheet → **Extensions** → **Apps Script**
2. Delete default `Code.gs` and create these files:
   - `Code.js` (paste contents)
   - `Sidebar.html` (paste contents)
   - `Logs.html` (paste contents)
3. Copy `appsscript.json` content to the manifest (View → Show manifest file)
4. Save and refresh your Sheet
5. Use **Jira Sync** menu to configure

### CLI Installation (Recommended)

```bash
# Clone the repository
git clone https://github.com/YOUR_USERNAME/jira-sheets-addon.git
cd jira-sheets-addon

# Install dependencies
npm install

# Login to Google
npm run login

# Create new Apps Script project
npm run create

# Push code to Apps Script
npm run push
```

## ⚙️ Configuration

1. **Jira Sync → JIRA Configuration**
   - Jira Domain: `yourcompany.atlassian.net`
   - Email: Your Jira email
   - API Token: [Generate here](https://id.atlassian.com/manage-profile/security/api-tokens)

2. **Jira Sync → JIRA Issues**
   - Enter JQL query
   - Select columns to display
   - Click "Fetch Issues"

## 📊 Custom Functions

### JIRA_COUNT
Returns the count of issues matching a JQL query or Filter ID.

```excel
# Using JQL
=JIRA_COUNT("resolution = Unresolved ORDER BY created DESC")

# Using Filter ID
=JIRA_COUNT(12345)
```

> ⚠️ **Important:** JQL strings must be enclosed in double quotes.

## 🛠️ Development

```bash
# Push changes to Apps Script
npm run push

# Pull remote changes
npm run pull

# Open in browser
npm run open

# View execution logs
npm run logs
```

## 🌐 Cloudflare Worker Backend

The add-on uses a Cloudflare Worker as a secure proxy to handle Jira API requests, license management, and payment processing.

### Deployment

1. **Prerequisites**: [Cloudflare account](https://dash.cloudflare.com/) and [Wrangler CLI](https://developers.cloudflare.com/workers/wrangler/install-and-setup/).
2. **Login**: `npx wrangler login`
3. **Deploy**:
   ```bash
   # Create KV namespaces
   npx wrangler kv namespace create LICENSES
   npx wrangler kv namespace create ASSETS

   # Update IDs in wrangler.toml
   npx wrangler deploy

   # Set proxy secret
   npx wrangler secret put PROXY_SECRET
   ```

### Architecture
- **Worker URL**: `https://jira-proxy.railflow.workers.dev`
- **KV LICENSES**: Stores license data for users and domains.
- **KV ASSETS**: (Optional) Serves static landing page assets.
- **PROXY_SECRET**: Shared secret between Apps Script and the Worker.

## 📈 Sales Reporting

The worker automatically records all successful payments in Cloudflare KV. You can export this data as a CSV file at any time using your admin secret.

### Download Sales CSV
Replace `YOUR_SECRET` with your `PROXY_SECRET` value:

```text
https://jira-proxy.railflow.workers.dev/api/admin/sales-csv?secret=YOUR_SECRET
```

**Metadata Captured:**
- Transaction Date
- User Email
- Subscription Plan (Personal/Unlimited)
- Organization Domain
- Stripe Customer & Subscription IDs
- Transaction Amount (USD)

## 📦 Publishing to Marketplace

See [PUBLISHING.md](./PUBLISHING.md) for detailed instructions on:

1. Setting up Google Cloud Project
2. Configuring OAuth consent screen
3. Creating Marketplace listing
4. Setting up GitHub Actions CI/CD
5. Submitting for review

### GitHub Actions Secrets Required

| Secret | Description |
|--------|-------------|
| `CLASP_CREDENTIALS` | Contents of `~/.clasprc.json` after `clasp login` |

## 📄 License

MIT License - see [LICENSE](./LICENSE) for details.

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Push and open a Pull Request

## 📞 Support

- [Report Issues](https://github.com/YOUR_USERNAME/jira-sheets-addon/issues)
- [Feature Requests](https://github.com/YOUR_USERNAME/jira-sheets-addon/issues/new)

