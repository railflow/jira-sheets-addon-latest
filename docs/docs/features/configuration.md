---
sidebar_position: 1
slug: /
---

# Configuration

Before you can sync data, you must connect your Jira account. The configuration process is secure and only takes a minute.

![Configuration Screen](/img/screenshot-config-1280x800.png)

## Setting Up Your Connection

1. Open the **Jira Sync** sidebar from **Extensions > Jira Sync > Open Sidebar**.
2. Navigate to the **Configuration** tab (the gear icon).
3. Enter your **Jira Domain** (e.g., `your-company.atlassian.net`). Do not include `https://`.
4. Enter your **Jira Email** associated with your account.
5. Provide your **API Token**.

:::tip Need an API Token?
You can generate a new API token from your [Atlassian Account Settings](https://id.atlassian.com/manage-profile/security/api-tokens). This token is more secure than using your password.
:::

6. Click **Save Configuration**.
7. Click **Test Connection** to verify everything is working correctly.

## Privacy & Security

Your API token is stored securely within your Google user properties and is never sent to our servers except to authorize requests directly to Jira. All communication happens directly between Google Apps Script and Atlassian's API.
