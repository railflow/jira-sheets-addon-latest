---
sidebar_position: 1
---

# Privacy Policy

**Effective Date**: February 8, 2026

**Jira Sync for Google Sheets** ("Add-on") is committed to protecting your privacy. This Privacy Policy explains how we handle your data when you use our Google Workspace Add-on.

## Information We Collect

### 1. Jira Credentials
To function, the Add-on requires your **Jira Domain**, **Email**, and **API Token**. 
- **Storage**: These credentials are stored securely within your Google User Properties (specifically, the `PropertiesService.getUserProperties()` storage provided by Google Apps Script).
- **Access**: The Add-on only accesses these credentials to authenticate API requests directly between Google's servers and Atlassian's Jira API.
- **Transmission**: We do not transmit your credentials to any third-party servers other than Atlassian for the purpose of authentication. We do not store your credentials on our own servers.

### 2. Sheet Data
The Add-on reads and writes data to the Google Sheet where it is active.
- **Read Access**: We read cell values only to identify changes you wish to sync back to Jira or to validate data structure.
- **Write Access**: We write data fetched from Jira into your sheet.
- **Retention**: We do not retain any of your sheet data. All data resides within your Google Sheet and Jira instance.

### 3. Usage Logs
We may collect anonymized usage statistics (e.g., "sync successful", "error encountered") to improve the reliability of the service. These logs do not contain personal data or sensitive project information.

## How We Use Your Information

- To provide the core functionality of syncing data between Jira and Google Sheets.
- To troubleshoot technical issues.
- To improve the user experience.

## Data Sharing

We do not sell, trade, or otherwise transfer your personally identifiable information to outside parties. Your data remains yours.

## Contact Us

If you have questions about this Privacy Policy, please contact us at [support@example.com](mailto:support@example.com).
