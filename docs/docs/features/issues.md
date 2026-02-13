---
sidebar_position: 2
---

# Issues Manager & Sync

The **Issues Manager** tab is the heart of Jira Sync. Here you can pull issues, update statuses, and create new tickets.

![Issues Manager Screen](/img/screenshot-issues-1280x800.png)

## Pulling Issues

Use **Jira Query Language (JQL)** to filter exactly the issues you need.

1. Navigate to the **Issues** tab (list icon).
2. Enter your JQL query in the text box.
   - Example: `project = PROJ AND status = "In Progress"`
3. Select your desired columns.
   - Click `Edit Columns` to choose from Summary, Status, Priority, Assignee, etc.
4. Click **Get Issues from Jira**.
5. The spreadsheet will populate with your data.

## Reviewing Changes & Syncing

Jira Sync supports bi-directional synchronization, meaning you can edit data in Sheets and push it back to Jira.

### Updating Existing Issues
1. Edit any cell in the spreadsheet (e.g., change Status from `To Do` to `Done`).
2. Changed rows will be highlighted.
3. Click **Update Jira** in the sidebar.
4. Confirm the changes.

### Deleting Issues
1. Select the rows or issues you wish to remove.
2. Click **Delete Issues** (trash icon) in the sidebar.
3. Confirm deletion. **Warning**: This action is permanent in Jira.

## Auto-Refresh

Keep your data up-to-date automatically using the refresh settings at the bottom of the sidebar.
- **Interval**: Choose how often to sync (e.g., every 30 seconds).

## Smart Column Discovery

The add-on now features **Smart Column Detection** to ensure smooth performance:
- **Automatic Fields**: If you attempt to generate a Roadmap or Metric report, the add-on will automatically verify if required fields like `Status`, `Assignee`, or `Created` are present in your sheet.
- **Zero-Config Updates**: If a required column is missing, the add-on will intelligently add it to your configuration and refresh the data automatically so you can proceed without manual technical setup.
- **Lowercase Normalization**: You can type columns with any capitalization (e.g., "Created" vs "created"); the add-on will correctly map them to the corresponding Jira data.
