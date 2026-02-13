---
sidebar_position: 7
---

# Label Management

Manage your Jira labels at scale across all issues in your spreadsheet.

## Organizing Labels

The **Label Management** tab provides a consolidated view of every label present in your current data fetch.

1. Go to the **Label Management** tab (tags icon 🏷️).
2. You will see a list of all unique labels and their frequency.

## Global Actions

### Renaming Labels
You can rename a label across every issue in your project:
1. Click the **Edit** (pencil) icon next to a label.
2. Enter the new name.
3. The add-on will find every issue with that label and update it in Jira automatically.

### Removing Labels
To remove a label entirely:
1. Click the **Trash** icon next to a label.
2. Confirm the deletion.
3. The label will be stripped from all issues currently in your dataset, and the change will be synced back to Jira.

:::warning
Removing or renaming a label affects live Jira data. Ensure you have the correct JQL filter active before performing bulk label operations.
:::
