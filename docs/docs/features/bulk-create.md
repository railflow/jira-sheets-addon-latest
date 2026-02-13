---
sidebar_position: 10
---

# Bulk Create Issues

The **Bulk Create** tool (plus-circle icon ⊕) allows you to turn multiple rows of spreadsheet data into Jira tickets in one click.

## How to Create Issues

1. **Prepare your spreadsheet**:
   - Add new rows to your spreadsheet.
   - Leave the **Key** column blank (this tells the add-on they are new).
   - Fill in the **Summary** and any other required fields (e.g., Description, Priority).
2. **Navigate to the Bulk Create tab**:
   - Click the plus-circle icon (⊕) in the sidebar.
3. **Select Target Project**:
   - Choose the Jira project where these issues should be created.
4. **Click Create**:
   - The add-on will process all rows without a Key and create them in Jira.
   - Once finished, the new Jira keys will be written back to the **Key** column automatically.

:::tip
Always make sure you have the required fields for the selected issue type. If a creation fails, check the **Logs** tab for specific Jira API error messages.
:::
