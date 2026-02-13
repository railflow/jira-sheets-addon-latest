---
sidebar_position: 8
---

# JQL Bulk Update (JQLU)

Perform powerful bulk updates to Jira issues using a natural, SQL-like syntax.

## JQLU Syntax

The **JQL Bulk Update** tool (lightning icon ⚡) allows you to modify fields for a large group of issues at once.

### Format
`UPDATE field = value WHERE jql_condition`

### Examples
- **Change Assignee**: `UPDATE assignee = "John Doe" WHERE status = "Todo"`
- **Bulk Close**: `UPDATE status = "Done" WHERE resolution = Unresolved AND created < -30d`
- **Set Priority**: `UPDATE priority = "High" WHERE issuetype = "Bug" AND summary ~ "Urgent"`

## How to execute
1. Type your command in the editor.
2. Click **Execute Bulk Update**.
3. Review the status logs to see the success/failure of each individual issue update.

:::tip
Always test your `WHERE` condition first in the **Jira Issues** tab to ensure you are targeting the correct issues.
:::
