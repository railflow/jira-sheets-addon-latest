---
sidebar_position: 6
---

# Roadmap Planning

Visualize your project timeline instantly with the **Roadmap** feature. This tool transforms your issue list into a Gantt-style chart, helping you plan sprints and deadlines effectively.

![Roadmap View](/img/screenshot-roadmap-1280x800.png)

## Generating a Roadmap

1. Ensure you have synced issues into your main sheet (so the add-on has data to work with).
2. Go to the **Roadmap** tab in the sidebar (Gantt chart icon).
3. Click **Generate Roadmap**.
4. A new sheet named **"Jira Roadmap"** will be created automatically.

## Understanding the View

- **Start Dates**: Based on `created` or custom start date fields.
- **Due Dates**: Based on `duedate`.
- **Progress**: Visualized based on issue status (Todo = 0%, In Progress = 50%, Done = 100%).
- **Assignees**: Color-coded bars help identify workload distribution.

:::note
The roadmap reflects the snapshot of data currently in your sheet. To update the roadmap, simply generate it again.
:::
