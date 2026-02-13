---
sidebar_position: 4
---

# Capacity Planning

Manage your team's workload and optimize resources by visualizing capacity vs. demand.

## Overview

The **Capacity Planning** feature (users icon 👥) calculates how much work your team can handle within a specific timeframe based on their availability, holidays, and point velocity.

## Configuration

1. **Estimation Type**: Choose between **Story Points** or **Hours**.
2. **Hours per Story Point**: (If using Points) The conversion factor for calculations.
3. **Hours/Week**: Standard working hours per team member (default 40).
4. **Capacity %**: Target utilization (default 80% to account for meetings/admin).
5. **Holidays**: Comma-separated dates (YYYY-MM-DD) to exclude from available time.
6. **Weeks Ahead**: The planning horizon for the report.

## Team Exclusions

You can exclude specific individuals from the capacity report (e.g., consultants or managers who don't take tasks):
1. Click **Sync Users** to find all assignees in your current sheet.
2. Uncheck the box next to any user you wish to exclude.

## Generating the Plan

Click **Generate Capacity Plan** to create a new sheet named **"Jira Capacity Plan"**. This sheet will show:
- **Total Capacity**: Total available hours for the team.
- **Assigned Work**: Committed hours based on current Jira issues.
- **Unassigned Work**: Remaining capacity or overallocation.
- **Visual Heatmap**: Color-coded cells indicating team health.
