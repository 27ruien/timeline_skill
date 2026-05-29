---
name: timeline-maker
description: Create Kivisense-style Excel project timeline and Gantt workbooks from simple task lists. Use when Codex needs to turn items such as "事项名称, Kivisense, brand, start date, 10 workdays" into a polished `.xlsx` timeline with China-business-day date columns, month headers, owner checkmarks, incomplete default status, square-like colored Gantt cells, end-marker stars, and Kivisense branding.
---

# Timeline Maker

## Overview

Create a client-facing Excel timeline from sparse task inputs. Match the Kivisense timeline style: one `Timeline` sheet, Kivisense/弥知科技 logo in the top-left header area, wide task description column, owner check columns, incomplete-by-default status, horizontal China-business-day calendar, merged month headers, colored bars for task durations, and an end-marker star at each task's final Gantt day.

When the user only provides content, infer the workbook structure and build the `.xlsx`; do not ask for extra fields unless the schedule cannot be inferred.

## Input Contract

Accept terse rows in Chinese or English. Normalize each row into:

```text
optional_model | task_name | owners | start_date | workday_duration | optional_color | optional_status
```

Examples:

```text
1. Project requirement, Kivisense, 2026-06-01, 5天
2. Creative Proposal - detail version, Kivisense, brand, 2026-06-08, 10天
3. Launch online, Kivisense + brand, 2026-08-18, 1 workday
4. 需求, Scope addendum, brand, 2026-06-18, 4天
```

Rules:

- `Kivisense` means check only the Kivisense owner column.
- `brand`, `Brands`, or `Client` means check only the Brands owner column.
- If both are present, check both owner columns.
- If no owner is present, leave both owner columns blank and continue.
- `5天`, `5 days`, and `5 workdays` all mean 5 business days.
- Durations are China business days. Exclude Saturdays, Sundays, and Chinese public holidays, but include official adjusted working weekends when the holiday calendar marks them as workdays.
- If a status is not explicitly provided, treat it as incomplete. In this visual style, leave `Status` blank rather than marking a completion check.
- If the user requests literal status labels, use `未完成` for default incomplete instead of blank.
- If `include_model` is enabled, treat the first field in each row as `Model`, group rows with the same model together even if they were entered non-adjacently, and merge the `Model` cells in the output.
- When grouping by Model, preserve the first-seen Model group order and sort tasks inside each Model by start date.
- If `include_status` is false, omit the `Status` column entirely.
- Support either `workdays` or `end` / `end_date`. If an end date is present, use the start/end date range for the Gantt bar. If only workdays are present, calculate the end date from workdays.

For more parsing examples, see `references/input-schema.md`.

## Workbook Layout

Default to the 0528-style streamlined layout unless the user asks for batch/phase structure:

```text
A: optional Model or Description
Next: Description, Kivisense, Brands, optional Status
Then: China-business-day timeline
```

Header rows:

- Row 1-4: logo, project title, and table headers must fit within these four rows.
- Row 1-2: Kivisense/弥知科技 logo anchored within `A1:A2`, left padded and vertically centered like the reference screenshot. Project title merged across `B1:D2`, centered, and kept within 20 Chinese characters when possible. Widen `B:D` enough to hold the title cleanly.
- Row 3: left headers and merged month headers.
- Row 4: day numbers for each China business-day date.
- Row 5 onward: task rows.
- When `Model` is enabled, group by first-seen model order, sort each model block by start date, and merge each model block vertically.

Freeze panes at the first date/data cell so owner columns remain visible.

## Timeline Rules

Build date columns from the earliest task start to the latest computed task end, plus a small right-side buffer of about 5 China business days. Use only China business days unless the user explicitly asks for calendar days.

For each task:

1. Parse start date.
2. Use explicit end date when provided; otherwise compute the end date by counting the start date as workday 1.
3. Fill every China business-day column between start and end inclusive.
4. Insert the small star marker from `assets/gantt-end-star.png` in the final workday cell of each task's Gantt bar.
5. Leave non-active timeline cells white.
6. Do not mark the `Status` column complete unless the user explicitly says complete/done/已完成.

If the start date is not a China business day, move the Gantt bar start to the next China business day and keep the requested workday duration.

## Visual Style

Match the Kivisense timeline style closely:

- Font: Microsoft YaHei or an available Chinese-compatible sans-serif.
- Header/task font size: 8-11 pt.
- Task description column: wide, left aligned, vertically centered.
- Owner/status columns: narrow, centered.
- Date columns: narrow, centered day numbers.
- Gantt/date cells should look close to square, not stretched rectangles. Keep timeline date column width and task row height visually balanced.
- Borders: thin light grid across the active range.
- Month headers: merged cells with centered month names.
- Insert the Kivisense/弥知科技 logo from `assets/kivisense-logo.png` at the top-left of the sheet.
- Insert the Gantt end star from `assets/gantt-end-star.png` at the final colored cell for every task. Example: a 10-workday task gets color across 10 China business-day cells and the star anchored on the 10th cell. Make the star prominent, nearly filling the square-like date cell, and visually centered inside the cell.
- Do not create a hidden pre-description column. `A3:A4` must be the `Description` header, and task names start in column `A`.
- Use checkmark `√` for owners and completed status.

Preferred colors:

- General/planning tasks: `A9D08E` green.
- Compliance/privacy: `F8CBAD` peach.
- Development/integration: use non-yellow colors such as `7030A0` purple or `5B9BD5` blue.
- Design/content tasks: `DC888E` rose or `757171` gray.
- Launch milestone: `00B050` green.
- Never use yellow for Gantt bars. Avoid `FFC000`, `FFFF00`, `FFFF99`, and other yellow-like fills even if the user asks casually for yellow; choose the nearest non-yellow alternative.
- Within any adjacent four task rows, do not repeat the same Gantt bar color. If category-based color choices would repeat, rotate to another approved non-yellow color.
- If the user gives no color/category, use the first approved color that does not conflict with the previous three rows.

## Creation Workflow

1. Load this skill and the `spreadsheets` skill.
2. Parse the user's list into structured tasks.
3. If the request is straightforward, use `scripts/build_timeline.py` to generate the workbook quickly.
4. If the user needs advanced formulas, dashboards, charts, or Google Sheets import, follow the `spreadsheets` skill workflow and recreate the same layout with the spreadsheet authoring tooling.
5. Visually verify the output by rendering or inspecting the first sheet. Check that:
   - Month headers align with dates.
   - Owner checkmarks match the provided owners.
   - Status is incomplete by default.
   - Gantt bars cover exactly the requested business-day duration.
   - Important text is not clipped.

## Script Usage

Use the bundled script for deterministic generation from JSON:

```bash
python3 scripts/build_timeline.py input.json output.xlsx
```

Input JSON shape:

```json
{
  "project_name": "AR Campaign",
  "include_model": false,
  "include_status": true,
  "literal_incomplete_status": false,
  "tasks": [
    {
      "model": "需求",
      "name": "Project requirement",
      "owners": ["Kivisense", "Brands"],
      "start": "2026-06-01",
      "end": "2026-06-05",
      "workdays": 5,
      "category": "planning"
    }
  ]
}
```

If building manually instead of using the script, preserve the same layout and rules above.
