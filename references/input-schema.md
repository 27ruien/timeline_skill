# Timeline Maker Input Schema

## Minimal Rows

These all parse to the same canonical fields:

```text
事项名称, kivisense, 2026-06-01, 5天
事项名称 / Kivisense / 2026-06-01 / 5 workdays
事项名称，Kivisense，开始日期 2026-06-01，周期 5天
```

Canonical form:

```json
{
  "name": "事项名称",
  "owners": ["Kivisense"],
  "start": "2026-06-01",
  "workdays": 5,
  "status": "incomplete"
}
```

## Stage / Model Rows

Stage is a first-class field in the workbench. In JSON input, pass the stage text as `model` or `stage`; the generator writes that value exactly into the `Model` / `工作内容` column. `include_model` is kept for older inputs, but current exports always include the stage column.

For raw text input, rows are parsed as:

```text
Model, 事项名称, 相关方, 开始日期, 工作日天数或结束日期
```

Example:

```text
1. 需求, Project requirement, Kivisense, 2026-06-01, 5天
14. 需求, Scope addendum, brand, 2026-06-18, 4天
2. 设计, Creative Proposal, Kivisense, brand, 2026-06-08, 10天
```

Output rules:

- Preserve the input `model` text exactly. Do not translate `Requirement` into `需求`, and do not translate `需求` into `Requirement`.
- If `model` is empty, use `未分类` for Chinese output or `Uncategorized` for English output.
- Group the same Model together even if the rows were entered non-adjacently.
- Preserve first-seen Model group order.
- Sort items inside each Model group by start date.
- Merge the Model cells vertically for each group.
- Keep Description as separate rows for each task.

If `include_status` is false, omit the Status column entirely.

## Owner Parsing

- `kivisense`, `Kivisense`, `KV`, `我方` -> `Kivisense`
- `brand`, `brands`, `client`, `客户`, `品牌方` -> `Brands`
- If both owner groups appear, check both columns.
- If only one appears, check only that owner column.

## Duration Parsing

- `10天`, `10 day`, `10 days`, `10 workdays`, `10个工作日` -> `10`
- Treat durations as business days by default.
- Count the start date as day 1 if it is a China business day.
- China business days exclude Saturdays, Sundays, and Chinese public holidays, but include official adjusted working weekends.
- If start date is not a China business day, move start to the next China business day.
- If an explicit end date is provided, use it for the Gantt range and derive workday count.
- If no end date is provided, derive the end date from workday duration.

## Status Parsing

Default status is incomplete.

- no status, `未完成`, `pending`, `todo`, `not done` -> incomplete
- `完成`, `已完成`, `done`, `complete`, `√` -> complete

For the default output style, incomplete usually displays as a blank `Status` cell. Only show the literal `未完成` when the user asks for literal status labels.

## Categories

Use category only for color selection; it is optional.

- `planning`, `kickoff`, `requirement`, `proposal`, `quotation` -> green `A9D08E`
- `privacy`, `compliance`, `legal` -> peach `F8CBAD`
- `development`, `integration`, `engineering`, `frontend`, `backend` -> purple `7030A0` or blue `5B9BD5`; do not use yellow.
- `design`, `content`, `creative`, `video`, `mural` -> rose `DC888E`
- `special`, `easter egg`, `coupon`, `middle office` -> purple `7030A0`
- `uat`, `testing`, `onsite`, `setup` -> green `A9D08E`
- `launch`, `online`, `go live` -> launch green `00B050`

If no category is present, infer from the task name; otherwise use green.

## Color Constraints

Never use yellow Gantt fills. Avoid `FFC000`, `FFFF00`, `FFFF99`, and similar yellow colors.

Within any four adjacent task rows, all Gantt fill colors should be different. When a category would repeat a recent color, rotate through the approved palette:

```text
A9D08E, F8CBAD, 7030A0, DC888E, 757171, 5B9BD5, 00B050
```
