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

## Owner Parsing

- `kivisense`, `Kivisense`, `KV`, `我方` -> `Kivisense`
- `brand`, `brands`, `client`, `客户`, `品牌方` -> `Brands`
- If both owner groups appear, check both columns.
- If only one appears, check only that owner column.

## Duration Parsing

- `10天`, `10 day`, `10 days`, `10 workdays`, `10个工作日` -> `10`
- Treat durations as business days by default.
- Count the start date as day 1 if it is a weekday.
- If start date is Saturday or Sunday, move start to the next Monday.

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
