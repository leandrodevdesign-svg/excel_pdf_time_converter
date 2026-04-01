# Excel Input Reference

## Supported File Types

- `.xlsx`
- `.xls`

The app reads only the first worksheet in the workbook.

## Preferred Columns

The dashboard works best when the spreadsheet contains these columns:

- `Date`
- `In`
- `Out`
- `Time`
- `Billable`
- `Customer`
- `Project`
- `Activity`
- `Description`
- `Username`

## Accepted Variants

The parser also accepts common Spanish or alternate naming patterns, including:

- `Fecha`, `Day`
- `Entrada`, `Start`, `Start Time`
- `Salida`, `End`, `End Time`
- `Horas`, `Hours`, `Total`, `Total Hours`
- `Facturable`, `Billable Hours`
- `Cliente`
- `Proyecto`
- `Actividad`
- `Descripcion`
- `Usuario`, `User`

The matching process is resilient because headers are normalized before parsing.

## Record Validation Rules

A row is considered valid only if:

- it contains a usable date, and
- it provides a positive duration, either directly or inferred from start/end times.

Rows that do not meet those conditions are ignored.

## Notes on Hours

- If `Time` or an equivalent hours column is present and greater than zero, that value is used.
- Otherwise, the application calculates the duration from `In` and `Out`.
- Negative or invalid durations are rejected.

## Billable Logic

- `Billable` is interpreted as a numeric value in hours.
- If the billable value is missing, the app assumes `0`.
- If billable hours exceed total hours, the value is capped at total hours.

## Example Workflow

1. Export a timesheet from your source system.
2. Ensure the relevant columns are present in the first worksheet.
3. Upload the file in the dashboard.
4. Review KPIs, charts, and record details.
5. Export the result as PDF if needed.
