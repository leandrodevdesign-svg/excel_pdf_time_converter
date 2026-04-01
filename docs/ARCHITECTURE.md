# Architecture Guide

## Overview

This project is a static frontend application. It does not require a backend, database, or build step. All processing happens in the browser.

The app has three responsibilities:

1. Read a spreadsheet from the user's machine.
2. Transform raw time-tracking rows into a normalized in-memory dataset.
3. Render metrics, charts, and an exportable detail report.

## File Responsibilities

### [index.html](/Users/1950LabsDesign/.gemini/antigravity/scratch/boceto/excel_pdf_time_converter/index.html)

- Defines the page structure and semantic sections.
- Declares the UI placeholders used by the JavaScript renderer.
- Loads third-party browser libraries and the local application code.

### [styles.css](/Users/1950LabsDesign/.gemini/antigravity/scratch/boceto/excel_pdf_time_converter/assets/css/styles.css)

- Holds the dashboard layout and visual system.
- Includes responsive rules for tablet and mobile widths.
- Includes print and PDF-specific overrides.

### [app.js](/Users/1950LabsDesign/.gemini/antigravity/scratch/boceto/excel_pdf_time_converter/assets/js/app.js)

- Handles file upload events.
- Parses and normalizes spreadsheet rows.
- Computes summary metrics.
- Renders bar charts, stat cards, and the detail table.
- Exports the current dashboard state to PDF.

## Data Flow

1. A user selects an Excel file.
2. `handleFileUpload()` reads the first worksheet as a two-dimensional array.
3. Headers are normalized so alternate field names still map to known keys.
4. `parseRow()` transforms each worksheet row into a consistent dashboard record.
5. Invalid or empty rows are discarded.
6. `renderDashboard()` computes KPIs and drives the UI rendering functions.

## Parsing Rules

- The app accepts common header variants in both English and Spanish.
- If the spreadsheet already includes total hours, that value is preferred.
- If total hours are missing, the app calculates hours from `In` and `Out`.
- Billable hours are clamped so they cannot exceed total hours.
- Empty text fields fall back to safe placeholder values.

## Security and Maintainability Choices

- The table and chart rows are rendered with DOM APIs instead of unsafe HTML interpolation.
- Spreadsheet values are treated as text when inserted into the UI.
- Date keys are built from local date parts instead of `toISOString()` to avoid timezone drift.
- Application state is centralized in a small in-memory object.
- Reusable formatting and parsing helpers keep business logic readable.

## PDF Export Strategy

PDF export uses `html2pdf.js`, which captures the already rendered dashboard. The CSS file contains a dedicated `.pdf-mode` section that:

- switches the dashboard to a white print-oriented theme,
- reduces layout complexity,
- prevents awkward page breaks in cards and rows,
- tightens spacing for tables on A4 pages.

## Recommended Next Improvements

- Add drag-and-drop upload support.
- Add filters by date range, customer, or project.
- Add automated tests around parsing helpers.
- Replace CDN dependencies with pinned local assets if offline usage is required.
