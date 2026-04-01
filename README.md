# excel_pdf_time_converter

`excel_pdf_time_converter` is a single-page dashboard that reads an Excel timesheet, transforms the raw rows into management-friendly metrics, and exports the resulting report to PDF.

## Project Goals

- Translate the original Spanish prototype to English.
- Keep the app fully client-side so it can run without a backend.
- Improve maintainability by separating markup, styles, and behavior.
- Document the code and expected spreadsheet format clearly.

## Project Structure

```text
excel_pdf_time_converter/
├── assets/
│   ├── css/
│   │   └── styles.css
│   └── js/
│       └── app.js
├── docs/
│   ├── ARCHITECTURE.md
│   └── EXCEL_INPUT.md
├── index.html
└── .gitignore
```

## How It Works

1. The user uploads an `.xlsx` or `.xls` file.
2. The browser reads the first worksheet through `SheetJS`.
3. The app normalizes headers, parses date/time values, and calculates hours.
4. The dashboard renders KPIs, charts, quick insights, and a detailed table.
5. The current dashboard view can be exported as a PDF through `html2pdf.js`.

## Run Locally

Because this is a static project, you can open [index.html](/Users/1950LabsDesign/.gemini/antigravity/scratch/boceto/excel_pdf_time_converter/index.html) directly in a browser. For a smoother local-development workflow, serve the folder with any static server.

Example:

```bash
python3 -m http.server 8000
```

Then open `http://localhost:8000` from the repository root.

## Main Dependencies

- `xlsx`: reads Excel workbooks in the browser.
- `html2pdf.js`: captures the rendered dashboard and generates a PDF file.

Both dependencies are loaded from public CDNs in [index.html](/Users/1950LabsDesign/.gemini/antigravity/scratch/boceto/excel_pdf_time_converter/index.html).

## Documentation

- [Architecture guide](/Users/1950LabsDesign/.gemini/antigravity/scratch/boceto/excel_pdf_time_converter/docs/ARCHITECTURE.md)
- [Excel input reference](/Users/1950LabsDesign/.gemini/antigravity/scratch/boceto/excel_pdf_time_converter/docs/EXCEL_INPUT.md)
