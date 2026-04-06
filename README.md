# excel_pdf_time_converter

`excel_pdf_time_converter` is a static dashboard suite that reads Excel timesheets, transforms raw rows into management-friendly metrics, exports reports to PDF, and now converts bulk invoice PDFs into a structured CSV.

## Project Goals

- Translate the original Spanish prototype to English.
- Keep the analytics dashboard fully client-side.
- Add a secure Vercel serverless endpoint for Gemini-powered invoice extraction.
- Improve maintainability by separating markup, styles, and behavior.
- Document the code and expected spreadsheet format clearly.

## Project Structure

```text
excel_pdf_time_converter/
├── api/
│   └── invoices-to-excel.js
├── assets/
│   ├── css/
│   │   ├── home.css
│   │   ├── invoices-to-excel.css
│   │   └── styles.css
│   └── js/
│       ├── app.js
│       └── invoices-to-excel.js
├── docs/
│   ├── ARCHITECTURE.md
│   └── EXCEL_INPUT.md
├── invoices-to-excel.html
├── index.html
└── .gitignore
```

## How It Works

1. The dashboard pages accept `.xlsx` or `.xls` files and process them entirely in the browser.
2. The invoices page accepts a bulk set of PDF invoices and extracts machine-readable text through `pdf.js`.
3. The browser builds an intermediate CSV with `Document Name` and `Content`.
4. The `/api/invoices-to-excel` Vercel function sends that CSV to Gemini using a server-side environment variable.
5. Gemini returns structured invoice rows, which the UI previews and exports as a final CSV with `Document Name`, `User Name`, and `Hours Billed`.

## Run Locally

The dashboard-only pages can be opened directly in a browser or served with any static server. The Gemini-powered invoice flow is different: it requires a backend runtime for `/api/invoices-to-excel`, so a plain `file://` preview or a static server alone is not enough for end-to-end testing.

Use a static server only for visual frontend review, for example:

```bash
python3 -m http.server 8000
```

Then open `http://localhost:8000` from the repository root.

To test the invoice conversion flow end to end, run it through Vercel locally or deploy it to Vercel so the API route is available.

To enable Gemini processing on Vercel, add the following environment variable:

```bash
GEMINI_API_KEY=your_key_here
```

Optional:

```bash
GEMINI_MODEL=gemini-2.0-flash
```

## Main Dependencies

- `xlsx`: reads Excel workbooks in the browser.
- `html2pdf.js`: captures the rendered dashboard and generates a PDF file.
- `pdf.js`: extracts text from invoice PDFs in the browser.
- Gemini API: structures invoice content into a downloadable CSV through the Vercel serverless function.

Both dependencies are loaded from public CDNs in [index.html](/Users/1950LabsDesign/.gemini/antigravity/scratch/boceto/excel_pdf_time_converter/index.html).

## Documentation

- [Architecture guide](/Users/1950LabsDesign/.gemini/antigravity/scratch/boceto/excel_pdf_time_converter/docs/ARCHITECTURE.md)
- [Excel input reference](/Users/1950LabsDesign/.gemini/antigravity/scratch/boceto/excel_pdf_time_converter/docs/EXCEL_INPUT.md)
