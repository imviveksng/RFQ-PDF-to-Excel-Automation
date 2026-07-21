# Quote PDF Automation

A Python utility that reads quote and enquiry PDFs, extracts key project details, and appends them to an Excel enquiry master. After a PDF is processed successfully, it is renamed with its RFQ number and moved to the processed folder.

## Features

- Extracts enquiry details such as RFQ number, customer, location, consultant, dates, sales engineer, designer, reviewer, revision, complexity, and design weight.
- Cleans text and normalizes supported dates before saving the data.
- Adds each processed PDF as a new row in the configured Excel worksheet.
- Organizes completed source files by renaming and moving them to `Processed PDFs`.
- Processes every `.pdf` file placed in the input folder in one run.

## Requirements

- Python 3.10 or later
- Microsoft Excel workbook at the configured output location
- PDF files with selectable text (image-only/scanned PDFs require OCR before they can be extracted reliably)

Python packages are pinned in [requirements.txt](requirements.txt):

- `pdfplumber` for PDF text extraction
- `openpyxl` for updating the Excel workbook

## Folder Structure

The application is configured to use the following layout:

```text
Quotes PDFs/
+-- Excel/
|   `-- Enquiry Master.xlsx
+-- New PDFs/                    # Place unprocessed quote PDFs here
+-- Processed PDFs/              # Processed PDFs are moved here
`-- Python Automation/
    +-- main.py                  # Application entry point
    +-- config.py                # Folder, workbook, and worksheet settings
    +-- pdf_reader.py            # PDF extraction rules
    +-- cleaner.py               # Text and date cleanup helpers
    +-- excel_writer.py          # Excel row writer
    +-- mappings.py              # Data-field to Excel-column mapping
    +-- utils.py                 # Shared extraction and file-move helpers
    +-- requirements.txt
    `-- README.md
```

## Installation

1. Clone or download this project into the `Quotes PDFs/Python Automation` folder.

2. Create and activate a virtual environment:

   ```powershell
   python -m venv venv
   .\venv\Scripts\Activate.ps1
   ```

3. Install the dependencies:

   ```powershell
   pip install -r requirements.txt
   ```

4. Create the folders shown above if they do not already exist.

5. Place `Enquiry Master.xlsx` in the `Excel` folder. Ensure it contains the worksheet specified by `SHEET_NAME` in [config.py](config.py).

6. Review [config.py](config.py) and update `ROOT_FOLDER`, `EXCEL_FILE`, or `SHEET_NAME` if your folders or workbook use different names.

## Usage

1. Copy one or more quote/enquiry PDFs into `New PDFs`.

2. Run the application from the `Python Automation` folder:

   ```powershell
   python main.py
   ```

3. For every PDF, the application will:

   1. Extract and clean the configured fields.
   2. Append the values to the next available row of the Excel worksheet.
   3. Rename the source PDF to `<RFQ_NUMBER>.pdf`.
   4. Move it to `Processed PDFs`.

4. Open `Excel/Enquiry Master.xlsx` to review the newly added rows.

Example console output:

```text
Processing : quotation.pdf
Excel Path: E:\Quotes PDFs\Excel\Enquiry Master.xlsx
[OK] Saved to Row 25
[OK] PDF moved.

Finished.
```

## Excel Mapping

The workbook columns are controlled in [mappings.py](mappings.py). The primary fields include:

| Excel column | Field |
| --- | --- |
| A | Month |
| B | Inquiry / RFQ number |
| C | ABR sales engineer |
| E | Client |
| F | Building |
| G | Consultant |
| H | Location |
| K-N | Key dates, designer, and checker |
| O-V | Weight, revision, complexity, status, and remarks |

## Screenshots

Add screenshots to this section to help new users verify the workflow. Recommended images:

1. **Input folder** - PDFs placed in `New PDFs` before processing.
2. **Terminal output** - a successful `python main.py` run.
3. **Excel result** - the new row in `Enquiry Master.xlsx`.
4. **Processed folder** - the renamed PDF after completion.

Once image files are added to the repository (for example, under `docs/screenshots/`), reference them here:

```markdown
![Excel output](docs/screenshots/excel-output.png)
```

## Configuration Notes

- Extraction patterns live in [pdf_reader.py](pdf_reader.py). Adjust them when a PDF template uses different labels or formatting.
- `ABR_SALES` is extracted from the text following `SALES ENGG.`. For multi-line PDF layouts, use a line-based pattern such as `r"SALES ENGG\.\s*([^\r\n]+)"`.
- The script appends to `sheet.max_row + 1`; keep the target worksheet available and avoid protected workbooks.
- Processed PDFs are renamed using the extracted RFQ number. Verify that each source PDF contains a valid RFQ number before running the script.

## Troubleshooting

| Issue | Suggested check |
| --- | --- |
| `No PDF Found.` | Add PDF files directly to `New PDFs`. |
| Workbook or worksheet error | Confirm the Excel file path and sheet name in `config.py`. |
| A field is blank in Excel | Inspect the extracted PDF text and update the matching pattern in `pdf_reader.py`. |
| PDF cannot be read | Confirm it is a valid, text-based PDF; scanned PDFs may require OCR. |

## License

This project is intended for internal use. Add a license file if it will be shared externally.
