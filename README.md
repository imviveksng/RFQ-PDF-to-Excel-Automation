# 📄 Quote PDF Automation — Gmail → Excel Enquiry Pipeline

![Python](https://img.shields.io/badge/Python-3.10+-3776AB?style=for-the-badge&logo=python&logoColor=white)
![Gmail API](https://img.shields.io/badge/Gmail%20API-EA4335?style=for-the-badge&logo=gmail&logoColor=white)
![Excel](https://img.shields.io/badge/Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![pdfplumber](https://img.shields.io/badge/pdfplumber-PDF%20Extraction-CC2927?style=for-the-badge&logo=adobeacrobatreader&logoColor=white)
![OAuth2](https://img.shields.io/badge/OAuth%202.0-Read--Only-4285F4?style=for-the-badge&logo=googlecloud&logoColor=white)

> A Python automation tool that pulls new EES quotation PDFs straight from Gmail, extracts key enquiry details, and logs them into a shared Excel enquiry master — with zero manual data entry.

---

## 📌 Table of Contents

- [Project Overview](#-project-overview)
- [Business Problem](#-business-problem)
- [Technology Stack](#-technology-stack)
- [Architecture](#-architecture)
- [Repository Structure](#-repository-structure)
- [Features](#-features)
- [Installation](#-installation)
- [Usage](#-usage)
- [Excel Mapping](#-excel-mapping)
- [Configuration Notes](#-configuration-notes)
- [Troubleshooting](#-troubleshooting)
- [Screenshots](#-screenshots)
- [Future Enhancements](#-future-enhancements)

---

## 📖 Project Overview

This project automates a task that would otherwise mean manually opening every quotation email, reading the PDF, and typing the details into an Excel tracker. Instead, it:

- Connects to Gmail (read-only) and pulls new EES quotation attachments
- Extracts structured fields (RFQ number, customer, dates, engineers, revision, etc.) from each PDF
- Appends a clean row to the enquiry master workbook
- Renames the source PDF to its RFQ number and archives it

The result is a self-running enquiry log that stays in sync with incoming quotation emails, with no copy-paste required.

---

## 🎯 Business Problem

Quotation tracking teams typically deal with:

- ❌ Manual, repetitive data entry from PDF quotations into Excel
- ❌ Inconsistent formatting and typos from human transcription
- ❌ Lost time hunting through inboxes for the latest RFQ PDFs
- ❌ No consistent naming/archiving convention for processed quotes

This tool solves that by turning "check inbox → open PDF → type into Excel → file the PDF" into a single command.

---

## 🛠️ Technology Stack

| Layer            | Technology                          | Purpose                              |
|------------------|--------------------------------------|---------------------------------------|
| Email Source     | Gmail API (OAuth 2.0, read-only)     | Fetch quotation emails & attachments   |
| PDF Extraction   | `pdfplumber`                          | Pull text from quotation PDFs          |
| Data Cleaning    | Custom cleaner module                | Normalize text & dates                 |
| Data Storage     | Excel (`openpyxl`)                    | Central enquiry master workbook        |
| File Management  | Python `os`/`shutil` helpers          | Rename & move processed PDFs           |
| Auth             | `google-auth`, `google-auth-oauthlib` | Gmail authentication & token refresh    |

---

## 🏗️ Architecture

```
┌──────────────────────────────┐
│           Gmail Inbox        │
│   Label: Quotation Emails    │
└───────────────┬───────────────┘
                │
                ▼
┌──────────────────────────────┐
│      gmail_downloader.py      │
│  Filters *ees* PDF attachments│
│      → saves to New PDFs      │
└───────────────┬───────────────┘
                │
                ▼
┌──────────────────────────────┐
│         pdf_reader.py         │
│   Extracts RFQ, dates, names, │
│   consultant, revision, etc.  │
└───────────────┬───────────────┘
                │
                ▼
┌──────────────────────────────┐
│          cleaner.py           │
│  Normalizes text & date fields│
└───────────────┬───────────────┘
                │
                ▼
┌──────────────────────────────┐
│       excel_writer.py         │
│  Appends row → Enquiry Master │
└───────────────┬───────────────┘
                │
                ▼
┌──────────────────────────────┐
│         utils.py              │
│  Renames PDF to RFQ number →  │
│      moves to Processed PDFs  │
└────────────────────────────────┘
```

---

## 📁 Repository Structure

```
Quotes PDFs/
├── Excel/
│   └── Enquiry Master.xlsx
├── New PDFs/                    # Drop unprocessed quote PDFs here
├── Processed PDFs/              # Processed PDFs land here, renamed
└── Python Automation/
    ├── main.py                  # Entry point
    ├── config.py                # Folder, workbook, worksheet settings
    ├── pdf_reader.py             # PDF extraction rules
    ├── cleaner.py                # Text & date cleanup
    ├── excel_writer.py           # Excel row writer
    ├── gmail_auth.py             # OAuth authentication & token refresh
    ├── gmail_downloader.py       # Gmail attachment downloader
    ├── mappings.py               # Field → Excel column mapping
    ├── utils.py                  # Shared extraction/file-move helpers
    ├── requirements.txt
    └── README.md
```

---

## ✅ Features

- 📥 Connects to Gmail via read-only OAuth — never marks messages read or modifies labels
- 🔍 Downloads only PDF attachments with `ees` in the filename (case-insensitive)
- 🧠 Extracts RFQ number, customer, location, consultant, dates, sales engineer, designer, reviewer, revision, complexity, and design weight
- 🧹 Cleans and normalizes text and dates before writing to Excel
- 📊 Appends each processed quote as a new row in the enquiry master
- 🗂️ Renames and archives every completed PDF into `Processed PDFs`
- 🔁 Also processes manually-added PDFs dropped into `New PDFs`

---

## 🚀 Installation

### Prerequisites
- Python 3.10+
- Microsoft Excel workbook at the configured path
- Text-based PDFs (scanned/image-only PDFs need OCR first)
- A Google Cloud OAuth Desktop-app credential with the Gmail API enabled

### Steps

```powershell
# 1. Clone into the Quotes PDFs/Python Automation folder
git clone https://github.com/YOUR-GITHUB-USERNAME/quote-pdf-automation.git

# 2. Create and activate a virtual environment
python -m venv venv
.\venv\Scripts\Activate.ps1

# 3. Install dependencies
pip install -r requirements.txt
pip install google-auth google-auth-oauthlib google-api-python-client
```

Then:
4. Create the folder structure shown above if it doesn't already exist.
5. Place `Enquiry Master.xlsx` inside `Excel/`, containing the worksheet named in `SHEET_NAME` (`config.py`).
6. Review `config.py` and update `ROOT_FOLDER`, `EXCEL_FILE`, or `SHEET_NAME` as needed.
7. Configure Gmail access:
   - Enable the Gmail API and create an OAuth 2.0 **Desktop app** client in Google Cloud.
   - Save the downloaded credential as `credentials.json` in `Python Automation`.
   - Set `LABEL_ID` in `gmail_downloader.py` to your quotation label.
   - On first run, complete the browser consent prompt — a `token.json` will be saved for future runs.

> 🔒 `credentials.json` and `token.json` grant access to your Gmail account — keep them private and never commit them.

---

## ▶️ Usage

```powershell
python main.py
```

The application will:
1. Authenticate with Gmail and fetch messages from the configured label
2. Download matching PDF attachments into `New PDFs`
3. Extract and clean the configured fields from every PDF found
4. Append the values to the next available row in Excel
5. Rename the source PDF to `<RFQ_NUMBER>.pdf`
6. Move it into `Processed PDFs`

**Example console output:**

```text
Downloading New EES PDFs from Gmail...
Found 2 email(s).
Downloaded : EES-quotation.pdf

Processing : EES-quotation.pdf
Excel Path: E:\Quotes PDFs\Excel\Enquiry Master.xlsx
[OK] Saved to Row 25
[OK] PDF moved.

Automation Completed Successfully.
```

Open `Excel/Enquiry Master.xlsx` afterward to review the newly added rows.

---

## 📊 Excel Mapping

Controlled in [`mappings.py`](mappings.py):

| Excel column | Field |
|---|---|
| A | Month |
| B | Inquiry / RFQ number |
| C | ABR sales engineer |
| E | Client |
| F | Building |
| G | Consultant |
| H | Location |
| K–N | Key dates, designer, checker |
| O–V | Weight, revision, complexity, status, remarks |

---

## ⚙️ Configuration Notes

- Gmail settings (`LABEL_ID`, `DOWNLOAD_FOLDER`) live in `gmail_downloader.py`.
- Gmail access is strictly **read-only** — no messages are marked read or relabeled.
- The downloader re-fetches all messages currently in the label and doesn't track history, so archive/remove processed emails from the label to avoid duplicate rows.
- Extraction patterns live in `pdf_reader.py` — adjust these when a PDF template changes.
- `ABR_SALES` extraction uses a pattern following `SALES ENGG.` (line-based, e.g. `r"SALES ENGG\.\s*([^\r\n]+)"`).
- New rows are appended at `sheet.max_row + 1` — keep the workbook unprotected and closed while running.
- Each processed PDF is renamed using its extracted RFQ number — verify PDFs contain a valid RFQ number before running.

---

## 🩺 Troubleshooting

| Issue | Suggested check |
|---|---|
| `No new PDFs found.` | Add matching Gmail attachments to the label, or drop PDFs directly into `New PDFs`. |
| Gmail sign-in opens every run | Delete an invalid `token.json` and re-run to complete OAuth again. |
| Gmail authentication fails | Confirm the Gmail API is enabled, `credentials.json` is a Desktop app credential, and the account can access the configured label. |
| No attachments downloaded | Confirm the email is labeled correctly and the filename is a PDF containing `ees`. |
| Workbook/worksheet error | Confirm the Excel path and sheet name in `config.py`. |
| Blank field in Excel | Inspect the extracted PDF text and adjust the pattern in `pdf_reader.py`. |
| PDF can't be read | Confirm it's a valid, text-based PDF — scanned PDFs may need OCR first. |

---

## 🖼️ Screenshots

1. **Input folder** — PDFs placed in `New PDFs` before processing
   `screenshots/Input Folder.PNG`
2. **Terminal output** — a successful `python main.py` run
   `screenshots/Terminal Output.PNG`
3. **Excel result** — the new row in `Enquiry Master.xlsx`
   `screenshots/output_excel.png`
4. **Processed folder** — the renamed PDF after completion
   `screenshots/Processed Folder.PNG`

---

## 🔮 Future Enhancements

- [ ] Track already-downloaded emails to avoid re-processing on repeated label runs
- [ ] OCR fallback for scanned/image-only PDFs
- [ ] Logging/audit trail of every run (success/failure per PDF)
- [ ] Config-driven extraction templates for multiple quotation formats
- [ ] Optional database backend alongside Excel for larger volumes

---

## 📬 Connect

Feel free to connect or reach out if you have questions about this project!

<!-- Replace these with your actual profiles -->
[![LinkedIn](https://img.shields.io/badge/LinkedIn-0077B5?style=for-the-badge&logo=linkedin&logoColor=white)](https://www.linkedin.com/in/YOUR-LINKEDIN-HANDLE)
[![GitHub](https://img.shields.io/badge/GitHub-181717?style=for-the-badge&logo=github&logoColor=white)](https://github.com/YOUR-GITHUB-USERNAME)

---

*Built as a workflow automation project — turning manual quote logging into a one-command pipeline.*
