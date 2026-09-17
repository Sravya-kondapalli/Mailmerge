# Certificate Generation Agent — Complete User Walkthrough

Welcome to the **Certificate Generation Agent**! This automated tool allows you to create hundreds or thousands of high-quality, personalized certificates in seconds using Microsoft Word certificate templates and Excel/CSV recipient lists.

---

## Quick Start (Double-Click Launcher)
Simply double-click:
```
run_agent.bat
```
This automatically verifies Python dependencies and opens the desktop application!

---

## Table of Contents
1. [A. First-Time Setup and Installation](#a-first-time-setup-and-installation)
2. [B. How to Prepare Your Excel / CSV File](#b-how-to-prepare-your-excel--csv-file)
3. [C. What Column Names are Required](#c-what-column-names-are-required)
4. [D. Where to Put the Certificate Template & How Placeholders Work](#d-where-to-put-the-certificate-template--how-placeholders-work)
5. [E. How to Run the Agent (GUI and Command Line)](#e-how-to-run-the-agent)
6. [F. Generating Large Batches (10, 100, 500+ Certificates)](#f-generating-large-batches)
7. [G. Where Certificates Are Saved & Filename Rules](#g-where-certificates-are-saved)
8. [H. How to Verify the Generated Batch](#h-how-to-verify-the-generated-batch)
9. [I. Troubleshooting and Error Handling](#i-troubleshooting-and-error-handling)

---

## A. First-Time Setup and Installation

### Prerequisites
- **Operating System:** Windows 10 or Windows 11 (64-bit).
- **Python:** Python 3.10 or higher.
  - Download from: [https://www.python.org/downloads/](https://www.python.org/downloads/)
  - **Important:** During Python setup, check the box **"Add Python to PATH"**.
- **Microsoft Word (Optional / Supported):**
  - The agent generates standard Word `.docx` certificates natively without needing Word open.
  - If Microsoft Word is installed, you can also export high-resolution PDF certificates.

### Automatic Setup (Easiest)
1. Open the `agent1` folder on your Windows laptop.
2. Double-click **`run_agent.bat`**.
3. The script will automatically detect Python, install any missing libraries (`python-docx`, `openpyxl`), and launch the graphical interface.

### Manual Setup (Command Prompt or PowerShell)
If you prefer running manual commands:
```powershell
cd C:\Users\thriv\Desktop\agent1
python -m pip install -r requirements.txt
```

---

## B. How to Prepare Your Excel / CSV File

You can supply your recipient list as either a Microsoft Excel workbook (`.xlsx`) or a comma-separated values file (`.csv`).

### Rules for the Spreadsheet:
1. **Header Row:** The very first row (Row 1) must contain the column headers (e.g. `Name`, `Course`, `Date`, etc.).
2. **One Recipient per Row:** Each subsequent row represents one individual certificate.
3. **Clean Data:** Avoid leaving entirely empty rows between recipients.
4. **Special Characters:** The agent safely handles special characters in names (accents, hyphens, etc.) and sanitizes forbidden Windows characters like `/ \ : * ? " < > |`.

---

## C. What Column Names are Required

The agent automatically matches column names in your spreadsheet to placeholders in your Word template.

### Standard Recommended Columns:
| Column Header | Description | Example Value |
|---|---|---|
| `Name` | Full name of the recipient (used for filename) | `Alice Johnson` |
| `Course` | Name of the program, workshop, or course | `Advanced Python & AI Mastery` |
| `Date` | Date of completion or issue | `September 16, 2026` |
| `Certificate_ID` | Unique identification or verification code | `CERT-2026-001` |
| `Grade` | Honors, score, or completion status | `Distinction` |
| `Organization` | Issuing organization or institution | `DeepMind Academy` |

> **Custom Columns:** You are free to add any additional columns (such as `Hours`, `Instructor`, `Roll_Number`). Any column you create can be inserted directly into your Word template!

---

## D. Where to Put the Certificate Template & How Placeholders Work

### Where to Store Templates
Place your Word template in the `templates/` folder:
```
agent1/
  └── templates/
        └── certificate_template.docx
```

### How Mail-Merge Fields Work in the Word Template
In your Word document, insert placeholders using either the traditional mail-merge chevron style `«FieldName»` or curly braces `{FieldName}`.

#### Supported Placeholder Formats:
- Chevron style: `«Name»`, `«Course»`, `«Date»`, `«Certificate_ID»`, `«Grade»`, `«Organization»`
- Curly brace style: `{Name}`, `{Course}`, `{Date}`, `{{Certificate_ID}}`
- Bracket style: `<Name>`, `<Course>`

#### How to Type Chevrons in Word:
- Option 1: Type `{Name}` directly into Word.
- Option 2: Copy and paste `«` and `»` from here: `«Name»`.
- Option 3: Press `Alt + 0171` for `«` and `Alt + 0187` for `»`.

#### Visual Styling:
- Whatever font, size, color, bold, or italic styling you apply to the placeholder in Word will be **strictly preserved** in the generated certificates.
- Placeholders work inside regular paragraphs, text boxes, headers, footers, and tables.

---

## E. How to Run the Agent

### Method 1: Graphical Desktop App (Recommended for Non-Programmers)
1. Double-click `run_agent.bat` (or run `python gui.py`).
2. The user interface will open:
   - **Recipient File:** Browse and select your `recipients.xlsx` or `recipients.csv`.
   - **Template:** Browse and select your `certificate_template.docx`.
   - **Output Directory:** Keep `output_certificates` or select a custom folder.
3. Click **"Check Fields & Preview"**:
   - The agent will scan your template and verify that all columns exist in your spreadsheet.
4. Click **"Generate Certificates"**:
   - The progress bar will fill in real-time as each certificate is crafted.
5. Click **"Open Output Folder"** to view your personalized documents!

### Method 2: Command Line (Fast & Automation-Friendly)
Run with default files:
```powershell
python mail_merge.py
```
Or specify custom paths:
```powershell
python mail_merge.py --recipients "recipients.xlsx" --template "templates/certificate_template.docx" --output "output_certificates"
```
Optional flags:
- `--overwrite`: Overwrites existing certificates instead of creating `Name (1).docx`.
- `--export-pdf`: Also exports PDF copies using Microsoft Word.

---

## F. Generating Large Batches (10, 100, 500+ Certificates)

The agent is optimized for high-volume batch generation:
- **Speed:** Generates ~25 to 50 certificates per second in native Word format. A batch of 500 certificates finishes in approximately 10 to 20 seconds.
- **Safety:** Each recipient document is isolated in memory so data from one person never leaks to another.
- **Progress Tracking:** The console and GUI report the exact progress: `[150/500] Successfully generated: John Doe.docx`.

---

## G. Where Certificates Are Saved & Filename Rules

Generated certificates are stored in:
```
agent1/
  └── output_certificates/
        ├── Alice Johnson.docx
        ├── Bob Smith.docx
        ├── Emily Davis.docx
        └── Emily Davis (1).docx
```

### Filename Rules & Safety:
1. **Sensible Naming:** Filenames are automatically created from the recipient's name (e.g. `Alice Johnson.docx`).
2. **Invalid Characters:** Windows does not allow `\ / : * ? " < > |` in filenames. The agent automatically cleans these characters so files save without errors (e.g., `Dr. Victor / Frankenstein` becomes `Dr. Victor Frankenstein.docx`).
3. **Duplicate Names:** If multiple recipients share the same name, the agent automatically appends `(1)`, `(2)`, etc. No previous certificate is accidentally overwritten.

---

## H. How to Verify the Generated Batch

To guarantee quality and ensure every recipient received their certificate:

### Using the GUI:
Click the **"Verify Batch"** button in the application.

### Using the Command Line:
```powershell
python verify_batch.py
```

### What Verification Checks:
1. **Total Count Match:** Checks that the number of generated certificates matches the number of valid rows in the Excel file.
2. **File Health:** Checks that every generated certificate file is greater than 0 bytes and not corrupted.
3. **Content Integrity:** Opens the Word documents and checks that the recipient's name is inside the document.
4. **Placeholder Check:** Confirms that no unmerged placeholder tags (like `«Name»`) remain in the final output.

---

## I. Troubleshooting and Error Handling

### 1. "Missing columns" Warning
- **Cause:** The Word template has a placeholder (e.g. `«Score»`) that does not exist in your Excel header row.
- **Fix:** Check spelling in Row 1 of Excel. Placeholders are matched case-insensitively.

### 2. "Permission Denied: output_certificates/..."
- **Cause:** A certificate file is currently open in Microsoft Word.
- **Fix:** Close Microsoft Word and run the agent again.

### 3. Log Files
All detailed activity and error messages are written automatically to the `logs/` directory:
- `logs/agent.log`: Complete log of every batch run, timestamps, and files generated.
- `logs/error.log`: Specific record of any row that failed or encountered an issue.

---

## 5-Step Summary for Daily Use

1. **Open** `recipients.xlsx` and add your recipient rows.
2. **Double-click** `run_agent.bat`.
3. **Click** `Check Fields & Preview` to confirm matching columns.
4. **Click** `Generate Certificates`.
5. **Click** `Open Output Folder` to inspect and print your certificates!
