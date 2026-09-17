import os
import sys
import re
import csv
import copy
import logging
import argparse
from pathlib import Path
from datetime import datetime

import docx
import openpyxl

# Setup logging
LOG_DIR = Path(__file__).resolve().parent / "logs"
LOG_DIR.mkdir(exist_ok=True)
AGENT_LOG = LOG_DIR / "agent.log"
ERROR_LOG = LOG_DIR / "error.log"

logger = logging.getLogger("CertificateAgent")
logger.setLevel(logging.INFO)

# Formatter
formatter = logging.Formatter("[%(asctime)s] [%(levelname)s] %(message)s", datefmt="%Y-%m-%d %H:%M:%S")

# File handlers
fh_agent = logging.FileHandler(AGENT_LOG, encoding="utf-8")
fh_agent.setLevel(logging.INFO)
fh_agent.setFormatter(formatter)
logger.addHandler(fh_agent)

fh_err = logging.FileHandler(ERROR_LOG, encoding="utf-8")
fh_err.setLevel(logging.WARNING)
fh_err.setFormatter(formatter)
logger.addHandler(fh_err)

# Console handler
ch = logging.StreamHandler(sys.stdout)
ch.setLevel(logging.INFO)
ch.setFormatter(formatter)
logger.addHandler(ch)


def check_word_installed():
    """Checks whether Microsoft Word is installed and COM automation is available."""
    standard_paths = [
        r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
        r"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE",
        r"C:\Program Files\Microsoft Office\Office16\WINWORD.EXE",
        r"C:\Program Files (x86)\Microsoft Office\Office16\WINWORD.EXE",
        r"C:\Program Files\Microsoft Office\Office15\WINWORD.EXE",
        r"C:\Program Files (x86)\Microsoft Office\Office15\WINWORD.EXE"
    ]
    for path in standard_paths:
        if os.path.exists(path):
            return True, path
    
    # Try importing win32com if available
    try:
        import win32com.client
        return True, "COM Automation Available"
    except Exception:
        pass
    return False, "Microsoft Word is not detected on standard paths."


def sanitize_filename(name, max_length=120):
    r"""
    Sanitizes string to be a safe Windows filename:
    Removes forbidden characters: < > : " / \ | ? *
    Trims trailing spaces and periods.
    """
    if not name:
        name = "Recipient"
    name = str(name).strip()
    # Replace forbidden chars with underscore or safe dash
    safe_name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', name)
    # Condense multiple underscores/spaces
    safe_name = re.sub(r'[\s_]+', ' ', safe_name).strip()
    # Windows doesn't allow filenames ending in dot or space
    safe_name = safe_name.rstrip('. ')
    if not safe_name:
        safe_name = "Certificate"
    return safe_name[:max_length]


def get_unique_filename(output_dir, base_name, extension=".docx", overwrite=False):
    """
    Determines output path. If overwrite is False and file exists,
    appends (1), (2), etc. to prevent silent overwriting.
    """
    target_path = output_dir / f"{base_name}{extension}"
    if overwrite or not target_path.exists():
        return target_path
    
    counter = 1
    while True:
        target_path = output_dir / f"{base_name} ({counter}){extension}"
        if not target_path.exists():
            return target_path
        counter += 1


def read_recipients_file(file_path):
    """
    Reads recipient data from either .xlsx or .csv.
    Returns list of dicts: [{"Name": "...", "Course": "..."}, ...]
    """
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"Recipient file not found: {path}")

    ext = path.suffix.lower()
    records = []

    if ext in [".xlsx", ".xlsm"]:
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active
        headers = []
        for col_idx, cell in enumerate(ws[1], start=1):
            val = str(cell.value).strip() if cell.value is not None else f"Column_{col_idx}"
            headers.append(val)
        
        for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if not row or all(v is None or str(v).strip() == "" for v in row):
                continue  # skip blank lines
            row_dict = {}
            for h, v in zip(headers, row):
                row_dict[h] = "" if v is None else str(v).strip()
            row_dict["_row_number"] = row_idx
            records.append(row_dict)

    elif ext in [".csv", ".txt"]:
        # Try multiple encodings
        encodings = ["utf-8-sig", "utf-8", "cp1252", "latin-1"]
        lines = None
        for enc in encodings:
            try:
                with open(path, mode="r", encoding=enc) as f:
                    lines = f.readlines()
                break
            except UnicodeDecodeError:
                continue
        if lines is None:
            raise ValueError(f"Could not decode CSV file {path} with supported encodings.")
        
        reader = csv.DictReader(lines)
        if reader.fieldnames:
            clean_fieldnames = [f.strip() if f else "" for f in reader.fieldnames]
        else:
            clean_fieldnames = []

        for row_idx, row in enumerate(reader, start=2):
            if not row or all(v is None or str(v).strip() == "" for v in row.values()):
                continue
            row_dict = {}
            for k, v in row.items():
                clean_k = k.strip() if k else ""
                row_dict[clean_k] = "" if v is None else str(v).strip()
            row_dict["_row_number"] = row_idx
            records.append(row_dict)
    else:
        raise ValueError(f"Unsupported file format '{ext}'. Please provide .xlsx or .csv file.")

    return records


def extract_template_placeholders(template_path):
    """
    Extracts all merge fields/placeholders from template docx.
    Detects «Field», {Field}, {{Field}}, and <Field>.
    """
    doc = docx.Document(template_path)
    text_content = []

    def get_text_from_paragraphs(paragraphs):
        for p in paragraphs:
            text_content.append(p.text)

    def get_text_from_tables(tables):
        for table in tables:
            for row in table.rows:
                for cell in row.cells:
                    get_text_from_paragraphs(cell.paragraphs)
                    get_text_from_tables(cell.tables)

    get_text_from_paragraphs(doc.paragraphs)
    get_text_from_tables(doc.tables)
    for section in doc.sections:
        get_text_from_paragraphs(section.header.paragraphs)
        get_text_from_tables(section.header.tables)
        get_text_from_paragraphs(section.footer.paragraphs)
        get_text_from_tables(section.footer.tables)

    combined = "\n".join(text_content)
    # Match «Field», {{Field}}, {Field}, <Field>
    patterns = [
        r'«([^»]+)»',
        r'\{\{([^}]+)\}\}',
        r'\{([A-Za-z0-9_]+)\}',
        r'<([A-Za-z0-9_]+)>'
    ]
    placeholders = set()
    for pat in patterns:
        matches = re.findall(pat, combined)
        for m in matches:
            placeholders.add(m.strip())
    return sorted(list(placeholders))


def build_replacement_dict(record):
    """
    Constructs comprehensive replacement dictionary covering
    «Field», {Field}, {{Field}}, and <Field> for all keys in record.
    """
    replacements = {}
    for key, val in record.items():
        if key.startswith("_"):
            continue
        clean_val = str(val) if val is not None else ""
        # Support various placeholder notations
        replacements[f"«{key}»"] = clean_val
        replacements[f"{{{key}}}"] = clean_val
        replacements[f"{{{{{key}}}}}"] = clean_val
        replacements[f"<{key}>"] = clean_val
        # Also support stripped/case-variant keys
        replacements[f"«{key.strip()}»"] = clean_val
        replacements[f"{{{key.strip()}}}"] = clean_val
    return replacements


def replace_in_paragraph(paragraph, replacements):
    """
    Replaces placeholders in paragraph while preserving font style, size, bold, color, etc.
    Correctly handles placeholders split across runs.
    """
    full_text = paragraph.text
    if not any(ph in full_text for ph in replacements):
        return

    # 1. Single run replacement (preserves exact character-level styles)
    for ph, val in replacements.items():
        if ph in full_text:
            for run in paragraph.runs:
                if ph in run.text:
                    run.text = run.text.replace(ph, str(val))

    # 2. Multi-run split replacement
    full_text = paragraph.text
    if any(ph in full_text for ph in replacements):
        new_text = full_text
        for ph, val in replacements.items():
            new_text = new_text.replace(ph, str(val))
        if paragraph.runs:
            paragraph.runs[0].text = new_text
            for r in paragraph.runs[1:]:
                r.text = ""


def replace_in_tables(tables, replacements):
    """Recursively replaces text inside all table cells."""
    for table in tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_in_paragraph(p, replacements)
                replace_in_tables(cell.tables, replacements)


def generate_certificate(template_path, record, output_path):
    """
    Generates a single personalized certificate docx file for a recipient record.
    """
    doc = docx.Document(template_path)
    replacements = build_replacement_dict(record)

    # Document body
    for p in doc.paragraphs:
        replace_in_paragraph(p, replacements)

    # Tables
    replace_in_tables(doc.tables, replacements)

    # Headers & Footers
    for section in doc.sections:
        for p in section.header.paragraphs:
            replace_in_paragraph(p, replacements)
        replace_in_tables(section.header.tables, replacements)
        for p in section.footer.paragraphs:
            replace_in_paragraph(p, replacements)
        replace_in_tables(section.footer.tables, replacements)

    doc.save(str(output_path))


def convert_docx_to_pdf_word(docx_path, pdf_path):
    """Converts a DOCX file to PDF using Microsoft Word COM automation."""
    import win32com.client
    import pythoncom
    pythoncom.CoInitialize()
    word = None
    doc = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        doc = word.Documents.Open(str(docx_path))
        # 17 = wdFormatPDF
        doc.SaveAs2(str(pdf_path), FileFormat=17)
    finally:
        if doc is not None:
            doc.Close(SaveChanges=False)
        if word is not None:
            word.Quit()
        pythoncom.CoUninitialize()


def run_batch_merge(
    recipients_path,
    template_path,
    output_dir,
    overwrite=False,
    export_pdf=False,
    progress_callback=None
):
    """
    Main batch processing function.
    Reads recipients, validates columns, generates certificates, and returns detailed summary.
    """
    start_time = datetime.now()
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    logger.info("=" * 60)
    logger.info("CERTIFICATE GENERATION AGENT STARTED")
    logger.info(f"Recipients file: {recipients_path}")
    logger.info(f"Template file:   {template_path}")
    logger.info(f"Output folder:   {output_dir}")
    logger.info(f"Overwrite:       {overwrite}")
    logger.info(f"Export PDF:      {export_pdf}")

    # Check Microsoft Word
    word_installed, word_info = check_word_installed()
    logger.info(f"Microsoft Word status: {'INSTALLED' if word_installed else 'NOT DETECTED'} ({word_info})")
    if export_pdf and not word_installed:
        logger.warning("PDF export was requested, but Microsoft Word is not installed. Will generate DOCX only.")
        export_pdf = False

    # Read data
    try:
        records = read_recipients_file(recipients_path)
    except Exception as e:
        logger.error(f"Failed to read recipient data: {e}")
        raise

    total_records = len(records)
    logger.info(f"Total recipient records loaded: {total_records}")

    if total_records == 0:
        logger.warning("No recipient records found in file!")
        return {
            "total": 0,
            "success": 0,
            "failed": 0,
            "failed_details": [],
            "generated_files": [],
            "duration": 0
        }

    # Extract template placeholders & validate
    template_placeholders = extract_template_placeholders(template_path)
    logger.info(f"Detected template placeholders: {template_placeholders}")

    sample_record = records[0]
    available_cols = set(sample_record.keys()) - {"_row_number"}
    logger.info(f"Available data columns: {sorted(list(available_cols))}")

    # Check required fields
    missing_fields = []
    for ph in template_placeholders:
        # Check if ph matches any column (exact or case-insensitive)
        matched = any(ph.lower() == col.lower() for col in available_cols)
        if not matched:
            missing_fields.append(ph)

    if missing_fields:
        msg = f"Warning: The following template placeholders are not present in recipient columns: {missing_fields}"
        logger.warning(msg)

    # Process each recipient
    success_count = 0
    failed_details = []
    generated_files = []

    for idx, record in enumerate(records, start=1):
        row_num = record.get("_row_number", idx + 1)
        # Find recipient name for filename
        name_key = next((k for k in ["Name", "Recipient", "Full Name", "Student", "Participant"] if k in record), None)
        if not name_key:
            name_key = next((k for k in record.keys() if "name" in k.lower()), None)
        
        recipient_name = record.get(name_key, "").strip() if name_key else f"Recipient_{idx}"
        if not recipient_name:
            recipient_name = f"Recipient_Row_{row_num}"

        safe_name = sanitize_filename(recipient_name)

        try:
            target_docx = get_unique_filename(output_dir, safe_name, extension=".docx", overwrite=overwrite)
            generate_certificate(template_path, record, target_docx)
            
            pdf_path = None
            if export_pdf:
                target_pdf = target_docx.with_suffix(".pdf")
                try:
                    convert_docx_to_pdf_word(target_docx, target_pdf)
                    pdf_path = target_pdf
                except Exception as pdf_err:
                    logger.warning(f"Could not convert {target_docx.name} to PDF: {pdf_err}")

            success_count += 1
            generated_files.append({
                "recipient": recipient_name,
                "docx": str(target_docx),
                "pdf": str(pdf_path) if pdf_path else None,
                "row": row_num
            })
            logger.info(f"[{idx}/{total_records}] Successfully generated: {target_docx.name}")

        except Exception as err:
            err_msg = f"Row {row_num} ('{recipient_name}'): {str(err)}"
            logger.error(f"Failed to generate certificate: {err_msg}")
            failed_details.append({
                "row": row_num,
                "recipient": recipient_name,
                "error": str(err)
            })

        if progress_callback:
            progress_callback(idx, total_records, recipient_name)

    duration = (datetime.now() - start_time).total_seconds()
    logger.info("=" * 60)
    logger.info(f"BATCH RUN COMPLETE in {duration:.2f} seconds")
    logger.info(f"Total Rows: {total_records} | Success: {success_count} | Failed: {len(failed_details)}")
    if failed_details:
        logger.warning(f"Failures ({len(failed_details)}):")
        for f in failed_details:
            logger.warning(f"  - Row {f['row']} ({f['recipient']}): {f['error']}")
    logger.info("=" * 60)

    return {
        "total": total_records,
        "success": success_count,
        "failed": len(failed_details),
        "failed_details": failed_details,
        "generated_files": generated_files,
        "duration": duration,
        "word_installed": word_installed
    }


def main():
    parser = argparse.ArgumentParser(description="Certificate Generation Mail Merge Agent")
    parser.add_argument("--recipients", "-r", default="recipients.xlsx", help="Path to recipients Excel or CSV file")
    parser.add_argument("--template", "-t", default="templates/certificate_template.docx", help="Path to certificate template docx")
    parser.add_argument("--output", "-o", default="output_certificates", help="Directory to save generated certificates")
    parser.add_argument("--overwrite", action="store_true", help="Overwrite existing certificates instead of appending (1), (2)")
    parser.add_argument("--export-pdf", action="store_true", help="Also export certificates as PDF using Microsoft Word")
    
    args = parser.parse_args()

    project_root = Path(__file__).resolve().parent
    recipients_file = (project_root / args.recipients).resolve() if not os.path.isabs(args.recipients) else Path(args.recipients)
    template_file = (project_root / args.template).resolve() if not os.path.isabs(args.template) else Path(args.template)
    output_directory = (project_root / args.output).resolve() if not os.path.isabs(args.output) else Path(args.output)

    if not recipients_file.exists():
        # Fallback check for csv if xlsx was default
        if recipients_file.suffix == ".xlsx" and (recipients_file.with_suffix(".csv")).exists():
            recipients_file = recipients_file.with_suffix(".csv")
        else:
            print(f"Error: Recipients file not found at {recipients_file}", file=sys.stderr)
            sys.exit(1)

    if not template_file.exists():
        print(f"Error: Template file not found at {template_file}", file=sys.stderr)
        sys.exit(1)

    summary = run_batch_merge(
        recipients_path=recipients_file,
        template_path=template_file,
        output_dir=output_directory,
        overwrite=args.overwrite,
        export_pdf=args.export_pdf
    )

    print("\n--- SUMMARY ---")
    print(f"Total Processed: {summary['total']}")
    print(f"Success:         {summary['success']}")
    print(f"Failed:          {summary['failed']}")
    print(f"Time Taken:      {summary['duration']:.2f}s")
    print(f"Certificates in: {output_directory}")

    if summary["failed"] > 0:
        sys.exit(1)
    else:
        sys.exit(0)


if __name__ == "__main__":
    main()

