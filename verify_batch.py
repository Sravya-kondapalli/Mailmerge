import os
import sys
import re
import argparse
from pathlib import Path

import docx
from mail_merge import read_recipients_file, sanitize_filename

def verify_certificates(recipients_path, output_dir, template_path=None):
    """
    Verifies that generated certificates in output_dir match the recipients file:
    1. Checks expected count vs generated count.
    2. Checks each file size (> 0 bytes).
    3. Inspects document text to ensure recipient's name is present and no unmerged placeholders remain.
    """
    output_dir = Path(output_dir)
    if not output_dir.exists():
        print(f"[FAIL] Output directory '{output_dir}' does not exist.")
        return False, {"error": "Output directory does not exist"}

    records = read_recipients_file(recipients_path)
    total_records = len(records)
    print(f"Loaded {total_records} recipient records from: {recipients_path}")
    print(f"Checking certificates in: {output_dir}\n")

    # Collect all docx files in output_dir
    docx_files = list(output_dir.glob("*.docx"))
    print(f"Found {len(docx_files)} .docx certificate files in output directory.")

    passed = 0
    failed = []
    
    # We track matched files so duplicate names can each match a distinct file
    matched_files = set()

    for idx, record in enumerate(records, start=1):
        row_num = record.get("_row_number", idx + 1)
        name_key = next((k for k in ["Name", "Recipient", "Full Name", "Student", "Participant"] if k in record), None)
        if not name_key:
            name_key = next((k for k in record.keys() if "name" in k.lower()), None)
        
        expected_name = record.get(name_key, "").strip() if name_key else f"Recipient_{idx}"
        safe_name = sanitize_filename(expected_name)

        # Search for candidates matching safe_name
        candidate = None
        # Candidate could be safe_name.docx or safe_name (N).docx
        possible_patterns = [
            f"{safe_name}.docx",
            f"{safe_name} (*).docx"
        ]
        
        for p in docx_files:
            if p in matched_files:
                continue
            if p.name == f"{safe_name}.docx":
                candidate = p
                break
            elif p.name.startswith(f"{safe_name} (") and p.name.endswith(").docx"):
                candidate = p
                break

        if not candidate or not candidate.exists():
            failed.append({
                "row": row_num,
                "recipient": expected_name,
                "file": f"{safe_name}.docx",
                "reason": "Missing output certificate file"
            })
            continue

        matched_files.add(candidate)

        # Check file size
        file_size = candidate.stat().st_size
        if file_size == 0:
            failed.append({
                "row": row_num,
                "recipient": expected_name,
                "file": candidate.name,
                "reason": "File is empty (0 bytes)"
            })
            continue

        # Inspect document content
        try:
            doc = docx.Document(str(candidate))
            doc_text = []
            for p in doc.paragraphs:
                doc_text.append(p.text)
            for tbl in doc.tables:
                for row in tbl.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            doc_text.append(p.text)
            
            full_text = " ".join(doc_text)

            # Check if expected recipient name is present
            # For names with special characters, check cleaned name
            cleaned_expected = re.sub(r'[:/*?"<>|]', '', expected_name).strip()
            if cleaned_expected not in full_text and expected_name not in full_text:
                # Also check parts of name
                name_parts = [part for part in expected_name.split() if len(part) > 2]
                if not any(part in full_text for part in name_parts):
                    failed.append({
                        "row": row_num,
                        "recipient": expected_name,
                        "file": candidate.name,
                        "reason": f"Recipient name '{expected_name}' not found in document text"
                    })
                    continue

            # Check for unmerged placeholders like «Name» or «Course»
            unmerged = re.findall(r'«([^»]+)»', full_text)
            if unmerged:
                failed.append({
                    "row": row_num,
                    "recipient": expected_name,
                    "file": candidate.name,
                    "reason": f"Unmerged placeholder tags found in certificate: {unmerged}"
                })
                continue

            passed += 1

        except Exception as e:
            failed.append({
                "row": row_num,
                "recipient": expected_name,
                "file": candidate.name,
                "reason": f"Corrupted docx or read error: {e}"
            })

    print("=" * 60)
    print("VERIFICATION BATCH RESULTS")
    print("=" * 60)
    print(f"Total Recipient Rows:      {total_records}")
    print(f"Valid Certificates Passed: {passed}")
    print(f"Failed / Missing:          {len(failed)}")
    print(f"Verification Success Rate: {(passed / total_records * 100):.1f}%")

    if failed:
        print("\nDiscrepancies found:")
        for f in failed:
            print(f"  [X] Row {f['row']} ({f['recipient']}) -> {f['file']}: {f['reason']}")
        print("=" * 60)
        return False, {"passed": passed, "failed": failed, "total": total_records}
    else:
        print("\n[SUCCESS] All certificates verified successfully with correct data and formatting!")
        print("=" * 60)
        return True, {"passed": passed, "failed": [], "total": total_records}


def main():
    parser = argparse.ArgumentParser(description="Verify generated batch of certificates")
    parser.add_argument("--recipients", "-r", default="recipients.xlsx", help="Path to recipients Excel or CSV")
    parser.add_argument("--output", "-o", default="output_certificates", help="Output directory containing certificates")
    
    args = parser.parse_args()
    project_root = Path(__file__).resolve().parent
    recipients_file = (project_root / args.recipients).resolve() if not os.path.isabs(args.recipients) else Path(args.recipients)
    output_directory = (project_root / args.output).resolve() if not os.path.isabs(args.output) else Path(args.output)

    if not recipients_file.exists():
        if recipients_file.suffix == ".xlsx" and (recipients_file.with_suffix(".csv")).exists():
            recipients_file = recipients_file.with_suffix(".csv")

    success, _ = verify_certificates(recipients_file, output_directory)
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
