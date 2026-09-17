import os
import sys
import threading
import subprocess
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

import mail_merge
import verify_batch

class CertificateAgentGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Certificate Generation Agent")
        self.root.geometry("780x680")
        self.root.minsize(700, 600)

        # Apply native Windows styling
        self.style = ttk.Style()
        try:
            self.style.theme_use("vista")
        except Exception:
            pass

        self.project_dir = Path(__file__).resolve().parent

        # Default paths
        default_recipients = self.project_dir / "recipients.xlsx"
        if not default_recipients.exists() and (self.project_dir / "recipients.csv").exists():
            default_recipients = self.project_dir / "recipients.csv"
        
        default_template = self.project_dir / "templates" / "certificate_template.docx"
        default_output = self.project_dir / "output_certificates"

        self.recipient_path_var = tk.StringVar(value=str(default_recipients) if default_recipients.exists() else "")
        self.template_path_var = tk.StringVar(value=str(default_template) if default_template.exists() else "")
        self.output_dir_var = tk.StringVar(value=str(default_output))
        self.overwrite_var = tk.BooleanVar(value=False)
        self.export_pdf_var = tk.BooleanVar(value=False)

        # Word status check
        self.word_installed, self.word_info = mail_merge.check_word_installed()

        self._build_ui()

    def _build_ui(self):
        # Header Banner
        banner_frame = ttk.Frame(self.root, padding=12)
        banner_frame.pack(fill=tk.X)

        title_label = ttk.Label(
            banner_frame,
            text="🎓 Certificate Generation Agent",
            font=("Segoe UI", 16, "bold")
        )
        title_label.pack(anchor=tk.W)

        subtitle_label = ttk.Label(
            banner_frame,
            text="Automated Microsoft Word mail merge certificate generator for Windows",
            font=("Segoe UI", 9)
        )
        subtitle_label.pack(anchor=tk.W)

        # Word detection status
        status_text = f"Microsoft Word: {'Installed' if self.word_installed else 'Not Detected'}"
        status_color = "#107c41" if self.word_installed else "#d83b01"
        word_lbl = tk.Label(
            banner_frame,
            text=f"● {status_text}",
            fg=status_color,
            font=("Segoe UI", 9, "bold")
        )
        word_lbl.pack(anchor=tk.E, pady=(0, 2))

        ttk.Separator(self.root, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=12)

        # Main configuration group
        config_frame = ttk.LabelFrame(self.root, text=" 1. Select Data & Template Files ", padding=12)
        config_frame.pack(fill=tk.X, padx=14, pady=8)

        # Recipients row
        ttk.Label(config_frame, text="Recipients File (.xlsx / .csv):", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky=tk.W, pady=4)
        rec_entry = ttk.Entry(config_frame, textvariable=self.recipient_path_var, width=60)
        rec_entry.grid(row=0, column=1, padx=6, pady=4, sticky=tk.EW)
        ttk.Button(config_frame, text="Browse...", command=self._browse_recipients).grid(row=0, column=2, padx=4, pady=4)

        # Template row
        ttk.Label(config_frame, text="Certificate Template (.docx):", font=("Segoe UI", 9, "bold")).grid(row=1, column=0, sticky=tk.W, pady=4)
        tmpl_entry = ttk.Entry(config_frame, textvariable=self.template_path_var, width=60)
        tmpl_entry.grid(row=1, column=1, padx=6, pady=4, sticky=tk.EW)
        ttk.Button(config_frame, text="Browse...", command=self._browse_template).grid(row=1, column=2, padx=4, pady=4)

        # Output folder row
        ttk.Label(config_frame, text="Output Directory:", font=("Segoe UI", 9, "bold")).grid(row=2, column=0, sticky=tk.W, pady=4)
        out_entry = ttk.Entry(config_frame, textvariable=self.output_dir_var, width=60)
        out_entry.grid(row=2, column=1, padx=6, pady=4, sticky=tk.EW)
        ttk.Button(config_frame, text="Browse...", command=self._browse_output).grid(row=2, column=2, padx=4, pady=4)

        config_frame.columnconfigure(1, weight=1)

        # Options group
        options_frame = ttk.LabelFrame(self.root, text=" 2. Options ", padding=10)
        options_frame.pack(fill=tk.X, padx=14, pady=4)

        ttk.Checkbutton(
            options_frame,
            text="Overwrite existing files (if unchecked, appends (1), (2) to prevent lost data)",
            variable=self.overwrite_var
        ).pack(anchor=tk.W, pady=2)

        pdf_cb = ttk.Checkbutton(
            options_frame,
            text="Also export as PDF certificates (requires Microsoft Word)",
            variable=self.export_pdf_var
        )
        pdf_cb.pack(anchor=tk.W, pady=2)
        if not self.word_installed:
            pdf_cb.config(state=tk.DISABLED)

        # Actions frame
        actions_frame = ttk.Frame(self.root, padding=8)
        actions_frame.pack(fill=tk.X, padx=14, pady=4)

        ttk.Button(actions_frame, text="🔍 Check Fields & Preview", command=self._check_fields).pack(side=tk.LEFT, padx=4)
        
        self.btn_generate = ttk.Button(
            actions_frame,
            text="🚀 Generate Certificates",
            command=self._start_generation
        )
        self.btn_generate.pack(side=tk.LEFT, padx=4)

        ttk.Button(actions_frame, text="✔ Verify Batch", command=self._verify_batch).pack(side=tk.LEFT, padx=4)
        ttk.Button(actions_frame, text="📂 Open Output Folder", command=self._open_output_folder).pack(side=tk.RIGHT, padx=4)

        # Progress bar & Status
        prog_frame = ttk.Frame(self.root, padding=8)
        prog_frame.pack(fill=tk.X, padx=14)

        self.prog_bar = ttk.Progressbar(prog_frame, orient=tk.HORIZONTAL, mode="determinate")
        self.prog_bar.pack(fill=tk.X, pady=2)

        self.status_var = tk.StringVar(value="Ready. Select recipient file and template, then click 'Generate Certificates'.")
        status_lbl = ttk.Label(prog_frame, textvariable=self.status_var, font=("Segoe UI", 9, "italic"))
        status_lbl.pack(anchor=tk.W, pady=2)

        # Console Log
        log_frame = ttk.LabelFrame(self.root, text=" Log & Progress Output ", padding=8)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=14, pady=6)

        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, font=("Consolas", 9), height=10)
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def log(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)

    def _browse_recipients(self):
        path = filedialog.askopenfilename(
            title="Select Recipients Excel or CSV File",
            filetypes=[("Excel & CSV Files", "*.xlsx *.csv"), ("Excel Files", "*.xlsx"), ("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        if path:
            self.recipient_path_var.set(path)

    def _browse_template(self):
        path = filedialog.askopenfilename(
            title="Select Certificate Template Word Document",
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")]
        )
        if path:
            self.template_path_var.set(path)

    def _browse_output(self):
        path = filedialog.askdirectory(title="Select Output Directory")
        if path:
            self.output_dir_var.set(path)

    def _check_fields(self):
        rec_path = self.recipient_path_var.get().strip()
        tmpl_path = self.template_path_var.get().strip()

        if not os.path.exists(rec_path):
            messagebox.showerror("File Error", f"Recipient file not found:\n{rec_path}")
            return
        if not os.path.exists(tmpl_path):
            messagebox.showerror("File Error", f"Template file not found:\n{tmpl_path}")
            return

        try:
            placeholders = mail_merge.extract_template_placeholders(tmpl_path)
            records = mail_merge.read_recipients_file(rec_path)
            if not records:
                messagebox.showwarning("Empty File", "Recipient file contains no data rows!")
                return
            
            cols = [k for k in records[0].keys() if not k.startswith("_")]
            
            missing = [ph for ph in placeholders if not any(ph.lower() == c.lower() for c in cols)]

            msg = f"=== FIELD CHECK & PREVIEW ===\n\n"
            msg += f"Recipient Records Found: {len(records)}\n"
            msg += f"Columns in Recipient File: {cols}\n\n"
            msg += f"Placeholders in Template: {placeholders}\n\n"
            
            if missing:
                msg += f"⚠️ MISSING COLUMNS ({len(missing)}):\n"
                for m in missing:
                    msg += f"  - Placeholder '«{m}»' is missing in recipient file!\n"
            else:
                msg += "✅ ALL template placeholders have corresponding columns in recipient data!\n"

            self.log(msg)
            messagebox.showinfo("Field Check Result", msg)

        except Exception as e:
            messagebox.showerror("Error Checking Fields", str(e))

    def _start_generation(self):
        rec_path = self.recipient_path_var.get().strip()
        tmpl_path = self.template_path_var.get().strip()
        out_dir = self.output_dir_var.get().strip()

        if not os.path.exists(rec_path):
            messagebox.showerror("Error", "Please select a valid recipient Excel or CSV file.")
            return
        if not os.path.exists(tmpl_path):
            messagebox.showerror("Error", "Please select a valid certificate template (.docx).")
            return
        if not out_dir:
            messagebox.showerror("Error", "Please select an output directory.")
            return

        self.btn_generate.config(state=tk.DISABLED)
        self.prog_bar["value"] = 0
        self.status_var.set("Starting certificate generation...")

        # Run in background thread to avoid freezing GUI
        thread = threading.Thread(
            target=self._run_generation_thread,
            args=(rec_path, tmpl_path, out_dir, self.overwrite_var.get(), self.export_pdf_var.get())
        )
        thread.daemon = True
        thread.start()

    def _run_generation_thread(self, rec_path, tmpl_path, out_dir, overwrite, export_pdf):
        def progress(current, total, name):
            pct = (current / total) * 100
            self.root.after(0, lambda: self.prog_bar.config(value=pct))
            self.root.after(0, lambda: self.status_var.set(f"Processing [{current}/{total}]: {name}"))
            self.root.after(0, lambda: self.log(f"[{current}/{total}] Generated certificate for: {name}"))

        try:
            summary = mail_merge.run_batch_merge(
                recipients_path=rec_path,
                template_path=tmpl_path,
                output_dir=out_dir,
                overwrite=overwrite,
                export_pdf=export_pdf,
                progress_callback=progress
            )

            # Verification
            self.root.after(0, lambda: self.status_var.set("Verifying generated certificates..."))
            v_success, v_data = verify_batch.verify_certificates(rec_path, out_dir)

            def finish_ui():
                self.btn_generate.config(state=tk.NORMAL)
                self.prog_bar["value"] = 100
                self.status_var.set(f"Completed! {summary['success']} of {summary['total']} generated.")
                
                result_msg = (
                    f"Certificate Generation Complete!\n\n"
                    f"Total Recipients: {summary['total']}\n"
                    f"Successfully Generated: {summary['success']}\n"
                    f"Failed Rows: {summary['failed']}\n"
                    f"Time Taken: {summary['duration']:.2f} seconds\n\n"
                    f"Verification: {'PASSED (100%)' if v_success else 'ISSUES FOUND'}\n"
                    f"Saved To: {out_dir}"
                )
                self.log("\n" + "=" * 50 + "\n" + result_msg + "\n" + "=" * 50)
                messagebox.showinfo("Batch Complete", result_msg)

            self.root.after(0, finish_ui)

        except Exception as e:
            def err_ui():
                self.btn_generate.config(state=tk.NORMAL)
                self.status_var.set("Error occurred during generation.")
                self.log(f"ERROR: {str(e)}")
                messagebox.showerror("Generation Error", f"An error occurred:\n{str(e)}")
            self.root.after(0, err_ui)

    def _verify_batch(self):
        rec_path = self.recipient_path_var.get().strip()
        out_dir = self.output_dir_var.get().strip()

        if not os.path.exists(rec_path):
            messagebox.showerror("Error", "Please select a valid recipient file.")
            return
        if not os.path.exists(out_dir):
            messagebox.showerror("Error", "Output directory does not exist.")
            return

        success, result = verify_batch.verify_certificates(rec_path, out_dir)
        msg = f"=== VERIFICATION RESULTS ===\n\n"
        msg += f"Total Expected: {result.get('total', 0)}\n"
        msg += f"Verified Valid: {result.get('passed', 0)}\n"
        msg += f"Failed/Missing: {len(result.get('failed', []))}\n\n"
        if success:
            msg += "✅ ALL certificates verified successfully with correct data!"
            self.log(msg)
            messagebox.showinfo("Verification Passed", msg)
        else:
            msg += "⚠️ Discrepancies detected. See log for details."
            self.log(msg)
            messagebox.showwarning("Verification Issues", msg)

    def _open_output_folder(self):
        out_dir = self.output_dir_var.get().strip()
        if os.path.exists(out_dir):
            os.startfile(out_dir)
        else:
            messagebox.showwarning("Folder Not Found", f"Output folder does not exist yet:\n{out_dir}")


def main():
    root = tk.Tk()
    app = CertificateAgentGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
