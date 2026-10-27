"""
main.py
-------
English-language Tkinter GUI for the ASL Inventory Parameter Optimization Tool.

The heavy lifting (reading, rule checking, writing) is delegated to
``processor`` and ``excel_writer``.  Processing runs on a background thread so
the window stays responsive and the progress bar can animate.

Build into a Windows executable with PyInstaller (see build_exe.bat / README).
"""

import os
import threading
import traceback

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import processor
import excel_writer


APP_TITLE = "ASL Inventory Parameter Optimization Tool"


class OptimizationApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("760x560")
        self.root.minsize(680, 520)

        self.input_path = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.status_text = tk.StringVar(value="Please upload an ASL Excel file.")

        self._build_ui()

    # ------------------------------------------------------------------ UI
    def _build_ui(self):
        pad = {"padx": 12, "pady": 6}

        header = tk.Frame(self.root, bg="#1F4E78")
        header.pack(fill="x")
        tk.Label(
            header, text=APP_TITLE, bg="#1F4E78", fg="white",
            font=("Segoe UI", 15, "bold"), pady=14,
        ).pack(fill="x")

        body = tk.Frame(self.root)
        body.pack(fill="both", expand=True, padx=14, pady=10)

        # --- Step 1: upload file ---
        tk.Label(body, text="1.  Select the ASL Excel file",
                 font=("Segoe UI", 10, "bold")).grid(row=0, column=0,
                                                     sticky="w", **pad)
        file_row = tk.Frame(body)
        file_row.grid(row=1, column=0, columnspan=2, sticky="we", padx=12)
        tk.Entry(file_row, textvariable=self.input_path, width=58).pack(
            side="left", fill="x", expand=True)
        tk.Button(file_row, text="Upload ASL File",
                  command=self.on_upload).pack(side="left", padx=8)

        # --- Step 2: output folder ---
        tk.Label(body, text="2.  Select the output folder (optional)",
                 font=("Segoe UI", 10, "bold")).grid(row=2, column=0,
                                                     sticky="w", **pad)
        out_row = tk.Frame(body)
        out_row.grid(row=3, column=0, columnspan=2, sticky="we", padx=12)
        tk.Entry(out_row, textvariable=self.output_folder, width=58).pack(
            side="left", fill="x", expand=True)
        tk.Button(out_row, text="Select Output Folder",
                  command=self.on_select_output).pack(side="left", padx=8)

        # --- Step 3: run ---
        tk.Label(body, text="3.  Run the optimization check",
                 font=("Segoe UI", 10, "bold")).grid(row=4, column=0,
                                                     sticky="w", **pad)
        self.run_button = tk.Button(
            body, text="Run Optimization Check", command=self.on_run,
            bg="#2E7D32", fg="white", font=("Segoe UI", 10, "bold"),
            height=2, width=24,
        )
        self.run_button.grid(row=5, column=0, sticky="w", padx=12, pady=4)

        # --- Progress bar ---
        self.progress = ttk.Progressbar(body, mode="determinate", maximum=100)
        self.progress.grid(row=6, column=0, columnspan=2, sticky="we",
                           padx=12, pady=(14, 4))

        # --- Status / log area ---
        tk.Label(body, text="Status", font=("Segoe UI", 10, "bold")).grid(
            row=7, column=0, sticky="w", padx=12, pady=(10, 0))
        self.log = tk.Text(body, height=10, width=86, state="disabled",
                           bg="#F7F7F7", relief="solid", borderwidth=1)
        self.log.grid(row=8, column=0, columnspan=2, sticky="nsew",
                      padx=12, pady=6)

        tk.Label(body, textvariable=self.status_text, fg="#1F4E78",
                 anchor="w").grid(row=9, column=0, columnspan=2, sticky="we",
                                  padx=12, pady=(0, 8))

        body.grid_columnconfigure(0, weight=1)
        body.grid_rowconfigure(8, weight=1)

    # -------------------------------------------------------------- helpers
    def log_message(self, message):
        self.status_text.set(message)
        self.log.configure(state="normal")
        self.log.insert("end", message + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")
        self.root.update_idletasks()

    def set_progress(self, value):
        self.progress["value"] = value
        self.root.update_idletasks()

    def set_busy(self, busy):
        self.run_button.configure(state="disabled" if busy else "normal")

    # --------------------------------------------------------------- events
    def on_upload(self):
        path = filedialog.askopenfilename(
            title="Select the ASL Excel file",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if not path:
            return

        ext = os.path.splitext(path)[1].lower()
        if ext not in (".xlsx", ".xls"):
            messagebox.showerror(
                "Invalid File",
                "The selected file is not an Excel file.\nPlease select a "
                ".xlsx or .xls file.",
            )
            return

        if "asl" not in os.path.basename(path).lower():
            proceed = messagebox.askyesno(
                "File Name Warning",
                "The selected file name does not contain 'ASL'.\n\n"
                "Do you want to continue anyway?",
            )
            if not proceed:
                return

        self.input_path.set(path)
        if not self.output_folder.get():
            self.output_folder.set(os.path.dirname(path))
        self.log_message("File uploaded successfully: {}".format(
            os.path.basename(path)))

    def on_select_output(self):
        folder = filedialog.askdirectory(title="Select the output folder")
        if folder:
            self.output_folder.set(folder)
            self.log_message("Output folder set to: {}".format(folder))

    def on_run(self):
        if not self.input_path.get():
            messagebox.showwarning(
                "No File", "Please upload an ASL Excel file first.")
            return
        if not os.path.exists(self.input_path.get()):
            messagebox.showerror(
                "File Not Found", "The selected file no longer exists.")
            return

        self.set_busy(True)
        self.set_progress(0)
        thread = threading.Thread(target=self._run_worker, daemon=True)
        thread.start()

    # --------------------------------------------------------------- worker
    def _run_worker(self):
        """Runs on a background thread; only logs back to the UI thread."""
        try:
            input_path = self.input_path.get()
            output_folder = self.output_folder.get() or None

            self.log_message("Reading ASL file...")
            self.set_progress(15)
            df = processor.read_asl_file(input_path)

            column_map, missing, distance_column = processor.resolve_columns(df)
            if missing:
                raise ValueError(
                    'Error: Required column "{}" was not found in the '
                    "uploaded ASL file.".format(missing[0])
                )
            if distance_column is None:
                self.log_message(
                    "Warning: no distance/KM column found. The KM-based Safety "
                    "Time check (Rule 2) will be skipped."
                )

            self.log_message("Filtering Type Trim = PG...")
            self.set_progress(35)
            pg_df = processor.filter_pg_rows(df, column_map["Type Trim"])
            if pg_df.empty:
                self.log_message("No PG rows found in the uploaded ASL file.")
                self._finish(success=False)
                return
            self.log_message("Found {} PG rows.".format(len(pg_df)))

            self.log_message("Checking inventory parameter rules...")
            self.set_progress(60)
            detail_df = processor.generate_detail_report(
                pg_df, column_map, distance_column)

            if detail_df.empty:
                self.log_message("No optimization issues found.")
                self._finish(success=False)
                return
            self.log_message("Found {} rows needing attention.".format(
                len(detail_df)))

            self.log_message("Generating optimization report...")
            self.set_progress(80)
            summary_df = processor.generate_summary_report(detail_df)

            output_path = excel_writer.build_output_path(
                input_path, output_folder)
            excel_writer.write_output_excel(summary_df, detail_df, output_path)

            self.set_progress(100)
            self.log_message("Output file saved successfully.")
            self.log_message("Processing completed successfully.")
            self.root.after(0, lambda: messagebox.showinfo(
                "Done",
                "Processing completed successfully.\n\nOutput file:\n{}".format(
                    output_path),
            ))
            self._finish(success=True)

        except (ValueError, PermissionError) as exc:
            self.root.after(0, lambda: messagebox.showerror("Error", str(exc)))
            self.log_message("Error: {}".format(exc))
            self._finish(success=False)
        except Exception as exc:  # unexpected
            traceback.print_exc()
            msg = "An unexpected error occurred:\n{}".format(exc)
            self.root.after(0, lambda: messagebox.showerror("Error", msg))
            self.log_message(msg)
            self._finish(success=False)

    def _finish(self, success):
        self.set_busy(False)
        if not success:
            self.set_progress(0)


def main():
    root = tk.Tk()
    OptimizationApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
