"""Desktop ETL mapping extractor and Copilot prompt builder.

Runs with WinPython Portable. Tkinter is part of Python; openpyxl is the only
third-party dependency. The workbook is read locally and never uploaded.
"""

import json
import re
import tkinter as tk
import zipfile
from datetime import date, datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from openpyxl import load_workbook


FIELD_ALIASES = {
    "source_system": {"src_system", "source_system", "source_application", "source_app"},
    "source_table": {"src_table", "source_table", "source_tablename"},
    "source_column": {"src_column", "source_column", "source_field", "source_attribute", "src_field"},
    "transformation_rule": {"transform_rule", "transformation_rule", "mapping", "transformation", "business_rule"},
    "target_system": {"tgt_system", "target_system"},
    "target_table": {"tgt_table", "target_table", "target_tablename"},
    "target_column": {"tgt_column", "target_column", "target_field", "target_attribute"},
    "data_type": {"data_type", "datatype", "target_data_type", "target_datatype"},
    "null_rule": {"null_rule", "null_handling", "nullable", "nullability"},
    "key_flag": {"key_flag", "key", "primary_key", "pk_flag"},
    "intermediate_staging_layer": {"intermediate_staging_layer", "staging_layer", "intermediate_layer"},
}
REQUIRED_HINTS = {"source_table", "target_column"}
ALL_FIELDS = list(FIELD_ALIASES)


def normalize(value):
    if value is None:
        return ""
    return re.sub(r"[^a-z0-9]+", "_", str(value).strip().lower()).strip("_")


def as_text(value):
    if value is None:
        return ""
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    return str(value).strip()


def header_fields(worksheet, row_number):
    result = {}
    extras = {}
    for col, cell in enumerate(worksheet[row_number], start=1):
        name = normalize(cell.value)
        if not name:
            continue
        found = next((field for field, aliases in FIELD_ALIASES.items() if name in aliases), None)
        if found:
            result[found] = col
        else:
            extras[name] = col
    return result, extras


def find_header(worksheet, limit=50):
    max_row = worksheet.max_row or limit
    for row_num in range(1, min(max_row, limit) + 1):
        fields, _ = header_fields(worksheet, row_num)
        if REQUIRED_HINTS.issubset(fields):
            return row_num, fields
    raise ValueError(
        "Could not identify the mapping headers. The sheet needs at least a "
        "Source Table and Target Column header."
    )


def extract_sheet(worksheet, source_file):
    header_row, fields = find_header(worksheet)
    mappings = []
    header_cells = tuple(worksheet[header_row])
    # Some valid .xlsx files omit worksheet dimension metadata. In read-only
    # mode max_column is then None; defaulting to 1 discards every target field.
    # Let openpyxl read the whole row instead of imposing a column limit.
    for row in worksheet.iter_rows(min_row=header_row + 1, values_only=True):
        entry = {field: "" for field in ALL_FIELDS}
        for field, col in fields.items():
            entry[field] = as_text(row[col - 1] if col <= len(row) else None)
        # Keep other populated workbook columns so no mapping detail is lost.
        extras = {}
        for col, cell in enumerate(header_cells, start=1):
            name = normalize(cell.value)
            known = any(name in aliases for aliases in FIELD_ALIASES.values())
            value = row[col - 1] if col <= len(row) else None
            if name and not known and value is not None and as_text(value):
                extras[str(cell.value).strip()] = as_text(value)
        if extras:
            entry["additional_fields"] = extras
        # Skip blank lines, merged section labels, and formatting rows.
        if not any(entry.get(k, "") for k in ("source_table", "source_column", "target_table", "target_column", "transformation_rule")):
            continue
        # A heading like D_CASA alone is a section label, not a mapping.
        if entry["source_table"] and not any(entry.get(k, "") for k in ("source_column", "target_table", "target_column", "transformation_rule")):
            continue
        mappings.append(entry)
    if not mappings:
        raise ValueError(
            "No mapping rows found below the detected header. Check that the "
            "worksheet contains populated source and target mapping columns."
        )
    return {
        "mapping_name": Path(source_file).stem,
        "source_file": Path(source_file).name,
        "worksheet": worksheet.title,
        "header_row": header_row,
        "mapping_count": len(mappings),
        "mappings": mappings,
    }


PROMPT_TEMPLATES = {
    "EDA": "Create EDA (data lake / ingestion) validation test cases from the ETL mapping below. Cover source ingestion, file/schema checks, staging, data quality, incremental loads, rejects, duplicates, nulls, and PII only where supported by the mapping. For each case provide ID, mapping reference, objective, validation steps or SQL/pseudocode, expected result, edge data, priority, and assumptions. Do not invent rules. Mark gaps as Clarification required.",
    "EDW": "Create EDW source-to-target transformation validation test cases from the ETL mapping below. Cover column mapping, cleansing, derivations, business rules, data types, null behavior, keys, and reconciliation where supported by the mapping. For each case provide ID, mapping reference, objective, validation steps or SQL/pseudocode, expected result, edge data, priority, and assumptions. Do not invent rules. Mark gaps as Clarification required.",
    "Datamart": "Create Datamart validation test cases from the ETL mapping below. Focus on mart-level transformations, aggregation, dimensional relationships, semantic outputs, and reconciliation only when the mapping provides enough evidence. For each case provide ID, mapping reference, objective, validation steps or SQL/pseudocode, expected result, edge data, priority, and assumptions. Do not invent business rules. Mark gaps as Clarification required.",
    "ETL Scripts": "Generate reviewable ETL script templates from the mapping below for the specified source and target technology. Include clear source/target references, transformations, null handling, datatype conversion, key handling, and error/reject handling only where defined. Add comments for assumptions and TODOs for missing details. Do not invent business rules. If the technology or rule is missing, ask for clarification instead of emitting executable assumptions.",
}


def make_prompt(layer, payload, technology="", notes=""):
    extra = []
    if technology.strip():
        extra.append(f"Technology / platform: {technology.strip()}")
    if notes.strip():
        extra.append(f"Layer notes: {notes.strip()}")
    context = "\n".join(extra)
    return (
        f"{PROMPT_TEMPLATES[layer]}\n\n"
        f"{context + chr(10) + chr(10) if context else ''}"
        "Mapping JSON:\n"
        f"{json.dumps(payload, indent=2, ensure_ascii=False)}"
    )


class MappingApp:
    def __init__(self, root):
        self.root = root
        root.title("ETL Mapping to Copilot Prompts")
        root.geometry("1000x760")
        root.minsize(820, 600)
        self.payload = None
        self.workbook_path = None
        self.sheet_names = []

        top = ttk.Frame(root, padding=12)
        top.pack(fill="x")
        ttk.Label(top, text="Excel mapping workbook:").grid(row=0, column=0, sticky="w")
        self.file_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.file_var).grid(row=0, column=1, sticky="ew", padx=8)
        ttk.Button(top, text="Browse…", command=self.browse).grid(row=0, column=2)
        ttk.Label(top, text="Worksheet:").grid(row=1, column=0, sticky="w", pady=(9, 0))
        self.sheet_var = tk.StringVar()
        self.sheet_combo = ttk.Combobox(top, textvariable=self.sheet_var, state="readonly", width=34)
        self.sheet_combo.grid(row=1, column=1, sticky="w", padx=8, pady=(9, 0))
        ttk.Button(top, text="Extract JSON", command=self.extract).grid(row=1, column=2, pady=(9, 0))
        top.columnconfigure(1, weight=1)

        opts = ttk.LabelFrame(root, text="Copilot prompt settings", padding=10)
        opts.pack(fill="x", padx=12, pady=(0, 8))
        ttk.Label(opts, text="Technology / platform (optional):").grid(row=0, column=0, sticky="w")
        self.tech_var = tk.StringVar()
        ttk.Entry(opts, textvariable=self.tech_var).grid(row=0, column=1, sticky="ew", padx=8)
        ttk.Label(opts, text="Layer notes (optional):").grid(row=1, column=0, sticky="w", pady=(7, 0))
        self.notes_var = tk.StringVar()
        ttk.Entry(opts, textvariable=self.notes_var).grid(row=1, column=1, sticky="ew", padx=8, pady=(7, 0))
        opts.columnconfigure(1, weight=1)

        actions = ttk.Frame(root, padding=(12, 0, 12, 8))
        actions.pack(fill="x")
        ttk.Label(actions, text="Prompt:").pack(side="left")
        self.layer_var = tk.StringVar(value="EDA")
        self.layer_combo = ttk.Combobox(actions, textvariable=self.layer_var, values=list(PROMPT_TEMPLATES), state="readonly", width=18)
        self.layer_combo.pack(side="left", padx=8)
        self.layer_combo.bind("<<ComboboxSelected>>", lambda _e: self.show_prompt())
        ttk.Button(actions, text="Generate / Preview", command=self.show_prompt).pack(side="left")
        ttk.Button(actions, text="Copy prompt", command=self.copy_prompt).pack(side="left", padx=6)
        ttk.Button(actions, text="Save prompt…", command=self.save_prompt).pack(side="left")
        ttk.Button(actions, text="Save all prompts ZIP…", command=self.save_prompt_zip).pack(side="left", padx=6)
        ttk.Button(actions, text="Save JSON…", command=self.save_json).pack(side="right")

        self.status_var = tk.StringVar(value="Choose an .xlsx workbook, select its mapping sheet, then click Extract JSON.")
        ttk.Label(root, textvariable=self.status_var, padding=(12, 2)).pack(fill="x")
        self.output = tk.Text(root, wrap="none", undo=False, font=("Consolas", 10))
        self.output.pack(fill="both", expand=True, padx=12, pady=(4, 12))
        scroll_y = ttk.Scrollbar(self.output, orient="vertical", command=self.output.yview)
        scroll_y.pack(side="right", fill="y")
        scroll_x = ttk.Scrollbar(root, orient="horizontal", command=self.output.xview)
        scroll_x.pack(fill="x", padx=12, pady=(0, 10))
        self.output.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

    def browse(self):
        path = filedialog.askopenfilename(title="Select mapping workbook", filetypes=[("Excel workbook", "*.xlsx *.xlsm"), ("All files", "*.*")])
        if not path:
            return
        self.workbook_path = Path(path)
        self.file_var.set(path)
        try:
            wb = load_workbook(path, read_only=True, data_only=True)
            self.sheet_names = wb.sheetnames
            wb.close()
            self.sheet_combo["values"] = self.sheet_names
            if self.sheet_names:
                self.sheet_var.set(self.sheet_names[0])
            self.status_var.set(f"Loaded workbook. Select the mapping worksheet and click Extract JSON.")
        except Exception as exc:
            messagebox.showerror("Workbook error", f"Could not open the workbook:\n{exc}")

    def extract(self):
        path = Path(self.file_var.get().strip())
        sheet = self.sheet_var.get().strip()
        if not path.is_file():
            messagebox.showwarning("Select workbook", "Browse to a valid Excel .xlsx or .xlsm workbook first.")
            return
        try:
            wb = load_workbook(path, read_only=True, data_only=True)
            if not sheet:
                sheet = wb.sheetnames[0]
            if sheet not in wb.sheetnames:
                raise ValueError(f"Worksheet '{sheet}' was not found.")
            self.payload = extract_sheet(wb[sheet], str(path))
            wb.close()
            self.workbook_path = path
            self.status_var.set(f"Extracted {self.payload['mapping_count']} mapping rows from '{sheet}'.")
            self.show_json()
        except Exception as exc:
            try:
                wb.close()
            except Exception:
                pass
            messagebox.showerror("Extraction error", str(exc))

    def show_json(self):
        if self.payload is None:
            return
        self.set_text(json.dumps(self.payload, indent=2, ensure_ascii=False))

    def show_prompt(self):
        if self.payload is None:
            messagebox.showinfo("Extract mapping first", "Select a workbook and click Extract JSON before generating prompts.")
            return
        layer = self.layer_var.get()
        self.set_text(make_prompt(layer, self.payload, self.tech_var.get(), self.notes_var.get()))

    def set_text(self, value):
        self.output.delete("1.0", "end")
        self.output.insert("1.0", value)

    def copy_prompt(self):
        if self.payload is None:
            messagebox.showinfo("Extract mapping first", "Extract the mapping before copying a prompt.")
            return
        self.show_prompt()
        value = self.output.get("1.0", "end-1c")
        self.root.clipboard_clear()
        self.root.clipboard_append(value)
        self.root.update()
        self.status_var.set(f"{self.layer_var.get()} prompt copied. Paste it into your approved Copilot environment.")

    def save_json(self):
        if self.payload is None:
            messagebox.showinfo("Extract mapping first", "Extract the mapping before saving JSON.")
            return
        path = filedialog.asksaveasfilename(title="Save mapping JSON", defaultextension=".json", filetypes=[("JSON", "*.json")])
        if path:
            Path(path).write_text(json.dumps(self.payload, indent=2, ensure_ascii=False), encoding="utf-8")
            self.status_var.set(f"JSON saved to {path}")

    def save_prompt(self):
        if self.payload is None:
            messagebox.showinfo("Extract mapping first", "Extract the mapping before saving a prompt.")
            return
        layer = self.layer_var.get()
        path = filedialog.asksaveasfilename(title="Save Copilot prompt", initialfile=f"{layer.replace(' ', '_')}_prompt.txt", defaultextension=".txt", filetypes=[("Text file", "*.txt")])
        if path:
            Path(path).write_text(make_prompt(layer, self.payload, self.tech_var.get(), self.notes_var.get()), encoding="utf-8")
            self.status_var.set(f"Prompt saved to {path}")

    def save_prompt_zip(self):
        if self.payload is None:
            messagebox.showinfo("Extract mapping first", "Extract the mapping before saving prompts.")
            return
        path = filedialog.asksaveasfilename(title="Save all Copilot prompts", initialfile="ETL_Copilot_Prompts.zip", defaultextension=".zip", filetypes=[("ZIP archive", "*.zip")])
        if path:
            with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
                for layer in PROMPT_TEMPLATES:
                    name = layer.replace(" ", "_")
                    zf.writestr(f"{name}_prompt.txt", make_prompt(layer, self.payload, self.tech_var.get(), self.notes_var.get()))
                zf.writestr("mapping.json", json.dumps(self.payload, indent=2, ensure_ascii=False))
            self.status_var.set(f"All prompts and mapping JSON saved to {path}")


if __name__ == "__main__":
    root = tk.Tk()
    MappingApp(root)
    root.mainloop()
