import os
import re
import glob
import datetime
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import openpyxl
from openpyxl.utils import column_index_from_string
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx2pdf import convert

try:
    import pythoncom
except ImportError:
    pythoncom = None


# Extraction rules for the 4 target sheets
# font_size applies to every cell in the table, including the header rows
SHEET_CONFIGS = [
    {
        "target_name": "Account Summary",
        "cols": list(range(1, column_index_from_string("T") + 1)),
        "pdf_prefix": "Account Summary",
        "font_size": 5.0,
    },
    {
        "target_name": "2026 Sequencing",
        "cols": list(range(1, column_index_from_string("H") + 1)),
        "pdf_prefix": "2026 Sequencing Detail",
        "font_size": 5.0,
    },
    {
        "target_name": "Oligo Account Summary",
        "cols": list(range(1, column_index_from_string("R") + 1)),
        "pdf_prefix": "Oligo Account Summary",
        "font_size": 5.0,
    },
    {
        "target_name": "Oligo 2026",
        "cols": [column_index_from_string(c) for c in ["A", "B", "C", "D", "E", "F", "J"]],
        "pdf_prefix": "Oligo 2026 Detail",
        "font_size": 5.0,
    },
]


# ==================== Core data processing and Word/PDF generation ====================

def extract_suffix_from_filename(filename: str) -> str:
    """Extract everything from 'BBAP-xxx' to the end of the Excel file name (without extension)."""
    stem = os.path.splitext(os.path.basename(filename))[0]
    match = re.search(r"(BBAP-\d+.*)", stem, re.IGNORECASE)
    if match:
        return match.group(1).strip()
    return stem.strip()


def format_cell_value(cell) -> str:
    """Format an Excel cell value: dates, float precision and empty values."""
    val = cell.value
    if val is None:
        return ""
    if isinstance(val, datetime.datetime):
        if val.hour == 0 and val.minute == 0 and val.second == 0:
            return val.strftime("%Y-%m-%d")
        return val.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(val, datetime.date):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, float):
        num_fmt = cell.number_format or ""
        if "0.00" in num_fmt:
            return f"{val:.2f}"
        if val.is_integer():
            return str(int(val))
        return f"{val:.2f}".rstrip("0").rstrip(".")
    return str(val).strip()


def set_cell_shading(cell, color_hex: str):
    """Set the cell background color."""
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading)


def set_row_cant_split(row):
    """Prevent a single row from being split across pages."""
    trPr = row._tr.get_or_add_trPr()
    trPr.append(OxmlElement("w:cantSplit"))


def set_cell_border_from_excel(docx_cell, excel_cell):
    """Map the borders of an Excel cell onto the Word cell."""
    tcPr = docx_cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    has_border = False

    borders_map = {
        'top': excel_cell.border.top,
        'bottom': excel_cell.border.bottom,
        'left': excel_cell.border.left,
        'right': excel_cell.border.right
    }

    for side, border_obj in borders_map.items():
        if border_obj and border_obj.style:
            has_border = True
            color_hex = "BFBFBF"
            if border_obj.color and border_obj.color.rgb:
                rgb_str = str(border_obj.color.rgb)
                if len(rgb_str) == 8:
                    color_hex = rgb_str[2:]
                elif len(rgb_str) == 6:
                    color_hex = rgb_str

            b_el = OxmlElement(f'w:{side}')
            b_el.set(qn('w:val'), 'single')
            b_el.set(qn('w:sz'), '4')  # 0.5 pt
            b_el.set(qn('w:space'), '0')
            b_el.set(qn('w:color'), color_hex)
            tcBorders.append(b_el)
        else:
            b_el = OxmlElement(f'w:{side}')
            b_el.set(qn('w:val'), 'none')
            tcBorders.append(b_el)

    if has_border:
        tcPr.append(tcBorders)


def get_valid_rows_and_merges(ws, target_cols: list, sheet_name: str):
    """Extract all non-empty rows in the target columns and build the merge list."""
    col_set = set(target_cols)
    col_to_new_idx = {orig_col: idx for idx, orig_col in enumerate(target_cols)}

    selected_rows = []
    for r in range(1, ws.max_row + 1):
        row_has_data = any(
            ws.cell(r, c).value is not None and str(ws.cell(r, c).value).strip() != ""
            for c in target_cols
        )
        if row_has_data:
            selected_rows.append(r)

    if not selected_rows:
        return [], [], [], []

    orig_row_to_new_idx = {orig_r: idx for idx, orig_r in enumerate(selected_rows)}

    merged_ranges = []
    for rng in ws.merged_cells.ranges:
        if rng.min_row in orig_row_to_new_idx and rng.max_row in orig_row_to_new_idx:
            intersect_cols = [c for c in range(rng.min_col, rng.max_col + 1) if c in col_set]
            if len(intersect_cols) > 1 or rng.min_row != rng.max_row:
                new_r1 = orig_row_to_new_idx[rng.min_row]
                new_r2 = orig_row_to_new_idx[rng.max_row]
                new_c1 = col_to_new_idx[min(intersect_cols)]
                new_c2 = col_to_new_idx[max(intersect_cols)]
                if new_r1 != new_r2 or new_c1 != new_c2:
                    merged_ranges.append((new_r1, new_c1, new_r2, new_c2))

    # Special rule 1: in "Oligo Account Summary", the first row containing "ACCOUNT SUMMARY"
    # is merged across the whole width (A to R)
    clean_sheet_name = sheet_name.strip()
    if "Oligo Account Summary" in clean_sheet_name:
        for idx, orig_r in enumerate(selected_rows):
            row_text = "".join([str(ws.cell(orig_r, c).value or "") for c in target_cols])
            if "ACCOUNT SUMMARY" in row_text.upper():
                merged_ranges.append((idx, 0, idx, len(target_cols) - 1))
                break

    # Special rule 2: Customer Information block in "Account Summary" and "Oligo Account Summary"
    if "Account Summary" in clean_sheet_name or "Oligo Account Summary" in clean_sheet_name:
        for idx, orig_r in enumerate(selected_rows):
            cell_val = str(ws.cell(orig_r, 1).value or "").strip()
            if "Customer Information" in cell_val:
                # Title row: merge columns A to C so it matches the label rows below
                merged_ranges.append((idx, 0, idx, 2))
            elif any(k in cell_val for k in ["Name", "Institution", "PI", "Email", "Contact", "Address", "Package", "Person", "Delivery", "Billing"]):
                # Label rows: merge columns A, B, C (index 0, 1, 2)
                merged_ranges.append((idx, 0, idx, 2))

    rows_data = []
    cell_objects = []
    for orig_r in selected_rows:
        row_items = []
        row_cells = []
        for orig_c in target_cols:
            cell = ws.cell(orig_r, orig_c)
            text = format_cell_value(cell)
            is_bold = bool(cell.font and cell.font.bold)
            fill_color = None
            if cell.fill and cell.fill.fill_type == "solid" and cell.fill.start_color:
                rgb = cell.fill.start_color.rgb
                if isinstance(rgb, str) and len(rgb) == 8 and rgb != "00000000":
                    fill_color = rgb[2:]
            row_items.append({
                "text": text,
                "bold": is_bold,
                "fill": fill_color,
            })
            row_cells.append(cell)
        rows_data.append(row_items)
        cell_objects.append(row_cells)

    col_widths = []
    for orig_c in target_cols:
        col_letter = openpyxl.utils.get_column_letter(orig_c)
        w = ws.column_dimensions[col_letter].width
        base_w = w if w and w > 0 else 10.0
        col_widths.append(base_w)

    return rows_data, col_widths, merged_ranges, cell_objects


def insert_sheet_into_template(template_path: str, output_docx_path: str, rows_data, col_widths, merged_ranges, cell_objects, font_size: float, sheet_name: str, log=print):
    """Insert the extracted table, centered, into the Word template and save it."""
    doc = Document(template_path)

    for section in doc.sections:
        section.left_margin = Inches(0.3)
        section.right_margin = Inches(0.3)

    section = doc.sections[0]
    usable_width_inches = (
        section.page_width.inches - section.left_margin.inches - section.right_margin.inches
    )

    num_rows = len(rows_data)
    num_cols = len(col_widths)

    table = doc.add_table(rows=num_rows, cols=num_cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    total_excel_width = sum(col_widths)
    scaled_widths = [
        Inches((w / total_excel_width) * usable_width_inches) for w in col_widths
    ]

    # Merge cells; failures are written to the log instead of being silently ignored
    for r1, c1, r2, c2 in merged_ranges:
        try:
            table.cell(r1, c1).merge(table.cell(r2, c2))
        except Exception as e:
            log(f"  [Merge failed] {sheet_name}: rows {r1}-{r2}, cols {c1}-{c2}: {e}")

    # Set widths AFTER merging:
    # 1) table grid column widths
    for i, gc in enumerate(table._tbl.tblGrid.gridCol_lst):
        gc.w = scaled_widths[i]

    # 2) each cell width = sum of the widths of the columns it spans
    for row in table.rows:
        col_idx = 0
        for tc in row._tr.tc_lst:
            span = tc.grid_span
            tc.width = sum(scaled_widths[col_idx:col_idx + span])
            col_idx += span

    for r_idx, row_items in enumerate(rows_data):
        row = table.rows[r_idx]
        set_row_cant_split(row)
        for c_idx, item in enumerate(row_items):
            cell = table.cell(r_idx, c_idx)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

            set_cell_border_from_excel(cell, cell_objects[r_idx][c_idx])

            if item["text"] and not cell.text:
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_before = Pt(1)
                p.paragraph_format.space_after = Pt(1)
                run = p.add_run(item["text"])
                run.font.name = "Arial"
                run.font.size = Pt(font_size)
                run.bold = item["bold"]

            if item["fill"] and item["fill"].upper() != "FFFFFF":
                set_cell_shading(cell, item["fill"])

    while len(doc.paragraphs) > 1 and not doc.paragraphs[-1].text.strip():
        p_element = doc.paragraphs[-1]._element
        p_element.getparent().remove(p_element)

    doc.save(output_docx_path)


# ==================== Tkinter GUI ====================

class AccountSummaryConverterApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Account Summary Batch Extractor & PDF Converter")
        self.root.geometry("760x520")
        self.root.resizable(True, True)

        self.input_dir_var = tk.StringVar()
        self.template_path_var = tk.StringVar()
        self.output_dir_var = tk.StringVar()
        self.keep_docx_var = tk.BooleanVar(value=False)
        self.status_var = tk.StringVar(value="Ready: select the Excel folder, Word template and output folder")

        self.is_running = False
        self._build_ui()

    def _build_ui(self):
        main_frame = ttk.Frame(self.root, padding="15 15 15 15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        path_group = ttk.LabelFrame(main_frame, text=" Files & Folders ", padding="12 10 12 10")
        path_group.pack(fill=tk.X, pady=(0, 10))
        path_group.columnconfigure(1, weight=1)

        ttk.Label(path_group, text="Excel folder:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(path_group, textvariable=self.input_dir_var).grid(row=0, column=1, sticky=tk.EW, padx=8, pady=5)
        ttk.Button(path_group, text="Browse folder...", width=16, command=self.select_input_dir).grid(row=0, column=2, pady=5)

        ttk.Label(path_group, text="Word template (.docx):").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(path_group, textvariable=self.template_path_var).grid(row=1, column=1, sticky=tk.EW, padx=8, pady=5)
        ttk.Button(path_group, text="Select template...", width=16, command=self.select_template_file).grid(row=1, column=2, pady=5)

        ttk.Label(path_group, text="PDF output folder:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(path_group, textvariable=self.output_dir_var).grid(row=2, column=1, sticky=tk.EW, padx=8, pady=5)
        ttk.Button(path_group, text="Select folder...", width=16, command=self.select_output_dir).grid(row=2, column=2, pady=5)

        opt_group = ttk.LabelFrame(main_frame, text=" Export Options ", padding="12 8 12 8")
        opt_group.pack(fill=tk.X, pady=(0, 10))

        ttk.Checkbutton(
            opt_group,
            text="Keep the intermediate Word (.docx) files (by default they are deleted after PDF conversion)",
            variable=self.keep_docx_var,
        ).pack(anchor=tk.W, pady=2)

        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill=tk.X, pady=(0, 10))

        self.start_btn = ttk.Button(
            action_frame, text="▶ Start batch extract & generate PDFs", command=self.start_processing
        )
        self.start_btn.pack(side=tk.LEFT, ipadx=12, ipady=4)

        self.progress = ttk.Progressbar(main_frame, mode="determinate")
        self.progress.pack(fill=tk.X, pady=(0, 4))

        ttk.Label(main_frame, textvariable=self.status_var, foreground="#333333").pack(anchor=tk.W, pady=(0, 8))

        log_group = ttk.LabelFrame(main_frame, text=" Run Log ", padding="8 8 8 8")
        log_group.pack(fill=tk.BOTH, expand=True)

        self.log_text = tk.Text(log_group, height=14, state=tk.DISABLED, wrap=tk.WORD, font=("Consolas", 9))
        scrollbar = ttk.Scrollbar(log_group, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)

        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def select_input_dir(self):
        folder = filedialog.askdirectory(title="Select the folder containing the Excel files")
        if folder:
            self.input_dir_var.set(folder)
            if not self.output_dir_var.get():
                self.output_dir_var.set(os.path.join(folder, "output_pdfs"))

    def select_template_file(self):
        filepath = filedialog.askopenfilename(
            title="Select the Word template",
            filetypes=[("Word documents", "*.docx"), ("All files", "*.*")],
        )
        if filepath:
            self.template_path_var.set(filepath)

    def select_output_dir(self):
        folder = filedialog.askdirectory(title="Select the PDF output folder")
        if folder:
            self.output_dir_var.set(folder)

    def log(self, message: str):
        def _append():
            self.log_text.configure(state=tk.NORMAL)
            self.log_text.insert(tk.END, message + "\n")
            self.log_text.see(tk.END)
            self.log_text.configure(state=tk.DISABLED)

        self.root.after(0, _append)

    def update_status(self, status_text: str, progress_val: float = None):
        def _update():
            self.status_var.set(status_text)
            if progress_val is not None:
                self.progress["value"] = progress_val

        self.root.after(0, _update)

    def start_processing(self):
        if self.is_running:
            return

        input_dir = self.input_dir_var.get().strip()
        template_path = self.template_path_var.get().strip()
        output_dir = self.output_dir_var.get().strip()

        if not input_dir or not os.path.isdir(input_dir):
            messagebox.showerror("Path error", "Please select a valid Excel input folder!")
            return
        if not template_path or not os.path.isfile(template_path):
            messagebox.showerror("Path error", "Please select a valid Word template (.docx) file!")
            return
        if not output_dir:
            messagebox.showerror("Path error", "Please specify the PDF output folder!")
            return

        self.log_text.configure(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        self.log_text.configure(state=tk.DISABLED)

        self.is_running = True
        self.start_btn.configure(state=tk.DISABLED)
        self.progress["value"] = 0

        worker = threading.Thread(
            target=self._run_batch_worker,
            args=(input_dir, template_path, output_dir, self.keep_docx_var.get()),
            daemon=True,
        )
        worker.start()

    def _run_batch_worker(self, input_dir: str, template_path: str, output_dir: str, keep_docx: bool):
        if pythoncom is not None:
            pythoncom.CoInitialize()

        try:
            os.makedirs(output_dir, exist_ok=True)
            temp_docx_dir = os.path.join(output_dir, "_temp_docx") if not keep_docx else output_dir
            os.makedirs(temp_docx_dir, exist_ok=True)

            all_xlsx = [
                f for f in glob.glob(os.path.join(input_dir, "*.xlsx"))
                if not os.path.basename(f).startswith("~$")
            ]

            if not all_xlsx:
                self.log(f"[Warning] No .xlsx files found in folder: {input_dir}")
                self.update_status("No Excel files found", 0)
                return

            total_steps = len(all_xlsx) * len(SHEET_CONFIGS)
            current_step = 0
            success_count = 0

            self.log(f"Found {len(all_xlsx)} Excel file(s). Starting batch processing...\n" + "=" * 55)

            for excel_path in all_xlsx:
                base_name = os.path.basename(excel_path)
                suffix_name = extract_suffix_from_filename(base_name)
                self.log(f"\nReading: {base_name}")
                self.log(f"Extracted suffix: {suffix_name}")

                try:
                    wb = openpyxl.load_workbook(excel_path, data_only=True)
                except Exception as e:
                    self.log(f"  [Error] Cannot open Excel file: {e}")
                    current_step += len(SHEET_CONFIGS)
                    continue

                sheet_map = {name.strip(): name for name in wb.sheetnames}

                for cfg in SHEET_CONFIGS:
                    current_step += 1
                    pct = (current_step / total_steps) * 100
                    target_key = cfg["target_name"].strip()

                    self.update_status(f"Processing ({current_step}/{total_steps}): {suffix_name} - {target_key}", pct)

                    if target_key not in sheet_map:
                        self.log(f"  [Skipped] Sheet not found: {cfg['target_name']}")
                        continue

                    ws = wb[sheet_map[target_key]]
                    rows_data, col_widths, merged_ranges, cell_objects = get_valid_rows_and_merges(ws, cfg["cols"], target_key)

                    if not rows_data:
                        self.log(f"  [Skipped] Sheet '{cfg['target_name']}' has no content in the selected columns")
                        continue

                    # File name: "<prefix> - <BBAP-xxx...>"
                    out_stem = f"{cfg['pdf_prefix']} - {suffix_name}"
                    out_stem = re.sub(r'[\\/:*?"<>|]', "_", out_stem)

                    docx_path = os.path.join(temp_docx_dir, f"{out_stem}.docx")
                    pdf_path = os.path.join(output_dir, f"{out_stem}.pdf")

                    try:
                        insert_sheet_into_template(
                            template_path=template_path,
                            output_docx_path=docx_path,
                            rows_data=rows_data,
                            col_widths=col_widths,
                            merged_ranges=merged_ranges,
                            cell_objects=cell_objects,
                            font_size=cfg["font_size"],
                            sheet_name=target_key,
                            log=self.log,
                        )
                        convert(docx_path, pdf_path)
                        success_count += 1
                        self.log(f"  [Done] {out_stem}.pdf")
                    except Exception as e:
                        self.log(f"  [Failed] {out_stem}: {e}")

                wb.close()

            if not keep_docx and os.path.isdir(temp_docx_dir):
                for f in glob.glob(os.path.join(temp_docx_dir, "*.docx")):
                    try:
                        os.remove(f)
                    except OSError:
                        pass
                try:
                    os.rmdir(temp_docx_dir)
                except OSError:
                    pass

            self.log("\n" + "=" * 55)
            self.log(f"All done! {success_count} PDF file(s) generated.")
            self.update_status(f"Finished: {success_count} PDF file(s) generated", 100)
            self.root.after(0, lambda: messagebox.showinfo("Finished", f"Batch conversion finished!\n{success_count} PDF file(s) generated."))

        except Exception as e:
            self.log(f"\n[Critical error] {e}")
            self.update_status("An error occurred. Please check the log", 0)
            self.root.after(0, lambda: messagebox.showerror("Error", str(e)))
        finally:
            if pythoncom is not None:
                pythoncom.CoUninitialize()
            self.is_running = False
            self.root.after(0, lambda: self.start_btn.configure(state=tk.NORMAL))


if __name__ == "__main__":
    root = tk.Tk()
    style = ttk.Style()
    if "vista" in style.theme_names():
        style.theme_use("vista")
    app = AccountSummaryConverterApp(root)
    root.mainloop()