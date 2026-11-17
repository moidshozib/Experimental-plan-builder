import datetime as datetime
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk

DEFAULT_COLUMNS = [
    "Run",
    "Sample",
    "Material",
    "Condition",
    "Temperature",
    "Time",
    "Concentration",
    "Notes",
]


def export_plan(path, columns, rows, details):
    """Write the plan and its metadata to a formatted XLSX workbook."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.worksheet.table import Table, TableStyleInfo
    from openpyxl.utils import get_column_letter

    workbook = Workbook()
    plan = workbook.active
    plan.title = "Experimental Plan"
    plan.append(columns)
    for row in rows:
        plan.append(row)

    header_fill = PatternFill("solid", fgColor="1F4E78")
    for cell in plan[1]:
        cell.font = Font(color="FFFFFF", bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for index, name in enumerate(columns, start=1):
        plan.column_dimensions[get_column_letter(index)].width = max(14, min(32, len(name) + 4))
    plan.freeze_panes = "A2"
    plan.auto_filter.ref = plan.dimensions
    if rows:
        table = Table(displayName="ExperimentalPlan", ref=plan.dimensions)
        table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        plan.add_table(table)

    details_sheet = workbook.create_sheet("Plan Details")
    details_sheet.append(["Field", "Value"])
    for key, value in details.items():
        details_sheet.append([key.replace("_", " ").title(), value])
    details_sheet.column_dimensions["A"].width = 20
    details_sheet.column_dimensions["B"].width = 42
    for cell in details_sheet[1]:
        cell.font = Font(color="FFFFFF", bold=True)
        cell.fill = header_fill

    workbook.save(path)


def import_plan(path):
    """Read the active worksheet as column headings plus plan rows."""
    from openpyxl import load_workbook

    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = workbook.active
        values = list(sheet.iter_rows(values_only=True))
    finally:
        workbook.close()
    if not values:
        raise ValueError("The selected workbook is empty.")
    columns = [str(value or f"Column {index + 1}") for index, value in enumerate(values[0])]
    rows = [["" if value is None else str(value) for value in row] for row in values[1:]]
    return columns, rows


class ExperimentalPlanApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Experimental Plan Builder")
        self.geometry("1180x700")
        self.minsize(850, 480)
        self.columns = list(DEFAULT_COLUMNS)
        self.rows = [[]]
        self.entries = {}
        self.details = {
            "plan_name": "",
            "researcher": "",
            "project": "",
            "created": datetime.date.today().isoformat(),
        }
        self._configure_style()
        self._build_interface()
        self._draw_grid()

    def _configure_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Title.TLabel", font=("Segoe UI", 18, "bold"))
        style.configure("Secondary.TLabel", foreground="#52606D")
        style.configure("Header.TLabel", font=("Segoe UI", 10, "bold"))
        style.configure("Export.TButton", font=("Segoe UI", 10, "bold"))

    def _build_interface(self):
        header = ttk.Frame(self, padding=(18, 16, 18, 8))
        header.pack(fill="x")
        ttk.Label(header, text="Experimental Plan Builder", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            header,
            text="Create a run sheet and export it as a formatted Excel workbook.",
            style="Secondary.TLabel",
        ).pack(anchor="w", pady=(2, 12))

        details = ttk.LabelFrame(self, text="Plan details", padding=10)
        details.pack(fill="x", padx=18, pady=(0, 10))
        self.detail_vars = {}
        fields = [("Plan name", "plan_name"), ("Researcher", "researcher"), ("Project", "project"), ("Created", "created")]
        for index, (label, key) in enumerate(fields):
            ttk.Label(details, text=label).grid(row=0, column=index * 2, sticky="w", padx=(0, 6))
            variable = tk.StringVar(value=self.details[key])
            self.detail_vars[key] = variable
            ttk.Entry(details, textvariable=variable, width=24).grid(
                row=0, column=index * 2 + 1, sticky="ew", padx=(0, 14)
            )
            details.columnconfigure(index * 2 + 1, weight=1)

        toolbar = ttk.Frame(self, padding=(18, 0, 18, 8))
        toolbar.pack(fill="x")
        ttk.Button(toolbar, text="+ Add run", command=self.add_run).pack(side="left")
        ttk.Button(toolbar, text="Remove last run", command=self.remove_last_run).pack(side="left", padx=6)
        ttk.Button(toolbar, text="Add column", command=self.add_column).pack(side="left")
        ttk.Button(toolbar, text="Remove column", command=self.remove_column).pack(side="left", padx=6)
        ttk.Button(toolbar, text="Import workbook", command=self.load_workbook).pack(side="right")
        ttk.Button(toolbar, text="Clear plan", command=self.clear_plan).pack(side="right", padx=6)

        frame = ttk.Frame(self, padding=(18, 0, 18, 10))
        frame.pack(fill="both", expand=True)
        self.canvas = tk.Canvas(frame, background="#F7F9FB", highlightthickness=0)
        vertical = ttk.Scrollbar(frame, orient="vertical", command=self.canvas.yview)
        horizontal = ttk.Scrollbar(frame, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        self.grid_frame = ttk.Frame(self.canvas, padding=8)
        self.canvas.create_window((0, 0), window=self.grid_frame, anchor="nw")
        self.grid_frame.bind("<Configure>", lambda _event: self.canvas.configure(scrollregion=self.canvas.bbox("all")))

        footer = ttk.Frame(self, padding=(18, 0, 18, 16))
        footer.pack(fill="x")
        self.status = tk.StringVar()
        ttk.Label(footer, textvariable=self.status, style="Secondary.TLabel").pack(side="left")
        ttk.Button(footer, text="Export XLSX", style="Export.TButton", command=self.save_workbook).pack(side="right")

    def _read_entries(self):
        values = []
        for row_index in range(len(self.rows)):
            values.append([self.entries[(row_index, column_index)].get() for column_index in range(len(self.columns))])
        return values

    def _draw_grid(self):
        for child in self.grid_frame.winfo_children():
            child.destroy()
        self.entries = {}
        ttk.Label(self.grid_frame, text="#", style="Header.TLabel", width=5).grid(row=0, column=0, padx=3, pady=3)
        for column_index, column in enumerate(self.columns, start=1):
            ttk.Label(self.grid_frame, text=column, style="Header.TLabel", width=17).grid(
                row=0, column=column_index, padx=3, pady=3, sticky="ew"
            )
        for row_index, row in enumerate(self.rows):
            ttk.Label(self.grid_frame, text=str(row_index + 1), width=5).grid(row=row_index + 1, column=0, padx=3, pady=2)
            for column_index in range(len(self.columns)):
                entry = ttk.Entry(self.grid_frame, width=17)
                entry.grid(row=row_index + 1, column=column_index + 1, padx=3, pady=2, sticky="ew")
                if column_index < len(row):
                    entry.insert(0, row[column_index])
                self.entries[(row_index, column_index)] = entry
        if self.rows and not self.entries[(0, 0)].get():
            self.entries[(0, 0)].insert(0, "1")
        count = len(self.rows)
        self.status.set(f"{count} run" + ("s" if count != 1 else ""))
        self.after_idle(lambda: self.canvas.configure(scrollregion=self.canvas.bbox("all")))

    def add_run(self):
        self.rows = self._read_entries()
        self.rows.append([""] * len(self.columns))
        self._draw_grid()

    def remove_last_run(self):
        if len(self.rows) > 1:
            self.rows = self._read_entries()[:-1]
            self._draw_grid()

    def add_column(self):
        name = simpledialog.askstring("Add column", "Column name:", parent=self)
        if name and name.strip() and name.strip() not in self.columns:
            self.rows = self._read_entries()
            self.columns.append(name.strip())
            self._draw_grid()

    def remove_column(self):
        if len(self.columns) == 1:
            messagebox.showwarning("Cannot remove column", "A plan needs at least one column.")
            return
        name = simpledialog.askstring("Remove column", "Type the exact column name:", parent=self)
        if name in self.columns:
            self.rows = self._read_entries()
            index = self.columns.index(name)
            self.columns.pop(index)
            for row in self.rows:
                row.pop(index)
            self._draw_grid()

    def clear_plan(self):
        if messagebox.askyesno("Clear plan", "Remove all runs and start over?", parent=self):
            self.rows = [[""] * len(self.columns)]
            self._draw_grid()

    def load_workbook(self):
        path = filedialog.askopenfilename(parent=self, filetypes=[("Excel workbook", "*.xlsx *.xlsm")])
        if not path:
            return
        try:
            self.columns, imported_rows = import_plan(path)
            self.rows = imported_rows or [[""] * len(self.columns)]
            self._draw_grid()
            messagebox.showinfo("Workbook imported", f"Loaded {Path(path).name}.", parent=self)
        except Exception as error:
            messagebox.showerror("Import failed", str(error), parent=self)

    def save_workbook(self):
        path = filedialog.asksaveasfilename(
            parent=self,
            defaultextension=".xlsx",
            initialfile="experimental_plan.xlsx",
            filetypes=[("Excel workbook", "*.xlsx")],
        )
        if not path:
            return
        try:
            details = {key: variable.get().strip() for key, variable in self.detail_vars.items()}
            export_plan(path, self.columns, self._read_entries(), details)
            messagebox.showinfo("Export complete", f"Saved {Path(path).name}.", parent=self)
        except ImportError:
            messagebox.showerror("Missing dependency", "Install dependencies with: python -m pip install -r requirements.txt", parent=self)
        except Exception as error:
            messagebox.showerror("Export failed", str(error), parent=self)


class MultiSheetPlanApp(tk.Tk):
    """Edit a multi-sheet plan while preserving the workbook as an XLSX template."""

    TEMPLATE_NAME = "Copy of MA_PPDT2FBT_PCBM.xlsx"

    def __init__(self):
        super().__init__()
        self.title("Experimental Plan Builder")
        self.geometry("1280x780")
        self.minsize(900, 560)
        self.workbook = None
        self.source_path = None
        self.sheet_entries = {}
        self.selection_anchor = None
        self.selected_range = None
        ttk.Style(self).configure("MergeSelection.TEntry", fieldbackground="#DCEBFA")
        self._build_multi_sheet_interface()
        template = Path(__file__).with_name(self.TEMPLATE_NAME)
        if template.exists():
            self._load_path(template)
        else:
            self._create_blank_workbook()

    def _build_multi_sheet_interface(self):
        header = ttk.Frame(self, padding=(18, 16, 18, 8))
        header.pack(fill="x")
        ttk.Label(header, text="Experimental Plan Builder", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            header,
            text="Edit the five planning sheets and export the complete experimental plan.",
            style="Secondary.TLabel",
        ).pack(anchor="w", pady=(2, 12))

        toolbar = ttk.Frame(self, padding=(18, 0, 18, 8))
        toolbar.pack(fill="x")
        ttk.Button(toolbar, text="Open workbook", command=self.open_workbook).pack(side="left")
        ttk.Button(toolbar, text="Reload template", command=self.reload_template).pack(side="left", padx=6)
        ttk.Button(toolbar, text="Merge cells", command=self.merge_selected_cells).pack(side="left", padx=6)
        ttk.Label(toolbar, text="Click a cell, then Shift-click the last cell to select a range.").pack(
            side="left", padx=6
        )
        ttk.Button(toolbar, text="Export XLSX", style="Export.TButton", command=self.export_workbook).pack(side="right")
        self.multi_status = tk.StringVar()
        ttk.Label(toolbar, textvariable=self.multi_status, style="Secondary.TLabel").pack(side="right", padx=12)

        content = ttk.Frame(self, padding=(18, 0, 18, 16))
        content.pack(fill="both", expand=True)
        self.notebook = ttk.Notebook(content)
        self.notebook.pack(fill="both", expand=True)
        self.notebook.bind("<<NotebookTabChanged>>", self._clear_selection)

    def _create_blank_workbook(self):
        from openpyxl import Workbook

        self.workbook = Workbook()
        self.workbook.active.title = "Ingredients"
        for name in ["Sample list", "Characterization", "Evaluation and Analysis", "Tasks & People involved"]:
            self.workbook.create_sheet(name)
        self._render_workbook()

    def _load_path(self, path):
        from openpyxl import load_workbook

        try:
            workbook = load_workbook(path)
            template = Path(__file__).with_name(self.TEMPLATE_NAME)
            if Path(path).resolve() == template.resolve():
                sheet = workbook["Ingredients"]
                title_range = "D1:H1"
                if title_range not in {str(merged) for merged in sheet.merged_cells.ranges}:
                    sheet.merge_cells(title_range)
                for cell_range in ("E69:E70", "G69:G70"):
                    if cell_range in {str(merged) for merged in sheet.merged_cells.ranges}:
                        sheet.unmerge_cells(cell_range)
            self.workbook = workbook
            self.source_path = Path(path)
            self._render_workbook()
        except Exception as error:
            messagebox.showerror("Workbook could not be opened", str(error), parent=self)

    def _render_workbook(self):
        for tab in self.notebook.tabs():
            self.notebook.nametowidget(tab).destroy()
        self.sheet_entries = {}
        self.selection_anchor = None
        self.selected_range = None
        for sheet in self.workbook.worksheets:
            self._render_sheet(sheet)
        self.multi_status.set(f"{len(self.workbook.worksheets)} sheets loaded")

    def _render_sheet(self, sheet):
        outer = ttk.Frame(self.notebook)
        outer.rowconfigure(0, weight=1)
        outer.columnconfigure(0, weight=1)
        canvas = tk.Canvas(outer, background="#F7F9FB", highlightthickness=0)
        vertical = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        horizontal = ttk.Scrollbar(outer, orient="horizontal", command=canvas.xview)
        canvas.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")
        grid = ttk.Frame(canvas, padding=10)
        canvas.create_window((0, 0), window=grid, anchor="nw")
        grid.bind("<Configure>", lambda _event: canvas.configure(scrollregion=canvas.bbox("all")))

        entries = {}
        self.sheet_entries[sheet.title] = entries
        merged_ranges = {}
        merged_cells = set()
        for merged in sheet.merged_cells.ranges:
            merged_ranges[(merged.min_row, merged.min_col)] = (
                merged.max_row - merged.min_row + 1,
                merged.max_col - merged.min_col + 1,
            )
            for row_index in range(merged.min_row, merged.max_row + 1):
                for column_index in range(merged.min_col, merged.max_col + 1):
                    merged_cells.add((row_index, column_index))
        for row_index in range(1, sheet.max_row + 1):
            ttk.Label(grid, text=str(row_index), width=5).grid(row=row_index, column=0, padx=2, pady=1)
            for column_index in range(1, sheet.max_column + 1):
                cell = sheet.cell(row_index, column_index)
                coordinate = (row_index, column_index)
                if coordinate in merged_ranges:
                    rowspan, columnspan = merged_ranges[coordinate]
                    widget = ttk.Entry(grid, width=18)
                    widget.insert(0, "" if cell.value is None else str(cell.value))
                    widget.grid(
                        row=row_index,
                        column=column_index,
                        rowspan=rowspan,
                        columnspan=columnspan,
                        padx=2,
                        pady=1,
                        sticky="ew",
                    )
                    entries[(row_index, column_index)] = widget
                elif coordinate not in merged_cells:
                    widget = ttk.Entry(grid, width=18)
                    widget.insert(0, "" if cell.value is None else str(cell.value))
                    widget.grid(row=row_index, column=column_index, padx=2, pady=1, sticky="ew")
                    entries[(row_index, column_index)] = widget
                else:
                    continue
                widget.bind(
                    "<Button-1>",
                    lambda event, sheet_name=sheet.title, cell_coordinate=coordinate: self._select_cell(
                        sheet_name, cell_coordinate, bool(event.state & 0x0001)
                    ),
                )
        self.notebook.add(outer, text=sheet.title)

    def _select_cell(self, sheet_name, coordinate, extend_selection):
        if extend_selection and self.selection_anchor and self.selection_anchor[0] == sheet_name:
            start_row, start_column = self.selection_anchor[1]
            end_row, end_column = coordinate
            self.selected_range = (
                sheet_name,
                min(start_row, end_row),
                min(start_column, end_column),
                max(start_row, end_row),
                max(start_column, end_column),
            )
        else:
            self.selection_anchor = (sheet_name, coordinate)
            self.selected_range = (sheet_name, coordinate[0], coordinate[1], coordinate[0], coordinate[1])

        _, min_row, min_column, max_row, max_column = self.selected_range
        for current_sheet, entries in self.sheet_entries.items():
            for (row_index, column_index), widget in entries.items():
                selected = (
                    current_sheet == sheet_name
                    and min_row <= row_index <= max_row
                    and min_column <= column_index <= max_column
                )
                widget.configure(style="MergeSelection.TEntry" if selected else "TEntry")

    def _clear_selection(self, _event=None):
        self.selection_anchor = None
        self.selected_range = None
        for entries in self.sheet_entries.values():
            for widget in entries.values():
                widget.configure(style="TEntry")

    def merge_selected_cells(self):
        if not self.selected_range:
            messagebox.showwarning(
                "Select cells first",
                "Click a cell, then Shift-click the last cell in the range you want to merge.",
                parent=self,
            )
            return

        sheet_name, min_row, min_column, max_row, max_column = self.selected_range
        if min_row == max_row and min_column == max_column:
            messagebox.showwarning(
                "Select multiple cells",
                "Select at least two adjacent cells before merging.",
                parent=self,
            )
            return

        sheet = self.workbook[sheet_name]
        for merged in sheet.merged_cells.ranges:
            overlaps = not (
                merged.max_row < min_row
                or merged.min_row > max_row
                or merged.max_col < min_column
                or merged.min_col > max_column
            )
            if overlaps:
                messagebox.showwarning(
                    "Range contains merged cells",
                    f"The selected range overlaps the existing merge {merged}. Choose a range without merged cells.",
                    parent=self,
                )
                return

        self._write_entries()
        other_values = [
            str(sheet.cell(row_index, column_index).value)
            for row_index in range(min_row, max_row + 1)
            for column_index in range(min_column, max_column + 1)
            if (row_index, column_index) != (min_row, min_column)
            and sheet.cell(row_index, column_index).value not in (None, "")
        ]
        if other_values and not messagebox.askyesno(
            "Merge cells",
            "Merging keeps the value in the top-left cell and removes the contents of the other selected cells. Continue?",
            parent=self,
        ):
            return

        sheet.merge_cells(
            start_row=min_row,
            start_column=min_column,
            end_row=max_row,
            end_column=max_column,
        )
        sheet_index = self.workbook.sheetnames.index(sheet_name)
        self._render_workbook()
        self.notebook.select(self.notebook.tabs()[sheet_index])

    def _write_entries(self):
        for sheet in self.workbook.worksheets:
            for (row_index, column_index), widget in self.sheet_entries.get(sheet.title, {}).items():
                sheet.cell(row_index, column_index).value = widget.get()

    def open_workbook(self):
        path = filedialog.askopenfilename(
            parent=self,
            filetypes=[("Excel workbook", "*.xlsx *.xlsm")],
        )
        if path:
            self._load_path(path)

    def reload_template(self):
        template = Path(__file__).with_name(self.TEMPLATE_NAME)
        if template.exists():
            self._load_path(template)
        else:
            messagebox.showwarning("Template not found", f"Place {self.TEMPLATE_NAME} beside main.py.", parent=self)

    def export_workbook(self):
        path = filedialog.asksaveasfilename(
            parent=self,
            defaultextension=".xlsx",
            initialfile="experimental_plan.xlsx",
            filetypes=[("Excel workbook", "*.xlsx")],
        )
        if not path:
            return
        try:
            self._write_entries()
            self.workbook.save(path)
            self.multi_status.set(f"Saved {Path(path).name}")
            messagebox.showinfo("Export complete", f"Saved {Path(path).name} with {len(self.workbook.worksheets)} sheets.", parent=self)
        except Exception as error:
            messagebox.showerror("Export failed", str(error), parent=self)


if __name__ == "__main__":
    MultiSheetPlanApp().mainloop()
