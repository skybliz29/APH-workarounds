import tkinter as tk
from tksheet import Sheet
from tkinter import messagebox
import openpyxl, os, re, platform

class SpreadsheetApp:
    def __init__(self, root):
        self.root = root
        self.os_type = platform.system()
        self.root.title(f"Professional Spreadsheet - {self.os_type}")
        self.root.geometry("1100x750")

        # --- Toolbar ---
        top_panel = tk.Frame(self.root, bg="#107c10")
        top_panel.pack(side="top", fill="x")
        
        btn_style = {"bg": "#ffffff", "fg": "#107c10", "relief": "flat", "font": ("Arial", 9, "bold")}
        if self.os_type == "Darwin": btn_style = {"highlightbackground": "#107c10"}
        
        tk.Button(top_panel, text="Save to Excel", command=self.save_to_excel, **btn_style).pack(side="left", padx=10, pady=10)
        tk.Button(top_panel, text="Recalculate All", command=self.force_recalc_all, **btn_style).pack(side="left", padx=5)

        # --- Status Bar ---
        self.status_bar = tk.Frame(self.root, bg="#f3f3f3", bd=1, relief=tk.SUNKEN)
        self.status_bar.pack(side="bottom", fill="x")
        self.stats_label = tk.Label(self.status_bar, text="Sum: 0.00 | Avg: 0.00 | Count: 0", bg="#f3f3f3")
        self.stats_label.pack(side="left", padx=10, pady=5)

        # --- Grid ---
        self.sheet = Sheet(self.root, data=[["" for _ in range(15)] for _ in range(50)])
        self.sheet.enable_bindings("all")
        self.sheet.extra_bindings([("cell_select", self.update_stats), 
                                  ("drag_select", self.update_stats),
                                  ("end_edit_cell", self.on_edit)])
        self.sheet.pack(expand=True, fill="both")

    def update_stats(self, event=None):
        selected = self.sheet.get_selected_cells()
        nums = []
        for r, c in selected:
            val = self.sheet.get_cell_data(r, c)
            try:
                if val is not None and str(val).strip(): nums.append(float(val))
            except: continue
        if nums:
            s, c = sum(nums), len(nums)
            self.stats_label.config(text=f"Sum: {s:,.2f} | Avg: {s/c:,.2f} | Count: {c}")
        else:
            self.stats_label.config(text="Sum: 0.00 | Avg: 0.00 | Count: 0")

    def on_edit(self, event):
        r, c = event[0], event[1]
        val = self.sheet.get_cell_data(r, c)
        if str(val).startswith("="):
            self.root.after(100, lambda: self.calculate(r, c, str(val)))

    def calculate(self, tr, tc, formula):
        try:
            expr = formula[1:].upper().replace(" ", "")
            # Range SUM: =SUM(A1:C1)
            sum_match = re.search(r'SUM\(([A-Z]+)(\d+):([A-Z]+)(\d+)\)', expr)
            if sum_match:
                c1, r1, c2, r2 = sum_match.groups()
                sc, ec, sr, er = ord(c1)-65, ord(c2)-65, int(r1)-1, int(r2)-1
                total = 0.0
                for r in range(min(sr, er), max(sr, er)+1):
                    for c in range(min(sc, ec), max(sc, ec)+1):
                        v = self.sheet.get_cell_data(r, c)
                        try: total += float(v or 0)
                        except: pass
                result = total
            else:
                # Math: =A1/B1
                refs = re.findall(r'[A-Z]+\d+', expr)
                for ref in refs:
                    cl, ri = re.search(r'[A-Z]+', ref).group(), int(re.search(r'\d+', ref).group())-1
                    expr = expr.replace(ref, str(self.sheet.get_cell_data(ri, ord(cl)-65) or 0))
                result = eval(expr, {"__builtins__": None}, {})
            self.sheet.set_cell_data(tr, tc, result)
        except ZeroDivisionError: self.sheet.set_cell_data(tr, tc, "#DIV/0!")
        except: self.sheet.set_cell_data(tr, tc, "#VALUE!")
        self.sheet.refresh()

    def force_recalc_all(self):
        for r, row in enumerate(self.sheet.get_sheet_data()):
            for c, val in enumerate(row):
                if str(val).startswith("="): self.calculate(r, c, str(val))

    def save_to_excel(self):
        wb = openpyxl.Workbook()
        for r_idx, row in enumerate(self.sheet.get_sheet_data()):
            for c_idx, val in enumerate(row):
                wb.active.cell(row=r_idx+1, column=c_idx+1, value=val)
        wb.save("spreadsheet_export.xlsx")
        messagebox.showinfo("Saved", "Exported to spreadsheet_export.xlsx")

if __name__ == "__main__":
    root = tk.Tk()
    app = SpreadsheetApp(root)
    root.mainloop()