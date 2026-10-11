#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
====================================================================
MATRIX SPREADSHEET TO JSON COMPILER v3.7.0 [ZERO-DEPENDENCY STABLE]
====================================================================
Author:       newddf
GitHub Repo:  https://github.com/newddf/excel2json
License:      MIT License
====================================================================
"""

import sys
import os
import json
import webbrowser
import zipfile
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import filedialog, messagebox

# ==========================================
# CORE GLOBAL CONFIGURATION & METADATA
# ==========================================
__AUTHOR__ = "newddf"
__VERSION__ = "v3.7.0"
__GITHUB__ = "https://github.com"

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def get_safe_json_path(source_path, strategy="overwrite"):
    base_path, _ = os.path.splitext(source_path)
    json_path = base_path + ".json"
    if strategy == "overwrite" or not os.path.exists(json_path):
        return json_path
    counter = 1
    while True:
        new_json_path = f"{base_path}_{counter}.json"
        if not os.path.exists(new_json_path):
            return new_json_path
        counter += 1

def parse_xlsx_universal(file_path):
    """Robust universal XML handler for any complex styles or system-generated xlsx"""
    with zipfile.ZipFile(file_path) as z:
        shared_strings = []
        # Find sharedStrings ignoring letter case variations
        ss_file = next((f for f in z.namelist() if f.lower() == "xl/sharedstrings.xml"), None)
        if ss_file:
            ss_root = ET.fromstring(z.read(ss_file))
            # Extract texts supporting both <t> and inline formatting trees
            for si in ss_root.findall(".//{*}si"):
                t_elements = si.findall(".//{*}t")
                text_val = "".join([t.text for t in t_elements if t.text])
                shared_strings.append(text_val)

        # Find the first primary worksheet data stream
        sheet_file = next((f for f in z.namelist() if "xl/worksheets/sheet" in f.lower()), None)
        if not sheet_file:
            return []
            
        sheet_root = ET.fromstring(z.read(sheet_file))
        rows_data = []
        
        for row_node in sheet_root.findall(".//{*}row"):
            row_cells = []
            for cell_node in row_node.findall("{*}c"):
                val_node = cell_node.find("{*}v")
                val = val_node.text if val_node is not None else None
                
                # Check formatting types supporting fallback attributes
                cell_type = cell_node.get("t")
                if cell_type == "s" and val is not None:
                    try:
                        val = shared_strings[int(val)]
                    except (IndexError, ValueError):
                        pass
                elif cell_node.find("{*}is/{*}t") is not None:
                    val = cell_node.find("{*}is/{*}t").text
                    
                row_cells.append(val)
                
            if any(v is not None for v in row_cells):
                rows_data.append(row_cells)
        return rows_data

def parse_ods_universal(file_path):
    """Robust universal XML handler for LibreOffice Calc sheets"""
    with zipfile.ZipFile(file_path) as z:
        content_root = ET.fromstring(z.read("content.xml"))
        table_node = content_root.find(".//{*}table")
        if table_node is None:
            return []
            
        rows_data = []
        for row_node in table_node.findall("{*}table-row"):
            row_cells = []
            for cell_node in row_node.findall("{*}table-cell"):
                # Support nested text paragraph elements inside LibreOffice cells
                p_nodes = cell_node.findall(".//{*}p")
                val = "".join([p.text for p in p_nodes if p.text]) if p_nodes else None
                row_cells.append(val)
                
            if any(v is not None and str(v).strip() != "" for v in row_cells):
                rows_data.append(row_cells)
        return rows_data

def convert_spreadsheet_to_json(source_path, json_path):
    try:
        _, ext = os.path.splitext(source_path.lower())
        if ext in ['.xlsx', '.xls']:
            raw_matrix = parse_xlsx_universal(source_path)
        elif ext == '.ods':
            raw_matrix = parse_ods_universal(source_path)
        else:
            return False, f"Unsupported matrix payload signature: {ext}"
            
        if not raw_matrix:
            return False, "Target sheet active data block is completely blank."
            
        headers = raw_matrix[0]
        data_dict = []
        
        for row in raw_matrix[1:]:
            row_data = {}
            for i, header in enumerate(headers):
                if header is not None and str(header).strip() != "":
                    val = row[i] if i < len(row) else None
                    row_data[__builtins__.str(header)] = None if (val is None or __builtins__.str(val).strip() == "") else val
            data_dict.append(row_data)
            
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(data_dict, f, ensure_ascii=False, indent=4)
        return True, json_path
    except Exception as e:
        return False, __builtins__.str(e)

class CyberpunkApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"MATRIX DISK TRANSLATOR CORE {__VERSION__}")
        self.root.geometry("580x420")
        self.root.configure(bg="#0D1117")
        self.root.resizable(False, False)
        
        ico_path = resource_path("logo.ico")
        if os.path.exists(ico_path):
            try:
                self.root.iconbitmap(ico_path)
            except Exception:
                pass
                
        self.bg_dark = "#0D1117"
        self.bg_panel = "#161B22"
        self.fg_cyan = "#00F0FF"
        self.fg_green = "#00FF66"
        self.fg_gray = "#8B949E"
        
        self.logo_canvas = tk.Canvas(root, width=120, height=120, bg=self.bg_dark, bd=0, highlightthickness=0)
        self.logo_canvas.pack(pady=(15, 2))
        self.logo_canvas.create_rectangle(10, 10, 110, 110, fill=self.bg_panel, outline="", width=0)
        
        polygons = [
            [49.26, 75.82, 34.04, 84.6, 34.04, 74.38, 20.83, 92.11, 34.49, 100, 51.33, 90.28, 62.29, 36.46, 49.26, 53.95],
            [89.16, 78.03, 89.16, 85.41, 77.67, 78.81, 77.67, 62.68, 67.96, 49.72, 67.96, 84.38, 98.26, 90.18, 95.21, 86.11],
            [89.16, 47.55, 89.16, 76.53, 95.99, 85.66, 98.5, 89.01, 88.39, 36.74, 68.17, 48.51, 77.67, 61.18, 77.67, 54.24],
            [98.04, 91.05, 70.05, 85.69, 88.71, 96.41],
            [98.86, 86.17, 98.86, 42.39, 96.37, 40.96, 89.34, 36.91],
            [20.61, 87.56, 33.86, 18.71, 20.61, 26.36],
            [34.04, 72.88, 34.04, 33.74, 49.26, 42.64, 49.26, 52.46, 55.17, 44.52, 62.47, 34.72, 34.81, 18.54, 20.97, 90.42],
            [62.68, 39.03, 52.38, 89.62, 62.68, 83.64]
        ]
        for poly in polygons:
            self.logo_canvas.create_polygon(poly, fill=self.fg_cyan, outline="")
            
        self.title_label = tk.Label(root, text="MATRIX COMPILER SYSTEM", font=("Consolas", 12, "bold"), fg=self.fg_cyan, bg=self.bg_dark)
        self.title_label.pack()
        
        self.main_frame = tk.Frame(root, bg=self.bg_panel, bd=1, relief=tk.SOLID, highlightbackground=self.fg_cyan, highlightthickness=1)
        self.main_frame.pack(pady=(10, 5), padx=25, fill=tk.BOTH, expand=True)
        
        self.path_var = tk.StringVar(value="[ STANDBY: WAITING FOR SOURCE DATA PATH... ]")
        self.path_entry = tk.Entry(self.main_frame, textvariable=self.path_var, font=("Consolas", 9), bg=self.bg_dark, fg=self.fg_gray, insertbackground=self.fg_cyan, bd=0, highlightthickness=1, highlightbackground="#30363D", state="readonly", justify="center")
        self.path_entry.pack(pady=(15, 10), padx=20, fill=tk.X, ipady=4)
        
        self.strategy_var = tk.StringVar(value="ask")
        self.radio_frame = tk.Frame(self.main_frame, bg=self.bg_panel)
        self.radio_frame.pack(pady=5)
        radio_opts = [("ASK ME", "ask"), ("OVERWRITE", "overwrite"), ("AUTO RENAME", "rename")]
        for text, val in radio_opts:
            rb = tk.Radiobutton(self.radio_frame, text=text, variable=self.strategy_var, value=val, bg=self.bg_panel, fg=self.fg_gray, activebackground=self.bg_panel, activeforeground=self.fg_cyan, selectcolor=self.bg_dark, font=("Consolas", 9, "bold"))
            rb.pack(side=tk.LEFT, padx=12)
            
        self.btn_frame = tk.Frame(self.main_frame, bg=self.bg_panel)
        self.btn_frame.pack(pady=15)
        self.btn_browse = tk.Button(self.btn_frame, text="LOAD FILE", font=("Consolas", 10, "bold"), bg=self.bg_dark, fg=self.fg_cyan, activebackground=self.fg_cyan, activeforeground=self.bg_dark, bd=1, relief=tk.SOLID, width=12, cursor="hand2", command=self.browse_file)
        self.btn_browse.pack(side=tk.LEFT, padx=15)
        self.btn_convert = tk.Button(self.btn_frame, text="EXECUTE", font=("Consolas", 10, "bold"), bg=self.bg_dark, fg="#404040", state="disabled", bd=1, relief=tk.SOLID, width=12, command=self.start_convert)
        self.btn_convert.pack(side=tk.LEFT, padx=15)
        
        self.dev_frame = tk.Frame(root, bg=self.bg_dark)
        self.dev_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=15, pady=(2, 8))
        self.dev_label = tk.Label(self.dev_frame, text=f"DEV: {__AUTHOR__} | {__VERSION__} | GET SOURCE", font=("Consolas", 8), fg="#30363D", bg=self.bg_dark, cursor="hand2")
        self.dev_label.pack(side=tk.RIGHT)
        self.dev_label.bind("<Button-1>", lambda e: webbrowser.open(__GITHUB__))
        self.dev_label.bind("<Enter>", lambda e: self.dev_label.configure(fg=self.fg_cyan))
        self.dev_label.bind("<Leave>", lambda e: self.dev_label.configure(fg="#30363D"))
        
        self.status_var = tk.StringVar(value="SYSTEM READY")
        self.status_bar = tk.Label(self.dev_frame, textvariable=self.status_var, font=("Consolas", 9), bg=self.bg_dark, fg=self.fg_green, anchor=tk.W)
        self.status_bar.pack(side=tk.LEFT)

    def browse_file(self):
        file_path = filedialog.askopenfilename(title="SELECT SOURCE DATA STREAM", filetypes=[("Spreadsheets Data", "*.xlsx *.xls *.ods")])
        if file_path:
            self.path_var.set(file_path)
            self.path_entry.configure(fg=self.fg_cyan)
            self.btn_convert.configure(state="normal", fg=self.fg_green, cursor="hand2")
            self.status_var.set("DATA STREAM CONNECTED.")

    def start_convert(self):
        excel_path = self.path_var.get()
        strategy = self.strategy_var.get()
        base_path, _ = os.path.splitext(excel_path)
        default_json_path = base_path + ".json"
        
        if os.path.exists(default_json_path):
            if strategy == "ask":
                response = messagebox.askyesnocancel("⚠️ FILE COLLISION DETECTED", f"Target file '{os.path.basename(default_json_path)}' already exists.\n\n[Yes] -> Overwrite\n[No] -> Auto-Rename\n[Cancel] -> Abort Transaction")
                if response is True: final_json_path = default_json_path
                elif response is False: final_json_path = get_safe_json_path(excel_path, strategy="rename")
                else:
                    self.status_var.set("TRANSACTION ABORTED.")
                    return
            else:
                final_json_path = get_safe_json_path(excel_path, strategy=strategy)
        else:
            final_json_path = default_json_path
            
        self.status_var.set("INJECTING AND PARSING CORRUPTED BLOCKS...")
        self.root.update_idletasks()
        
        success, result = convert_spreadsheet_to_json(excel_path, final_json_path)
        if success:
            self.status_var.set("✨ TRANSLATION COMPILED.")
            messagebox.showinfo("TRANSACTION SUCCESSFUL", f"Payload compiled perfectly into:\n{os.path.basename(result)}")
        else:
            self.status_var.set("❌ CRITICAL ERROR INJECTED.")
            messagebox.showerror("TRANSACTION FAILED", f"Parser breakdown:\n{result}")

if __name__ == "__main__":
    root = tk.Tk()
    app = CyberpunkApp(root)
    root.mainloop()
