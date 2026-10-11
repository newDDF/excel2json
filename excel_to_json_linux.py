#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
====================================================================
MATRIX SPREADSHEET TO JSON COMPILER v3.7.0 [LINUX CLI EDITION]
====================================================================
Author:       newddf
GitHub Repo:  https://github.com/newddf/excel2json
License:      MIT License
====================================================================
"""

import sys
import os
import json
import zipfile
import xml.etree.ElementTree as ET

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
    with zipfile.ZipFile(file_path) as z:
        shared_strings = []
        ss_file = next((f for f in z.namelist() if f.lower() == "xl/sharedstrings.xml"), None)
        if ss_file:
            ss_root = ET.fromstring(z.read(ss_file))
            for si in ss_root.findall(".//{*}si"):
                t_elements = si.findall(".//{*}t")
                text_val = "".join([t.text for t in t_elements if t.text])
                shared_strings.append(text_val)

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
    with zipfile.ZipFile(file_path) as z:
        content_root = ET.fromstring(z.read("content.xml"))
        table_node = content_root.find(".//{*}table")
        if table_node is None:
            return []
        rows_data = []
        for row_node in table_node.findall("{*}table-row"):
            row_cells = []
            for cell_node in row_node.findall("{*}table-cell"):
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
            print(f"❌ [ERROR] Unsupported extension: {ext}")
            return False
            
        if not raw_matrix:
            print("❌ [ERROR] Sheet active data block is completely blank.")
            return False
            
        headers = raw_matrix
        data_dict = []
        for row in raw_matrix[1:]:
            row_data = {}
            for i, header in enumerate(headers):
                if header is not None and str(header).strip() != "":
                    val = row[i] if i < len(row) else None
                    row_data[str(header)] = None if (val is None or str(val).strip() == "") else val
            data_dict.append(row_data)
            
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(data_dict, f, ensure_ascii=False, indent=4)
        print(f"✨ [SUCCESS] Compiled safely -> {json_path}")
        return True
    except Exception as e:
        print(f"❌ [CRITICAL PARSER BREAKDOWN] {e}")
        return False

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("\n=======================================================")
        print("📊 MATRIX SPREADSHEET TO JSON LINUX CORE")
        print("=======================================================")
        print("Usage:   ./Matrix_Converter <file1.xlsx> <file2.ods> [--overwrite/--rename]")
        print("Options: --overwrite : Directly overwrite target files if exist (Default)")
        print("         --rename    : Auto rename collision targets with suffixes\n")
        sys.exit(1)
        
    strategy = "overwrite"
    files_to_process = []
    
    for arg in sys.argv[1:]:
        if arg == "--overwrite":
            strategy = "overwrite"
        elif arg == "--rename":
            strategy = "rename"
        else:
            files_to_process.append(arg)
            
    for file_path in files_to_process:
        if not os.path.exists(file_path):
            print(f"❌ [FILE NOT FOUND] Skipping paths -> {file_path}")
            continue
        print(f"⚡ [PARSING TRANSACTION INJECTED] Processing: {os.path.basename(file_path)}")
        final_json_path = get_safe_json_path(file_path, strategy=strategy)
        convert_spreadsheet_to_json(file_path, final_json_path)
