#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Scania Spec Index (ttkinter — meanings, dynamic 3113..3148, responsive, search)

CLI options (similar spirit to Full SOPS Editor):
  --open PATH           Path to an XML to work with
  --year YEAR           Production year to store for the XML (used with --add)
  --add                 Non-interactive add: copy XML into store + append to CSV
  --analyze             After start, refresh the table and attempt to select the row
  --set-editor PATH     Persist the editor script path to config.json (for "Open in Editor")

Examples (Windows):
  py -3 Scania_Spec_Index_CLI_ready.py --open C:\data\truck.xml --year 2017 --add --analyze
  py -3 Scania_Spec_Index_CLI_ready.py --set-editor "C:\\Users\\Kstore\\Documents\\GitHub\\scania_SOPS\\Full_SOPS_Editor_v8_14.py"
"""

import os
import csv
import shutil
import time
import sys
import subprocess
import json
import argparse
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import xml.etree.ElementTree as ET

APP_TITLE = "Scania Spec Index (ttkinter — meanings, dynamic 3113..3148, responsive, search)"

# Base paths
BASE_DIR = os.path.join(os.path.expanduser("~"), r"Documents\GitHub\scania_SOPS\classifier")
CSV_PATH = os.path.join(BASE_DIR, "spec_index.csv")
XML_STORE_DIR = os.path.join(BASE_DIR, "stored_xml")
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")  # store editor path

# Mapping candidates
MAPPING_CANDIDATES = [
    r"C:\Users\Kstore\Documents\GitHub\scania_SOPS\sops_fpc_mapping.csv",
    os.path.join(os.path.expanduser("~"), r"Documents\GitHub\scania_SOPS\sops_fpc_mapping.csv"),
    os.path.join(os.path.expanduser("~"), r"Documents\GitHub\scania_SOPS\classifier\sops_fpc_mapping.csv"),
    os.path.join(os.path.dirname(__file__), "sops_fpc_mapping.csv"),
]

# Remove any columns whose name contains this (case-insensitive)
FORBIDDEN_SUBSTR = "environment"

# Explicitly excluded column names (case-insensitive)
EXCLUDED_COLUMNS = {"engine management system"}  # dropped entirely

# Fixed fields (Descriptions only). Order defines column order in table/CSV.
FIXED_FIELDS = [
    ("448",   "Axles"),
    ("3127",  "EngineECU"),
    ("408",   "Engine Version"),
    ("142",   "EngineSize"),

    # Emissions & related additions
    ("2344",  "Immobiliser"),
    ("2409",  "EGR"),
    ("4306",  "NOx sensor"),
    ("4280",  "Exhaust Emission Control"),
    ("2471",  "Emission level"),

    # COO + Gearbox cluster
    ("3113",  "COO"),
    ("3129",  "GearboxECU"),  # if code == 'Z' -> 'GMan'
    ("17",    "Gearbox Type"),
]
FIXED_IDS = {fid for fid, _ in FIXED_FIELDS}

# Dynamic FPC range (short-named columns, dedup by label, excluding forbidden names)
DYN_START = 3113
DYN_END   = 3148
DYN_RANGE = [str(i) for i in range(DYN_START, DYN_END + 1)]

# Mapping stores
LONG_BY_ID   = {}
SHORT_BY_ID  = {}
DESC_BY_PAIR = {}
MAPPING_PATH = ""

# ---------- Helpers ----------
def ensure_dirs():
    os.makedirs(BASE_DIR, exist_ok=True)
    os.makedirs(XML_STORE_DIR, exist_ok=True)

def local_name(tag: str) -> str:
    if "}" in tag:
        return tag.split("}", 1)[1]
    return tag

def _clean(s: str) -> str:
    return (s or "").strip()

def _is_intlike(s: str) -> bool:
    try:
        int(s.strip())
        return True
    except Exception:
        return False

def unique_store_path(original_path: str) -> str:
    base = os.path.basename(original_path)
    name, ext = os.path.splitext(base)
    dest = os.path.join(XML_STORE_DIR, base)
    if not os.path.exists(dest):
        return dest
    suffix = time.strftime("%Y%m%d_%H%M%S")
    dest = os.path.join(XML_STORE_DIR, f"{name}_{suffix}{ext}")
    c = 1
    while os.path.exists(dest):
        dest = os.path.join(XML_STORE_DIR, f"{name}_{suffix}_{c}{ext}")
        c += 1
    return dest

# ---- Editor discovery ----
def _sops_root() -> str:
    return os.path.join(os.path.expanduser("~"), "Documents", "GitHub", "scania_SOPS")

def find_editor_script() -> str:
    env = os.environ.get("FULL_SOPS_EDITOR")
    if env and os.path.isfile(env):
        return env

    base = _sops_root()
    candidates = []
    try:
        for fn in os.listdir(base):
            low = fn.lower()
            if low.startswith("full_sops_editor") and low.endswith(".py"):
                candidates.append(os.path.join(base, fn))
    except Exception:
        pass

    if candidates:
        candidates.sort(key=lambda p: os.path.getmtime(p), reverse=True)
        return candidates[0]

    fallback = os.path.join(base, "Full_SOPS_Editor_v8_13_EN_Ready_Custom_v4_ecu_vin_diag_bottombar_fixedfinal.py")
    if os.path.exists(fallback):
        return fallback

    return ""

# ---------- Config for editor path ----------
def load_config() -> dict:
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {"editor_path": ""}

def save_config(cfg: dict) -> None:
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception as e:
        messagebox.showwarning("Config", f"Failed to save config:\n{e}")

def ask_editor_path(parent=None) -> str:
    messagebox.showinfo(
        "Choose Editor",
        "Please select the editor script (Full_SOPS_Editor*.py). The path will be saved for future runs."
    )
    initial = _sops_root()
    p = filedialog.askopenfilename(
        parent=parent,
        title="Select Editor script (Full_SOPS_Editor*.py)",
        initialdir=initial if os.path.isdir(initial) else None,
        filetypes=[("Python files", "*.py"), ("All files", "*.*")]
    )
    if not p:
        return ""
    cfg = load_config()
    cfg["editor_path"] = p
    save_config(cfg)
    return p

def get_editor_path(parent=None) -> str:
    env = os.environ.get("FULL_SOPS_EDITOR")
    if env and os.path.isfile(env):
        return env

    cfg = load_config()
    p = (cfg.get("editor_path") or "").strip()
    if p and os.path.isfile(p):
        return p

    auto = find_editor_script()
    if auto:
        cfg["editor_path"] = auto
        save_config(cfg)
        return auto

    return ask_editor_path(parent)

# ---------- Mapping ----------
def resolve_mapping_path() -> str:
    for p in MAPPING_CANDIDATES:
        if p and os.path.exists(p):
            return p
    return ""

def load_mapping():
    global LONG_BY_ID, SHORT_BY_ID, DESC_BY_PAIR, MAPPING_PATH
    LONG_BY_ID, SHORT_BY_ID, DESC_BY_PAIR = {}, {}, {}
    MAPPING_PATH = resolve_mapping_path()
    if not MAPPING_PATH:
        return
    try:
        with open(MAPPING_PATH, "r", encoding="utf-8", newline="") as f:
            rows = list(csv.reader(f))
        if not rows:
            return
        headers = rows[0]
        norm = {"".join(ch for ch in (h or "").lower() if ch.isalnum()): i for i, h in enumerate(headers)}
        def idx_of(*names):
            for n in names:
                k = "".join(ch for ch in (n or "").lower() if ch.isalnum())
                if k in norm: return norm[k]
            return None
        idx_fpc  = idx_of("FPC- ID","FPC ID","fpc id","fpc","id")
        idx_short= idx_of("Short")
        idx_long = idx_of("Long","Name","Parameter","Title")
        idx_code = idx_of("code","Value","val")
        idx_desc = idx_of("Description","Desc","Meaning","Explanation")
        if idx_fpc is None:
            for k in ["fpcid","fpc","id"]:
                if k in norm: idx_fpc = norm[k]; break
        for r in rows[1:]:
            if not r: continue
            def at(i): return _clean(r[i]) if (i is not None and i < len(r)) else ""
            fpc_raw = at(idx_fpc)
            short   = at(idx_short) if idx_short is not None else ""
            long_   = at(idx_long) if idx_long is not None else ""
            code    = at(idx_code)
            desc    = at(idx_desc)
            if not fpc_raw: continue
            keys = {fpc_raw}
            if _is_intlike(fpc_raw): keys.add(str(int(fpc_raw)))
            for fk in keys:
                if long_ and fk not in LONG_BY_ID: LONG_BY_ID[fk] = long_
                if short and fk not in SHORT_BY_ID: SHORT_BY_ID[fk] = short
                if code and desc: DESC_BY_PAIR[(fk, code)] = desc
    except Exception as e:
        messagebox.showerror("Mapping load error", f"Failed to read mapping CSV:\n{e}")

def meaning_for(fpc_id: str, code: str) -> str:
    if not fpc_id: return ""
    cands = [fpc_id]
    if _is_intlike(fpc_id): cands.append(str(int(fpc_id)))
    for fk in cands:
        d = DESC_BY_PAIR.get((fk, code))
        if d: return d
    return ""

def short_for(fpc_id: str) -> str:
    if not fpc_id: return ""
    cands = [fpc_id]
    if _is_intlike(fpc_id): cands.append(str(int(fpc_id)))
    for fk in cands:
        s = SHORT_BY_ID.get(fk)
        if s: return s
    return f"FPC_{fpc_id}"

# ---------- CSV/columns filtering ----------
def _remove_forbidden(name: str) -> bool:
    if not name:
        return False
    low = (name or "").lower()
    if FORBIDDEN_SUBSTR and FORBIDDEN_SUBSTR.lower() in low:
        return True
    if low in EXCLUDED_COLUMNS:
        return True
    return False

def compute_csv_columns():
    cols = ["Year"]
    for _, col_name in FIXED_FIELDS:
        if not _remove_forbidden(col_name):
            cols.append(col_name)

    dynamic_labels = []
    seen = set(cols)
    for fid in DYN_RANGE:
        if fid in FIXED_IDS:
            continue
        s = short_for(fid)
        if _remove_forbidden(s):
            continue
        if s in seen:
            continue
        seen.add(s)
        dynamic_labels.append((fid, s))

    cols.extend([label for _, label in dynamic_labels])
    cols.append("XML_Path")
    return cols, dynamic_labels

def upgrade_csv_if_needed(target_header):
    if not os.path.exists(CSV_PATH):
        with open(CSV_PATH, "w", encoding="utf-8", newline="") as f:
            csv.writer(f).writerow(target_header)
        return
    with open(CSV_PATH, "r", encoding="utf-8", newline="") as f:
        old_rows = list(csv.reader(f))
    if not old_rows:
        with open(CSV_PATH, "w", encoding="utf-8", newline="") as f:
            csv.writer(f).writerow(target_header)
        return
    old_header = old_rows[0]
    if old_header == target_header:
        return
    old_idx = {name: i for i, name in enumerate(old_header)}
    tmp = CSV_PATH + ".tmp"
    with open(tmp, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(target_header)
        for r in old_rows[1:]:
            new_row = []
            for col in target_header:
                if col in old_idx and old_idx[col] < len(r):
                    new_row.append(r[old_idx[col]])
                else:
                    new_row.append("")
            w.writerow(new_row)
    shutil.move(tmp, CSV_PATH)

# ---------- Build & save ----------
def parse_fpc_map(xml_path: str) -> dict:
    fpc = {}
    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()
    except Exception as e:
        raise RuntimeError(f"XML parse failed: {e}")
    for el in root.iter():
        if local_name(el.tag).upper() == "FPC":
            name = el.attrib.get("Name")
            val  = el.attrib.get("Value")
            if name is not None and val is not None:
                fpc[name.strip()] = val.strip()
    return fpc

def build_record(xml_path: str, year: str, stored_xml_path: str, dynamic_labels):
    fpc_map = parse_fpc_map(xml_path)
    header, _ = compute_csv_columns()
    rec = {col: "" for col in header}
    rec["Year"] = (year or "").strip()

    for fpc_id, col_name in FIXED_FIELDS:
        if _remove_forbidden(col_name):
            continue
        code = fpc_map.get(fpc_id, "")
        if fpc_id == "3129" and code == "Z":
            rec[col_name] = "GMan"
        else:
            rec[col_name] = meaning_for(fpc_id, code)

    for fpc_id, col_label in dynamic_labels:
        if _remove_forbidden(col_label):
            continue
        code = fpc_map.get(fpc_id, "")
        rec[col_label] = meaning_for(fpc_id, code)

    rec["XML_Path"] = stored_xml_path
    return rec

def append_record(rec, header):
    upgrade_csv_if_needed(header)
    with open(CSV_PATH, "a", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        row = [rec.get(col, "") for col in header]
        w.writerow(row)

def read_all_records(header):
    if not os.path.exists(CSV_PATH):
        return []
    rows = []
    with open(CSV_PATH, "r", encoding="utf-8", newline="") as f:
        rdr = csv.DictReader(f)
        for r in rdr:
            rows.append(r)
    return rows

# ---------- Search Window ----------
class SearchWindow(tk.Toplevel):
    def __init__(self, parent, columns, rows, current_criteria):
        super().__init__(parent)
        self.title("Search")
        self.transient(parent)
        self.grab_set()
        self.resizable(True, True)

        self.parent = parent
        self.columns = columns
        self.widgets = {}

        frm = ttk.Frame(self, padding=8)
        frm.pack(fill="both", expand=True)

        frm.columnconfigure(1, weight=1)
        for i, col in enumerate(columns):
            ttk.Label(frm, text=f"{col}:").grid(row=i, column=0, sticky="w", padx=4, pady=3)
            values = sorted({ (r.get(col,"") or "") for r in rows if (r.get(col,"") or "") })
            options = ["(Any)"] + values
            cb = ttk.Combobox(frm, values=options, state="readonly")
            sel = current_criteria.get(col, "(Any)")
            if sel not in options:
                sel = "(Any)"
            cb.set(sel)
            cb.grid(row=i, column=1, sticky="ew", padx=4, pady=3)
            self.widgets[col] = cb

        btns = ttk.Frame(frm)
        btns.grid(row=len(columns), column=0, columnspan=2, sticky="e", pady=(10,0))
        ttk.Button(btns, text="Clear", command=self.on_clear).pack(side="left", padx=4)
        ttk.Button(btns, text="Apply", command=self.on_apply).pack(side="left", padx=4)
        ttk.Button(btns, text="Close", command=self.destroy).pack(side="left", padx=4)

    def on_clear(self):
        for cb in self.widgets.values():
            cb.set("(Any)")

    def on_apply(self):
        crit = {}
        for col, cb in self.widgets.items():
            val = cb.get()
            if val and val != "(Any)":
                crit[col] = val
        self.parent.set_search_criteria(crit)
        self.destroy()

# ---------- UI ----------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1280x700")

        ensure_dirs()
        load_mapping()

        self.header, self.dynamic_labels = compute_csv_columns()
        upgrade_csv_if_needed(self.header)

        # Menubar with Settings -> Change Editor Path…
        menubar = tk.Menu(self)
        settings_menu = tk.Menu(menubar, tearoff=False)
        settings_menu.add_command(label="Change Editor Path…", command=self.on_change_editor_path)
        menubar.add_cascade(label="Settings", menu=settings_menu)
        self.config(menu=menubar)

        # Top bar
        top = ttk.Frame(self, padding=8)
        top.pack(fill="x")

        ttk.Button(top, text="Add XML(s)", command=self.on_add_xmls).pack(side="left")
        ttk.Button(top, text="Refresh", command=self.refresh_table).pack(side="left", padx=(8, 0))
        ttk.Button(top, text="Search…", command=self.open_search_window).pack(side="left", padx=(8, 0))

        # Filter
        self.filter_var = tk.StringVar()
        ent = ttk.Entry(top, textvariable=self.filter_var, width=40)
        ent.pack(side="left", padx=(8, 0))
        ent.bind("<KeyRelease>", lambda e: self.apply_filter())

        # Right-side actions
        ttk.Button(top, text="Copy Filtered to Downloads", command=self.copy_filtered).pack(side="right")
        ttk.Button(top, text="Delete Selected", command=self.delete_selected).pack(side="right", padx=(8, 0))
        ttk.Button(top, text="Copy Selected to Downloads", command=self.copy_selected).pack(side="right")
        ttk.Button(top, text="Open in Editor", command=self.open_in_editor).pack(side="right", padx=(8, 0))

        # Treeview
        self.visible_columns = []      # computed after loading rows (no XML_Path)
        self.tree = ttk.Treeview(self, columns=(), show="headings")
        self.row_by_iid = {}           # iid -> full record (for hidden XML_Path)

        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscroll=vsb.set, xscroll=hsb.set)

        self.tree.pack(fill="both", expand=True, side="left")
        vsb.pack(fill="y", side="right")
        hsb.pack(fill="x", side="bottom")

        # Auto-fit on resize
        self.bind("<Configure>", lambda e: self.auto_fit_columns())

        # Search criteria (single choice per column)
        self.search_criteria = {}

        self.all_rows = []
        self.refresh_table()

        # First run: ensure editor path is set (will prompt once if missing)
        _ = get_editor_path(parent=self)

    # Settings action
    def on_change_editor_path(self):
        p = ask_editor_path(parent=self)
        if p:
            messagebox.showinfo("Editor Path", f"Editor path set:\n{p}")

    # ----- Visible columns -----
    def compute_visible_columns(self, rows):
        fixed_names = [name for _, name in FIXED_FIELDS if not _remove_forbidden(name)]
        must_keep = {"Year"} | set(fixed_names)

        cols = [c for c in self.header if c != "XML_Path" and not _remove_forbidden(c)]

        nonempty = {c: False for c in cols}
        for r in rows:
            for c in cols:
                if (r.get(c) or "").strip():
                    nonempty[c] = True

        visible = []
        for c in cols:
            if c in must_keep or nonempty[c]:
                visible.append(c)
        return visible

    def rebuild_tree_columns(self):
        self.tree["columns"] = self.visible_columns
        for c in self.visible_columns:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=120, anchor="w", stretch=True)

    # ----- Responsive columns -----
    def auto_fit_columns(self):
        self.update_idletasks()
        if not self.visible_columns:
            return
        tree_w = self.tree.winfo_width()
        if tree_w <= 1:
            return
        vertical_bar_width = 18
        padding = 20
        avail = max(200, tree_w - vertical_bar_width - padding)
        weights = {c: 1 for c in self.visible_columns}
        for c in ("Engine Version", "Exhaust Emission Control"):
            if c in weights:
                weights[c] = 2
        total_weight = sum(weights.values()) or 1
        min_widths = {c: 80 for c in self.visible_columns}
        for c in ("Engine Version","EngineECU","EngineSize","GearboxECU","Gearbox Type","Exhaust Emission Control"):
            if c in min_widths:
                min_widths[c] = 120
        remaining = avail - sum(min_widths.get(c, 0) for c in self.visible_columns)
        if remaining < 0:
            remaining = 0
        for c in self.visible_columns:
            portion = (weights[c] / total_weight)
            extra = int(remaining * portion)
            width = min_widths.get(c, 80) + extra
            self.tree.column(c, width=max(60, width), stretch=True)

    # ----- Search handling -----
    def open_search_window(self):
        SearchWindow(self, self.visible_columns, self.all_rows, self.search_criteria)

    def set_search_criteria(self, criteria: dict):
        self.search_criteria = criteria
        self.apply_filter()

    # ----- Actions -----
    def on_add_xmls(self):
        paths = filedialog.askopenfilenames(
            title="Select Scania XML file(s)",
            filetypes=[("XML files", "*.xml"), ("All files", "*.*")]
        )
        if not paths:
            return
        added = 0
        for p in paths:
            year = simpledialog.askstring("Year", f"Enter production year for:\n{os.path.basename(p)}", parent=self)
            if year is None:
                continue
            year = year.strip()
            try:
                dest = unique_store_path(p)
                shutil.copy2(p, dest)
                rec = build_record(p, year, dest, self.dynamic_labels)
                append_record(rec, self.header)
                added += 1
            except Exception as e:
                messagebox.showerror("Error", f"Failed to add {os.path.basename(p)}:\n{e}")
        if added:
            messagebox.showinfo("Done", f"Added {added} file(s) to CSV.")
        self.refresh_table()

    def refresh_table(self):
        # Recompute header each load to ensure removed columns are dropped
        self.header, self.dynamic_labels = compute_csv_columns()
        upgrade_csv_if_needed(self.header)

        self.all_rows = read_all_records(self.header)
        self.visible_columns = self.compute_visible_columns(self.all_rows)
        self.rebuild_tree_columns()
        self.render_rows(self.all_rows)

    def render_rows(self, rows):
        self.row_by_iid.clear()
        for i in self.tree.get_children():
            self.tree.delete(i)
        for r in rows:
            values = [r.get(col, "") for col in self.visible_columns]
            iid = self.tree.insert("", "end", values=values)
            self.row_by_iid[iid] = r
        self.auto_fit_columns()

    def apply_filter(self):
        rows = list(self.all_rows)

        # Apply single-choice criteria (AND)
        if self.search_criteria:
            filtered = []
            for r in rows:
                ok = True
                for col, val in self.search_criteria.items():
                    if (r.get(col, "") or "") != val:
                        ok = False
                        break
                if ok:
                    filtered.append(r)
            rows = filtered

        # Free-text filter (visible columns only)
        q = (self.filter_var.get() or "").strip().lower()
        if q:
            filtered = []
            for r in rows:
                row_text = " | ".join(str(r.get(col, "")) for col in self.visible_columns).lower()
                if q in row_text:
                    filtered.append(r)
            rows = filtered

        self.render_rows(rows)

    # ----- Copy helpers -----
    def _copy_paths_to_downloads(self, xml_paths):
        if not xml_paths:
            messagebox.showwarning("No files", "Nothing to copy.")
            return
        downloads = os.path.join(os.path.expanduser("~"), "Downloads")
        os.makedirs(downloads, exist_ok=True)
        copied = 0
        failed = 0
        for xml_path in xml_paths:
            if not xml_path or not os.path.exists(xml_path):
                failed += 1
                continue
            base = os.path.basename(xml_path)
            dest = os.path.join(downloads, base)
            if os.path.exists(dest):
                name, ext = os.path.splitext(base)
                suffix = time.strftime("%Y%m%d_%H%M%S")
                dest = os.path.join(downloads, f"{name}_{suffix}{ext}")
            try:
                shutil.copy2(xml_path, dest)
                copied += 1
            except Exception:
                failed += 1
        messagebox.showinfo("Copy", f"Copied: {copied}\nFailed: {failed}")

    def copy_selected(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("No selection", "Please select a row to copy its XML.")
            return
        xml_paths = []
        for iid in sel:
            row = self.row_by_iid.get(iid, {})
            p = row.get("XML_Path", "")
            if p:
                xml_paths.append(p)
        self._copy_paths_to_downloads(xml_paths)

    def copy_filtered(self):
        # copy ALL currently visible rows (after current filters)
        xml_paths = []
        for iid in self.tree.get_children():
            row = self.row_by_iid.get(iid, {})
            p = row.get("XML_Path", "")
            if p:
                xml_paths.append(p)
        self._copy_paths_to_downloads(xml_paths)

    # ----- Delete selected -----
    def _rewrite_csv_excluding(self, xml_paths_to_remove: set):
        # Ensure header is up-to-date
        upgrade_csv_if_needed(self.header)
        try:
            with open(CSV_PATH, "r", encoding="utf-8", newline="") as f:
                rdr = csv.DictReader(f)
                rows = list(rdr)
        except FileNotFoundError:
            return

        keep = [r for r in rows if (r.get("XML_Path", "") not in xml_paths_to_remove)]

        tmp = CSV_PATH + ".tmpdel"
        with open(tmp, "w", encoding="utf-8", newline="") as f:
            w = csv.DictWriter(f, fieldnames=self.header)
            w.writeheader()
            for r in keep:
                w.writerow({col: r.get(col, "") for col in self.header})
        shutil.move(tmp, CSV_PATH)

    def delete_selected(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("No selection", "Please select at least one row to delete.")
            return

        xml_paths = []
        for iid in sel:
            row = self.row_by_iid.get(iid, {})
            p = row.get("XML_Path", "")
            if p:
                xml_paths.append(p)
        if not xml_paths:
            messagebox.showwarning("No files", "Selected rows have no stored XML paths.")
            return

        ans = messagebox.askyesnocancel(
            "Delete confirmation",
            "Delete selected rows from index?\n\nYes = Remove from index AND delete stored XML files.\nNo = Remove from index ONLY.\nCancel = Abort."
        )
        if ans is None:
            return  # Cancel
        delete_from_disk = bool(ans)

        to_remove = set(xml_paths)
        # Remove from CSV
        try:
            self._rewrite_csv_excluding(to_remove)
        except Exception as e:
            messagebox.showerror("Delete failed", f"Failed updating CSV:\n{e}")
            return

        # Optionally delete files
        deleted = 0
        failed = 0
        if delete_from_disk:
            for p in to_remove:
                try:
                    if p and os.path.exists(p):
                        os.remove(p)
                        deleted += 1
                except Exception:
                    failed += 1

        self.refresh_table()
        if delete_from_disk:
            messagebox.showinfo("Deleted", f"Removed from index: {len(to_remove)}\nDeleted files: {deleted}\nFailed: {failed}")
        else:
            messagebox.showinfo("Deleted", f"Removed from index: {len(to_remove)}")

    # ---- Open selected XML directly in the Editor ----
    def open_in_editor(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("No selection", "Please select a row first.")
            return
        if len(sel) > 1:
            messagebox.showinfo(APP_TITLE, "Only the first selected row will be opened.")

        row = self.row_by_iid.get(sel[0], {})
        xml_path = row.get("XML_Path", "")
        if not xml_path or not os.path.exists(xml_path):
            messagebox.showerror("Missing file", f"Stored XML not found:\n{xml_path}")
            return

        editor_path = get_editor_path(parent=self)
        if not editor_path or not os.path.exists(editor_path):
            messagebox.showerror(
                "Editor not found",
                "Could not find the editor. Use Settings → Change Editor Path… to set it."
            )
            return

        py = sys.executable or shutil.which("python") or shutil.which("py")
        if not py:
            messagebox.showerror("Python not found", "Could not locate Python to run the editor.")
            return

        try:
            subprocess.Popen([py, editor_path, "--open", xml_path, "--analyze"], close_fds=True)
        except Exception as e:
            messagebox.showerror("Launch failed", f"Failed to launch editor:\n{e}")

    # ====== NEW: programmatic helpers to support CLI ======
    def add_xml_programmatically(self, xml_path: str, year: str) -> str:
        """Add a single XML without dialogs. Returns stored destination path."""
        dest = unique_store_path(xml_path)
        shutil.copy2(xml_path, dest)
        rec = build_record(xml_path, year or "", dest, self.dynamic_labels)
        append_record(rec, self.header)
        self.refresh_table()
        try:
            self.select_row_by_xml(dest)
        except Exception:
            pass
        return dest

    def select_row_by_xml(self, path: str) -> bool:
        """Try to select the row whose XML_Path matches the given path."""
        want = os.path.normcase(os.path.abspath(path or ""))
        for iid, row in self.row_by_iid.items():
            p = os.path.normcase(os.path.abspath(row.get("XML_Path", "") or ""))
            if p and p == want:
                self.tree.selection_set(iid)
                self.tree.see(iid)
                return True
        # fallback by basename if absolute paths differ
        base = os.path.basename(want)
        for iid, row in self.row_by_iid.items():
            if os.path.basename(row.get("XML_Path", "") or "") == base:
                self.tree.selection_set(iid)
                self.tree.see(iid)
                return True
        return False

# ====== CLI glue ======

def parse_args():
    ap = argparse.ArgumentParser(description=APP_TITLE)
    ap.add_argument("--open", dest="open_path", help="Path to an XML file to add/select")
    ap.add_argument("--year", dest="year", help="Production year for the XML (used with --add)")
    ap.add_argument("--add", action="store_true", help="Copy XML into store and append to CSV")
    ap.add_argument("--analyze", action="store_true", help="Refresh & select row after startup")
    ap.add_argument("--set-editor", dest="set_editor", help="Persist editor path to config.json")
    return ap.parse_args()


def main():
    ensure_dirs()
    load_mapping()
    args = parse_args()

    # Prepare Tk app
    app = App()

    # Optionally set editor path from CLI
    if args.set_editor:
        cfg = load_config()
        cfg["editor_path"] = args.set_editor
        try:
            with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
            try:
                messagebox.showinfo("Editor Path", f"Editor path saved to config:\n{args.set_editor}")
            except Exception:
                pass
        except Exception as e:
            try:
                messagebox.showerror("Config", f"Failed to write config.json:\n{e}")
            except Exception:
                pass

    # If an XML is provided on CLI
    if args.open_path:
        p = os.path.abspath(args.open_path)
        try:
            if args.add:
                year = args.year or ""
                try:
                    # If no year provided, ask interactively (keeps behavior similar to UI)
                    if not year:
                        year = simpledialog.askstring("Year", f"Enter production year for:\n{os.path.basename(p)}", parent=app) or ""
                except Exception:
                    pass
                dest = app.add_xml_programmatically(p, year)
                # Explicit analyze step isn't different from refresh in this app, but keep flag parity
                if args.analyze:
                    app.refresh_table()
                    app.select_row_by_xml(dest)
            else:
                # Not adding: just try to select existing row by basename
                app.refresh_table()
                app.select_row_by_xml(p)
        except Exception as e:
            try:
                messagebox.showerror(APP_TITLE, f"CLI operation failed:\n{e}")
            except Exception:
                pass

    app.mainloop()


if __name__ == "__main__":
    main()
