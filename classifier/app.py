import os
import csv
import shutil
import time
import sys
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import xml.etree.ElementTree as ET

APP_TITLE = "Scania Spec Index (ttkinter — meanings, dynamic 3113..3148, responsive, search)"

# Base paths
BASE_DIR = os.path.join(os.path.expanduser("~"), r"Documents\GitHub\scania_SOPS\classifier")
CSV_PATH = os.path.join(BASE_DIR, "spec_index.csv")
XML_STORE_DIR = os.path.join(BASE_DIR, "stored_xml")

# Mapping candidates
MAPPING_CANDIDATES = [
    r"C:\Users\Kstore\Documents\GitHub\scania_SOPS\sops_fpc_mapping.csv",
    os.path.join(os.path.expanduser("~"), r"Documents\GitHub\scania_SOPS\sops_fpc_mapping.csv"),
    os.path.join(os.path.expanduser("~"), r"Documents\GitHub\scania_SOPS\classifier\sops_fpc_mapping.csv"),
    os.path.join(os.path.dirname(__file__), "sops_fpc_mapping.csv"),
]

# Remove any columns whose name contains this (case-insensitive)
FORBIDDEN_SUBSTR = "environment"

# Fixed fields (Descriptions only). Order defines column order in table/CSV.
# NOTE: Power_kW (3266) was removed earlier by request.
FIXED_FIELDS = [
    ("448",   "Axles"),
    ("3127",  "EngineECU"),
    ("408",   "Engine Version"),             # renamed from EngineSoftware
    ("142",   "EngineSize"),

    # Emissions & related additions
    ("2344",  "Immobiliser"),
    ("2409",  "EGR"),
    ("4306",  "NOx sensor"),
    ("4280",  "Exhaust Emission Control"),
    ("2471",  "Emission level"),

    # COO + Gearbox cluster
    ("3113",  "COO"),                        # brought as fixed to control placement
    ("3129",  "GearboxECU"),                 # special rule: if code == 'Z' -> 'GMan'
    ("17",    "Gearbox Type"),               # placed right after GearboxECU
]

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

# ---- Editor discovery (NEW) ----
def _sops_root() -> str:

    return os.path.join(os.path.expanduser("~"), "Documents", "GitHub", "scania_SOPS")

def find_editor_script() -> str:
    """
    يبحث عن ملف المحرّر Full_SOPS_Editor*.py تحت مجلد scania_SOPS.
    يمكن تجاوز المسار عبر متغير البيئة FULL_SOPS_EDITOR.
    """
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

    # fallback إلى اسم قديم إن وُجد
    fallback = os.path.join(base, "Full_SOPS_Editor_v8_13_EN_Ready_Custom_v4_ecu_vin_diag_bottombar_fixedfinal.py")
    if os.path.exists(fallback):
        return fallback

    return ""  # لم يُعثر على المحرّر

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

# ---------- XML ----------
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

# ---------- CSV header ----------
def _remove_forbidden(name: str) -> bool:
    return FORBIDDEN_SUBSTR and (FORBIDDEN_SUBSTR.lower() in (name or "").lower())

def compute_csv_columns():
    """
    Header = Year + fixed (incl. new requested) + dynamic (3113..3148 by Short, dedup, no 'environment') + XML_Path.
    """
    cols = ["Year"]

    # Add fixed fields in desired order, skipping forbidden names
    for _, col_name in FIXED_FIELDS:
        if not _remove_forbidden(col_name):
            cols.append(col_name)

    # Dynamic by Short, dedup by label, skip 'environment'
    dynamic_labels = []
    seen = set(cols)
    for fid in DYN_RANGE:
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
def build_record(xml_path: str, year: str, stored_xml_path: str, dynamic_labels):
    fpc_map = parse_fpc_map(xml_path)
    header, _ = compute_csv_columns()
    rec = {col: "" for col in header}
    rec["Year"] = (year or "").strip()

    # Fill fixed fields (Descriptions only)
    for fpc_id, col_name in FIXED_FIELDS:
        if _remove_forbidden(col_name):
            continue
        code = fpc_map.get(fpc_id, "")
        if fpc_id == "3129" and code == "Z":
            rec[col_name] = "GMan"
        else:
            rec[col_name] = meaning_for(fpc_id, code)

    # Dynamic (3113..3148)
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

        # Top bar
        top = ttk.Frame(self, padding=8)
        top.pack(fill="x")

        ttk.Button(top, text="Add XML(s)", command=self.on_add_xmls).pack(side="left")
        ttk.Button(top, text="Refresh", command=self.refresh_table).pack(side="left", padx=(8, 0))
        ttk.Button(top, text="Search…", command=self.open_search_window).pack(side="left", padx=(8, 0))

        map_status = MAPPING_PATH if MAPPING_PATH else "NOT FOUND (meanings will be blank)"
        self.map_status_var = tk.StringVar(value=f"Mapping: {map_status}")
        

        
        self.filter_var = tk.StringVar()
        ent = ttk.Entry(top, textvariable=self.filter_var, width=40)
        
        ent.bind("<KeyRelease>", lambda e: self.apply_filter())

        # NEW: Open in Editor button (on the right)
        ttk.Button(top, text="Open in Editor", command=self.open_in_editor).pack(side="right", padx=(8, 0))
        ttk.Button(top, text="Copy Selected to Downloads", command=self.copy_selected).pack(side="right")

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

    # ----- Visible columns -----
    def compute_visible_columns(self, rows):
        """
        Hide XML_Path; drop columns empty across all rows; keep Year + fixed always;
        also exclude any column containing 'environment'.
        """
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
        # give more space to long ones
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

    def copy_selected(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("No selection", "Please select a row to copy its XML.")
            return
        downloads = os.path.join(os.path.expanduser("~"), "Downloads")
        os.makedirs(downloads, exist_ok=True)
        copied = 0
        for iid in sel:
            row = self.row_by_iid.get(iid, {})
            xml_path = row.get("XML_Path", "")
            if not xml_path or not os.path.exists(xml_path):
                messagebox.showwarning("Missing file", f"Stored XML path not found:\n{xml_path}")
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
            except Exception as e:
                messagebox.showerror("Copy failed", f"Failed to copy to Downloads:\n{e}")
        if copied:
            messagebox.showinfo("Copied", f"Copied {copied} file(s) to Downloads.")

    # ---- NEW: open selected XML directly in the Editor ----
    def open_in_editor(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("No selection", "Please select a row first.")
            return
        if len(sel) > 1:
            messagebox.showinfo(APP_TITLE, "سيتم فتح أوّل صف محدّد فقط.")

        row = self.row_by_iid.get(sel[0], {})
        xml_path = row.get("XML_Path", "")
        if not xml_path or not os.path.exists(xml_path):
            messagebox.showerror("Missing file", f"Stored XML not found:\n{xml_path}")
            return

        editor_path = find_editor_script()
        if not editor_path:
            messagebox.showerror(
                "Editor not found",
                "لم أجد ملف المُحرّر تحت:\nDocuments\\GitHub\\scania_SOPS\n"
                "يمكنك أيضًا تحديد المسار بدقّة عبر متغير البيئة FULL_SOPS_EDITOR."
            )
            return

        py = sys.executable or shutil.which("python") or shutil.which("py")
        if not py:
            messagebox.showerror("Python not found", "لم أستطع العثور على Python لتشغيل المُحرّر.")
            return

        try:
            subprocess.Popen([py, editor_path, "--open", xml_path, "--analyze"], close_fds=True)
        except Exception as e:
            messagebox.showerror("Launch failed", f"فشل تشغيل المُحرّر:\n{e}")

if __name__ == "__main__":
    load_mapping()
    app = App()
    app.mainloop()
