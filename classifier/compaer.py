import os
import csv
import time
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import xml.etree.ElementTree as ET
from collections import defaultdict, Counter

APP_TITLE = "SOPS Delta Viewer (Unique FPC — centered columns & header separators)"

# ----- Locations -----
BASE_DIR = os.path.join(os.path.expanduser("~"), r"Documents\GitHub\scania_SOPS\delta_viewer")
os.makedirs(BASE_DIR, exist_ok=True)

# Mapping CSV candidates
MAPPING_CANDIDATES = [
    r"C:\Users\Kstore\Documents\GitHub\scania_SOPS\sops_fpc_mapping.csv",
    os.path.join(os.path.expanduser("~"), r"Documents\GitHub\scania_SOPS\sops_fpc_mapping.csv"),
    os.path.join(os.path.expanduser("~"), r"Documents\GitHub\scania_SOPS\classifier\sops_fpc_mapping.csv"),
    os.path.join(os.path.dirname(__file__), "sops_fpc_mapping.csv"),
]

# ----- Mapping stores -----
LONG_BY_ID   = {}
SHORT_BY_ID  = {}
DESC_BY_PAIR = {}
MAPPING_PATH = ""

# ---------- Helpers ----------
def _clean(s: str) -> str:
    return (s or "").strip()

def _is_intlike(s: str) -> bool:
    try:
        int(s.strip()); return True
    except Exception:
        return False

def local_name(tag: str) -> str:
    if "}" in tag:
        return tag.split("}", 1)[1]
    return tag

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

def short_for(fpc_id: str) -> str:
    if not fpc_id: return ""
    cands = [fpc_id]
    if _is_intlike(fpc_id): cands.append(str(int(fpc_id)))
    for fk in cands:
        s = SHORT_BY_ID.get(fk)
        if s: return s
    return f"FPC_{fpc_id}"

def long_for(fpc_id: str) -> str:
    if not fpc_id: return ""
    cands = [fpc_id]
    if _is_intlike(fpc_id): cands.append(str(int(fpc_id)))
    for fk in cands:
        s = LONG_BY_ID.get(fk)
        if s: return s
    return ""

def desc_for(fpc_id: str, code: str) -> str:
    if not fpc_id: return ""
    cands = [fpc_id]
    if _is_intlike(fpc_id): cands.append(str(int(fpc_id)))
    for fk in cands:
        s = DESC_BY_PAIR.get((fk, code))
        if s: return s
    return ""

def parse_fpc_map(xml_path: str) -> dict:
    """Return {FPC_ID: code} from a Scania XML."""
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

# ---------- Details Window ----------
class DetailsWindow(tk.Toplevel):
    """Tabbed details for a single FPC_ID: Summary (code counts A/B) + Files A + Files B."""
    def __init__(self, parent, fid, details_A, details_B):
        super().__init__(parent)
        self.title(f"Details — FPC_ID {fid}")
        self.geometry("1100x620")

        # Style: center headings + solid header borders
        style = ttk.Style(self)
        # Prefer Windows 'vista' for clearer separators when available
        try:
            if 'vista' in style.theme_names():
                style.theme_use('vista')
        except Exception:
            pass
        style.configure('Grid.Treeview', rowheight=24)
        style.configure('Grid.Treeview.Heading', anchor='center', relief='solid', borderwidth=1)

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True)

        # Build code-count summaries
        cntA = Counter([d.get("Code","") for d in details_A if d.get("Code","")])
        cntB = Counter([d.get("Code","") for d in details_B if d.get("Code","")])

        def build_summary_rows():
            rows = []
            # union of all codes
            all_codes = sorted(set(cntA.keys()) | set(cntB.keys()))
            for code in all_codes:
                desc = desc_for(fid, code)
                rows.append({"Side":"A", "Code":code, "Count":cntA.get(code,0), "Description":desc})
                rows.append({"Side":"B", "Code":code, "Count":cntB.get(code,0), "Description":desc})
            return rows

        # Tab: Summary
        tab1 = ttk.Frame(nb, padding=8); nb.add(tab1, text="Summary")
        cols1 = ["Side", "Code", "Count", "Description"]
        tree1 = ttk.Treeview(tab1, columns=cols1, show="headings", style='Grid.Treeview')
        vsb1 = ttk.Scrollbar(tab1, orient="vertical", command=tree1.yview)
        hsb1 = ttk.Scrollbar(tab1, orient="horizontal", command=tree1.xview)
        tree1.configure(yscroll=vsb1.set, xscroll=hsb1.set)
        tree1.pack(fill="both", expand=True, side="left")
        vsb1.pack(fill="y", side="right")
        hsb1.pack(fill="x", side="bottom")
        for c in cols1:
            tree1.heading(c, text=c, anchor="center")
            tree1.column(c, width=160, stretch=True, anchor="center")
        for r in build_summary_rows():
            tree1.insert("", "end", values=[r.get(c,"") for c in cols1])

        # Tab: Files A
        tab2 = ttk.Frame(nb, padding=8); nb.add(tab2, text="Files A")
        cols2 = ["File", "FPC_ID", "Code", "Description"]
        tree2 = ttk.Treeview(tab2, columns=cols2, show="headings", style='Grid.Treeview')
        vsb2 = ttk.Scrollbar(tab2, orient="vertical", command=tree2.yview)
        hsb2 = ttk.Scrollbar(tab2, orient="horizontal", command=tree2.xview)
        tree2.configure(yscroll=vsb2.set, xscroll=hsb2.set)
        tree2.pack(fill="both", expand=True, side="left")
        vsb2.pack(fill="y", side="right")
        hsb2.pack(fill="x", side="bottom")
        for c in cols2:
            tree2.heading(c, text=c, anchor="center")
            tree2.column(c, width=200, stretch=True, anchor="center")
        for d in details_A:
            tree2.insert("", "end", values=[d.get("File",""), d.get("FPC_ID",""), d.get("Code",""), d.get("Description","")])

        # Tab: Files B
        tab3 = ttk.Frame(nb, padding=8); nb.add(tab3, text="Files B")
        tree3 = ttk.Treeview(tab3, columns=cols2, show="headings", style='Grid.Treeview')
        vsb3 = ttk.Scrollbar(tab3, orient="vertical", command=tree3.yview)
        hsb3 = ttk.Scrollbar(tab3, orient="horizontal", command=tree3.xview)
        tree3.configure(yscroll=vsb3.set, xscroll=hsb3.set)
        tree3.pack(fill="both", expand=True, side="left")
        vsb3.pack(fill="y", side="right")
        hsb3.pack(fill="x", side="bottom")
        for c in cols2:
            tree3.heading(c, text=c, anchor="center")
            tree3.column(c, width=200, stretch=True, anchor="center")
        for d in details_B:
            tree3.insert("", "end", values=[d.get("File",""), d.get("FPC_ID",""), d.get("Code",""), d.get("Description","")])

# ---------- App ----------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1400x840")

        load_mapping()

        # Style: center headings + solid header borders
        style = ttk.Style(self)
        try:
            if 'vista' in style.theme_names():
                style.theme_use('vista')
        except Exception:
            pass
        style.configure('Grid.Treeview', rowheight=26)
        style.configure('Grid.Treeview.Heading', anchor='center', relief='solid', borderwidth=1)

        # State
        self.files_A = []
        self.files_B = []
        self.fpccodes_A = {}   # file -> {id: code}
        self.fpccodes_B = {}
        self.only_changes = tk.BooleanVar(value=True)
        self.sort_mode = tk.StringVar(value="similarity")  # "similarity" or "category"

        # compare rows cache (unique per FPC_ID)
        self.rows_cmp = []

        # ----- Top controls -----
        top = ttk.Frame(self, padding=8)
        top.pack(fill="x")

        a_box = ttk.LabelFrame(top, text="Baseline (A)", padding=8)
        a_box.pack(side="left", padx=(0, 8))
        ttk.Button(a_box, text="Add XML…", command=lambda: self.on_add_files("A")).pack(side="left")
        ttk.Button(a_box, text="Clear", command=lambda: self.on_clear("A")).pack(side="left", padx=(6,0))

        b_box = ttk.LabelFrame(top, text="Target (B)", padding=8)
        b_box.pack(side="left", padx=(8, 8))
        ttk.Button(b_box, text="Add XML…", command=lambda: self.on_add_files("B")).pack(side="left")
        ttk.Button(b_box, text="Clear", command=lambda: self.on_clear("B")).pack(side="left", padx=(6,0))

        opts = ttk.Frame(top, padding=0)
        opts.pack(side="left", padx=(8, 8))
        ttk.Checkbutton(opts, text="Show only changes", variable=self.only_changes, command=self.apply_filter).pack(side="left")

        sort_box = ttk.Frame(top, padding=0)
        sort_box.pack(side="left", padx=(8, 8))
        ttk.Label(sort_box, text="Sort by:").pack(side="left", padx=(0,4))
        sort_cb = ttk.Combobox(sort_box, state="readonly",
                               values=["Similarity", "Category"],
                               width=12)
        sort_cb.set("Similarity")
        sort_cb.pack(side="left")
        sort_cb.bind("<<ComboboxSelected>>", lambda e: self.on_sort_changed(sort_cb.get()))

        actions = ttk.Frame(top, padding=0)
        actions.pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Compare", command=self.recompute).pack(side="left")
        ttk.Button(actions, text="Export CSV…", command=lambda: self.export_csv()).pack(side="left", padx=(6,0))
        ttk.Button(actions, text="Export HTML…", command=lambda: self.export_html()).pack(side="left", padx=(6,0))

        ttk.Label(top, text="Filter:").pack(side="left", padx=(16,4))
        self.filter_var = tk.StringVar()
        ent = ttk.Entry(top, textvariable=self.filter_var, width=40)
        ent.pack(side="left")
        ent.bind("<KeyRelease>", lambda e: self.apply_filter())

        self.status_var = tk.StringVar(value="Ready.")
        ttk.Label(top, textvariable=self.status_var).pack(side="right")

        # Table
        self.tree = ttk.Treeview(self, columns=(), show="headings", style='Grid.Treeview')
        self.tree.bind("<Double-1>", self.on_row_double_click)
        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscroll=vsb.set, xscroll=hsb.set)
        self.tree.pack(fill="both", expand=True, side="left")
        vsb.pack(fill="y", side="right")
        hsb.pack(fill="x", side="bottom")

        # row tag styles (body backgrounds)
        self.tree.tag_configure("cat_both_same",     background="#F5F5F5")   # gray
        self.tree.tag_configure("cat_both_inc",      background="#E6FFED")   # light green
        self.tree.tag_configure("cat_both_dec",      background="#FFEEEF")   # light red
        self.tree.tag_configure("cat_only_a",        background="#FDECEC")   # removed
        self.tree.tag_configure("cat_only_b",        background="#E8F8E8")   # added

        self.rebuild_columns()

    def on_sort_changed(self, label):
        self.sort_mode.set("similarity" if label.lower().startswith("similar") else "category")
        self.sort_rows_in_place()
        self.apply_filter()

    def rebuild_columns(self):
        # Unique row per FPC_ID (no repetition)
        cols = ["Category", "Similarity", "FPC_ID", "Short", "Long",
                "Count_A", "Count_B", "Δ (B−A)",
                "Codes_A", "Codes_B", "∩ Codes", "A−B", "B−A"]
        self.tree["columns"] = cols
        for c in cols:
            self.tree.heading(c, text=c, anchor="center")             # <-- center header
            self.tree.column(c, width=140, anchor="center", stretch=True)  # <-- center cells

    def refresh_table(self, rows):
        self.tree.delete(*self.tree.get_children())
        cols = self.tree["columns"]
        for r in rows:
            values = [r.get(c, "") for c in cols]
            self.tree.insert("", "end", values=values, tags=(r.get("_tag",""),))
        # auto stretch a bit
        self.update_idletasks()
        tw = self.tree.winfo_width()
        if tw > 200 and len(cols) > 0:
            extra = max(0, tw - 200)
            per = extra // len(cols)
            for c in cols:
                curw = self.tree.column(c, option="width")
                self.tree.column(c, width=curw + per)

    # ----- Files -----
    def on_add_files(self, which):
        paths = filedialog.askopenfilenames(
            title=f"Select XML files for Group {which}",
            filetypes=[("XML files", "*.xml"), ("All files", "*.*")]
        )
        if not paths: return

        target_files = self.files_A if which == "A" else self.files_B
        target_map   = self.fpccodes_A if which == "A" else self.fpccodes_B

        added = 0
        for p in paths:
            if p not in target_files:
                try:
                    m = parse_fpc_map(p)
                    target_files.append(p)
                    target_map[p] = m
                    added += 1
                except Exception as e:
                    messagebox.showerror("Parse error", f"Failed to parse {os.path.basename(p)}:\n{e}")

        self.status_var.set(f"Loaded A: {len(self.files_A)} | B: {len(self.files_B)}")
        if added:
            messagebox.showinfo("Added", f"Group {which}: added {added} file(s).")

    def on_clear(self, which):
        if which == "A":
            self.files_A.clear()
            self.fpccodes_A.clear()
        else:
            self.files_B.clear()
            self.fpccodes_B.clear()
        self.rows_cmp = []
        self.refresh_table([])
        self.status_var.set(f"Cleared Group {which}.")

    # ----- Compute compare (unique per FPC_ID) -----
    def recompute(self):
        self.rebuild_columns()
        if not self.files_A and not self.files_B:
            messagebox.showwarning("No files", "Add files to Group A and/or B first.")
            return

        # Groups by FPC_ID only (unique parameter row)
        groups_A = defaultdict(set)   # fid -> set(files)
        groups_B = defaultdict(set)
        details_A = defaultdict(list) # fid -> list of dicts (File,FPC_ID,Code,Description)
        details_B = defaultdict(list)

        for fpath, m in self.fpccodes_A.items():
            for fid, code in m.items():
                groups_A[fid].add(fpath)
                details_A[fid].append({"File": fpath, "FPC_ID": fid, "Code": code, "Description": desc_for(fid, code)})

        for fpath, m in self.fpccodes_B.items():
            for fid, code in m.items():
                groups_B[fid].add(fpath)
                details_B[fid].append({"File": fpath, "FPC_ID": fid, "Code": code, "Description": desc_for(fid, code)})

        # Code sets per FID
        codes_A_by_fid = defaultdict(set)
        codes_B_by_fid = defaultdict(set)
        for fid, lst in details_A.items():
            for d in lst: 
                if d.get("Code",""): codes_A_by_fid[fid].add(d["Code"])
        for fid, lst in details_B.items():
            for d in lst: 
                if d.get("Code",""): codes_B_by_fid[fid].add(d["Code"])

        def join_lim(iterable, limit=10):
            items = sorted([x for x in iterable if x])
            return ", ".join(items[:limit]) + (" ..." if len(items) > limit else "")

        rows = []
        all_fids = set(groups_A.keys()) | set(groups_B.keys())

        for fid in all_fids:
            setA = groups_A.get(fid, set())
            setB = groups_B.get(fid, set())
            cntA, cntB = len(setA), len(setB)
            delta = cntB - cntA

            # Categorize
            if cntA == 0 and cntB > 0:
                cat, tag, arrow = "Only B", "cat_only_b", "▲"
            elif cntB == 0 and cntA > 0:
                cat, tag, arrow = "Only A", "cat_only_a", "▼"
            elif cntA == cntB:
                cat, tag, arrow = "Both (same)", "cat_both_same", "="
            elif cntB > cntA:
                cat, tag, arrow = "Both (increased)", "cat_both_inc", "▲"
            else:
                cat, tag, arrow = "Both (decreased)", "cat_both_dec", "▼"

            # Similarity: 60% count similarity + 40% Jaccard of codes
            codesA = codes_A_by_fid.get(fid, set())
            codesB = codes_B_by_fid.get(fid, set())
            inter = codesA & codesB
            union = codesA | codesB
            jacc = (len(inter) / len(union)) if len(union) > 0 else 0.0
            count_sim = (min(cntA, cntB) / max(cntA, cntB)) if max(cntA, cntB) > 0 else 0.0
            sim = 0.6 * count_sim + 0.4 * jacc
            support = cntA + cntB

            rows.append({
                "Category": cat,
                "Similarity": f"{sim*100:.1f}%",
                "_similarity_val": sim,
                "_support_val": support,
                "FPC_ID": fid,
                "Short": short_for(fid),
                "Long": long_for(fid),
                "Count_A": cntA,
                "Count_B": cntB,
                "Δ (B−A)": f"{arrow} {delta:+d}",
                "Codes_A": join_lim(codesA),
                "Codes_B": join_lim(codesB),
                "∩ Codes": join_lim(inter),
                "A−B": join_lim(codesA - codesB),
                "B−A": join_lim(codesB - codesA),
                "_details_A": details_A.get(fid, []),
                "_details_B": details_B.get(fid, []),
                "_fid": fid,
                "_tag": tag
            })

        self.rows_cmp = rows
        self.sort_rows_in_place()
        self.apply_filter()
        self.status_var.set(f"Compared FPCs: {len(rows)} | Files A: {len(self.fpccodes_A)} | Files B: {len(self.fpccodes_B)}")

    def sort_rows_in_place(self):
        if not self.rows_cmp: return
        if self.sort_mode.get() == "similarity":
            self.rows_cmp.sort(
                key=lambda r: (-float(r.get("_similarity_val",0.0)),
                               -int(r.get("_support_val",0)),
                               r.get("FPC_ID",""))
            )
        else:
            cat_order = {"Only B": 0, "Only A": 1, "Both (increased)": 2, "Both (decreased)": 3, "Both (same)": 4}
            self.rows_cmp.sort(key=lambda r: (cat_order.get(r["Category"], 9), r["FPC_ID"]))

    # ----- Filter -----
    def apply_filter(self):
        if not self.rows_cmp:
            self.refresh_table([])
            return
        q = (self.filter_var.get() or "").lower().strip()
        rows = self.rows_cmp
        if self.only_changes.get():
            rows = [r for r in rows if r["Category"] != "Both (same)"]
        if q:
            cols = self.tree["columns"]
            out = []
            for r in rows:
                joined = " | ".join(str(r.get(c, "")) for c in cols).lower()
                if q in joined:
                    out.append(r)
            rows = out
        self.refresh_table(rows)

    # ----- Details -----
    def on_row_double_click(self, _event):
        item = self.tree.focus()
        if not item: return
        values = self.tree.item(item, "values")
        columns = list(self.tree["columns"])
        row = {c: values[i] if i < len(values) else "" for i, c in enumerate(columns)}
        fid = row.get("FPC_ID", "")
        # find original
        matches = [r for r in self.rows_cmp if r.get("_fid") == fid]
        if not matches: return
        r = matches[0]
        DetailsWindow(self, fid, r.get("_details_A", []), r.get("_details_B", []))

    # ----- Export -----
    def export_csv(self):
        if not self.rows_cmp:
            messagebox.showwarning("No data", "Run Compare first."); return
        folder = filedialog.askdirectory(title="Choose folder to save CSV", mustexist=True)
        if not folder: return
        ts = time.strftime("%Y%m%d_%H%M%S")
        path = os.path.join(folder, f"delta_uniqueFPC_{self.sort_mode.get()}_{ts}.csv")
        cols = list(self.tree["columns"])
        try:
            with open(path, "w", encoding="utf-8", newline="") as f:
                w = csv.writer(f)
                w.writerow(cols)
                for r in self.rows_cmp:
                    w.writerow([r.get(c, "") for c in cols])
        except Exception as e:
            messagebox.showerror("Export failed", f"Failed to write CSV:\n{e}")
            return
        messagebox.showinfo("Exported", f"Saved: {os.path.basename(path)}")

    def export_html(self):
        if not self.rows_cmp:
            messagebox.showwarning("No data", "Run Compare first."); return
        folder = filedialog.askdirectory(title="Choose folder to save HTML", mustexist=True)
        if not folder: return
        ts = time.strftime("%Y%m%d_%H%M%S")
        path = os.path.join(folder, f"delta_uniqueFPC_{self.sort_mode.get()}_{ts}.html")
        cols = list(self.tree["columns"])
        cat_color = {
            "Both (same)": "#F5F5F5",
            "Both (increased)": "#E6FFED",
            "Both (decreased)": "#FFEEEF",
            "Only A": "#FDECEC",
            "Only B": "#E8F8E8",
        }
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write("<!doctype html><html><head><meta charset='utf-8'><title>Delta Report</title>")
                f.write("<style>body{font-family:Segoe UI,Arial,sans-serif;padding:16px}table{border-collapse:collapse;width:100%}th,td{border:1px solid #ddd;padding:6px 8px;font-size:13px}th{background:#fafafa;text-align:center}td{text-align:center}tr:hover{filter:brightness(0.98)}</style>")
                f.write(f"<h2>Delta Report — Unique FPC | Sort: {self.sort_mode.get()}</h2>")
                f.write("<table><thead><tr>")
                for c in cols: f.write(f"<th>{c}</th>")
                f.write("</tr></thead><tbody>")
                for r in self.rows_cmp:
                    bg = cat_color.get(r.get("Category",""), "#fff")
                    f.write(f"<tr style='background:{bg}'>")
                    for c in cols:
                        val = (r.get(c, "") or "")
                        val = str(val).replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
                        f.write(f"<td>{val}</td>")
                    f.write("</tr>")
                f.write("</tbody></table></body></html>")
        except Exception as e:
            messagebox.showerror("Export failed", f"Failed to write HTML:\n{e}")
            return
        messagebox.showinfo("Exported", f"Saved: {os.path.basename(path)}")

if __name__ == "__main__":
    app = App()
    app.mainloop()
