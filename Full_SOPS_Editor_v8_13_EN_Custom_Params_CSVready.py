#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
full sops editor — READY v8.12 (compact preview, Euro6 bulk-select)
- Preview shows: FPC_ID, LONG, CODE, DESCRIPTION
- CODE/DESCRIPTION displayed as "Before → After" when changed
- "adblue euro 6 fix" uses exact user mapping and auto-selects all affected rows present
"""

import os, re, csv, traceback, sys
from pathlib import Path
from typing import Dict, Any, List, Optional, Tuple
import xml.etree.ElementTree as ET

from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

APP_TITLE = "Full SOPS Parameter Editor — v8.13 EN (Red Save)"
APP_GEOM  = "1200x680"

# ================== Fixed controller lists ==================
ENGINE_LIST  = [11, 32, 2250, 142, 408, 3127]
GEARBOX_LIST = [17, 37, 5731, 2519, 3129, 3575, 4099, 3508]
SCR_LIST     = [3636, 3719, 9452, 2471, 3459, 4149, 2409, 4280]  # ADBLUE = SCR

# Placeholders (tweakable later)
PLACE_SMS      = [86,1020,1169,1170,1175,1176,1528,3285,3645,4031,5413,5688,5739,5740,6246,7141,7277,7555,7556,8402,33,34,122,257,2973,2974,3135,3226,3663,3912,8233]
PLACE_CHASSIS  = [22, 68, 1019, 2987, 8074, 448, 1334, 3226]
PLACE_ICL      = [442, 2301, 3117, 3118, 3888, 6681, 8303]
PLACE_RETARDER = [111, 510, 3131, 3132, 5056, 8096]
PLACE_CLUTCH   = [89, 2999, 3173, 3517, 3575, 4099, 4939]
PLACE_HEATER   = [9501, 9502]
PLACE_BRAKE    = [27, 125, 1255, 5848, 5909, 4551, 3875, 3874, 3872, 3759, 3478, 3465, 3134, 3133, 2452, 2440, 1521, 5921, 6146, 6147, 6171, 6220, 6302, 6303, 7131, 7575, 8053, 8076, 8501, 8646, 9403]
PLACE_APS      = [9701]
PLACE_BWE      = [9801, 9802]
PLACE_BWS      = [9901]
PLACE_PTO      = [3203, 3204, 3205, 3462, 3468, 3469, 3498, 3514, 3515, 3543, 3544, 3545, 3853, 3877, 3896, 4303, 4787, 5877, 5950, 5951, 7623]
PLACE_EGR      = [2409, 2471, 3636, 3782]
PLACE_LAS      = [9808]

CONTROLLERS = [
    "ALL",
    "ENGINE", "LAS",
    "GEARBOX",
    "SCR", "ADBLUE", "EGR",
    "SMS", "CHASSIS", "ICL", "RETARDER",
    "CLUTCH", "HEATER", "BRAKE", "APS", "BWE", "BWS", "PTO"
]

CONTROLLER_PARAMS: Dict[str, List[int]] = {
    "ENGINE":   ENGINE_LIST,
    "GEARBOX":  GEARBOX_LIST,
    "SCR":      SCR_LIST,
    "ADBLUE":   SCR_LIST,      # = SCR
    "SMS":      PLACE_SMS,
    "CHASSIS":  PLACE_CHASSIS,
    "ICL":      PLACE_ICL,
    "RETARDER": PLACE_RETARDER,
    "CLUTCH":   PLACE_CLUTCH,
    "HEATER":   PLACE_HEATER,
    "BRAKE":    PLACE_BRAKE,
    "APS":      PLACE_APS,
    "BWE":      PLACE_BWE,
    "BWS":      PLACE_BWS,
    "PTO":      PLACE_PTO,
    "EGR":      PLACE_EGR,
    "LAS":      PLACE_LAS,
}

def norm(code: str) -> str:
    return re.sub(r"[^A-Za-z0-9]+", "", str(code or "")).upper()

def letter_prefix(code: str) -> str:
    m = re.match(r"[A-Za-z]+", str(code or ""))
    return (m.group(0) if m else "").upper()

def _localname(tag: str) -> str:
    return tag.split('}', 1)[-1] if '}' in tag else tag

# ---------- XML loader (encoding tolerant) ----------
def load_xml_root_bytesafe(xml_path: Path):
    try:
        with open(xml_path, "rb") as f:
            return ET.parse(f).getroot()
    except Exception:
        pass
    data = xml_path.read_bytes()
    try:
        import chardet  # optional
        det = chardet.detect(data or b"")
        encoding = (det.get("encoding") if det else None) or "utf-8"
    except Exception:
        encoding = "utf-8"
    text = data.decode(encoding, errors="replace").replace("\x00", "")
    return ET.fromstring(text.encode("utf-8"))

# ---------- mapping loaders (CSV / XLSX / PDF) ----------
def _norm_header(h: str) -> str:
    h = (h or "").strip().lower()
    h = re.sub(r"[\s_\-]+", "", h)
    return h

def _pick_column(headers, want: str, aliases: List[str]) -> Optional[str]:
    want_n = _norm_header(want)
    norm_map = {_norm_header(h): h for h in headers}
    if want_n in norm_map: return norm_map[want_n]
    for a in aliases:
        a_n = _norm_header(a)
        if a_n in norm_map: return norm_map[a_n]
    return None

def load_mapping_csv_any(path: Path, log=None):
    data = path.read_text(encoding="utf-8-sig", errors="replace")
    try:
        sample = "\n".join(data.splitlines()[:3]) or ","
        dialect = csv.Sniffer().sniff(sample)
        delim = dialect.delimiter
    except Exception:
        delim = ";" if data.count(";") > data.count(",") else ","
    rows = list(csv.reader(data.splitlines(), delimiter=delim))
    if not rows: raise ValueError("CSV seems empty.")
    headers = rows[0]
    if log: log(f"[INFO] CSV headers: {headers} | delimiter='{delim}'")

    col_id    = _pick_column(headers, "FPC_ID",      ["fpc","fpcid","id","name"])
    col_code  = _pick_column(headers, "CODE",        ["code","value","val"])
    col_desc  = _pick_column(headers, "DESCRIPTION", ["description","desc","meaning"])
    col_long  = _pick_column(headers, "LONG",        ["long","parameter","group","category"])
    col_short = _pick_column(headers, "SHORT",       ["short","label"])
    if not col_id: raise ValueError(f"Missing FPC_ID column. Got: {headers}")
    idx = {h: headers.index(h) for h in headers}

    mapping_desc: Dict[int, Dict[str, str]] = {}
    mapping_long: Dict[int, str] = {}
    accepted = 0
    for r in rows[1:]:
        if not r: continue
        try: fid = int(str(r[idx[col_id]]).strip())
        except Exception: continue

        long_val = ""
        if col_long and idx[col_long] < len(r):
            long_val = str(r[idx[col_long]] or "").strip()
        if not long_val and col_short and idx[col_short] < len(r):
            long_val = str(r[idx[col_short]] or "").strip()
        if long_val and fid not in mapping_long:
            mapping_long[fid] = long_val

        if col_code and col_desc and idx[col_code] < len(r) and idx[col_desc] < len(r):
            code = str(r[idx[col_code]] or "").strip()
            desc = str(r[idx[col_desc]]).strip()
            if code and desc:
                b = mapping_desc.setdefault(fid, {})
                up = code.upper()
                b.setdefault(up, desc)
                b.setdefault(norm(code), desc)
                pref = letter_prefix(code)
                if pref: b.setdefault(pref, desc)
                accepted += 1

    if log:
        log(f"[OK] CSV parsed: rows={len(rows)-1} | accepted={accepted} "
            f"| FPCs(long)={len(mapping_long)} | FPCs(desc)={len(mapping_desc)}")
    if not mapping_long and not mapping_desc:
        raise ValueError("CSV mapping parsed empty after normalization.")
    return mapping_desc, mapping_long

def load_mapping_xlsx(path: Path, log=None):
    try:
        import pandas as pd  # type: ignore
    except Exception:
        raise RuntimeError("Reading XLSX requires pandas + openpyxl.")
    try:
        xl = pd.ExcelFile(path)
    except Exception as e:
        raise RuntimeError(f"Cannot open XLSX: {e}")

    last_err = None
    for sheet in xl.sheet_names:
        try:
            df = xl.parse(sheet).fillna("")
            headers = list(df.columns)
            if log: log(f"[INFO] XLSX sheet '{sheet}' headers: {headers}")
            col_id    = _pick_column(headers, "FPC_ID",      ["fpc","fpcid","id","name"])
            col_code  = _pick_column(headers, "CODE",        ["code","value","val"])
            col_desc  = _pick_column(headers, "DESCRIPTION", ["description","desc","meaning"])
            col_long  = _pick_column(headers, "LONG",        ["long","parameter","group","category"])
            col_short = _pick_column(headers, "SHORT",       ["short","label"])
            if not col_id: continue

            mapping_desc: Dict[int, Dict[str, str]] = {}
            mapping_long: Dict[int, str] = {}
            accepted = 0
            for _, row in df.iterrows():
                try: fid = int(str(row[col_id]).strip())
                except Exception: continue
                long_val = str(row[col_long]).strip() if col_long else ""
                if not long_val and col_short:
                    long_val = str(row[col_short]).strip()
                if long_val and fid not in mapping_long:
                    mapping_long[fid] = long_val
                if col_code and col_desc:
                    code = str(row[col_code]).strip()
                    desc = str(row[col_desc]).strip()
                    if code and desc:
                        b = mapping_desc.setdefault(fid, {})
                        up = code.upper()
                        b.setdefault(up, desc)
                        b.setdefault(norm(code), desc)
                        pref = letter_prefix(code)
                        if pref: b.setdefault(pref, desc)
                        accepted += 1
            if mapping_long or mapping_desc:
                if log: log(f"[OK] XLSX parsed '{sheet}': accepted={accepted}")
                return mapping_desc, mapping_long
        except Exception as e:
            last_err = e; continue
    raise RuntimeError(f"No valid sheet (needs FPC_ID & Long/Short). Last error: {last_err}")

def load_mapping_pdf(path: Path):
    mapping_desc: Dict[int, Dict[str, str]] = {}
    mapping_long: Dict[int, str] = {}
    def set_long(fid: int, v: str):
        if v and fid not in mapping_long: mapping_long[fid] = v
    def add(fid: int, code_raw: str, desc: str):
        if not code_raw or not desc: return
        b = mapping_desc.setdefault(fid, {})
        up = code_raw.strip().upper()
        b.setdefault(up, desc)
        b.setdefault(norm(code_raw), desc)
        pref = letter_prefix(code_raw)
        if pref: b.setdefault(pref, desc)
    try:
        import pdfplumber  # type: ignore
        with pdfplumber.open(str(path)) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables(table_settings={
                    "vertical_strategy": "lines",
                    "horizontal_strategy": "lines",
                    "intersection_tolerance": 5,
                    "snap_tolerance": 3,
                    "join_tolerance": 3,
                }) or []
                for tbl in tables:
                    for row in tbl or []:
                        cells = [(c or "").strip() for c in row]
                        fidx = None
                        for i,c in enumerate(cells):
                            if re.fullmatch(r"\d{1,4}", c):
                                fidx = i; break
                        if fidx is None: continue
                        try: fid = int(cells[fidx])
                        except: continue
                        tail = [c for c in cells[fidx+1:] if c]
                        if not tail: continue
                        cand = [c for c in tail if re.fullmatch(r"[A-Za-z0-9/\+\-]{1,32}", c)]
                        if not cand: continue
                        code_raw = min(cand, key=len)
                        others = [c for c in tail if c != code_raw]
                        desc = max(others, key=len) if others else ""
                        maybe_long = min(others, key=len) if others else ""
                        if maybe_long: set_long(fid, maybe_long)
                        add(fid, code_raw, desc)
        if mapping_desc or mapping_long: return mapping_desc, mapping_long
    except Exception:
        pass
    try:
        import PyPDF2  # type: ignore
        txt = ""
        with open(path, "rb") as f:
            reader = PyPDF2.PdfReader(f)
            txt = "\n".join((p.extract_text() or "") for p in reader.pages)
        lines = [ln.strip() for ln in txt.splitlines() if ln.strip()]
        row_re = re.compile(r"^(\d+)\s+([A-Za-z0-9\.,\-/\+\(\) ]+?)\s+([A-Za-z0-9\.,\-/\+\(\) ]+?)\s+([A-Za-z0-9/\+\-]{1,32})\s+(.*)$")
        for ln in lines:
            m = row_re.match(ln)
            if not m: continue
            fpc_id, short, longn, code_raw, desc = m.groups()
            fid = int(fpc_id)
            if longn: set_long(fid, longn.strip())
            add(fid, code_raw, desc.strip())
    except Exception:
        pass
    if not mapping_desc and not mapping_long:
        raise RuntimeError("Failed to parse PDF mapping. Prefer CSV/XLSX.")
    return mapping_desc, mapping_long

def discover_library() -> Optional[Path]:
    env = os.environ.get("FULL_SOPS_LIB")
    if env:
        p = Path(env.replace("\\","/")).expanduser()
        if p.exists(): return p
    here = Path(__file__).resolve().parent
    for name in ("sops_fpc_mapping.csv", "sops_fpc_mapping.xlsx", "sopslist.html.pdf"):
        p = here / name
        if p.exists(): return p
    desk = Path.home() / "Desktop" / "SOPS app"
    for name in ("sops_fpc_mapping.csv", "sops_fpc_mapping.xlsx", "sopslist.html.pdf"):
        p = desk / name
        if p.exists(): return p
    return None

def _load_library_with_logger(p: Path, logger=None):
    ext = p.suffix.lower()
    if ext == ".csv":
        mapping_desc, mapping_long = load_mapping_csv_any(p, log=logger)
    elif ext in (".xlsx", ".xls"):
        mapping_desc, mapping_long = load_mapping_xlsx(p, log=logger)
    elif ext == ".pdf":
        mapping_desc, mapping_long = load_mapping_pdf(p)
        if logger: logger("[WARN] Using PDF mapping; CSV/XLSX is more reliable.")
    else:
        raise ValueError(f"Unsupported library type: {p.suffix}")
    return mapping_desc, mapping_long, str(p)

def read_fpcs_from_xml(xml_path: Path) -> List[Dict[str, Optional[str]]]:
    root = load_xml_root_bytesafe(xml_path)
    items: List[Dict[str, Optional[str]]] = []
    for e in root.iter():
        if _localname(e.tag).lower() == "fpc":
            name = e.attrib.get("Name") or e.attrib.get("name")
            value = e.attrib.get("Value") or e.attrib.get("value")
            upd   = e.attrib.get("Updated") or e.attrib.get("updated") or ""
            if not name or not value: continue
            try: fid = int(str(name))
            except: continue
            items.append({"id": fid, "code": value, "updated": upd})
    if not items:
        for e in root.iter():
            name = e.attrib.get("id") or e.attrib.get("fpc")
            value = e.attrib.get("value") or e.attrib.get("val")
            upd   = e.attrib.get("updated") or e.attrib.get("Updated") or ""
            if not name or not value: continue
            try: fid = int(str(name))
            except: continue
            items.append({"id": fid, "code": value, "updated": upd})
    return items

def resolve_long_and_desc(mapping_desc, mapping_long, fid: int, code_display: str):
    long_generic = mapping_long.get(fid, "")
    bucket = mapping_desc.get(fid, {})
    if not code_display: return long_generic, ""
    up = code_display.strip().upper()
    for key in (up, norm(code_display), letter_prefix(code_display)):
        if not key: continue
        desc = bucket.get(key)
        if desc: return long_generic, desc
    return long_generic, ""

# -------------------- Diagnostics --------------------
def _check_libs() -> Dict[str, Any]:
    out = {}
    for mod in ["pandas", "openpyxl", "pdfplumber", "PyPDF2", "chardet"]:
        try:
            __import__(mod)
            out[mod] = "OK"
        except Exception as e:
            out[mod] = f"Missing: {e.__class__.__name__}"
    try:
        from tkinter import ttk as _ttk
        _ = _ttk.Style().theme_names()
        out["tkinter"] = "OK"
    except Exception as e:
        out["tkinter"] = f"Error: {e.__class__.__name__}"
    return out

def _format_env_report(mapping_path: Optional[str], mapping_desc, mapping_long, xml_path: Optional[Path]) -> str:
    lines = []
    lines.append(f"[ENV] FULL_SOPS_LIB={os.environ.get('FULL_SOPS_LIB')!r}")
    lines.append(f"[ENV] Mapping file detected: {mapping_path or 'None'}")
    lines.append(f"[ENV] FPCs(long)={len(mapping_long)} | FPCs(desc)={len(mapping_desc)}")
    for k,v in _check_libs().items():
        lines.append(f"[DEP] {k}: {v}")
    if xml_path:
        lines.append(f"[XML] Selected: {xml_path}")
    else:
        lines.append("[XML] None selected.")
    return "\n".join(lines)

# ========================= GUI =========================

# ---------- custom packs (editable) ----------
CUSTOM_PACKS = {
    "custom pack 1": {2409: "Z", 3459: "Z", 3636: "Z"},
    "custom pack 2": {2471: "A", 3719: "Z"},
    "custom pack 3": {4149: "Z", 4295: "Z", 4374: "Z"},
    "custom pack 4": {4280: "A", 4290: "B"},
    "custom pack 5": {2331592: "Z", 2331593: "Z"},
    "custom pack 6": {2402250: "A"},
    "custom pack 7": {2200001: "B", 2200002: "C"},
    "custom pack 8": {3456: "Z", 3476: "Z", 3488: "Z"},
    "custom pack 9": {3510: "A", 3511: "A"},
    "custom pack 10": {3600: "Z", 3610: "Z", 3620: "Z"},
}


class App(ttk.Frame):
    def __init__(self, master: tk.Tk) -> None:
        super().__init__(master)
        self.master = master
        master.title(APP_TITLE); master.geometry(APP_GEOM)

        self.xml_path: Optional[Path] = None
        # (id,long,code,desc,updated,flase)
        self.rows_all: List[Tuple[Any, Any, Any, Any, Any, Any]] = []
        self.modified_codes: Dict[int, str] = {}
        self._typing_job: Optional[str] = None

        # Theme
        style = ttk.Style()
        themes = style.theme_names()
        if sys.platform.startswith("win") and "vista" in themes:
            style.theme_use("vista")
        elif "clam" in themes:
            style.theme_use("clam")

        
        self._apply_styles(style)
        # ---- Default parameters for placeholder actions (editable) ----
        self._default_params = {
            "immo fix":                 {2250: "Z"},
            "sms disable":              {3285: "Z"},
            "retarder disable":         {3131: "A", 3132: "A"},
            "clutch off":               {3173: "Z", 3517: "Z"},
            "pto type":                 {3468: "A"},
            "cylinder off match off":   {2409: "Z"},
            "ess off":                  {3636: "Z"},
            "flight off":               {9808: "A"},
        }
        # === Display names -> internal keys (edit LEFT freely) ===
        self._modify_display_to_key = {
            "AdBlue Euro 5 Off":        "adblue euro 5 fix",
            "AdBlue Euro 6 Off":        "adblue euro 6 fix",
            "IMMO Fix":                 "immo fix",
            "Disable SMS":              "sms disable",
            "Disable Retarder":         "retarder disable",
            "Clutch Off":               "clutch off",
            "PTO Type":                 "pto type",
            "Cylinder Off Match Off":   "cylinder off match off",
            "ESS Off":                  "ess off",
            "Flight Mode Off":          "flight off",
            "Custom Pack 1":            "custom pack 1",
            "Custom Pack 2":            "custom pack 2",
            "Custom Pack 3":            "custom pack 3",
            "Custom Pack 4":            "custom pack 4",
            "Custom Pack 5":            "custom pack 5",
            "Custom Pack 6":            "custom pack 6",
            "Custom Pack 7":            "custom pack 7",
            "Custom Pack 8":            "custom pack 8",
            "Custom Pack 9":            "custom pack 9",
            "Custom Pack 10":           "custom pack 10",
        }

        # === External CSV parameters override ===
        self._csv_params = {}
        try:
            import csv, os
            csv_path = os.path.join(os.path.dirname(__file__), "sops_fpc_mapping.csv")
            if os.path.isfile(csv_path):
                with open(csv_path, "r", encoding="utf-8-sig", newline="") as f:
                    reader = csv.DictReader(f)
                    rows = list(reader)
                cols = {c.strip().lower(): c for c in (reader.fieldnames or [])}
                col_action = cols.get("action")
                col_fpc = cols.get("fpc") or cols.get("fpc_id") or cols.get("fpc- id") or cols.get("fpc id")
                col_value = cols.get("value") or cols.get("code")
                if col_fpc and col_value:
                    if col_action:
                        for r in rows:
                            act = (r.get(col_action) or "").strip().lower()
                            if not act:
                                continue
                            try:
                                fid = int(str(r.get(col_fpc)).strip())
                            except Exception:
                                continue
                            val = str(r.get(col_value)).strip()
                            self._csv_params.setdefault(act, {})[fid] = val
                    else:
                        g = {}
                        for r in rows:
                            try:
                                fid = int(str(r.get(col_fpc)).strip())
                            except Exception:
                                continue
                            val = str(r.get(col_value)).strip()
                            g[fid] = val
                        if g:
                            self._csv_params["*"] = g
                self._log(f"[OK] CSV loaded: {csv_path}")
        except Exception as e:
            self._log(f"[WARN] CSV load failed: {e}")
    
        self._build_ui()
        # Flush any boot logs captured before the text widget existed
        try:
            if getattr(self, "_boot_logs", None):
                for _m in self._boot_logs:
                    try:
                        self.txt.insert("end", _m + "\n")
                    except Exception:
                        print(_m)
                self.txt.see("end")
                self._boot_logs.clear()
        except Exception:
            pass


        # Modify actions
        self._modify_registry = {
            "adblue euro 5 off":        self._action_adblue_euro5_fix,
            "adblue euro 6 fix":        self._action_adblue_euro6_fix,
            "immo fix":                 self._action_immo_fix,
            "sms disable":              self._action_sms_disable,
            "retarder disable":         self._action_retarder_disable,
            "clutch off":               self._action_clutch_off,
            "pto type":                 self._action_pto_type,
            "cylinder off match off":   self._action_cyl_off_match_off,
            "ess off":                  self._action_ess_off,
            "flight off":               self._action_flight_off,
            "custom pack 1":            lambda: self._run_custom_pack("custom pack 1"),
            "custom pack 2":            lambda: self._run_custom_pack("custom pack 2"),
            "custom pack 3":            lambda: self._run_custom_pack("custom pack 3"),
            "custom pack 4":            lambda: self._run_custom_pack("custom pack 4"),
            "custom pack 5":            lambda: self._run_custom_pack("custom pack 5"),
            "custom pack 6":            lambda: self._run_custom_pack("custom pack 6"),
            "custom pack 7":            lambda: self._run_custom_pack("custom pack 7"),
            "custom pack 8":            lambda: self._run_custom_pack("custom pack 8"),
            "custom pack 9":            lambda: self._run_custom_pack("custom pack 9"),
            "custom pack 10":           lambda: self._run_custom_pack("custom pack 10"),
        }

        try:
            p = discover_library()
            if not p: raise FileNotFoundError("Mapping library not found (CSV/XLSX/PDF).")
            self.mapping_desc, self.mapping_long, self.mapping_path = _load_library_with_logger(p, self._log)
            lib_msg = f"[OK] Library: {Path(self.mapping_path).name} | FPCs(long)={len(self.mapping_long)} | FPCs(desc)={len(self.mapping_desc)}"
        except Exception as e:
            self.mapping_desc, self.mapping_long, self.mapping_path = {}, {}, None
            lib_msg = f"[ERROR] {e}"

        self._log("Ready. Open XML → Analyze. Change Controller or Search. Right-click table to copy.")
        self._log(lib_msg)

    def _apply_styles(self, style: ttk.Style):
        style.configure("Treeview", font=("Segoe UI", 10), rowheight=24)
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))
        style.configure("TButton", padding=(10, 4))
        style.configure("TLabel", font=("Segoe UI", 10))
        style.configure("TEntry", padding=2)
        style.configure("TCombobox", padding=2)

    # --- UI construction ---
    def _build_ui(self) -> None:
        top = ttk.Frame(self.master, padding=(8, 6, 8, 4))
        top.pack(fill="x")

        ttk.Button(top, text="Open XML", command=self.on_open_xml).pack(side="left")
        ttk.Button(top, text="Analyze",  command=self.on_analyze).pack(side="left", padx=(6,0))
        ttk.Button(top, text="Save XML", command=self.on_save_xml).pack(side="left", padx=(6,0))
        ttk.Button(top, text="Export CSV", command=self.on_export_csv).pack(side="left", padx=(6,0))
        
        ttk.Label(top, text="Modify:").pack(side="left", padx=(6,4))
        self.var_modify = tk.StringVar(value="— select action —")
        self.cmb_modify = ttk.Combobox(
            top,
            textvariable=self.var_modify,
            state="readonly",
            width=24,
            values=["— select action —"] + list(self._modify_display_to_key.keys()),
        )
        self.cmb_modify.pack(side="left", padx=(0,10))

        self.btn_modify_apply = ttk.Button(top, text="Apply", command=self.on_modify_choice)
        self.btn_modify_apply.pack(side="left", padx=(6,10))
        ttk.Label(top, text="Controller:").pack(side="left")
        self.var_controller = tk.StringVar(value="ALL")
        self.cmb_controller = ttk.Combobox(top, textvariable=self.var_controller, values=CONTROLLERS, state="readonly", width=16)
        self.cmb_controller.pack(side="left", padx=(4,8))
        self.cmb_controller.bind("<<ComboboxSelected>>", lambda e: self.apply_filter())

        ttk.Label(top, text="Search:").pack(side="left")
        self.var_query = tk.StringVar()
        self.ent_query = ttk.Entry(top, textvariable=self.var_query, width=22)
        self.ent_query.pack(side="left", padx=(4,0))
        self.ent_query.bind("<KeyRelease>", lambda e: self.apply_filter())

        paned = ttk.PanedWindow(self.master, orient="horizontal")
        paned.pack(fill="both", expand=True)

        left = ttk.Frame(paned, padding=(8, 4, 4, 8))
        paned.add(left, weight=4)

        cols = ("fpc_id","long","code","description")
        headers = ("FPC_ID","LONG","CODE","DESCRIPTION")
        self.tree = ttk.Treeview(left, columns=cols, show="headings", height=18, selectmode="extended")
        # highlight modified rows in red
        self.tree.tag_configure("modified", foreground="red")
        for c, label in zip(cols, headers):
            self.tree.heading(c, text=label, anchor="center")
            self.tree.column(c, anchor="center", width=120)
        self.tree.grid(row=0, column=0, sticky="nsew")

        self.tree.bind('<<TreeviewSelect>>', self._update_preview)
        self.tree.bind("<<TreeviewSelect>>", self.on_row_select)

        left.columnconfigure(0, weight=1); left.rowconfigure(0, weight=1)
        ttk.Scrollbar(left, orient="vertical", command=self.tree.yview).grid(row=0, column=1, sticky="ns")
        ttk.Scrollbar(left, orient="horizontal", command=self.tree.xview).grid(row=1, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=lambda *args: None, xscrollcommand=lambda *args: None)

        self.menu = tk.Menu(self.master, tearoff=0)
        self.menu.add_command(label="Copy row(s)", command=self.copy_rows)
        self.menu.add_command(label="Copy cell", command=self.copy_cell)
        self.tree.bind("<Button-3>", self._popup)

        right = ttk.Frame(paned, padding=(4, 8, 8, 8))
        paned.add(right, weight=2)

        editor = ttk.Frame(right)
        editor.pack(fill="x")
        ttk.Label(editor, text="Parameter Editor", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(0,6))

        form = ttk.Frame(right)
        form.pack(fill="x")
        for i in range(2): form.columnconfigure(i, weight=1)

        ttk.Label(form, text="FPC_ID:").grid(row=0, column=0, sticky="w", padx=(0,6), pady=(0,6))
        self.var_fpc = tk.StringVar(value="")
        self.ent_fpc = ttk.Entry(form, textvariable=self.var_fpc, width=18)
        self.ent_fpc.grid(row=0, column=1, sticky="ew", pady=(0,6))
        self.ent_fpc.bind("<KeyRelease>", self.on_fpc_typed)
        self.ent_fpc.bind("<Return>", self.on_fpc_enter)

        ttk.Label(form, text="Current CODE:").grid(row=1, column=0, sticky="w", padx=(0,6), pady=(0,6))
        self.var_code_cur = tk.StringVar(value="")
        self.ent_code_cur = ttk.Entry(form, textvariable=self.var_code_cur, state="readonly")
        self.ent_code_cur.grid(row=1, column=1, sticky="ew", pady=(0,6))

        ttk.Label(form, text="Options (from library):").grid(row=2, column=0, sticky="w", padx=(0,6), pady=(0,6))
        self.var_code_new = tk.StringVar(value="")
        self.cmb_code_new = ttk.Combobox(form, textvariable=self.var_code_new, values=[], state="readonly", height=12, width=28)
        self.cmb_code_new.grid(row=2, column=1, sticky="ew", pady=(0,6))
        self.cmb_code_new.bind('<<ComboboxSelected>>', self._update_preview)

        row_tools = ttk.Frame(form)
        row_tools.grid(row=3, column=0, columnspan=2, sticky='ew', pady=(6,0))
        ttk.Button(row_tools, text='Preview…', command=self.on_preview_popup).pack(side='left')
        ttk.Button(row_tools, text='Apply', command=self.on_apply_code).pack(side='right')

        self.var_b_code = tk.StringVar(); self.var_b_desc = tk.StringVar(); self.var_b_long = tk.StringVar()
        self.var_a_code = tk.StringVar(); self.var_a_desc = tk.StringVar(); self.var_a_long = tk.StringVar()

        logf = ttk.Frame(self.master, padding=(8, 0, 8, 6))
        logf.pack(fill="both", expand=False)
        self.txt = tk.Text(logf, height=6, bg="#101315", fg="#a7ffb5", insertbackground="white", font=("Consolas", 10), borderwidth=0, highlightthickness=0)
        self.txt.pack(fill="both", expand=True)

        self.status = ttk.Label(self.master, text="Ready", anchor="w", padding=(8,2))
        self.status.pack(fill="x")

        self.master.after(0, self._update_preview)

    # ---------- helpers ----------
    def _popup(self, event):
        try:
            self.menu.tk_popup(event.x_root, event.y_root, 0)
        finally:
            self.menu.grab_release()

    def copy_rows(self):
        sel = self.tree.selection()
        if not sel: return
        lines = []
        for iid in sel:
            vals = self.tree.item(iid, "values")
            lines.append("\t".join(str(v) for v in vals))
        self.master.clipboard_clear()
        self.master.clipboard_append("\n".join(lines))
        self.status.configure(text=f"Copied {len(lines)} row(s)")

    def copy_cell(self):
        sel = self.tree.selection()
        if not sel: return
        iid = sel[0]
        vals = self.tree.item(iid, "values")
        cell = vals[2] if len(vals) > 2 else ""
        self.master.clipboard_clear()
        self.master.clipboard_append(str(cell))
        self.status.configure(text="Copied CODE cell")

    def _log(self, msg: str) -> None:
        self.txt.insert("end", msg + "\n"); self.txt.see("end")
        self.status.configure(text=msg.splitlines()[-1][:160])

    # ----- Modify handler -----
    def on_modify_choice(self, event=None):
        choice_raw = (self.var_modify.get() or "").strip()
        if not choice_raw or choice_raw.startswith("—"):
            return
        choice = self._modify_display_to_key.get(choice_raw, choice_raw).lower()
        fn = getattr(self, "_modify_registry", {}).get(choice)
        if fn is None:
            self._log(f"[MODIFY] No handler for: {choice}")
        else:
            try:
                fn()
                self._log(f"[MODIFY] Executed: {choice}")
            except Exception as e:
                self._log(f"[ERROR] Modify '{choice}': {e}")

    # ----- Placeholder actions -----
    def _action_adblue_euro5_fix(self):
        mapping = {2471: "A", 3459: "Z", 3476: "Z", 3488: "Z", 3636: "Z", 3719: "Z"}
        if not self.rows_all:
            self._log("[EDIT] No data loaded. Please Analyze first."); return
        self._bulk_preview_and_apply(mapping, "adblue euro 5 fix")

    
    def _action_adblue_euro6_fix(self):
        """Preview-first Euro6 mapping: target A for all mapped IDs; 2471 stays A."""
        mapping = {
    2471: 'A',   # Emission level stays A
    3459: 'Z',
    3476: 'Z',
    3488: 'Z',
    3636: 'Z',
    3719: 'Z',
    4149: 'Z',
    4295: 'Z',
    4374: 'Z',
}
        if not self.rows_all:
            self._log("[EDIT] No data loaded. Please Analyze first."); return
        self._bulk_preview_and_apply(mapping, "adblue euro 6 fix")

    

    def _action_immo_fix(self):
        mapping = self._default_params.get("immo fix", {})
        if mapping: self._bulk_preview_and_apply(mapping, "immo fix")
        else: self._log("[ACTION] immo fix (no defaults set)")
    def _action_sms_disable(self):
        mapping = self._default_params.get("sms disable", {})
        if mapping: self._bulk_preview_and_apply(mapping, "sms disable")
        else: self._log("[ACTION] sms disable (no defaults set)")
    def _action_retarder_disable(self):
        mapping = self._default_params.get("retarder disable", {})
        if mapping: self._bulk_preview_and_apply(mapping, "retarder disable")
        else: self._log("[ACTION] retarder disable (no defaults set)")
    def _action_clutch_off(self):
        mapping = self._default_params.get("clutch off", {})
        if mapping: self._bulk_preview_and_apply(mapping, "clutch off")
        else: self._log("[ACTION] clutch off (no defaults set)")
    def _action_pto_type(self):
        mapping = self._default_params.get("pto type", {})
        if mapping: self._bulk_preview_and_apply(mapping, "pto type")
        else: self._log("[ACTION] pto type (no defaults set)")
    def _action_cyl_off_match_off(self):
        mapping = self._default_params.get("cylinder off match off", {})
        if mapping: self._bulk_preview_and_apply(mapping, "cylinder off match off")
        else: self._log("[ACTION] cylinder off match off (no defaults set)")
    def _action_ess_off(self):
        mapping = self._default_params.get("ess off", {})
        if mapping: self._bulk_preview_and_apply(mapping, "ess off")
        else: self._log("[ACTION] ess off (no defaults set)")
    def _action_flight_off(self):
        mapping = self._default_params.get("flight off", {})
        if mapping: self._bulk_preview_and_apply(mapping, "flight off")
        else: self._log("[ACTION] flight off (no defaults set)")

    def _auto_fit_columns(self) -> None:
        for col in self.tree["columns"]:
            data = [len(str(self.tree.set(k, col))) for k in self.tree.get_children('')]
            data.append(len(col))
            max_len = max(data) if data else len(col)
            width = max(90, min(540, int(max_len * 7.0)))
            self.tree.column(col, width=width, anchor="center")

    def _allowed_ids_for_controller(self, ctrl: str) -> Optional[set]:
        ctrl = (ctrl or "").upper()
        if ctrl == "ALL":
            return None
        ids = CONTROLLER_PARAMS.get(ctrl, [])
        return set(int(x) for x in ids)

    def _build_display_options(self, fid: int, current_code: str) -> List[str]:
        bucket = (self.mapping_desc.get(fid, {}) or {})
        seen = {}
        for k, v in bucket.items():
            key = (k or "").strip()
            if not key or not key.isalnum():
                continue
            ku = key.upper()
            if ku not in seen:
                seen[ku] = (key, v or "")
        items = sorted(seen.values(), key=lambda kv: (len(kv[0]), kv[0].upper()))
        display = [f"{k} — {d}" if d else k for (k, d) in items]
        if current_code:
            cur_up = current_code.strip().upper()
            if cur_up not in seen:
                display = [current_code] + display
        return display

    
    def _bulk_preview_and_apply(self, mapping: dict[int, str], title: str):
        import tkinter as tk
        from tkinter import ttk, messagebox

        all_rows = self.rows_all or []
        if not all_rows:
            messagebox.showinfo(APP_TITLE, "Analyze an XML first."); return

        id_to_idx = {int(r[0]): i for i, r in enumerate(all_rows)}
        # CSV override per action or global
        action_key = (title or '').strip().lower()
        if getattr(self, '_csv_params', None):
            if action_key in self._csv_params:
                mapping.update(self._csv_params[action_key])
            elif '*' in self._csv_params:
                mapping.update(self._csv_params['*'])
        preview_rows = []

        for fid, target_code in mapping.items():
            idx = id_to_idx.get(fid)
            if idx is None:
                continue
            fid0, long0, code_before, desc_before, upd0, _mark = all_rows[idx]
            long_after, desc_after = resolve_long_and_desc(self.mapping_desc, self.mapping_long, fid0, str(target_code))
            long_used = long0 or long_after
            code_show = f"{code_before} \u2192 {target_code}" if str(code_before) != str(target_code) else str(code_before)
            desc_show = (f"{desc_before} \u2192 {desc_after}" if desc_after and desc_after != desc_before
                         else (desc_before or desc_after or ""))
            preview_rows.append((fid0, long_used, code_show, desc_show, str(target_code), desc_after))

        if not preview_rows:
            messagebox.showinfo(APP_TITLE, "None of the mapped FPCs exist in this XML."); return

        win = tk.Toplevel(self.master)
        win.title(f"{title} — Preview"); win.transient(self.master); win.grab_set(); win.geometry("980x560"); win.minsize(860, 440)
        ttk.Label(win, text=f"Selected rows: {len(preview_rows)}",
                  font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=10, pady=(10,6))

        cols = ("fpc","long","code","description")
        headers = ("FPC_ID","LONG","CODE","DESCRIPTION")
        tv = ttk.Treeview(win, columns=cols, show="headings", selectmode="none", height=18)
        for c, h in zip(cols, headers):
            tv.heading(c, text=h, anchor="center")
            tv.column(c, anchor="center", width=150 if c!="description" else 320)
        tv.pack(fill="both", expand=True, padx=10, pady=(0,6))

        for fid0, longv, code_ba, desc_ba, _tgt, _desc_after in preview_rows:
            tv.insert("", "end", values=(fid0, longv, code_ba, desc_ba))

        def _apply_now():
            changed = 0
            for fid0, _longv, _code_ba, _desc_ba, tgt, desc_after in preview_rows:
                idx = id_to_idx.get(fid0)
                if idx is None:
                    continue
                f, longn, cur, desc, upd, _mark = all_rows[idx]
                if str(cur).upper() != str(tgt).upper():
                    all_rows[idx] = (f, longn, tgt, (desc_after or desc), upd, tgt)
                    self.modified_codes[int(fid0)] = tgt
                    changed += 1
            self.apply_filter()
            self._log(f"[EDIT] Applied '{title}' to {changed} row(s).")
            win.destroy()


    def _run_custom_pack(self, name: str):
        mapping = CUSTOM_PACKS.get(name, {})
        if not mapping:
            self._log(f"[EDIT] '{name}': mapping is empty."); return
        self._bulk_preview_and_apply(mapping, name)

        fx = ttk.Frame(win); fx.pack(side="bottom", fill="x", padx=10, pady=(8,10))
        ttk.Button(fx, text="Close", command=win.destroy).pack(side="right", padx=6)
        ttk.Button(fx, text="Apply", command=_apply_now).pack(side="right", padx=6)
# ---------- actions ----------
    def on_open_xml(self) -> None:
        path = filedialog.askopenfilename(title="Select SOPS XML", filetypes=[("XML files","*.xml"), ("All files","*.*")])
        if not path: return
        self.xml_path = Path(path)
        self.modified_codes.clear()
        self._log(f"[OK] XML selected: {self.xml_path}")

    def on_analyze(self) -> None:
        has_lib = (self.mapping_desc or self.mapping_long)
        if not has_lib:
            messagebox.showwarning(APP_TITLE, "No mapping library loaded."); return
        if not self.xml_path:
            messagebox.showwarning(APP_TITLE, "Select an XML file first."); return
        try:
            fpcs = read_fpcs_from_xml(self.xml_path)
            decoded: List[Tuple[Any,Any,Any,Any,Any,Any]] = []
            resolved = 0
            for it in fpcs:
                fid = it["id"]; code_disp = it["code"]; upd = it["updated"]
                longn, desc = resolve_long_and_desc(self.mapping_desc, self.mapping_long, fid, code_disp)
                if desc: resolved += 1
                decoded.append((fid, longn, code_disp, desc, upd, ""))  # flase empty
            self.rows_all = decoded
            self.apply_filter()
            self._log(f"[OK] XML parsed. items={len(decoded)} | resolved={resolved} | unresolved={len(decoded)-resolved}")
        except Exception as e:
            self._log(f"[ERROR] XML parse error:\n{e}")
            self._log(traceback.format_exc())

    def apply_filter(self) -> None:
        prev_id = self._current_selected_id()
        q = (self.var_query.get() or "").strip().lower()
        ctrl = self.var_controller.get().strip().upper()
        allowed = self._allowed_ids_for_controller(ctrl)

        def match(r: Tuple[Any,Any,Any,Any,Any,Any]) -> bool:
            if allowed is not None and int(r[0]) not in allowed:
                return False
            if not q: return True
            return any(q in str(x).lower() for x in r)

        rows = [r for r in self.rows_all if match(r)]
        self._refresh_table(rows)

        target = None
        if prev_id is not None and any(int(r[0]) == prev_id for r in rows):
            target = prev_id
        elif rows:
            target = int(rows[0][0])
        if target is not None:
            self._select_and_scroll_to(target)
            self.on_row_select()

    def _refresh_table(self, rows: List[Tuple[Any,Any,Any,Any,Any,Any]]) -> None:
        self.tree.delete(*self.tree.get_children())
        for idx, r in enumerate(rows):
            fid = int(r[0])
            tags = []
            try:
                flase_val = str(r[5]).strip()
            except Exception:
                flase_val = ''
            if flase_val and flase_val.lower() not in ('false','0','no','none'):
                tags.append('flase_alert')
            if fid in self.modified_codes: tags.append('modified')
            self.tree.insert("", "end", values=r[:4], tags=tuple(tags))
        self._auto_fit_columns()

    def _current_selected_id(self) -> Optional[int]:
        sel = self.tree.selection()
        if not sel: return None
        try:
            vals = self.tree.item(sel[0], "values")
            return int(vals[0])
        except Exception:
            return None

    def _visible_ids(self) -> List[int]:
        return [int(self.tree.item(i, "values")[0]) for i in self.tree.get_children()]

    def _select_and_scroll_to(self, fid: int) -> bool:
        for iid in self.tree.get_children():
            vals = self.tree.item(iid, "values")
            if not vals: continue
            if int(vals[0]) == fid:
                self.tree.selection_set(iid)
                self.tree.see(iid)
                return True
        return False

    def on_row_select(self, event=None):
        sel = self.tree.selection()
        if not sel: return
        item = sel[0]
        vals = self.tree.item(item, "values") or []
        if not vals: return
        fpc_id = vals[0]
        try:
            fid = int(fpc_id)
        except Exception:
            return
        id_to_idx = {int(r[0]): i for i, r in enumerate(self.rows_all)}
        if fid not in id_to_idx:
            return
        longn, code, desc = "", "", ""
        try:
            _, longn, code, desc, _, _ = self.rows_all[id_to_idx[fid]]
        except Exception:
            pass
        self.var_fpc.set(str(fpc_id))
        self.var_code_cur.set(str(code))
        options = self._build_display_options(fid, str(code))
        self.cmb_code_new["values"] = options
        if options: self.var_code_new.set(options[0])

    def on_fpc_typed(self, event=None):
        if hasattr(self, "_typing_job") and self._typing_job:
            try: self.master.after_cancel(self._typing_job)
            except Exception: pass
        self._typing_job = self.master.after(300, self._perform_fpc_lookup)

    def on_fpc_enter(self, event=None):
        if hasattr(self, "_typing_job") and self._typing_job:
            try: self.master.after_cancel(self._typing_job)
            except Exception: pass
        self._perform_fpc_lookup()

    def _perform_fpc_lookup(self):
        text = (self.var_fpc.get() or "").strip()
        if not text.isdigit(): return
        fid = int(text)
        all_ids = {int(r[0]) for r in self.rows_all}
        if fid not in all_ids:
            return
        if fid not in self._visible_ids():
            self.var_controller.set("ALL")
            self.apply_filter()
        self._select_and_scroll_to(fid)
        self.on_row_select()

    def on_preview_popup(self):
        """Compact preview for selected rows: FPC_ID, LONG, CODE (Before→After), DESCRIPTION (Before→After)."""
        try:
            idx = {'FPC_ID': 0, 'LONG': 1, 'CODE': 2, 'DESCRIPTION': 3}
            if hasattr(self, "tree") and self.tree is not None:
                cols = list(self.tree["columns"]) if self.tree["columns"] else []
                name_to_i = {}
                for i, c in enumerate(cols):
                    try:
                        label = (self.tree.heading(c, "text") or "").strip().upper()
                    except Exception:
                        label = ''
                    if label: name_to_i[label] = i
                for k in list(idx.keys()):
                    if k in name_to_i: idx[k] = name_to_i[k]

            sel = list(self.tree.selection())
            if not sel:
                kids = self.tree.get_children('')
                if kids: sel = [kids[0]]
            if not sel:
                self._log("[Preview] No rows to preview."); return

            try:
                opt_text = self.var_code_new.get() or ''
            except Exception:
                opt_text = ''

            arrow = " \u2192 "
            preview_rows = []
            for iid in sel:
                vals = self.tree.item(iid, "values") or []
                seq = vals.get('values', []) if isinstance(vals, dict) else vals
                fid = str(seq[idx['FPC_ID']]) if len(seq) > idx['FPC_ID'] else ''
                longv = seq[idx['LONG']] if len(seq) > idx['LONG'] else ''
                code_b = seq[idx['CODE']] if len(seq) > idx['CODE'] else ''
                desc_b = seq[idx['DESCRIPTION']] if len(seq) > idx['DESCRIPTION'] else ''
                code_a, desc_a = self._predict_after(code_b, desc_b, opt_text)
                code_b_a = f"{code_b}{arrow}{code_a}" if code_b != code_a else code_b
                desc_b_a = f"{desc_b}{arrow}{desc_a}" if desc_b != desc_a else desc_b
                preview_rows.append((fid, longv, code_b_a, desc_b_a))

            win = tk.Toplevel(self.master)
            win.title("Preview")
            win.transient(self.master); win.grab_set(); win.geometry("900x500")
            ttk.Label(win, text=f"Selected rows: {len(preview_rows)}", font=("Segoe UI", 10, "bold")).pack(anchor='w', padx=10, pady=(10,6))

            cols = ("fpc","long","code","description")
            headers = ("FPC_ID","LONG","CODE","DESCRIPTION")
            tv = ttk.Treeview(win, columns=cols, show='headings', selectmode='none', height=16)
            for c, h in zip(cols, headers):
                tv.heading(c, text=h, anchor='center')
                tv.column(c, anchor='center', width=140 if c!="description" else 260)
            tv.pack(fill='both', expand=True, padx=10)

            tv.tag_configure('odd', background='#f7f7f7')
            tv.tag_configure('even', background='#ffffff')
            for i, r in enumerate(preview_rows):
                tv.insert('', 'end', values=r, tags=('odd' if i%2==0 else 'even',))

            fx = ttk.Frame(win); fx.pack(fill='x', padx=10, pady=10)
            ttk.Button(fx, text='Apply', command=lambda: (self.on_apply_code(), win.destroy())).pack(side='right', padx=6)
            ttk.Button(fx, text='Close', command=win.destroy).pack(side='right')
        except Exception:
            traceback.print_exc()

    def on_apply_code(self):
        selected = (self.var_code_new.get() or "").strip()
        for sep in (" — ", " - ", " – "):
            if sep in selected:
                selected = selected.split(sep, 1)[0].strip()
                break
        new_code = selected
        if not new_code:
            messagebox.showinfo(APP_TITLE, "Select a code option first."); return
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo(APP_TITLE, "Select a row first."); return

        id_to_idx = {int(r[0]): i for i, r in enumerate(self.rows_all)}
        changed = 0
        for iid in sel:
            vals = self.tree.item(iid, "values")
            if not vals: continue
            fid = int(vals[0])
            idx = id_to_idx.get(fid)
            if idx is None: continue
            fid0, long0, code0, desc0, upd0, flase0 = self.rows_all[idx]
            if str(code0).upper() == new_code.upper(): continue
            long_after, desc_after = resolve_long_and_desc(self.mapping_desc, self.mapping_long, fid0, new_code)
            new_desc = (desc_after or desc0)
            new_long = (long0 or long_after)
            self.rows_all[idx] = (fid0, new_long, new_code, new_desc, upd0, new_code)
            self.modified_codes[fid] = new_code
            changed += 1

        if changed:
            self.apply_filter()
            self._log(f"[EDIT] Applied code '{new_code}' to {changed} row(s).")
            self.var_code_cur.set(new_code)

    def on_export_csv(self) -> None:
        if not self.rows_all:
            messagebox.showinfo(APP_TITLE, "Nothing to export yet."); return
        path = filedialog.asksaveasfilename(title="Save results as CSV",
                                            defaultextension=".csv",
                                            filetypes=[("CSV files",".csv")])
        if not path: return
        p = Path(path)
        try:
            with open(p, "w", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                w.writerow(["fpc_id","long","code","description","updated","flase"])
                for r in self.rows_all: w.writerow(list(r))
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Failed to save CSV:\n{e}"); return
        self._log(f"[OK] Saved CSV: {p}")

    
    
    def on_save_xml(self) -> None:
        from tkinter import filedialog, messagebox
        import re, traceback
        from pathlib import Path
        import xml.etree.ElementTree as ET

        if not self.xml_path:
            messagebox.showinfo(APP_TITLE, "Choose an XML first."); return
        if not self.modified_codes:
            messagebox.showinfo(APP_TITLE, "No modifications to save."); return

        src = Path(self.xml_path)

        # --- Suggest incremented filename for ANY XML ---
        full_name = src.name
        if full_name.lower().endswith('.xml'):
            base_no_xml = full_name[:-4]
        else:
            base_no_xml = src.stem
        m = re.search(r'(.*?)(?:\.)(\d+)(.*)$', base_no_xml)
        if m:
            prefix, num, tail = m.groups()
            new_num = str(int(num) + 1)
            suggested_no_xml = f"{prefix}.{new_num}{tail}"
        else:
            suggested_no_xml = base_no_xml + ".1"
        suggested_name = suggested_no_xml + ".xml"

        out_path = filedialog.asksaveasfilename(
            title="Save modified XML AS (original will remain unchanged)",
            defaultextension=".xml",
            initialfile=suggested_name,
            initialdir=str(src.parent),
            filetypes=[("XML files",".xml"), ("All files","*.*")]
        )
        if not out_path:
            return
        outp = Path(out_path)

        try:
            root = load_xml_root_bytesafe(self.xml_path)
            updated_cnt = 0
            for e in root.iter():
                if not e.attrib: continue
                name = e.attrib.get("Name") or e.attrib.get("name") or e.attrib.get("id") or e.attrib.get("fpc")
                if not name: continue
                try:
                    fid = int(str(name))
                except:
                    continue
                if fid in self.modified_codes:
                    new_code = self.modified_codes[fid]
                    if "Value" in e.attrib: e.attrib["Value"] = new_code
                    if "value" in e.attrib: e.attrib["value"] = new_code
                    if "Updated" in e.attrib: e.attrib["Updated"] = "true"
                    if "updated" in e.attrib: e.attrib["updated"] = "true"
                    if not any(k in e.attrib for k in ("Updated","updated")):
                        e.attrib["Updated"] = "true"
                    updated_cnt += 1

            ET.ElementTree(root).write(outp, encoding="utf-8", xml_declaration=True)

            try:
                txt = outp.read_text(encoding="utf-8", errors="ignore")
                def bump(mv): return f'MajorVersion="{int(mv.group(1))+1}"'
                txt2, nsub = re.subn(r'MajorVersion="(\d+)"', bump, txt, count=1)
                if nsub == 0:
                    def bump2(mv): return f'MajorVersion="{int(mv.group(1))+1}"'
                    txt2, nsub2 = re.subn(r'\bMajorVersion\s*=\s*"(\d+)"', bump2, txt, count=1)
                    nsub += nsub2
                if nsub > 0:
                    outp.write_text(txt2, encoding="utf-8")
                    self._log(f"[SAVE] Bumped MajorVersion by +1 in output.")
                else:
                    self._log("[SAVE] MajorVersion not found; left unchanged.")
            except Exception as e_bump:
                self._log(f"[WARN] Could not bump MajorVersion: {e_bump}")

            # --- Preserve original XML declaration (line 1) and newline style ---
            try:
                orig_text = Path(self.xml_path).read_text(encoding="utf-8", errors="ignore")
                # detect newline style from original
                orig_eol = "\r\n" if "\r\n" in orig_text else "\n"
                first_line = orig_text.splitlines()[0] if orig_text else ""
                if first_line.strip().startswith("<?xml"):
                    new_text = outp.read_text(encoding="utf-8", errors="ignore")
                    # splitlines(keepends=False) then rejoin with original EOL
                    lines = new_text.splitlines()
                    if lines:
                        lines[0] = first_line.strip()
                        new_text2 = orig_eol.join(lines) + (orig_eol if new_text.endswith(('\n', '\r', '\r\n')) else "")
                        outp.write_text(new_text2, encoding="utf-8")
                        self._log("[SAVE] Preserved original XML declaration and EOL style.")
            except Exception as e_decl:
                self._log(f"[WARN] Could not preserve XML declaration: {e_decl}")

            self._log(f"[SAVE] Saved XML: {outp} | modified FPCs: {updated_cnt}")
            messagebox.showinfo(APP_TITLE, f"Saved: {outp}\nModified FPCs: {updated_cnt}\n(Original kept: {src.name})")
        except Exception as e:
            self._log(f"[ERROR] Save failed:\n{e}")
            self._log(traceback.format_exc())
            messagebox.showerror(APP_TITLE, f"Save failed:\n{e}")

    def on_diag(self):
        try:
            rep = _format_env_report(getattr(self, "mapping_path", None), getattr(self, "mapping_desc", {}), getattr(self, "mapping_long", {}), self.xml_path)
            self._log(rep)
        except Exception as e:
            self._log(f"[DIAG-ERROR] {e}\n{traceback.format_exc()}")

    def _update_preview(self, event=None):
        idx = {'LONG': 1, 'CODE': 2, 'DESCRIPTION': 3}
        try:
            if hasattr(self, "tree") and self.tree is not None:
                cols = list(self.tree["columns"]) if self.tree["columns"] else []
                name_to_i = {}
                for i, c in enumerate(cols):
                    try:
                        label = (self.tree.heading(c, "text") or "").strip().upper()
                    except Exception:
                        label = ""
                    if label: name_to_i[label] = i
                for k in list(idx.keys()):
                    if k in name_to_i: idx[k] = name_to_i[k]
        except Exception:
            pass

        long_now = code_now = desc_now = ""
        try:
            if hasattr(self, "tree") and self.tree is not None:
                sel = self.tree.selection()
                if sel:
                    vals = self.tree.item(sel[0], "values")
                    seq = vals.get('values', []) if isinstance(vals, dict) else vals
                    if isinstance(seq, (list, tuple)):
                        long_now = seq[idx['LONG']] if len(seq) > idx['LONG'] else ""
                        code_now = seq[idx['CODE']] if len(seq) > idx['CODE'] else ""
                        desc_now = seq[idx['DESCRIPTION']] if len(seq) > idx['DESCRIPTION'] else ""
        except Exception:
            pass

        self.var_b_code.set(code_now); self.var_b_desc.set(desc_now); self.var_b_long.set(long_now)
        try:
            opt_text = self.var_code_new.get() or ''
        except Exception:
            opt_text = ''
        new_code, new_desc = self._predict_after(code_now, desc_now, opt_text)
        self.var_a_code.set(new_code); self.var_a_desc.set(new_desc); self.var_a_long.set(long_now)

    def _predict_after(self, code_now, desc_now, opt_text):
        if not opt_text:
            return code_now, desc_now
        try:
            parts = re.split(r'[—-]+', opt_text, maxsplit=1)
            left  = (parts[0].strip() if parts else '').split()
            right = (parts[1].strip() if len(parts) > 1 else '')
            new_code = left[0] if left else code_now
            new_desc = right if right else desc_now
            return new_code, new_desc
        except Exception:
            return code_now, desc_now

# ---------- main ----------
def main() -> None:
    root = tk.Tk()
    root.option_add('*TCombobox*Listbox*Font', ('Segoe UI', 10))
    root.option_add('*Font', ('Segoe UI', 10))
    root.title(APP_TITLE); root.geometry(APP_GEOM)
    App(root)
    root.mainloop()

if __name__ == "__main__":
    main()

    # ----- Custom packs actions -----
    def _run_custom_pack(self, name: str):
        mapping = CUSTOM_PACKS.get(name, {})
        if not mapping:
            self._log(f"[EDIT] '{name}': mapping is empty."); return
        self._bulk_preview_and_apply(mapping, name)

    def _action_custom_pack_1(self): self._run_custom_pack("custom pack 1")
    def _action_custom_pack_2(self): self._run_custom_pack("custom pack 2")
    def _action_custom_pack_3(self): self._run_custom_pack("custom pack 3")
    def _action_custom_pack_4(self): self._run_custom_pack("custom pack 4")
    def _action_custom_pack_5(self): self._run_custom_pack("custom pack 5")
    def _action_custom_pack_6(self): self._run_custom_pack("custom pack 6")
    def _action_custom_pack_7(self): self._run_custom_pack("custom pack 7")
    def _action_custom_pack_8(self): self._run_custom_pack("custom pack 8")
    def _action_custom_pack_9(self): self._run_custom_pack("custom pack 9")
    def _action_custom_pack_10(self): self._run_custom_pack("custom pack 10")