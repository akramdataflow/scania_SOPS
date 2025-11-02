
import os
import csv
import re
from typing import List, Dict, Tuple, Optional

import streamlit as st
from lxml import etree

APP_TITLE = "Scania FPC Spec Viewer — XML + CSV mapping (no pandas/numpy)"
DEFAULT_FPC_TAG = "FPC"  # element name; attributes are exactly Name/Value

def fixed_csv_path() -> str:
    home = os.path.expanduser("~")
    return os.path.join(home, r"Documents\GitHub\scania_SOPS\classifier\sops_fpc_mapping.csv")

def local_csv_path() -> str:
    return os.path.join(os.path.dirname(__file__), "sops_fpc_mapping.csv")

def find_csv_mapping() -> str:
    for p in [fixed_csv_path(), local_csv_path()]:
        if os.path.exists(p):
            return p
    return ""

def _clean(s: str) -> str:
    s = (s or "").strip()
    s = re.sub(r"\s+", " ", s)
    return s

def _norm_header(h: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", (h or "").lower())

def _is_intlike(s: str) -> bool:
    try:
        int(s.strip())
        return True
    except Exception:
        return False

@st.cache_data(show_spinner=False)
def load_csv_mapping(path: str) -> Tuple[Dict[str, str], Dict[Tuple[str,str], str]]:
    long_by_id: Dict[str, str] = {}
    desc_by_pair: Dict[Tuple[str,str], str] = {}
    if not path or not os.path.exists(path):
        raise FileNotFoundError("sops_fpc_mapping.csv not found")
    with open(path, "r", encoding="utf-8", newline="") as f:
        rows = list(csv.reader(f))
    if not rows:
        return long_by_id, desc_by_pair
    headers = rows[0]
    norm_headers = { _norm_header(h): i for i, h in enumerate(headers) }
    def idx_of(*aliases):
        for a in aliases:
            key = _norm_header(a)
            if key in norm_headers:
                return norm_headers[key]
        return None
    idx_fpc = idx_of("FPC- ID", "FPC ID", "fpc id", "fpc", "id")
    idx_long = idx_of("Long", "Name", "Parameter", "Title")
    idx_code = idx_of("code", "Value", "val")
    idx_desc = idx_of("Description", "Desc", "Meaning", "Explanation")
    if idx_fpc is None:
        for k in ["fpcid","fpc","id"]:
            if k in norm_headers: idx_fpc = norm_headers[k]; break
    if idx_long is None:
        for k in ["long","name","parameter","title"]:
            if k in norm_headers: idx_long = norm_headers[k]; break
    if idx_code is None:
        for k in ["code","value","val"]:
            if k in norm_headers: idx_code = norm_headers[k]; break
    if idx_desc is None:
        for k in ["description","desc","meaning","explanation"]:
            if k in norm_headers: idx_desc = norm_headers[k]; break
    for r in rows[1:]:
        if not r: continue
        def at(i): return _clean(r[i]) if (i is not None and i < len(r)) else ""
        fpc_raw = at(idx_fpc); long_txt = at(idx_long); code = at(idx_code); desc = at(idx_desc)
        if not fpc_raw: continue
        keys_for_fpc = {fpc_raw}
        if _is_intlike(fpc_raw): keys_for_fpc.add(str(int(fpc_raw)))
        for fpc_key in keys_for_fpc:
            if long_txt and fpc_key not in long_by_id: long_by_id[fpc_key] = long_txt
            if code and desc: desc_by_pair[(fpc_key, code)] = desc
    return long_by_id, desc_by_pair

def parse_xml_fpc(xml_bytes: bytes, tag_guess: str = DEFAULT_FPC_TAG) -> List[Dict[str, str]]:
    parser = etree.XMLParser(recover=True, huge_tree=True)
    root = etree.fromstring(xml_bytes, parser=parser)
    xpath = f"//*[translate(local-name(), '{tag_guess.lower()}', '{tag_guess.upper()}')='{tag_guess.upper()}']"
    elems = root.xpath(xpath)
    rows, seen = [], set()
    for el in elems:
        name  = el.attrib.get("Name")
        value = el.attrib.get("Value")
        if name is None or value is None: continue
        fpc_id = _clean(name); code = _clean(value)
        key = (fpc_id, code)
        if key in seen: continue
        seen.add(key)
        rows.append({"fpc_id": fpc_id, "code": code})
    return rows

def merge_specs(xml_rows: List[Dict[str, str]], long_by_id: Dict[str,str], desc_by_pair: Dict[Tuple[str,str], str]) -> List[Dict[str,str]]:
    out = []
    for r in xml_rows:
        fpc_raw = r["fpc_id"]; code = r["code"]
        candidates = [fpc_raw] + ([str(int(fpc_raw))] if _is_intlike(fpc_raw) else [])
        long_txt = ""; desc_txt = ""
        for fk in candidates:
            if not long_txt and fk in long_by_id: long_txt = long_by_id[fk]
            if not desc_txt and (fk, code) in desc_by_pair: desc_txt = desc_by_pair[(fk, code)]
            if long_txt and desc_txt: break
        out.append({"FPC ID": fpc_raw, "Long": long_txt, "Value": code, "Description": desc_txt})
    def _sort_key(it):
        f = it["FPC ID"]
        try: f2 = int(f)
        except: f2 = f
        return (f2, it["Value"])
    out.sort(key=_sort_key)
    return out

def render_table(rows: List[Dict[str, str]]):
    css = """
    <style>
    table.fpc { width:100%; border-collapse: collapse; }
    table.fpc th, table.fpc td { border: 1px solid #ddd; padding: 8px; font-size: 14px; }
    table.fpc th { background: #f2f2f2; text-align: left; }
    tr.unknown td { background: #fff3cd; }
    </style>
    """
    html = ["<table class='fpc'>"]
    html.append("<thead><tr><th>FPC ID</th><th>Long</th><th>Value</th><th>Description</th></tr></thead><tbody>")
    for r in rows:
        klass = "unknown" if not r.get("Description") else ""
        html.append(f"<tr class='{klass}'><td>{r['FPC ID']}</td><td>{r['Long']}</td><td>{r['Value']}</td><td>{r['Description']}</td></tr>")
    html.append("</tbody></table>")
    st.markdown(css + "\n" + "\n".join(html), unsafe_allow_html=True)

st.set_page_config(page_title=APP_TITLE, layout="wide")
st.title("📄 Scania FPC Spec Viewer — XML + CSV mapping")
st.caption("يرفع XML فقط. يقرأ sops_fpc_mapping.csv تلقائيًا (مسار ثابت أو نفس مجلد التطبيق). لا يعتمد على pandas/numpy.")

csv_path = find_csv_mapping()
if not csv_path:
    st.error("لم يتم العثور على ملف الخريطة sops_fpc_mapping.csv.\nضعه في: "
             "`~/Documents/GitHub/scania_SOPS/classifier/` أو بجانب app.py.")
    st.stop()

try:
    long_by_id, desc_by_pair = load_csv_mapping(csv_path)
except Exception as e:
    st.error(f"تعذر تحميل/قراءة sops_fpc_mapping.csv: {e}")
    st.stop()

with st.sidebar:
    st.header("المدخل الوحيد")
    xml_file = st.file_uploader("XML ملف الشاحنة", type=["xml"])
    with st.expander("إعدادات متقدمة", expanded=False):
        tag_guess = st.text_input("اسم عنصر FPC داخل XML (غالبًا: FPC)", value=DEFAULT_FPC_TAG)
    st.caption(f"خريطة المواصفات: `{csv_path}`")

st.subheader("النتيجة: FPC ID | Long | Value | Description")
if not xml_file:
    st.info("الرجاء رفع ملف XML.")
else:
    try:
        xml_bytes = xml_file.read()
        fpc_rows = parse_xml_fpc(xml_bytes, tag_guess=tag_guess)
    except Exception as e:
        st.error(f"فشل تحليل XML: {e}")
        st.stop()

    if not fpc_rows:
        st.warning("لم يتم العثور على مدخلات FPC في XML.")
    else:
        merged = merge_specs(fpc_rows, long_by_id, desc_by_pair)
        c1,c2,c3 = st.columns(3)
        c1.metric("FPC entries", len(merged))
        c2.metric("Long found", sum(1 for r in merged if r["Long"]))
        c3.metric("Description found", sum(1 for r in merged if r["Description"]))
        render_table(merged)
        st.caption("الصفوف المظللة تعني عدم العثور على Description لهذه القيمة في ملف الخريطة.")
