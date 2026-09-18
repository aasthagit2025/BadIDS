"""
Bad ID Report - Streamlit app
=============================

Upload the raw data file and (optionally) the reports already sent to the
client. The app flags straightliners, speeders, junk open ends and duplicate
IPs, drops anything reported before, and builds the client report workbook
with the usual highlighting.

Run locally:
    pip install -r requirements.txt
    streamlit run app.py
"""

import io
import re
import collections

import numpy as np
import pandas as pd
import streamlit as st
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.styles.colors import Color
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Bad ID Report", page_icon="🧹", layout="wide")

FACTOR_ORDER = ["LOI", "High LOI", "Junk OE", "Straight Liner", "Duplicate IP"]

# ============================ core logic ====================================

def read_table(upload):
    """Read an uploaded csv/xlsx into a string DataFrame."""
    name = upload.name.lower()
    data = upload.getvalue()
    if name.endswith((".xlsx", ".xlsm")):
        return pd.read_excel(io.BytesIO(data), dtype=str)
    for enc in ("utf-8", "latin-1"):
        try:
            return pd.read_csv(io.BytesIO(data), dtype=str, low_memory=False, encoding=enc)
        except UnicodeDecodeError:
            continue
    return pd.read_csv(io.BytesIO(data), dtype=str, low_memory=False,
                       encoding="utf-8", encoding_errors="replace")


def read_previous(uploads, id_col="sys_RespNum", factor_col="Factor"):
    """{resp_id: factor} across every previous report uploaded."""
    prev = {}
    for up in uploads or []:
        wb = openpyxl.load_workbook(io.BytesIO(up.getvalue()), data_only=True)
        for ws in wb.worksheets:
            head = [c.value for c in ws[1]]
            if id_col not in head or factor_col not in head:
                continue
            i_id, i_f = head.index(id_col), head.index(factor_col)
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row[i_f] not in (None, ""):
                    prev[str(row[i_id]).strip()] = row[i_f]
            break
    return prev


def detect_grids(columns):
    """Group columns that look like GRID_r1, GRID_r2, ... into grids."""
    found = collections.defaultdict(list)
    for c in columns:
        m = re.fullmatch(r"(.+?)_[rR](\d+)", str(c))
        if m:
            found[m.group(1)].append((int(m.group(2)), c))
    out = {}
    for name, items in found.items():
        if len(items) >= 3:
            out[name] = [c for _, c in sorted(items)]
    return out


def straightliner_flags(df, grids, min_items):
    """Per-row count of straightlined grids and which ones."""
    counts, which = [], []
    sub = {g: [c for c in cols if c in df.columns] for g, cols in grids.items()}
    for _, row in df.iterrows():
        n, hit = 0, []
        for g, cols in sub.items():
            vals = [row[c] for c in cols
                    if pd.notna(row[c]) and str(row[c]).strip() != ""]
            need = max(2, int(min_items.get(g, 5)))
            if len(vals) >= need and len(set(vals)) == 1:
                n += 1
                hit.append(g)
        counts.append(n)
        which.append("; ".join(hit))
    return counts, which


STOP = set("""yes no na n/a none nothing nope idk dk dunno good nice ok okay great fine cool
test asdf asdfg qwerty abc xyz nil null blank same maybe sure yeah yep hmm hi hello
dont know unknown whatever anything everything thanks thank""".split())

PHRASES = {"no comment", "no reason", "no idea", "i just did", "i did", "same as above",
           "see above", "not sure", "no answer", "no opinion", "as above", "n a",
           "no thoughts", "i dont know", "i don't know", "just because"}


def junk_oe_reason(text):
    """Reason string if the verbatim reads as a non-answer, else None."""
    if text is None or (isinstance(text, float) and np.isnan(text)):
        return None
    s = re.sub(r"\s+", " ", str(text)).strip()
    if s == "" or s.lower() == "nan":
        return None
    low = re.sub(r"\s+", " ", re.sub(r"[^a-z0-9\s']", " ", s.lower())).strip()
    words = [w for w in low.split() if w]
    if len(s) < 3:
        return "Too short"
    if not re.search(r"[A-Za-z]", s):
        return "No text - digits or symbols only"
    if low in PHRASES:
        return "Non-answer phrase"
    if len(words) <= 2 and all(w in STOP for w in words):
        return "Non-answer / filler"
    if len(words) == 1 and len(words[0]) <= 4:
        return "Single short word"
    for w in words:
        if len(w) >= 5 and not re.search(r"[aeiou]", w):
            return "Gibberish"
    if re.search(r"(.)\1{3,}", low.replace(" ", "")):
        return "Repeated characters"
    if len(words) > 1 and len(set(words)) == 1:
        return "Same word repeated"
    return None


def duplicate_ips(df, ip_col):
    counts = df[ip_col].value_counts()
    out = {}
    for ip in counts[counts > 1].index:
        ids = sorted(df.loc[df[ip_col] == ip, "RespNum"], key=lambda x: int(x)
                     if str(x).isdigit() else 0)
        for i in ids:
            out[i] = [x for x in ids if x != i]
    return out


def build_factor(row, junk_ids, dup_ids, sl_min, speeder, high_loi):
    parts = []
    if pd.notna(row.LOI) and row.LOI < speeder:
        parts.append("LOI")
    if high_loi and pd.notna(row.LOI) and row.LOI > high_loi:
        parts.append("High LOI")
    if row.RespNum in junk_ids:
        parts.append("Junk OE")
    if row.SL_count >= sl_min:
        parts.append("Straight Liner")
    if row.RespNum in dup_ids:
        parts.append("Duplicate IP")
    return " + ".join(sorted(parts, key=FACTOR_ORDER.index))


def write_report(bad, grids, cfg):
    """Build the client workbook and return it as bytes."""
    grid_cols = [c for cols in grids.values() for c in cols if c in bad.columns]
    loi_head = f"LOI ({cfg['speeder']} Mins)"
    head = [cfg["id_col"], cfg["psid_col"], "Factor", "Total_StraighLiner",
            cfg["ip_col"], loi_head, cfg["oe_col"]] + grid_cols
    col = {h: i + 1 for i, h in enumerate(head)}

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data Reports"

    hf = Font(name="Calibri", size=10, bold=True)
    bf = Font(name="Calibri", size=10)
    cen = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin = Side(style="thin", color="D9D9D9")
    bd = Border(left=thin, right=thin, top=thin, bottom=thin)

    hfill = PatternFill("solid", start_color="D9E1F2")
    yellow = PatternFill("solid", start_color="FFFF00")
    tint = 0.7999816888943144
    oe_fill = PatternFill("solid", fgColor=Color(theme=4, tint=tint))
    loi_fill = PatternFill("solid", fgColor=Color(theme=5, tint=tint))
    ip_fill = PatternFill("solid", start_color="E2EFDA")

    for j, h in enumerate(head, 1):
        c = ws.cell(row=1, column=j, value=h)
        c.font, c.alignment, c.fill, c.border = hf, cen, hfill, bd

    for i, (_, r) in enumerate(bad.iterrows(), 2):
        vals = [r["RespNum"], r.get(cfg["psid_col"]), r["Factor"], int(r["SL_count"]),
                r.get(cfg["ip_col"]),
                None if pd.isna(r["LOI"]) else round(float(r["LOI"]), 2),
                None if pd.isna(r.get(cfg["oe_col"])) else str(r.get(cfg["oe_col"]))]
        try:
            vals[0] = int(vals[0])
        except (TypeError, ValueError):
            pass
        for c_ in grid_cols:
            v = r.get(c_)
            if pd.isna(v):
                vals.append(None)
            else:
                try:
                    vals.append(int(v))
                except (TypeError, ValueError):
                    vals.append(str(v))
        for j, v in enumerate(vals, 1):
            c = ws.cell(row=i, column=j, value=v)
            c.font, c.alignment, c.border = bf, cen, bd
            if j == col[loi_head]:
                c.number_format = "0.00"

        if r["SL_count"] >= cfg["sl_min"]:
            for g in [x for x in str(r["SL_which"]).split("; ") if x]:
                for c_ in grids.get(g, []):
                    if c_ in col:
                        ws.cell(row=i, column=col[c_]).fill = yellow
        if "Junk OE" in r["Factor"]:
            ws.cell(row=i, column=col[cfg["oe_col"]]).fill = oe_fill
        if "LOI" in r["Factor"]:
            ws.cell(row=i, column=col[loi_head]).fill = loi_fill
        if "Duplicate IP" in r["Factor"]:
            ws.cell(row=i, column=col[cfg["ip_col"]]).fill = ip_fill

    for k, w in {"A": 15.8, "B": 33.5, "C": 26, "D": 19.3,
                 "E": 34.2, "F": 16.7, "G": 96.3}.items():
        ws.column_dimensions[k].width = w
    for j in range(8, len(head) + 1):
        ws.column_dimensions[get_column_letter(j)].width = 10.5
    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(head))}{len(bad) + 1}"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ================================ UI ========================================

st.title("Bad ID Report")
st.caption("Straightliners, speeders, junk open ends and duplicate IPs, in the "
           "client report format. IDs reported in earlier rounds are excluded.")

with st.sidebar:
    st.header("1. Files")
    data_up = st.file_uploader("Raw data (.csv or .xlsx)", type=["csv", "xlsx", "xlsm"])
    prev_up = st.file_uploader("Earlier report(s) to exclude", type=["xlsx"],
                               accept_multiple_files=True)

if not data_up:
    st.info("Upload a raw data file in the sidebar to begin.")
    st.stop()

df = read_table(data_up)
cols = list(df.columns)


def pick(label, options, prefer, key, allow_none=False):
    opts = (["(none)"] if allow_none else []) + options
    default = next((p for p in prefer if p in options), opts[0])
    return st.selectbox(label, opts, index=opts.index(default), key=key)


with st.sidebar:
    st.header("2. Columns")
    id_col = pick("Respondent ID", cols, ["sys_RespNum", "RespNum", "respid"], "id")
    psid_col = pick("Panel ID", cols, ["psid", "PSID"], "psid", allow_none=True)
    ip_col = pick("IP address", cols, ["sys_IPAddress", "IPAddress"], "ip", allow_none=True)
    time_col = pick("Time in seconds (LOI source)", cols,
                    ["sys_SumPageTimes", "sys_ElapsedTime"], "time")
    oe_col = pick("Open end to check", cols, ["Q115"], "oe", allow_none=True)

    st.header("3. Criteria")
    sl_min = st.number_input("Straightliner: report when grids flagged is at least",
                             min_value=1, max_value=6, value=2, step=1,
                             help="2 means Total_StraighLiner greater than 1.")
    speeder = st.number_input("Speeder: LOI under (mins)", min_value=0.0,
                              value=4.32, step=0.1, format="%.2f")
    use_high = st.checkbox("Also flag High LOI", value=True)
    high_loi = st.number_input("High LOI: over (mins)", min_value=1.0, value=50.0,
                               step=5.0, disabled=not use_high)
    check_ip = st.checkbox("Flag duplicate IP addresses", value=True,
                           disabled=(ip_col == "(none)"))

df["RespNum"] = df[id_col].astype(str).str.strip()
df["LOI"] = pd.to_numeric(df[time_col], errors="coerce") / 60.0

# ---- grids ----
auto = detect_grids(cols)
st.subheader("Grids checked for straightlining")
if not auto:
    st.warning("No columns matching a GRID_r1, GRID_r2 pattern were found. "
               "Straightlining cannot be checked on this file.")
    chosen, min_items = {}, {}
else:
    names = st.multiselect("Grids", sorted(auto), default=sorted(auto))
    grid_cfg = pd.DataFrame(
        [{"Grid": g, "Items": len(auto[g]), "Min answered items": 5} for g in names])
    if len(grid_cfg):
        grid_cfg = st.data_editor(grid_cfg, hide_index=True, use_container_width=True,
                                  disabled=["Grid", "Items"], key="grids")
    chosen = {g: auto[g] for g in names}
    min_items = dict(zip(grid_cfg["Grid"], grid_cfg["Min answered items"])) if len(grid_cfg) else {}
    st.caption("Min answered items mirrors the \"count greater than 4\" guard in the "
               "syntax: a grid only counts when at least this many items are "
               "answered and every answer is identical.")

df["SL_count"], df["SL_which"] = straightliner_flags(df, chosen, min_items)

# ---- previous rounds ----
prev = read_previous(prev_up, id_col=id_col)
new = df[~df.RespNum.isin(prev)].copy()

# ---- junk OE ----
junk_auto = {}
if oe_col != "(none)":
    for idx, v in new[oe_col].items():
        why = junk_oe_reason(v)
        if why:
            junk_auto[new.at[idx, "RespNum"]] = (str(v).strip(), why)

st.subheader("Junk open ends")
if oe_col == "(none)":
    st.caption("No open end selected.")
    junk = set()
else:
    if junk_auto:
        pick_df = pd.DataFrame(
            [{"Report": True, "ID": k, "Verbatim": v[0], "Why": v[1]}
             for k, v in sorted(junk_auto.items(), key=lambda x: int(x[0])
                                if x[0].isdigit() else 0)])
        pick_df = st.data_editor(pick_df, hide_index=True, use_container_width=True,
                                 disabled=["ID", "Verbatim", "Why"], key="oe")
        junk = set(pick_df.loc[pick_df["Report"], "ID"])
    else:
        st.caption("Nothing picked up automatically.")
        junk = set()
    extra = st.text_input("Add IDs by hand (comma separated)",
                          help="For verbatims that read as nonsense but pass the "
                               "automatic rules.")
    junk |= {x.strip() for x in extra.split(",") if x.strip()}
    with st.expander("Browse the shortest verbatims"):
        look = new[["RespNum", oe_col, "LOI"]].dropna(subset=[oe_col]).copy()
        look["Length"] = look[oe_col].str.len()
        st.dataframe(look.sort_values("Length").head(40), hide_index=True,
                     use_container_width=True)

dup = duplicate_ips(new, ip_col) if (check_ip and ip_col != "(none)") else {}

# ---- build ----
cfg = {"id_col": id_col, "psid_col": psid_col, "ip_col": ip_col, "oe_col": oe_col,
       "sl_min": sl_min, "speeder": speeder,
       "high_loi": high_loi if use_high else None}

new["Factor"] = new.apply(
    lambda r: build_factor(r, junk, dup, sl_min, speeder, cfg["high_loi"]), axis=1)
bad = new[new.Factor != ""].copy()
bad["_sort"] = pd.to_numeric(bad.RespNum, errors="coerce")
bad = bad.sort_values("_sort")

st.subheader("Result")
c = st.columns(5)
c[0].metric("Records", len(df))
c[1].metric("Excluded (reported before)", len(prev))
c[2].metric("Reviewed", len(new))
c[3].metric("Bad IDs", len(bad))
c[4].metric("Median LOI", f"{df.LOI.median():.2f}")

counts = {
    "Straight Liner": int((bad.SL_count >= sl_min).sum()),
    "LOI (speeders)": int((bad.LOI < speeder).sum()),
    "High LOI": int((bad.LOI > cfg["high_loi"]).sum()) if cfg["high_loi"] else 0,
    "Junk OE": int(bad.RespNum.isin(junk).sum()),
    "Duplicate IP": int(bad.RespNum.isin(set(dup)).sum()),
}
st.dataframe(pd.DataFrame([{"Criterion": k, "IDs": v} for k, v in counts.items()]),
             hide_index=True)

show = [c_ for c_ in ["RespNum", "Factor", "SL_count", "SL_which", "LOI",
                      ip_col, oe_col] if c_ in bad.columns or c_ == "RespNum"]
st.dataframe(bad[show], hide_index=True, use_container_width=True)

if len(bad):
    st.download_button(
        "Download report",
        data=write_report(bad, chosen, cfg),
        file_name="Bad_IDs_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary")

with st.expander("IDs for the SPSS syntax"):
    def sps(ids):
        ids = [str(i) for i in ids]
        return ", ".join(sorted(ids, key=lambda x: int(x) if x.isdigit() else 0))
    st.code(
        f"IF (any(sys_RespNum, {sps(bad.loc[bad.LOI < speeder, 'RespNum'])}))FlagLOI=1.\n"
        f"IF (any(sys_RespNum, {sps(bad.loc[bad.RespNum.isin(junk), 'RespNum'])}))FlagQ115=1.\n"
        f"IF (any(sys_RespNum, {sps(bad.loc[bad.RespNum.isin(set(dup)), 'RespNum'])}))Dup_IPAddress=1.",
        language="text")

with st.expander("Borderline cases not reported"):
    edge = new[(new.SL_count == sl_min - 1) & (new.Factor == "")]
    st.write(f"**{len(edge)} IDs with exactly {sl_min - 1} straightlined grid(s)**")
    st.code(", ".join(sorted(edge.RespNum, key=lambda x: int(x) if x.isdigit() else 0))
            or "none")
    half = df.LOI.median() / 2
    band = new[(new.LOI >= speeder) & (new.LOI < half) & (new.Factor == "")]
    st.write(f"**{len(band)} IDs between {speeder:.2f} mins and half the median "
             f"({half:.2f} mins)**")
    st.dataframe(band[["RespNum", "LOI", "SL_count"]].sort_values("LOI"),
                 hide_index=True)
