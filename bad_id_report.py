"""
Bad ID detection and client report builder
==========================================

Mirrors KESHO014_Data_Reports_Syntax.sps and writes the report in the same
layout as the Data Reports workbook sent to the client.

What it does
------------
1. Reads the raw data file and computes LOI = sys_SumPageTimes / 60.
2. Flags straightliners per grid, using the same rule as the .sps:
       E11    - more than 4 answered items AND SD = 0
       Q103   - more than 4 answered items AND SD = 0
       Q102   - SD = 0
       Q112   - SD = 0
       Q116   - SD = 0
   Total_StraighLiner = the number of grids flagged.
3. Flags speeders (low LOI), high LOI, junk open ends, and duplicate IPs.
4. Carries forward every ID reported in earlier rounds, so the report stays
   cumulative. IDs new to this round are marked red on sys_RespNum and psid.
5. Stamps each row with the date it was first reported, so counts can be
   pulled by round.
6. Writes the report sheet with the client's columns and highlighting.

Usage
-----
    python bad_id_report.py

Edit the CONFIG block below for each project/wave. Everything else can stay.

Dependencies: pandas, numpy, openpyxl
"""

import os
import re
import collections
import datetime as dt

import numpy as np
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.styles.colors import Color
from openpyxl.utils import get_column_letter

# ============================== CONFIG ======================================

DATA_FILE = "KESHO014_data__2_.csv"          # raw data (.csv or .xlsx)
# Reports already sent, with the date each one went out. Rows are carried
# forward as originally sent. If a file already has a "Reported Date" column,
# that is used instead of the date given here.
PREV_REPORTS = {
    "KESHO014_ZEISS_US_Regimen_Claims_Validation_Research_Data_Reports_16092026.xlsx":
        dt.date(2026, 9, 16),
}
PREV_SHEET = "Data Reports"
THIS_DATE = dt.date.today()          # date this round is reported
OUT_FILE = "KESHO014_Bad_IDs_Report.xlsx"

ID_COL = "sys_RespNum"
PSID_COL = "psid"
IP_COL = "sys_IPAddress"
TIME_COL = "sys_SumPageTimes"                 # seconds; LOI = this / 60
OE_COL = "Q115"                               # open end checked for junk

# Grids to test for straightlining.
# min_items = the ">4 answered" guard in the .sps; use None where the syntax
# applies no guard (Q102, Q112, Q116 are always asked in full).
GRIDS = {
    "E11":  {"cols": [f"E11_r{i}"  for i in range(1, 13)], "min_items": 5},
    "Q102": {"cols": [f"Q102_r{i}" for i in range(1, 9)],  "min_items": None},
    "Q103": {"cols": [f"Q103_r{i}" for i in range(1, 11)], "min_items": 5},
    "Q112": {"cols": [f"Q112_r{i}" for i in range(1, 9)],  "min_items": None},
    "Q116": {"cols": [f"Q116_r{i}" for i in range(1, 9)],  "min_items": None},
}

SL_MIN = 2          # report when Total_StraighLiner is greater than 1
SPEEDER_MINS = 4.32  # report when LOI is below this
HIGH_LOI_MINS = 50   # report when LOI is above this; set to None to switch off
CHECK_DUP_IP = True

# Junk OE is auto-detected, then reviewed by eye. Put confirmed IDs here to
# force them in, and false positives in OE_EXCLUDE to force them out.
# 6330 ("Bwest") reads as nonsense but passes the automatic rules, so it is
# forced in. This is how any eyeballed verbatim gets added or removed.
OE_FORCE = ["6330"]
OE_EXCLUDE = []

# Factor names and the order they appear in the string, per the .sps
FACTOR_ORDER = ["LOI", "High LOI", "Junk OE", "Straight Liner", "Duplicate IP"]

# ============================== LOAD ========================================

def load_data(path):
    if path.lower().endswith((".xlsx", ".xlsm")):
        return pd.read_excel(path, dtype=str)
    return pd.read_csv(path, dtype=str, low_memory=False)


def load_previous(reports, sheet):
    """Rows from every earlier report, exactly as they were sent.

    Returns a list of dicts (one per reported ID) carrying a _date key.
    """
    carried = []
    for path, sent in (reports or {}).items():
        if not os.path.exists(path):
            print(f"  ! not found, skipped: {path}")
            continue
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb[sheet] if sheet in wb.sheetnames else wb.worksheets[0]
        head = [c.value for c in ws[1]]
        if "Factor" not in head:
            print(f"  ! no Factor column, skipped: {path}")
            continue
        for row in ws.iter_rows(min_row=2, values_only=True):
            rec = dict(zip(head, row))
            if not rec.get("Factor"):
                continue
            d = rec.get("Reported Date") or sent
            rec["_date"] = d.date() if isinstance(d, dt.datetime) else d
            carried.append(rec)
    seen, out = set(), []
    for rec in carried:                      # keep the earliest sighting of an ID
        rid = str(rec.get("sys_RespNum"))
        if rid in seen:
            continue
        seen.add(rid)
        out.append(rec)
    return out


# ============================== CHECKS ======================================

def straightliner_flags(df):
    """Per-grid straightliner flags, the count, and which grids were flagged."""
    counts, which = [], []
    for _, row in df.iterrows():
        n, hit = 0, []
        for name, spec in GRIDS.items():
            vals = [row[c] for c in spec["cols"]
                    if c in df.columns and pd.notna(row[c]) and str(row[c]).strip() != ""]
            need = spec["min_items"] or 2          # SD needs at least 2 values
            if len(vals) >= need and len(set(vals)) == 1:
                n += 1
                hit.append(name)
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
    """Return a reason string if the verbatim looks like a non-answer, else None."""
    if text is None or (isinstance(text, float) and np.isnan(text)):
        return None
    s = re.sub(r"\s+", " ", str(text)).strip()
    if s == "" or s.lower() == "nan":
        return None                                   # not asked / left blank
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


def duplicate_ips(df):
    """IDs sharing an IP address with at least one other respondent."""
    counts = df[IP_COL].value_counts()
    shared = counts[counts > 1].index
    out = collections.defaultdict(list)
    for ip in shared:
        ids = sorted(df.loc[df[IP_COL] == ip, "RespNum"], key=lambda x: int(x))
        for i in ids:
            out[i] = [x for x in ids if x != i]
    return out

# ============================== BUILD =======================================

def build_factor(row, junk_ids, dup_ids):
    parts = []
    if row.LOI < SPEEDER_MINS:
        parts.append("LOI")
    if HIGH_LOI_MINS and row.LOI > HIGH_LOI_MINS:
        parts.append("High LOI")
    if row.RespNum in junk_ids:
        parts.append("Junk OE")
    if row.SL_count >= SL_MIN:
        parts.append("Straight Liner")
    if row.RespNum in dup_ids:
        parts.append("Duplicate IP")
    return " + ".join(sorted(parts, key=FACTOR_ORDER.index))


def write_report(rows, grids, cfg, path):
    """Build the client workbook.

    rows is a list of {"id", "new", "src", "rec"}, where rec is either a
    carried-forward dict from an earlier report or a row of the current data.
    """
    grid_cols = [c for cols in grids.values() for c in cols]
    loi_head = f"LOI ({cfg['speeder']} Mins)"
    head = [cfg["id_col"], cfg["psid_col"], "Factor", "Reported Date",
            "Total_StraighLiner", cfg["ip_col"], loi_head, cfg["oe_col"]] + grid_cols
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
    yellow = PatternFill("solid", start_color="FFFF00")     # straightlined cells
    red = PatternFill("solid", start_color="FF0000")        # new this round
    tint = 0.7999816888943144
    oe_fill = PatternFill("solid", fgColor=Color(theme=4, tint=tint))
    loi_fill = PatternFill("solid", fgColor=Color(theme=5, tint=tint))
    ip_fill = PatternFill("solid", start_color="E2EFDA")

    for j, h in enumerate(head, 1):
        c = ws.cell(row=1, column=j, value=h)
        c.font, c.alignment, c.fill, c.border = hf, cen, hfill, bd

    def as_cell(v):
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return None
        try:
            return int(v)
        except (TypeError, ValueError):
            return str(v)

    for i, item in enumerate(rows, 2):
        r = item["rec"]
        if item["src"] == "prev":
            vals = [r["_date"] if h == "Reported Date" else r.get(h) for h in head]
            sl_count = r.get("Total_StraighLiner") or 0
            factor_txt = r.get("Factor") or ""
            sl_which = [g for g, cols in grids.items()
                        if len([r.get(c) for c in cols if r.get(c) not in (None, "")]) >= 5
                        and len({r.get(c) for c in cols if r.get(c) not in (None, "")}) == 1]
        else:
            vals = [r["RespNum"], r.get(cfg["psid_col"]), r["Factor"], cfg["this_date"],
                    int(r["SL_count"]), r.get(cfg["ip_col"]),
                    None if pd.isna(r["LOI"]) else round(float(r["LOI"]), 2),
                    None if pd.isna(r.get(cfg["oe_col"])) else str(r.get(cfg["oe_col"]))]
            vals += [r.get(c) for c in grid_cols]
            sl_count = int(r["SL_count"])
            factor_txt = r["Factor"]
            sl_which = [x for x in str(r["SL_which"]).split("; ") if x]

        for j, v in enumerate(vals, 1):
            if j == col[loi_head] and v is not None:
                v = round(float(v), 2)
            keep = j in (col["Factor"], col["Reported Date"], col[loi_head])
            c = ws.cell(row=i, column=j, value=v if keep else as_cell(v))
            c.font, c.alignment, c.border = bf, cen, bd
            if j == col["Reported Date"]:
                c.number_format = "DD-MMM-YYYY"
            if j == col[loi_head]:
                c.number_format = "0.00"

        if sl_count and "Straight Liner" in factor_txt:
            for g in sl_which:
                for c_ in grids.get(g, []):
                    if c_ in col:
                        ws.cell(row=i, column=col[c_]).fill = yellow
        if "Junk OE" in factor_txt and cfg["oe_col"] in col:
            ws.cell(row=i, column=col[cfg["oe_col"]]).fill = oe_fill
        if "LOI" in factor_txt:
            ws.cell(row=i, column=col[loi_head]).fill = loi_fill
        if "Duplicate IP" in factor_txt and cfg["ip_col"] in col:
            ws.cell(row=i, column=col[cfg["ip_col"]]).fill = ip_fill

        if item["new"]:                      # red on the ID columns
            for h in (cfg["id_col"], cfg["psid_col"]):
                if h not in col:
                    continue
                c = ws.cell(row=i, column=col[h])
                c.fill = red
                c.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")

    for k, w in {"A": 15.8, "B": 33.5, "C": 26, "D": 15,
                 "E": 19.3, "F": 34.2, "G": 16.7, "H": 96.3}.items():
        ws.column_dimensions[k].width = w
    for j in range(9, len(head) + 1):
        ws.column_dimensions[get_column_letter(j)].width = 10.5
    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(head))}{len(rows) + 1}"

    # by-date summary
    ws2 = wb.create_sheet("By Date")
    tally = collections.Counter(
        (r["rec"]["_date"] if r["src"] == "prev" else cfg["this_date"]) for r in rows)
    for j, h in enumerate(["Reported Date", "Bad IDs", "Cumulative"], 1):
        c = ws2.cell(row=1, column=j, value=h)
        c.font = hf
    run = 0
    for i, (d, n) in enumerate(sorted(tally.items()), 2):
        run += n
        ws2.cell(row=i, column=1, value=d).number_format = "DD-MMM-YYYY"
        ws2.cell(row=i, column=2, value=n)
        ws2.cell(row=i, column=3, value=run)
        for j in range(1, 4):
            ws2.cell(row=i, column=j).font = bf
    for k, w in {"A": 16, "B": 12, "C": 12}.items():
        ws2.column_dimensions[k].width = w

    wb.save(path)


def main():
    df = load_data(DATA_FILE)
    df["RespNum"] = df[ID_COL].astype(str).str.strip()
    df["LOI"] = pd.to_numeric(df[TIME_COL], errors="coerce") / 60.0
    df["SL_count"], df["SL_which"] = straightliner_flags(df)

    # junk open ends
    junk = {}
    for idx, v in df[OE_COL].items():
        reason = junk_oe_reason(v)
        if reason:
            junk[df.at[idx, "RespNum"]] = (str(v).strip(), reason)
    for i in OE_FORCE:
        junk.setdefault(str(i), ("", "added by hand"))
    for i in OE_EXCLUDE:
        junk.pop(str(i), None)

    dup = duplicate_ips(df) if CHECK_DUP_IP else {}

    carried = load_previous(PREV_REPORTS, PREV_SHEET)
    prev_ids = {str(r.get("sys_RespNum")) for r in carried}

    fresh = df[~df.RespNum.isin(prev_ids)].copy()
    fresh["Factor"] = fresh.apply(lambda r: build_factor(r, junk, dup), axis=1)
    fresh = fresh[fresh.Factor != ""]

    rows = [{"id": str(r.get("sys_RespNum")), "new": False, "src": "prev", "rec": r}
            for r in carried]
    rows += [{"id": r["RespNum"], "new": True, "src": "new", "rec": r}
             for _, r in fresh.iterrows()]
    rows.sort(key=lambda x: int(x["id"]) if str(x["id"]).isdigit() else 0)

    cfg = {"id_col": ID_COL, "psid_col": PSID_COL, "ip_col": IP_COL, "oe_col": OE_COL,
           "sl_min": SL_MIN, "speeder": SPEEDER_MINS, "this_date": THIS_DATE}
    write_report(rows, {g: spec["cols"] for g, spec in GRIDS.items()}, cfg, OUT_FILE)

    print(f"Records in data file   : {len(df)}")
    print(f"Carried forward        : {len(carried)}")
    print(f"New this round (in red): {len(fresh)}")
    print(f"Rows in report         : {len(rows)}")
    print(f"Median LOI             : {df.LOI.median():.2f} mins")
    print()
    print("  Straight Liner :", int((fresh.SL_count >= SL_MIN).sum()))
    print("  LOI (speeders) :", int((fresh.LOI < SPEEDER_MINS).sum()))
    if HIGH_LOI_MINS:
        print("  High LOI       :", int((fresh.LOI > HIGH_LOI_MINS).sum()))
    print("  Junk OE        :", sum(1 for i in fresh.RespNum if i in junk))
    print("  Duplicate IP   :", sum(1 for i in fresh.RespNum if i in dup))
    print()
    print("Junk OE detected (check these by eye before sending):")
    for rid, (txt, why) in junk.items():
        where = "new" if rid in set(fresh.RespNum) else "reported earlier"
        print(f"   {rid:>6}  {why:<22} \"{txt[:45]}\"  [{where}]")
    print()
    print("Borderline, not reported:")
    edge = fresh_edge = df[(~df.RespNum.isin(prev_ids)) & (df.SL_count == SL_MIN - 1)
                           & (~df.RespNum.isin(set(fresh.RespNum)))]
    print(f"   {len(edge)} IDs with exactly {SL_MIN - 1} straightlined grid(s): "
          + ", ".join(sorted(edge.RespNum, key=int)))
    print()
    print("Saved:", OUT_FILE)

    def sps(ids):
        return ", ".join(sorted((str(i) for i in ids), key=int))
    print()
    print("For the syntax (new IDs only):")
    print("  FlagLOI       :", sps(fresh.loc[fresh.LOI < SPEEDER_MINS, "RespNum"]))
    print("  FlagQ115      :", sps([i for i in fresh.RespNum if i in junk]))
    print("  Dup_IPAddress :", sps([i for i in fresh.RespNum if i in dup]))


if __name__ == "__main__":
    main()
