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
4. Drops any ID that was already reported in an earlier round.
5. Writes the report sheet with the client's columns and highlighting.

Usage
-----
    python bad_id_report.py

Edit the CONFIG block below for each project/wave. Everything else can stay.

Dependencies: pandas, numpy, openpyxl
"""

import os
import re
import collections

import numpy as np
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.styles.colors import Color
from openpyxl.utils import get_column_letter

# ============================== CONFIG ======================================

DATA_FILE = "KESHO014_data__2_.csv"          # raw data (.csv or .xlsx)
PREV_REPORT = "KESHO014_ZEISS_US_Regimen_Claims_Validation_Research_Data_Reports_16092026.xlsx"
PREV_SHEET = "Data Reports"
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


def load_previous(path, sheet):
    """Return {resp_id: factor} for every ID reported in an earlier round."""
    if not path or not os.path.exists(path):
        print("No previous report found - nothing will be excluded.")
        return {}
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet]
    head = [c.value for c in ws[1]]
    i_id, i_fac = head.index("sys_RespNum"), head.index("Factor")
    prev = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[i_fac]:
            prev[str(row[i_id])] = row[i_fac]
    return prev

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


def write_report(bad, path):
    grid_cols = [c for spec in GRIDS.values() for c in spec["cols"]]
    head = [ID_COL, PSID_COL, "Factor", "Total_StraighLiner", IP_COL,
            f"LOI ({SPEEDER_MINS} Mins)", OE_COL] + grid_cols
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
    yellow = PatternFill("solid", start_color="FFFF00")            # straightlined cells
    tint = 0.7999816888943144
    oe_fill = PatternFill("solid", fgColor=Color(theme=4, tint=tint))   # light blue
    loi_fill = PatternFill("solid", fgColor=Color(theme=5, tint=tint))  # light red
    ip_fill = PatternFill("solid", start_color="E2EFDA")                # light green

    for j, h in enumerate(head, 1):
        c = ws.cell(row=1, column=j, value=h)
        c.font, c.alignment, c.fill, c.border = hf, cen, hfill, bd

    for i, (_, r) in enumerate(bad.iterrows(), 2):
        vals = [int(r.RespNum), r[PSID_COL], r.Factor, int(r.SL_count), r[IP_COL],
                round(float(r.LOI), 2),
                None if pd.isna(r[OE_COL]) else str(r[OE_COL])]
        for c_ in grid_cols:
            v = r[c_] if c_ in bad.columns else None
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
            if j == col[f"LOI ({SPEEDER_MINS} Mins)"]:
                c.number_format = "0.00"

        # highlight only what the ID is actually reported for
        if r.SL_count >= SL_MIN:
            for g in [x for x in str(r.SL_which).split("; ") if x]:
                for c_ in GRIDS[g]["cols"]:
                    ws.cell(row=i, column=col[c_]).fill = yellow
        if "Junk OE" in r.Factor:
            ws.cell(row=i, column=col[OE_COL]).fill = oe_fill
        if "LOI" in r.Factor:
            ws.cell(row=i, column=col[f"LOI ({SPEEDER_MINS} Mins)"]).fill = loi_fill
        if "Duplicate IP" in r.Factor:
            ws.cell(row=i, column=col[IP_COL]).fill = ip_fill

    for k, w in {"A": 15.8, "B": 33.5, "C": 26, "D": 19.3,
                 "E": 34.2, "F": 16.7, "G": 96.3}.items():
        ws.column_dimensions[k].width = w
    for j in range(8, len(head) + 1):
        ws.column_dimensions[get_column_letter(j)].width = 10.5
    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(head))}{len(bad) + 1}"
    wb.save(path)


def main():
    df = load_data(DATA_FILE)
    df["RespNum"] = df[ID_COL].astype(str)
    df["LOI"] = pd.to_numeric(df[TIME_COL], errors="coerce") / 60.0

    df["SL_count"], df["SL_which"] = straightliner_flags(df)

    # junk open ends
    junk = {}
    for idx, v in df[OE_COL].items():
        reason = junk_oe_reason(v)
        if reason:
            junk[df.at[idx, "RespNum"]] = (str(v).strip(), reason)
    for i in OE_FORCE:
        junk.setdefault(str(i), ("", "manually added"))
    for i in OE_EXCLUDE:
        junk.pop(str(i), None)

    dup = duplicate_ips(df) if CHECK_DUP_IP else {}

    prev = load_previous(PREV_REPORT, PREV_SHEET)
    new = df[~df.RespNum.isin(prev)].copy()
    new["Factor"] = new.apply(lambda r: build_factor(r, junk, dup), axis=1)

    bad = new[new.Factor != ""].copy()
    bad["_sort"] = pd.to_numeric(bad.RespNum, errors="coerce")
    bad = bad.sort_values("_sort")

    write_report(bad, OUT_FILE)

    print(f"Records in data file      : {len(df)}")
    print(f"Reported earlier (skipped): {len(prev)}")
    print(f"Reviewed this round       : {len(new)}")
    print(f"Bad IDs reported          : {len(bad)}")
    print(f"Median LOI                : {df.LOI.median():.2f} mins")
    print()
    print("  Straight Liner :", int((bad.SL_count >= SL_MIN).sum()))
    print("  LOI (speeders) :", int((bad.LOI < SPEEDER_MINS).sum()))
    if HIGH_LOI_MINS:
        print("  High LOI       :", int((bad.LOI > HIGH_LOI_MINS).sum()))
    print("  Junk OE        :", sum(1 for i in bad.RespNum if i in junk))
    print("  Duplicate IP   :", sum(1 for i in bad.RespNum if i in dup))
    print()
    print("Junk OE detected (check these by eye before sending):")
    for rid, (txt, why) in junk.items():
        mark = "reported" if rid in set(bad.RespNum) else "already reported earlier"
        print(f"   {rid:>6}  {why:<22} \"{txt[:45]}\"  [{mark}]")
    print()
    print("Borderline, not reported:")
    edge = new[(new.SL_count == SL_MIN - 1) & (new.Factor == "")]
    print(f"   {len(edge)} IDs with exactly {SL_MIN - 1} straightlined grid(s): "
          + ", ".join(sorted(edge.RespNum, key=int)))
    band = new[(new.LOI >= SPEEDER_MINS) & (new.LOI < df.LOI.median() / 2) & (new.Factor == "")]
    print(f"   {len(band)} IDs between {SPEEDER_MINS} mins and half the median LOI")
    print()
    print("Saved:", OUT_FILE)

    # IDs formatted for pasting into the .sps IF (any(sys_RespNum, ...)) lines
    def sps(ids):
        return ", ".join(sorted((str(i) for i in ids), key=int))
    print()
    print("For the syntax:")
    print("  FlagLOI       :", sps(bad.loc[bad.LOI < SPEEDER_MINS, "RespNum"]))
    print("  FlagQ115      :", sps([i for i in bad.RespNum if i in junk]))
    print("  Dup_IPAddress :", sps([i for i in bad.RespNum if i in dup]))


if __name__ == "__main__":
    main()
