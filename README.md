# Bad ID Report

Flags straightliners, speeders, junk open ends and duplicate IPs, drops IDs
already reported in earlier rounds, and builds the client report workbook.

## Files

| File | What it is |
|---|---|
| `app.py` | Streamlit app - upload files in the browser, set the criteria, download the report |
| `bad_id_report.py` | The same checks as a plain script, for running on the desktop or scheduling |
| `requirements.txt` | Dependencies |

## Run it locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Opens at http://localhost:8501.

## Put it on Streamlit Community Cloud

1. Push `app.py` and `requirements.txt` to a GitHub repo (no data files - the
   app takes uploads, so nothing confidential needs to sit in the repo).
2. Go to share.streamlit.io, sign in with GitHub, and click **New app**.
3. Pick the repo, branch, and `app.py` as the main file, then **Deploy**.
4. Under **Settings > Sharing**, set the app to private and invite only the
   people who should see it. Client data passes through the app, so it should
   not be left public.

Uploads live in memory for the session only and are not written to disk.

## Using the app

1. **Files** - upload the raw data (.csv or .xlsx), and the earlier report(s)
   to carry forward. Any workbook with `sys_RespNum` and `Factor` columns
   works, and several can be uploaded at once. Set the date each one was sent
   (guessed from a ddmmyyyy filename) and the date for this round.
2. **Columns** - the ID, panel ID, IP, timing and open-end columns are
   pre-selected by name where possible; change them if the layout differs.
3. **Criteria** - straightliner threshold, speeder cut-off, High LOI on or off,
   duplicate IP on or off.
4. **Grids** - detected automatically from any `NAME_r1, NAME_r2, ...` columns.
   Untick the ones that are not rating grids. "Min answered items" mirrors the
   `count greater than 4` guard in the syntax: 5 for E11 and Q103, and it does
   no harm on the grids that are always asked in full.
5. **Junk open ends** - candidates are listed with the reason; untick a false
   positive, or type in IDs the rules missed. There is also a panel showing the
   shortest verbatims so nothing obvious gets past.
6. **Download** the workbook, and copy the ready-made `IF (any(sys_RespNum,
   ...))` lines into the SPSS syntax.

## Cumulative reporting

IDs from earlier rounds are **not** dropped. Their rows are carried forward
exactly as they were sent, and the IDs new to this round are marked red on
`sys_RespNum` and `psid`. Every row carries a **Reported Date**, and a
**By Date** tab totals the bad IDs per round with a running cumulative count.

## Straightliner rule

Matches `KESHO014_Data_Reports_Syntax.sps`: a grid counts when at least the
minimum number of items are answered and every answer is identical.
`Total_StraighLiner` is the number of grids flagged. Re-running this against
the 16 Sep report reproduces all 520 of its `Total_StraighLiner` values.

## Highlighting

Red on `sys_RespNum` and `psid` for IDs new this round, yellow on the
straightlined grid cells, light blue on a junk Q115, light red on an LOI
outside the thresholds, light green on a duplicate IP. All but the last are
taken from the colours in the client report.
