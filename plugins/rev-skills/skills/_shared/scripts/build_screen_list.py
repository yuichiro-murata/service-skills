"""Screen-name list for design-doc-writing-rules W5: 画面名 -> 画面ID for every visible 画面設計書 sheet
under 01_Doc\\08_機能定義書 (all WGs), minus the agent-guide folder exclusions.

The header row is found by its LABEL: the row (3-5) whose col 5 reads 画面ID gives the ID in col 10 and
the name in col 15. Row 4 is not always that row — in 04_品質管理 / 05_受注出荷, row 3 is 機能ID and
row 4 is ﾌﾟﾛｸﾞﾗﾑID (PXJDO102_010_測定.xlsx), so a fixed [4,10]/[4,15] read returns program IDs.
Sheets with no 画面ID label are skipped. Cache freshness: compare the `# source-count` and
`# source-newest-mtime-utc` header lines with the folder (rebuild when either differs).

usage: python build_screen_list.py <root: ...\\01_Doc\\08_機能定義書> <out.tsv>
"""
import datetime, os, re, sys, time, warnings
import openpyxl
warnings.filterwarnings("ignore")

EXCL = re.compile(r"(90_Branches|開発DDL作成用|\\11_単体テスト\\.*\\(サンプルデータ|参考データ)\\|90_JAGUR各管理台帳|jira_slack_notifier)")

def main():
    root, out = sys.argv[1], sys.argv[2]
    files = []
    for dp, _dn, fn in os.walk(root):
        for f in fn:
            p = os.path.join(dp, f)
            if f.lower().endswith((".xlsx", ".xlsm")) and not f.startswith("~$") and not EXCL.search(p):
                files.append(p)
    newest = max(os.path.getmtime(p) for p in files)
    t0 = time.time(); rows = []; errs = []; skipped = 0
    for p in sorted(files):
        try:
            wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
        except Exception as e:
            errs.append(f"{p}\t{type(e).__name__}: {e}"); continue
        for ws in wb.worksheets:
            if not ws.title.startswith("画面設計書") or ws.sheet_state != "visible": continue
            hit = None
            for r in ws.iter_rows(min_row=3, max_row=5, max_col=20, values_only=True):
                if len(r) > 14 and str(r[4] or "").strip() == "画面ID":
                    hit = r; break
            if not hit: skipped += 1; continue
            sid = str(hit[9] or "").strip(); name = str(hit[14] or "").strip()
            if not sid or sid == "-" or not name or name == "-": skipped += 1; continue
            if not sid.startswith("G"):   # e.g. GXJD601A's 画面ID cell holds PXJDO601 — a doc defect
                print(f"NON-SCREEN-ID\t{sid}\t{ws.title}\t{os.path.relpath(p, root)}"); skipped += 1; continue
            rows.append((name, sid, ws.title, os.path.relpath(p, root)))
        wb.close()
    with open(out, "w", encoding="utf-8") as f:
        f.write("# lookup: screen-name -> screen-id (row with col-5 label 画面ID in rows 3-5; id col 10, name col 15)\n")
        f.write("# scope: 01_Doc\\08_機能定義書 all WGs minus agent-guide exclusions\n")
        f.write(f"# source-count: {len(files)}\n")
        f.write(f"# source-newest-mtime-utc: {datetime.datetime.fromtimestamp(newest, datetime.timezone.utc).isoformat()}\n")
        f.write("# columns: name<TAB>screen_id<TAB>sheet<TAB>workbook\n")
        for r in rows:
            f.write("\t".join(x.replace("\t", " ").replace("\n", " ") for x in r) + "\n")
    print(f"files={len(files)} screens={len(rows)} skipped={skipped} errors={len(errs)} secs={round(time.time() - t0)}")
    for e in errs: print("ERR", e)

if __name__ == "__main__":
    main()
