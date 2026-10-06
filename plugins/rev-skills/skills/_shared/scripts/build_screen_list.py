"""Screen-name list for design-doc-writing-rules W5: 画面名 -> 画面ID for every visible 画面設計書 sheet
under 01_Doc\\08_機能定義書 (all WGs).

Scope per WG folder (the same rule as build_kind_table.py): where a WG folder has PHASE* sub-folders,
only the workbooks directly inside them; otherwise the folder recursively. Stale copies are skipped —
folders bk / draw.io / 高S<n>対応 / 90_Branches / 開発DDL作成用 / 90_JAGUR各管理台帳 / jira_slack_notifier,
files 完了_* and the XXXXX000 self-check sample, and sheets other than `画面設計書(<ID>)` (so
`…_bak`, `_BK`, `_OLD`, `_20220706`, ` 旧` sheets are out). Cells are read live (struck/gray text
dropped, via live_dump.classify).

The header row is found by its LABEL: the row (3-5) whose col 5 reads 画面ID gives the ID in col 10
and the name in col 15 (row 4 is ﾌﾟﾛｸﾞﾗﾑID in 04_品質管理 / 05_受注出荷). The sheet-name ID is written
too: when several sheets carry the same header ID (PSJCO403 B/C/D all say GSJC403A), W5 must use the
sheet-name ID, and header != sheet-name is naming-standard-compliance's finding (`HEADER-MISMATCH`
lines). A 画面ID cell holding a non-screen ID prints `NON-SCREEN-ID`.

Output TSV: name, screen_id (header), sheet, workbook, sheet_id. Freshness: `# source-count` and
`# source-newest-mtime-utc` (compare as values).

usage: python build_screen_list.py <...\\01_Doc\\08_機能定義書> <out.tsv> [--check]
       --check: print FRESH/STALE for an existing <out.tsv> (format-version, file count with this
       script's own scope rules, newest mtime as a value) and exit 0/1 without rebuilding.
"""
import datetime, glob, os, re, sys, time, warnings
import openpyxl
warnings.filterwarnings("ignore")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from live_dump import classify

STALE_DIR = re.compile(r"\\(bk|draw\.io|高S\d+対応|90_Branches|開発DDL作成用|90_JAGUR各管理台帳|jira_slack_notifier)\\", re.I)
STALE_FILE = re.compile(r"^(~\$|完了_)|XXXXX000")
SHEET = re.compile(r"^画面設計書\(([^)]+)\)\s*$")

FORMAT_VERSION = "2"

def check_fresh(out, files):
    """--check: FRESH only when format-version, source-count and newest mtime (as a value) all match."""
    try:
        head = {}
        for line in open(out, encoding="utf-8-sig"):
            if not line.startswith("#"): break
            if ":" in line: k, v = line[1:].split(":", 1); head[k.strip()] = v.strip()
        newest = max(os.path.getmtime(p) for p in files)
        ok = (head.get("format-version") == FORMAT_VERSION and head.get("source-count") == str(len(files))
              and abs(datetime.datetime.fromisoformat(head["source-newest-mtime-utc"]).timestamp() - newest) < 1)
    except Exception:
        ok = False
    print("FRESH" if ok else "STALE"); sys.exit(0 if ok else 1)

def workbooks(root):
    files = []
    for wg in sorted(d for d in glob.glob(os.path.join(root, "*")) if os.path.isdir(d)):
        phases = [d for d in glob.glob(os.path.join(wg, "PHASE*")) if os.path.isdir(d)]
        cands = [p for d in phases for p in glob.glob(os.path.join(d, "*.xls[xm]"))] if phases else \
                [os.path.join(dp, f) for dp, _dn, fn in os.walk(wg) for f in fn if f.lower().endswith((".xlsx", ".xlsm"))]
        files += [p for p in cands if not STALE_DIR.search(p) and not STALE_FILE.search(os.path.basename(p))]
    return files

def main():
    args = [a for a in sys.argv[1:] if a != "--check"]
    if len(args) != 2: sys.exit(__doc__)
    root, out = args
    files = workbooks(root)
    if not files: sys.exit(f"no workbooks found under {root}")
    if "--check" in sys.argv: check_fresh(out, files)
    newest = max(os.path.getmtime(p) for p in files)
    t0 = time.time(); rows = []; errs = []; skipped = 0
    for p in sorted(files):
        try:
            wb = openpyxl.load_workbook(p, read_only=True, data_only=True, rich_text=True)
        except Exception as e:
            errs.append(f"{p}\t{type(e).__name__}: {e}"); continue
        for ws in wb.worksheets:
            m = SHEET.match(ws.title)
            if not m or ws.sheet_state != "visible": continue
            hit = None
            for r in ws.iter_rows(min_row=3, max_row=5, max_col=20):
                v = {c.column: (classify(c)[0] or "").strip() for c in r if getattr(c, "value", None) is not None}
                if v.get(5) == "画面ID": hit = v; break
            if not hit: skipped += 1; continue
            sid, name = hit.get(10, ""), hit.get(15, "")
            if not sid or sid == "-" or not name or name == "-": skipped += 1; continue
            rel = os.path.relpath(p, root)
            tok = re.match(r"[A-Z]{3,4}\d{3}[A-Z0-9]?", m.group(1).strip())   # GSJA241B-基本情報 -> GSJA241B
            sheet_id = tok.group(0) if tok else m.group(1).strip()
            if not sid.startswith("G"):   # e.g. GXJD601A's 画面ID cell holds PXJDO601 — a doc defect
                print(f"NON-SCREEN-ID\t{sid}\t{ws.title}\t{rel}"); skipped += 1; continue
            if sid != sheet_id:
                print(f"HEADER-MISMATCH\theader={sid}\t{ws.title}\t{rel}")
            rows.append((name, sid, ws.title, rel, sheet_id))
        wb.close()
    with open(out, "w", encoding="utf-8") as f:
        f.write("# lookup: screen-name -> screen-id (header row by its 画面ID label; sheet-name ID in the last column)\n")
        f.write("# scope: 01_Doc\\08_機能定義書, per WG: PHASE* top level where present, else recursive; stale copies skipped\n")
        f.write(f"# format-version: {FORMAT_VERSION}\n")
        f.write(f"# source-count: {len(files)}\n")
        f.write(f"# source-newest-mtime-utc: {datetime.datetime.fromtimestamp(newest, datetime.timezone.utc).isoformat()}\n")
        f.write("# columns: name<TAB>screen_id<TAB>sheet<TAB>workbook<TAB>sheet_id\n")
        for r in rows:
            f.write("\t".join(x.replace("\t", " ").replace("\n", " ") for x in r) + "\n")
    print(f"files={len(files)} screens={len(rows)} skipped={skipped} errors={len(errs)} secs={round(time.time() - t0)}")
    for e in errs: print("ERR", e)

if __name__ == "__main__":
    main()
