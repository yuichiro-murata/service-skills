"""Parameter index for design-doc-internal-consistency check 18 (calls across workbooks).

Reads every workbook of ROOT (a WG folder such as 01_Doc\\08_機能定義書\\11_工程管理) in build_screen_list.py's scope
— the PHASE* folders' top level when the WG has them, else recursive; stale folders/files skipped, from the zip only (seconds per book, no Excel, never saves), live
text only (struck runs/cells dropped), visible sheets only, and writes JSON:

    {"format-version": 2, "root": ..., "books": {file: [mtime, size]},
     "programs": {<program ID>: {"book", "sheet", "args": [{"r", "arg", "name", "omit"}],
                                 "rets": [{"r", "ret", "name"}]}},
     "screens":  {<screen ID>: {"book", "sheet", "params": [{"r", "name", "common", "desc"}] | null}},
     "callers":  {<program ID>: [{"book", "sheet", "r"}]}}   — every row of another book's 機能定義書/画面設計書
                                                             that names the ID (check 18 run in reverse)

programs = the callee side of a sub-program call: the 機能定義書 tables `No.|名前|引数名|説明|省略時`
(variants `項目|引数名|…|省略時`, and `項目|説明|省略時` with no 引数名 column — then "arg" is "" and
"name" carries the item) and `No.|名前|戻り値|説明`. A caller's own table (`…|引数名|設定値`) is not a
definition and is skipped. screens = 画面設計書(<ID>) Ⅱ `(1) ﾊﾟﾗﾒｰﾀ定義` rows (`No.|ﾊﾟﾗﾒｰﾀ|説明`, `※共通`
anywhere in the row → common); null when the sheet has no (1).

usage: python build_param_index.py <ROOT> <out.json>
       python build_param_index.py <ROOT> <out.json> --check     (prints FRESH or STALE)
"""
import glob, json, os, re, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlsx_zip import iter_sheets  # noqa: E402

FORMAT = 2
STALE_DIR = re.compile(r"\\(bk|draw\.io|高S\d+対応|90_Branches|開発DDL作成用|90_JAGUR各管理台帳|jira_slack_notifier)\\", re.I)
DOC = re.compile(r"^(機能定義書|画面設計書)")
PID = re.compile(r"(?<![A-Z0-9])([PS][XS]J[A-Z][OB][0-9A-Z]{3})(?![0-9A-Z])")


def ns(s):
    return re.sub(r"\s", "", s or "")


def sheets(path):
    """yield (name, rows{r: {c: text}}) for visible 機能定義書/画面設計書 sheets, live text, cols < 100."""
    for name, _state, rows in iter_sheets(path, name_filter=DOC, max_col=99):
        yield name, rows


def tables(rows, need):
    """yield (header_row, {label: col}, [(r, row)]) for header rows holding every label in need."""
    rs = sorted(rows)
    for i, r in enumerate(rs):
        lab = {ns(v): c for c, v in rows[r].items()}
        if not all(n in lab for n in need):
            continue
        noc = lab.get("No.", lab.get("No", min(lab.values())))
        out = []
        for r2 in rs[i + 1:]:
            row = rows[r2]
            if not ns(row.get(noc, "")).isdigit():
                if not set(row) <= set(lab.values()):
                    break      # a row outside the table's columns ends it
            out.append((r2, row))
        yield r, lab, out


def program_defs(rows):
    args, rets = [], []
    seen = set()
    for need in (("名前", "引数名", "省略時"), ("項目", "引数名", "省略時"), ("項目", "説明", "省略時"), ("名前", "説明", "省略時")):
        for hr, lab, out in tables(rows, need):
            if "設定値" in lab or hr in seen:
                continue
            seen.add(hr)
            nc = lab.get("名前", lab.get("項目"))
            ac = lab.get("引数名")
            for r, row in out:
                arg = ns(row.get(ac, "")) if ac else ""
                name = (row.get(nc) or row.get(nc + 1) or "").strip()
                if arg or (not ac and name):
                    args.append({"r": r, "arg": arg, "name": name, "omit": ns(row.get(lab.get("省略時", -1), ""))})
    for hr, lab, out in tables(rows, ("名前", "戻り値")):
        for r, row in out:
            ret = ns(row.get(lab["戻り値"], ""))
            if ret:
                rets.append({"r": r, "ret": ret, "name": (row.get(lab["名前"]) or "").strip()})
    return args, rets


def screen_params(rows):
    rs = sorted(rows)
    # the (1) heading sits in col 3 on most sheets, col 4 on ~50 (GXJC124A [90,4], GXJC165A [57,4])
    hit = next(((r, c) for r in rs for c in (3, 4, 5) if re.match(r"^\(1\)\s*ﾊﾟﾗﾒｰﾀ", ns(rows[r].get(c, "")))), None)
    if hit is None:
        return None
    start, hc = hit
    # (1) ends at the next (n) heading in cols 3-5 or the next Roman section (Ⅲ．…) in col 2 — the next heading
    # need not share (1)'s column (GSJA601A: (1) in col 3, Ⅲ's (1)… in col 4)
    end = next((r for r in rs if r > start and (re.match(r"^[ⅠⅡⅢⅣⅤⅥⅦⅧ]", ns(rows[r].get(2, "")))
                or any(re.match(r"^\(\d+\)", ns(rows[r].get(c, ""))) for c in (3, 4, 5)))), start + 60)
    params, pc, dc, nc = [], None, None, None
    for r in rs:
        if not start < r < end:
            continue
        row = rows[r]
        lab = {ns(v): c for c, v in row.items()}
        if "ﾊﾟﾗﾒｰﾀ" in lab:
            pc = lab["ﾊﾟﾗﾒｰﾀ"]
            dc = next((c for l, c in lab.items() if l.startswith("説明")), pc + 8)
            nc = lab.get("No.", lab.get("No", pc - 2))
            continue
        if pc and ns(row.get(nc, "")).isdigit() and ns(row.get(pc, "")):
            params.append({"r": r, "name": row[pc].strip(), "common": "※共通" in "".join(row.values()),
                           "desc": (row.get(dc) or "").strip()})
    return params


def books(root):
    """build_screen_list.py's per-WG scope: the PHASE* folders' top level when the WG has them, else recursive."""
    phases = [d for d in glob.glob(os.path.join(root, "PHASE*")) if os.path.isdir(d)]
    cands = [q for d in phases for q in glob.glob(os.path.join(d, "*.xls[xm]"))] if phases else \
            [os.path.join(dp, f) for dp, _dn, fn in os.walk(root) for f in fn if f.lower().endswith((".xlsx", ".xlsm"))]
    return sorted(q for q in cands
                  if not STALE_DIR.search(q) and not re.search(r"^(~\$|完了_)|XXXXX000", os.path.basename(q)))


def main():
    if len(sys.argv) < 3:
        sys.exit(__doc__)
    root, outp = sys.argv[1], sys.argv[2]
    found = {os.path.basename(p): [int(os.path.getmtime(p)), os.path.getsize(p)] for p in books(root)}
    if "--check" in sys.argv:
        try:
            old = json.load(open(outp, encoding="utf-8"))
            ok = old.get("format-version") == FORMAT and old.get("books") == found
        except (OSError, ValueError):
            ok = False
        print("FRESH" if ok else "STALE")
        return
    progs, screens, callers = {}, {}, {}
    for p in books(root):
        b = os.path.basename(p)
        pid = b.split("_")[0]
        try:
            for name, rows in sheets(p):
                for r, row in rows.items():   # who mentions which program ID (for check 18's reverse run)
                    for cid in set(PID.findall(" ".join(row.values()))) - {pid}:
                        callers.setdefault(cid, []).append({"book": b, "sheet": name, "r": r})
                if name.startswith("機能定義書"):
                    a, r = program_defs(rows)
                    if a or r:
                        d = progs.setdefault(pid, {"book": b, "sheet": name, "args": [], "rets": []})
                        d["args"] += a
                        d["rets"] += r
                m = re.match(r"^画面設計書\(([GZ][XS]J[A-Z]\d{3}[A-Z])\)", name)
                if m:
                    screens[m.group(1)] = {"book": b, "sheet": name, "params": screen_params(rows)}
        except Exception as e:  # one unreadable book must not stop the index
            print(f"WARNING: {b}: {e}", file=sys.stderr)
    json.dump({"format-version": FORMAT, "root": root, "books": found, "programs": progs, "screens": screens,
               "callers": callers},
              open(outp, "w", encoding="utf-8"), ensure_ascii=False)
    sys.stdout.reconfigure(encoding="utf-8")
    print(f"books={len(found)} programs-with-defs={len(progs)} screens={len(screens)} "
          f"with-(1)={sum(1 for v in screens.values() if v['params'] is not None)}")


if __name__ == "__main__":
    main()
