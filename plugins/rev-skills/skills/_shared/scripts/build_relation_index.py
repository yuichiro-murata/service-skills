"""Relation index for design-doc-internal-consistency check 19 (program <-> XJC_ｼｽﾃﾑ共通設計書 common processes).

Reads ROOT\\XJC_ｼｽﾃﾑ共通設計書.xlsx and every program workbook of ROOT in build_param_index.py's scope (the PHASE*
folders' top level), from the zip only, live text only, hidden sheets included (the 共通機能利用一覧 sheet is hidden).
A program relates to a common process by

    呼び出し — a live cell (詳細設計* sheets skipped) holding 共通設計 and the process's sheet name ("A"/"A'"), the
               利用一覧 旧名称 「ﾛｯﾄ停止」 with 共通設計 ("A2"), `(<name>)` ("B"), a message code only that process
               defines, XJZ- system messages excluded ("E"), or its program ID ("C" — never seen in 2026-10); a bare name ("D") is recorded but is
               not a relation;
    データ   — a table of one of the program's 更新条件表(<ID>) sheets that exactly one process names (tables named
               by >= 2 processes go to "data_shared" and are ignored).

Also records each process's 呼出元 list (Ⅰ-1, program ID in col 5, 「※呼出なし」 under it) for the list cross-check.

    {"format-version": 1, "root": ..., "books": {file: [mtime, size]}, "xjc": [mtime, size],
     "proc_id": {proc: ID}, "proc_tables": {proc: [tables]}, "tbl_single": {table: proc}, "tbl_shared": {table: [procs]},
     "code2proc": {message code: proc}, "callers_listed": {proc: {program ID: "" | "※呼出なし…"}},
     "programs": {PID: {"book", "updated_tables": [...], "calls": [{"proc","kind","sheet","state","r","c","text"}],
                        "data": [{"proc","table","sheet","state"}], "data_shared": [...]}}}

usage: python build_relation_index.py <ROOT> <out.json>
       python build_relation_index.py <ROOT> <out.json> --check     (prints FRESH or STALE)
ROOT = the WG folder holding XJC_ｼｽﾃﾑ共通設計書.xlsx (01_Doc\\08_機能定義書\\11_工程管理). ~150 s for 121 books.
"""
import collections, json, os, re, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlsx_zip import iter_sheets  # noqa: E402
from build_param_index import books  # noqa: E402

FORMAT = 1
XJC_NAME = "XJC_ｼｽﾃﾑ共通設計書.xlsx"
PROC = ["前後工程設定", "着手完了判断", "加工期限ﾁｪｯｸ", "ﾛｯﾄ停止ﾁｪｯｸ", "QA判定", "実績表項目設定", "個数計算", "製造ﾛｯﾄNo自動採番"]
ALIAS = {"ﾛｯﾄ停止ﾁｪｯｸ": re.compile("「ﾛｯﾄ停止」")}   # 共通機能利用一覧 col 31 旧名称; counted only beside 共通設計
TBL = re.compile(r"\b([TV][SX]J[A-Z][A-Z]\d{3}(?:_\d+)?)\b")
MSG = re.compile(r"\b[SX]J[CZ]-\d{6}\b")
PID = re.compile(r"[PS][SX]J[A-Z]{2}\w{3}$")


def base_of(sheet):
    """ﾁｪｯｸ処理設計書(製造ﾛｯﾄNo自動採番_SMD) belongs to 製造ﾛｯﾄNo自動採番."""
    return re.sub(r"^ﾁｪｯｸ処理設計書\((製造ﾛｯﾄNo自動採番)_.*\)$", r"\1", sheet)


def read_xjc(path):
    proc_tables, proc_id, uselist, codes, callers = {}, {}, {}, {}, {}
    for name, _state, rows in iter_sheets(path, include_hidden=True, max_col=104):
        if name == "共通機能利用一覧":
            for r in range(21, 40):
                nm = rows.get(r, {}).get(14, "").strip()
                if nm:
                    uselist[nm] = rows.get(r, {}).get(7, "").strip()
            continue
        base = base_of(name)
        if base not in PROC:
            continue
        proc_id.setdefault(base, rows.get(4, {}).get(10, "").strip())
        tabs = proc_tables.setdefault(base, set())
        for r, cs in rows.items():
            for t in cs.values():
                if r >= 6:
                    tabs.update(tb.split("_")[0] for tb in TBL.findall(t))
                for m in MSG.findall(t):
                    codes.setdefault(m, set()).add(base)
        if name == base:   # Ⅰ-1 呼出元: program IDs in col 5 above the ﾊﾟﾗﾒｰﾀ定義 heading
            lst = {}
            for r in sorted(rows):
                if r > 10 and any("ﾊﾟﾗﾒｰﾀ定義" in t for t in rows[r].values()):
                    break
                v = rows[r].get(5, "").strip()
                if PID.match(v):
                    nxt = rows.get(r + 1, {}).get(10, "")
                    lst[v] = nxt if "呼出なし" in nxt else ""
            callers[base] = lst
    id2proc = {i: p for p, i in proc_id.items() if re.match(r"S[SX]JCB\d{3}$", i)}
    for nm, i in uselist.items():
        if nm in PROC and re.match(r"S[SX]JCB\d{3}$", i):
            id2proc.setdefault(i, nm)
    count = collections.Counter(t for s in proc_tables.values() for t in s)
    single = {t: p for p, s in proc_tables.items() for t in s if count[t] == 1}
    shared = {t: sorted(p for p, s in proc_tables.items() if t in s) for t in count if count[t] >= 2}
    # XJZ-xxxxxx are the system-wide messages every screen uses — never a sign of a call
    code2proc = {m: next(iter(ps)) for m, ps in codes.items() if len(ps) == 1 and not m.startswith("XJZ")}
    return {"proc_id": proc_id, "id2proc": id2proc, "proc_tables": {p: sorted(s) for p, s in proc_tables.items()},
            "tbl_single": single, "tbl_shared": shared, "code2proc": code2proc, "callers_listed": callers}


def scan_program(path, x):
    names = {p: re.compile(re.escape(p)) for p in PROC}
    ids = re.compile("|".join(sorted(x["id2proc"], key=len, reverse=True))) if x["id2proc"] else None
    rec = {"book": os.path.basename(path), "updated_tables": [], "calls": [], "data": [], "data_shared": []}
    for name, state, rows in iter_sheets(path, include_hidden=True, max_col=104):
        if name.startswith("詳細設計"):
            continue
        m = re.match(r"\s*更新条件表\(([^)]+)\)", name)
        if m:
            rec["updated_tables"].append({"table": m.group(1).strip(), "sheet": name, "state": state})
        for r, cs in rows.items():
            for c, t in cs.items():
                hits = []
                for p, rx in names.items():
                    if not rx.search(t):
                        continue
                    if "共通設計" in t:
                        hits.append((p, "A" if ("「" + p) in t or ("." + p) in t else "A'"))
                    elif "(" + p in t or "（" + p in t:
                        hits.append((p, "B"))
                    else:
                        hits.append((p, "D"))
                if "共通設計" in t:
                    hits += [(p, "A2") for p, rx in ALIAS.items() if rx.search(t) and not names[p].search(t)]
                hits += [(x["code2proc"][m], "E " + m) for m in set(MSG.findall(t)) if m in x["code2proc"]]
                if ids:
                    hits += [(x["id2proc"][i], "C " + i) for i in ids.findall(t)]
                rec["calls"] += [{"proc": p, "kind": k, "sheet": name, "state": state, "r": r, "c": c, "text": t[:160]}
                                 for p, k in hits]
    for ut in rec["updated_tables"]:
        tb = ut["table"].split("_")[0]
        if tb in x["tbl_single"]:
            rec["data"].append({"proc": x["tbl_single"][tb], "table": tb, "sheet": ut["sheet"], "state": ut["state"]})
        elif tb in x["tbl_shared"]:
            rec["data_shared"].append({"table": tb, "procs": x["tbl_shared"][tb], "sheet": ut["sheet"]})
    return rec


def main():
    if len(sys.argv) < 3:
        sys.exit(__doc__)
    root, outp = sys.argv[1], sys.argv[2]
    xjc = os.path.join(root, XJC_NAME)
    if not os.path.exists(xjc):
        sys.exit(f"{XJC_NAME} not found under {root}")
    progs_paths = [p for p in books(root) if PID.match(os.path.basename(p).split("_")[0])]
    found = {os.path.basename(p): [int(os.path.getmtime(p)), os.path.getsize(p)] for p in progs_paths}
    xstamp = [int(os.path.getmtime(xjc)), os.path.getsize(xjc)]
    if "--check" in sys.argv:
        try:
            old = json.load(open(outp, encoding="utf-8"))
            ok = old.get("format-version") == FORMAT and old.get("books") == found and old.get("xjc") == xstamp
        except (OSError, ValueError):
            ok = False
        print("FRESH" if ok else "STALE")
        return
    x = read_xjc(xjc)
    programs = {}
    for p in progs_paths:
        try:
            programs[os.path.basename(p).split("_")[0]] = scan_program(p, x)
        except Exception as e:  # one unreadable book must not stop the index
            print(f"WARNING: {os.path.basename(p)}: {e}", file=sys.stderr)
    json.dump({"format-version": FORMAT, "root": root, "books": found, "xjc": xstamp, **x, "programs": programs},
              open(outp, "w", encoding="utf-8"), ensure_ascii=False)
    sys.stdout.reconfigure(encoding="utf-8")
    related = sum(1 for r in programs.values()
                  if r["data"] or any(c["kind"][0] in "ABCE" for c in r["calls"]))
    print(f"books={len(found)} related={related} single-process tables={len(x['tbl_single'])} "
          f"shared tables={len(x['tbl_shared'])} process codes={len(x['code2proc'])}")


if __name__ == "__main__":
    main()
