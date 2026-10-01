# REV agent guide — what every check agent needs, and nothing else

This is the one shared document every single-program check skill reads. It replaces the
boilerplate each SKILL.md used to repeat, and it carries the parts of the dump and index docs that a
check agent actually uses. The heavy material — how to *produce* a dump, how to recover a crashed
Excel, how to *build* an index — stays in the orchestrator docs, so an agent handed a ready dump does
not pay for it.

## Which document to read

| You are… | Read |
|---|---|
| A check agent handed a ready dump (normal case under `rev-program-review`) | **This file only.** |
| A check agent that must read a **reference master** (ﾃｰﾌﾞﾙﾚｲｱｳﾄ, `04.ﾒｯｾｰｼﾞ管理_*`, `05.ｼｽﾃﾑ共通設計書`, `06-0x_*一覧_*`, `07.共通項目取得`, …) | This file. Read the master **live** with `scripts/live_dump.py <master> <out_dir> --sheets <regex>` (see "Reading a reference file yourself"). **Never** use the COM cross-session cache in `xlsx-excel-com-dump.md` for a master — it writes raw `Value2` with struck text included, and layouts carry struck cells too (TXJCM006 `[81,40]`). One out_dir per master (or `--prefix`, which also names the digest `_DELETED_DIGEST_<prefix>.txt`); anchor `--sheets` (`^ﾃｰﾌﾞﾙﾚｲｱｳﾄ$` — unanchored also matches `旧ﾃｰﾌﾞﾙﾚｲｱｳﾄ`, `JAG_…`, `…_20251002時点`); `--max-col` for very wide registries. `ｼｽﾃﾑ共通設計書.工程管理共通ﾙｰﾙ` → `05.ｼｽﾃﾑ共通設計書`. Prebuilt-lookup builders: `scripts/build_kbn_index.py` (区分名称), `scripts/build_screen_list.py` (screen names). |
| The orchestrator, or a standalone run that must dump the **target workbook** itself | This file **and** `xlsx-excel-com-dump.md` in full. |
| Anyone building a reference index (画面項目辞書 …) | `reference-index.md` (builders, freshness, subsetting). Reading an index is covered below. |
| `design-doc-formatting-consistency` | Additionally `xlsx-formatting-scan.md`. |

All `_shared/*.md` files ship **inside this plugin** (`<plugin root>/skills/_shared/`), not under
`~/.claude/skills/_shared/`. If a path doesn't resolve, glob `**/rev-skills/**/skills/_shared/<file>`
and read the **marketplace** copy (`.claude/plugins/marketplaces/rev-skills/...`) — the
`.claude/plugins/cache/...` copies lag behind it.

## Rules common to every check skill

**Scope selection comes first.** When the user asks to REV a single program's design-doc workbook
without naming specific checks, `rev-program-review` is the entry point: it presents the
single-program checks as a checkbox list, runs only the selected ones as one combined pass, dumps the
workbook once up front and shares the text with every check. Do not unconditionally launch all of
them yourself, and do not post a per-check status update — the combined report is posted once, after
every selected check has finished. A check skill runs on its own only when it was one of the selected
checks, or when the user asked for that check by name.

**Called from `rev-program-review`?** The workbook is already dumped and the scope already chosen:
read the dump paths you were given, don't re-dump, don't launch other checks. **Standalone?** Then
you are your own orchestrator: dump the target per `xlsx-excel-com-dump.md`, and build any index you
need per `reference-index.md`.

**Environment (this machine).**
- The Bash tool is broken (`add_item ("\??\C:\Program Files\Git", "/", ...) failed`) — use PowerShell.
- `python` (CPython 3.13, openpyxl 3.1.5) works; `python3` is a Store stub. Write scripts with the
  Write tool and run `python <path>`.
- Cylance blocks running a `.ps1` file: pass PowerShell inline, or `Invoke-Expression (Get-Content … -Raw)`.
- Read dump `.txt` files with the Read tool; **parse them in code only through `scripts/dump_cells.py`**
  (`load()` groups tokens by `[r,c]`). A hand-written line-by-line parser drops every token after an
  in-cell newline — that produced a false 中 finding in a full REV (PSJCO501 `TSJCA311` row 70).
- Set `$env:PYTHONIOENCODING='utf-8'` before running python from PowerShell, or printing half-width
  katakana fails under cp932.
- **openpyxl cannot open every workbook**: `PXJCO130_ﾛｯﾄﾄﾚｰｽ.xlsx` raises
  `There is no item named 'xl/drawings/NULL' in the archive` (a dangling drawing relationship) while
  Excel opens it fine. If the fallback dies this way, use the COM dump for that workbook — never skip it.
- Excel COM dump is the default reader for the target; the openpyxl live dump
  (`scripts/live_dump.py`, same output format, digest one line per cell) is the **validated** fallback
  and the standalone route — it reproduced an independent openpyxl implementation cell for cell on
  SXJCB147 (1,451 struck / 396 partial).

**Don't dump or read `詳細設計*` / `*画面ｲﾒｰｼﾞ*` sheets** — no check reads them (the shared dump skips
them and records only their size). `design-doc-formatting-consistency` documents its own deliberate
exception note.

**The dump is live** — struck/gray content is already gone; **do NOT run your own strikethrough/gray
scan** (font-size/merge scanning is `design-doc-formatting-consistency`'s separate job). The
orchestrator passes the per-sheet `dead=`/`partial=` counts in your prompt; quote them for the one-line
"excluded at dump time" note, and omit that line if none were given. See "Excluding struck-through / grayed-out rows" below for what that means and when to open
`_DELETED_DIGEST.txt`.

**Reporting conventions (all checks).** The output goes to the designer who owns the doc. Per finding:
the defect in one or two sentences, the exact sheet and cell (`[row,col]` in chat; the Excel 指摘一覧
converts to A1), the master file it was checked against when there is one, and a concrete fix when
obvious. Order by impact. Mark a judgment call (a convention that might be deliberate, a cross-WG
reference you couldn't verify) as needing the designer's confirmation rather than asserting it.
Omit: "確認済み・問題なし" lines, tallies of what was checked, and narration of your own process. The
one exception: a section or master you could not check at all — one line, because that is a coverage
gap the designer should know about.

## Common reference files (where things live)

| Master | Path (under the Unicorn root) |
|---|---|
| ｼｽﾃﾑ共通設計書 (各ID採番, ﾌｧｲﾙ出力制御, 一覧表示制御, 区分 rules …) | `01_Doc\04_共通設計\05.ｼｽﾃﾑ共通設計書.xlsx` |
| 共通項目取得 (delegated get-item blocks) | `01_Doc\04_共通設計\07.共通項目取得.xlsx` — sheet `共通項目取得` plus WG sheets such as `共通項目取得(工程管理)`; search both, items migrate between them |
| 区分名称 | `01_Doc\04_共通設計\09.区分名称_step2.xlsx`, sheet `区分名称_STEP2～` only (see "区分名称/区分コード lookups" below) |
| ﾒｯｾｰｼﾞ管理 | `01_Doc\04_共通設計\04.ﾒｯｾｰｼﾞ管理_<WG>.xlsx` (`_共通` for XJZ/SJZ) |
| 画面項目辞書 | `01_Doc\04_共通設計\82.画面項目辞書_<WG>.xlsx` — read via an index |
| 機能一覧 / 帳票一覧 / ﾌｧｲﾙ一覧 / DB一覧 / ﾚｽﾎﾟﾝｽ一覧 | `01_Doc\06_システム設計書（一覧、管理台帳）\06-01_*` / `06-03_*` / `06-04_*` / `06-06_*` / `06-09_*` |
| ﾃｰﾌﾞﾙﾚｲｱｳﾄ / ﾌｧｲﾙﾚｲｱｳﾄ | `<nn>_<WG名>WG\07_データベース・ファイル設計書(仮)\` (e.g. `11_工程管理WG\…`, with `PH2`/`PH3`/`ファイルレイアウト` subfolders) and `01_Doc\07_データベース・ファイル設計書\` (flat level; `07_03_ファイルレイアウト\<nn>_<WG名>\` for files) |

Header labels in these registries are often spaced out with full-width spaces (`フ　ァ　イ　ル　名　称`):
**strip all whitespace, half- and full-width, before matching a label.**

## Reading a reference file yourself (openpyxl live read)

When a check reads a master directly (a registry, `07.共通項目取得`, `09.区分名称`), apply the same
"live" semantics as the dump — never compare against struck or gray text. **The easy way is the
script**, which implements every rule below and writes the shared-dump format:

    python <plugin>/skills/_shared/scripts/live_dump.py <master.xlsx> <out_dir> --sheets "<regex>"

and then `dump_cells.py` to read the result. Write your own code only when you need cells the
script skips (hidden sheets: `--hidden`). The rules, for that case:

- `openpyxl.load_workbook(path, data_only=True, rich_text=True)` (not `read_only` — it drops rich text).
- **Merged ranges**: openpyxl returns `None` for every non-anchor cell. Keep it that way — Excel can
  store hidden values there (SXJCB147 `G704:P704` all hold `YOTO`), and the COM dump now drops them too.
- **Cell-level**: `cell.font.strike` → dead; font colour gray (`rgb` with r = g = b and 80 < r < 220)
  → dead. Theme-colour grays are not detected by this test; that is an accepted gap.
- **Rich text** (`CellRichText`): a `TextBlock` with its own font **overrides** the cell font
  (`strike is None` there means *not* struck); a bare `str` part has no run properties and
  **inherits** the cell font. Keep the live parts only.
- **A partially-struck cell is a rename in place, not a dead row** — keep its live remainder; drop a
  registry row only when its **key** cell (the ID) has no live text left.
- Dates come back as `datetime`; compare as values, not as the Excel serial.

## Reading a reference index or a prebuilt lookup

The orchestrator builds indexes (`reference-index.md`) and folder-wide lookups (the 区分名称 index,
the screen-name list) and names their paths in your prompt. For the 画面項目辞書 you may get a
**subset** for this program (read whole), the **full** index (grep it), or both. An ID absent from
the subset is the "unregistered" finding — confirm with one grep of the full index before reporting.
**If your prompt names a prebuilt file and it is missing or stale, say so and stop.** If your prompt
says nothing was prebuilt (or you run standalone), build what your check needs yourself, as the
check describes.

### Which dictionary file: route by the ID's own prefix, not by the reviewing WG

`82.画面項目辞書_*.xlsx` is split by JOBコード, not by who is doing the review. A 工程管理 program that
puts a shared item on its screen cites an `XJZ…`/`SJZ…` ID, and that ID is registered in
`82.画面項目辞書_共通.xlsx` — it is not in `_工程管理` at all. Look for it there and you report a
registered item as unregistered.

| ID prefix | Dictionary file (under `01_Doc/04_共通設計/`) |
|---|---|
| `XJZ` / `SJZ` | `82.画面項目辞書_共通.xlsx` |
| `XJA` / `SJA` | `82.画面項目辞書_基準情報.xlsx` |
| `XJB` / `SJB` | `82.画面項目辞書_受注出荷.xlsx` |
| `XJC` / `SJC` | `82.画面項目辞書_工程管理.xlsx` |
| `XJD` / `SJD` | `82.画面項目辞書_品質管理.xlsx` |

A `MENU`-prefixed ID carries its routing prefix immediately after `MENU` — `MENUXJCP00` → 工程管理,
`MENUXJZP00` → 共通.

This is the routing rule `04.ﾒｯｾｰｼﾞ管理_*.xlsx` already follows (see `xlsx-excel-com-dump.md`): a
foreign prefix routes to that other WG's own file, **never** to a same-named sheet embedded in the
copy you happen to have open. In this file that matters twice over, because the sheets the 品質管理
and 受注出荷 copies carry under other WGs' names are not dictionaries at all — see the builder's
sheet table in `reference-index.md`.

**Build one index per dictionary file the program's IDs route to — never one merged index.** The
freshness contract below is single-source: a merged index has no meaningful `# source-mtime-utc`.
Collect the distinct prefixes from the program's dump, build an index for each file they name,
subset each, and hand the agent all of them. Most programs touch one or two files. The builder needs
no per-file configuration — it locates dictionary sheets by their header, so the same call handles
every one of the five.

### Reading the index back: one record is not always one line

Because the format is lossless, a value containing a newline makes **one record span several
physical lines**. In the 工程管理 dictionary index that is 2 records (`SJC7009`, `SJC0633`) spread
over 6 extra physical lines, and one more record has a comma inside its quoted name
(`SJC0600,"0001～ZZZZ (0～9,A～Z(I,O,Q除く))"`).

So a line-oriented read of the index is wrong in three specific ways, all silent:

- **Counting records by counting lines over-counts.** The 工程管理 index has 3,288 records on 3,294
  data lines. (Confirmed the hard way: the discrepancy looked like a builder bug and cost a round of
  debugging.)
- **`$_ -split ','` splits inside quotes**, so the `name` field comes back truncated for the record
  above.
- **A continuation line matches no id and does not start with `#`**, so a naive filter drops it and
  leaves the record's first line with an unterminated quote — a malformed CSV handed to an agent.

Grep/`Select-String` for a single id is fine (an id never contains a newline, and it is the first
field). Anything that walks records must track quote state — see the subsetting snippet in
`reference-index.md`.

## Folders to always exclude from project-wide searches

When any skill searches broadly across the whole `Unicorn` project root (Glob/Grep for a table ID,
a design-doc filename, a DB layout file, etc.), always exclude these path patterns up front —
each confirmed out of scope by the user directly. A file found only under one of these should never
be treated as a real hit (e.g. a table's layout "found" only in `90_Branches/` is not the same as it
being found in the real WG folder structure).

**Quick reference — exclude these from every recursive search's scope:**

- `jira_slack_notifier/**`
- `90_Branches/**`
- `01_Doc/04_共通設計/90_JAGUR各管理台帳/**`
- `*_*WG/開発DDL作成用*/**` (any WG, any date suffix)
- `**/11_単体テスト/**/サンプルデータ/**`
- `**/11_単体テスト/**/参考データ/**`

**Why** (background, skip on a routine run):

- `jira_slack_notifier/` — unrelated internal tooling repo, not a design-doc/DB-layout folder.
- `90_Branches/` — personal/WIP branch copies, not authoritative.
- `01_Doc/04_共通設計/90_JAGUR各管理台帳/` — superseded pre-migration ledger copies.
  `.claude/settings.json` denies `Read` here, but Glob/Grep and the Excel-COM dump script (opens by
  raw path, bypassing the `Read` deny) still see these files — exclude explicitly regardless.
- `<WG>WG\開発DDL作成用<date>\` — always excluded; a former exception for cross-WG table-layout
  lookup no longer applies.
- `.../11_単体テスト/.../サンプルデータ/` and `.../参考データ/` — test-data files whose names match
  real table IDs (e.g. `JAGUR.TXJAM008.csv`); confirmed never real evidence across multiple reviews.

## Don't parallelize a large table/sheet list across your own sub-agents

However large the list (e.g. `PSJCO308_焼成ｻﾔ詰め組み.xlsx`'s 29 更新条件表 sheets, or its 2092-row
画面設計書), work through it yourself, sequentially — do not fan out to your own sub-agents to split
it. Confirmed real failure on `PSJCO308`: both `design-doc-io-table-check` and `xlsx-db-column-check`
split their 29 tables across 4-7 sub-agents each; several got broken/empty prompts and returned
nothing, others were still running when the parent's own turn ended (forcing the orchestrating
session to step in and force a synthesis), and the extra request volume materially contributed to
the session hitting its rate limit. None of it was faster or more thorough than working the list
directly — it only added failure points and cost.

If a list is too large for one pass, work it in batches within your own turn (e.g. 5-10 tables at a
time, repeat) and report any remaining gap explicitly as a coverage limitation — don't silently drop
scope, and don't spin up sub-agents to cover the rest. Only the orchestrating session decides whether
to split a REV across multiple parallel agents.

## Excluding struck-through / grayed-out rows from review

**This is already done for you. Do not run your own strikethrough or gray-out scan.** The dump
script resolves both while the workbook is open: fully-struck and gray cells are absent from the
`.txt`, and partially-struck cells carry only their live text. A row you can see in the dump is
live; a row that was struck simply is not there. Re-deriving this per agent was the single largest
source of duplicated cost (~49k tokens × 5-6 agents) **and** of false findings — see "What this
prevents" below.

Read this section to understand what the dump has already decided on your behalf, and when to reach
for `_DELETED_DIGEST.txt`.

### Why the content is excluded at all

Design docs are often copied from an older workbook and edited in place. This project's writing-rule
checklist (`01_Doc/99.共通資料/設計書記述ルール/05.設計書記述ルール_チェックリスト.xlsx`, sheet
"ﾁｪｯｸﾘｽﾄ", item 1-5) says authors should delete struck-through rows entirely and turn red
"to-be-fixed" text back to black when reusing a workbook this way — but that cleanup is often
skipped, leaving struck-through/grayed rows still physically present. **Content marked this way is
inactive/deleted-in-spirit and must never be treated as something the program actually
references** — don't count it toward a "referenced columns" list, don't flag it as a
naming/consistency violation, don't cite it as evidence of what the current design does. Note that
red alone does *not* mean deleted (red is the "to-be-fixed" marker and is still live); only
strikethrough and gray are.

A one-line "N struck-through/grayed cells excluded at dump time" note in the final report is enough —
never enumerate them as findings. The per-sheet `dead=`/`partial=` counts the dump script prints are
exactly that number.

**An empty or near-empty live dump for a sheet is meaningful, not a dump failure.** When an entire
更新条件表/画面設計書 sheet comes back struck top to bottom — including its own header — the correct
read is that the **whole sheet is superseded** (commonly: an old JAGUR-era 更新条件表(<TableID>)
sheet that a newer 更新条件表(<TableID>WF) sheet has replaced). Exclude the whole thing from the
review and say so. Per explicit user direction, red+strikethrough always means "exclude from REV
scope" in this project, with no exception for how broadly it is applied — do not reason your way out
of it ("this can't be meaningful, it's on every row including the header, must be leftover noise").
That reasoning has been tried and explicitly overruled.

### When to read `_DELETED_DIGEST.txt`

**A digest with no entries means nothing on those sheets was struck or gray** (the COM digest is then
0 bytes; `live_dump.py`'s still has its two `#` header lines). **Its line format depends on who
produced it**: the COM dump's grouped blocks (next paragraph, bare numbers omitted as filler) or
`live_dump.py`'s one line per cell (`<sheet> [r,c] DEL|GRAY|PART:
value`, numbers kept). A row that *looks* deleted in the digest may have been moved — confirm against
the live dump before calling something removed. Never rely on a struck *number* being
listed: to decide whether a gap in a numbered list is explained by a deletion, look for **any deleted
content on the rows between** the two live numbers.

**COM digest only:** it holds what was removed, grouped into contiguous row blocks, with `DEL` (whole cell),
`GRAY` (whole cell, gray) and `PART` (`raw=` vs `live=` for a partially-struck cell) entries. Only
structural filler is left out — a bare row/condition number, a hyphen placeholder, a comparison
operator — and each block still reports how many it dropped (`+ N filler cells omitted`). Short but
meaningful cells are listed in full: `(5)`, `※2`, an alias `Y`, `引数`, `ﾗﾝｸ` and the like all
survive the filter, because deletion asymmetries usually hang on exactly those.

Most checks never need it — they are asking "is what the doc *currently* says correct?", and the live
dump answers that directly. Reach for the digest only when a finding turns on **what was deleted**,
which in practice is `design-doc-internal-consistency`:

- A section whose counterpart was deleted while it stayed live — e.g. a ※-note whose only reference
  was removed, or a "(通常/識別ｶｰﾄﾞ)" wording left behind after the 識別ｶｰﾄﾞ output rows were deleted.
  The live dump shows the survivor; only the digest shows that its partner is gone.
- A step-number or branch table with a hole in it, where the missing branch was struck rather than
  renumbered.
- A "削除" revision note sitting next to content that is *not* struck — a real and load-bearing
  ambiguity (confirmed on `SXJCB147`'s `帳票設計書(RXJC042)` 検索条件 No.5, where the surrounding
  same-day deletions *were* struck and this one was not).

If you are about to report a section as "a leftover that should have been deleted but wasn't cleaned
up," check the digest first — it may already be marked dead, in which case there is nothing to fix.
Confirmed for real on `XJC_ｼｽﾃﾑ共通設計書.xlsx`, sheet `ﾛｯﾄ停止ﾁｪｯｸ`, rows 829-832 (a
"結合条件(A INNER JOIN B)" block left over after its alias B was reassigned from a real table to a
delegated get-item sub-block): a review reported this as a live, not-yet-cleaned-up defect, but every
cell in the block was `Font.Strikethrough=True` — it had already been correctly deleted, and the
"defect" was a pure false positive.

### What this prevents

Every one of these was a real false finding produced by an agent doing its own scan, or skipping it:

- **Six tables reported as "used but not declared in the I/O table"** (PXJCO193), entirely because
  the scan was skipped on a 1000+ row sheet whose 参照ｴﾝﾃｨﾃｨ blocks were red-and-struck-through 100+
  rows deep. Sheet size was exactly why the scan mattered most, and exactly why it got shortcut.
- **`帳票設計書(RSJC035)`'s `注意事項備考` reported as a column that doesn't exist in `TXJCM137`**
  (`SXJCB147`, cells `F202`/`Q202`): raw text was `注意事項備考` / `A.注意事項備考`, but only
  `注意事項` was live. TXJCM137 has `備考` and `注意事項` as separate columns — the doc was correctly
  referencing the latter with a stale, not-fully-deleted `備考` glued on.
- **`COUNT(A.層数層No)` reported as referencing a nonexistent column** and "fixed" to
  `COUNT(A.層No)` (`PXJCO124_ﾛｯﾄ振向け.xlsx`, `ﾁｪｯｸ処理設計書(GXJC124A)`, `[280,17]`) — `層No` was
  already the live text; `層数` was the struck leftover. The doc was already correct.
- **`"GSXJC205A"` reported as a typo of the screen ID `GSJC205A`** (`PSJCO205_返品処置指示発行.xlsx`,
  `画面設計書(GSJC205A)`, `[333,28]`) — only the `X` was struck; live text was an exact match.
- **`③④` treated as two live entity references** (`PSJCO205`, `更新条件表(TXJCM006)`, rows 74/84/95,
  col 16 取得元) when only `④` was live — the sheet's own row-9 legend documents the renumbering that
  stranded `③`.

The last four share one shape: a **mixed-formatting cell**, where `Font.Strikethrough` returns
`System.DBNull` rather than `True`/`False`. That is why the dump script drops to per-character
reconstruction on exactly those cells. The trap is that `System.DBNull` interpolates into a
PowerShell string as an *empty string* — `Write-Output "strike=$whole"` prints `strike=` whether the
value is `False` or `DBNull` — so an ad-hoc scan can read "mixed" as "not struck, live" and never
notice. Test the type explicitly (`-is [System.DBNull]`) and never eyeball a printed value. The dump
script does this; a hand-rolled scan usually doesn't.

### One judgment call the dump can't make for you

**Parallel `処理区分="..."の場合` (or similarly-named) branch variants inside one sheet.** A common
shape in this project's 画面設計書 is two or more parallel sub-blocks handling different values of the
same discriminator (e.g. `品目コード` vs `管理No`, one per branch), often followed by several more
sub-blocks that logically belong to just one of those branches. When one branch is deprecated, the
strikethrough frequently covers not just that one block but every subsequent block that depended on
it too, spanning hundreds of rows and several numbered sub-sections in a row.

In the live dump this appears as a *gap* — a discriminator with no handling. Whether that gap is a
correct deprecation or a genuine design hole is the finding, and it needs the digest plus your own
reading of the surrounding branch structure. `SXJCB147`'s 部門GRP="SC200"(SMD) row, left with no
帳票 after RSJC034/RSJC036 were deleted, is the canonical example of the "genuine hole" case.

## Parsing a dump line

**A logical row is NOT always one physical line.** Cell values in these workbooks routinely contain
newlines — a 属性 cell reading `TextBox⏎※5`, a 初期値 cell reading `新規:空白⏎ｺﾋﾟｰ登録・解除:…` — and
the dump writes the value verbatim, so the rest of that row's `[row,col]=value` tokens continue on
the next physical line(s). Reading the file with `ReadAllLines`/`Get-Content` and treating each line
as a row therefore **silently loses every token after the first embedded newline**.

Confirmed for real, and it produced a wrong conclusion before it was understood: on
`PXJCO128_ﾛｯﾄ停止指示登録.xlsx`'s `画面設計書(GXJC128B)`, row 509's 属性 cell is `TextBox⏎※5`, so the
physical line ends at `[509,18]=TextBox` and `[509,22]`, `[509,24]=30`, `[509,31]`, `[509,34]`,
`[509,36]` all sit on the following line. A line-oriented check read the row as having an empty 桁数,
concluded a real finding was a hallucination, and was only corrected by reading the cell back through
Excel COM. On that one sheet this affects most of the `Ⅴ．画面項目定義` rows.

**So group tokens by the row number inside each `[r,c]=` token, never by physical line.** Scan the
whole file for `\[(\d+),(\d+)\]=` matches and bucket them by `r`; a token's value runs from after its
`]=` up to the next `" | "` or the next `[r,c]=` token, whichever comes first. A quick
`Select-String`/grep for one coordinate is still fine — it is only row-at-a-time reading that breaks.

Within a physical line, split on the literal `" | "` first, then take everything after `]=`
**verbatim** — do not trim the value. Cell values in these workbooks routinely carry significant leading or trailing
spaces (150 of the 980 non-empty cells on `PXJCO192_処置指示登録.xlsx`'s `ﾛｯﾄ管理C　処置指示登録`
sheet alone, including cells that are a single space used as a spacer), which makes a regex like
`\[(\d+),(\d+)\]=([^|]*)` plus a trim silently corrupt them. Verified: parsing verbatim reproduces
all 980 cells of that sheet exactly; the trimming version reported 149 false differences.

**Before concluding a section or ID is missing, check the dump's cap.** The COM dump caps at 3000
rows / 220 columns (`live_dump.py` has no cap unless `--max-col` was given); compare the `rows x cols`
the dump printed for the sheet against that. A sheet larger than the cap needs a targeted re-dump,
not a "not found" finding.

**Large dumps exceed one Read.** A big 画面設計書 dump (~70 KB+) is over the Read tool's limit — page it
with `offset`/`limit`, or pull the rows you need with `dump_cells.py <dump> <row> [row2]`.
**On Windows PowerShell 5.1, `Set-Content -Encoding utf8` / `Out-File` write a BOM**; a list file fed
to python then starts with `\ufeff`. Write such files from python, or read them with `utf-8-sig`.

## Program structure variants to expect

Every REV skill's procedure is written assuming the "default" shape — one プログラムID, one 画面ID,
one 画面設計書 sheet. Two real variants recur often enough in this project that they're worth
checking for up front, before assuming a section is simply missing or that a skill's default
evidence-source doesn't apply:

**A program can have NO 画面設計書 sheet at all** — this happens for a pure batch program (プログラムID
ends in `B`, e.g. `PXJCB102`, `PXJCB134`) or a report-only サブプロ (`S`-prefixed, e.g. `SXJCB147`).
Confirm this from 表紙's "Ⅱ．設計書構成" list (画面設計書 row marked `-`/`-`) before concluding
anything about screen-item checks doesn't apply. When there's no 画面設計書, the column-level
get-item evidence (参照ｴﾝﾃｨﾃｨ blocks, 取得項目/検索条件/結合条件) that a screen program would carry
in 画面設計書's "Ⅲ．画面表示仕様" instead lives in one of two places, and both need checking:
- **機能定義書's own "Ⅳ．機能処理概要"** — a pure batch program (no 帳票設計書 either, e.g.
  `PXJCB102`, `PXJCB134`) writes its 参照ｴﾝﾃｨﾃｨ blocks directly inline in this narrative section,
  numbered as sub-steps (e.g. "(2).共通ｺｰﾄﾞﾏｽﾀより完了予定日に加算する月数を取得する。" followed by
  its own 参照ｴﾝﾃｨﾃｨ/取得項目/検索条件 block).
- **Each 帳票設計書(<ReportID>) sheet's own get-item blocks** — a report-generating サブプロ
  (`SXJCB147` is the confirmed example) delegates almost all of its actual table access to its
  帳票設計書 sheets' own "Ⅲ．編集仕様" sections, each with the same 参照ｴﾝﾃｨﾃｨ/取得項目/検索条件/
  結合条件/ｿｰﾄ順 shape as a screen's get-item block — 機能定義書 itself may carry only one or two
  such blocks (or none) and just says "帳票設計書通りとする。" for the rest. Don't stop at "this
  program's 機能定義書 barely references any tables" — check every 帳票設計書 sheet too before
  concluding a declared table is unused, or that a table used in a report isn't declared.

**A single プログラムID can pair with MORE than one 画面ID** — confirmed for real on
`PSJCO304_着手ﾒｯｾｰｼﾞﾒﾝﾃﾅﾝｽ.xlsx`: one プログラムID (`PSJCO304`) pairs with two screens, `GSJC304A`
and `GSJC304B`, each with its own `画面設計書(GSJC304A)`/`画面設計書(GSJC304B)` sheet and its own
`ﾁｪｯｸ処理設計書(GSJC304A)`/`ﾁｪｯｸ処理設計書(GSJC304B)` sheet — while sharing a single set of
更新条件表 sheets and a single Ⅲ．入出力定義 table in one 機能定義書 sheet. Confirm the actual sheet
list before assuming "one 画面設計書" — if there's more than one, run every per-screen check (event↔
processing-overview, screen-item ID dictionary, message ID, three-way match, item-order consistency,
DB-column extraction) **separately for each screen**, and when verifying I/O-table usage, search
*all* screens' get-item blocks (not just the first one you find) before calling a table "unused."
Also expect a `ﾌｧｲﾙ出力仕様書(<FileID>)` sheet to show up alongside multi-screen programs like this —
treat its own get-item blocks (if any) the same way as a 帳票設計書's, per the batch-program note
above.

## 区分名称/区分コード lookups — always use the STEP2 file directly, and match the group name exactly

A design doc's ※-note frequently delegates a code→label lookup to `ｼｽﾃﾑ共通設計書.区分名称`,
naming a specific group (e.g. "区分名称 ﾛｯﾄ停止区分(ﾛｯﾄ停止指示登録) を参照"). **Always go straight
to `01_Doc/04_共通設計/09.区分名称_step2.xlsx`, sheet `区分名称_STEP2～`, as ground truth** — per
explicit user direction, this is a standing rule, not a per-run judgment call. Do not open the
unsuffixed `01_Doc/04_共通設計/09.区分名称.xlsx` at all (it's a superseded older version), and don't
probe the folder for some other/higher `_stepN` variant each time — STEP2 is the one to use, always.

**Independently of file version, a single 区分名称 file can carry more than one group with very
similar names — an old group and its renamed/expanded replacement coexisting side by side** — so a
keyword search alone (e.g. searching for "ﾛｯﾄ停止区分") can silently match the wrong one. Confirmed
for real: `09.区分名称_step2.xlsx`'s "区分名称_STEP2～" sheet contains BOTH a plain **`ﾛｯﾄ停止区分`**
group (row 846, an old code scheme: `4`=ﾃｰﾌﾟﾛｯﾄNo, `5`=波及範囲検索, `6`=製造ﾛｯﾄNo, no `C`/`D`) AND a
separately-named **`ﾛｯﾄ停止区分(ﾛｯﾄ停止指示登録)`** group (row 2249, the current scheme actually used
by `XJC_ｼｽﾃﾑ共通設計書.xlsx`'s `ﾛｯﾄ停止ﾁｪｯｸ` sheet: `4`=品目ｺｰﾄﾞ+ﾃｰﾌﾟﾛｯﾄNo, `5`=製造ﾛｯﾄNo,
`6`=品目ｺｰﾄﾞ+製造ﾛｯﾄNo(部分一致), through `C`/`D`). A review that matched on the shorter/plainer name
without checking the design doc's own ※-note for the EXACT full group name in parentheses reported a
"stale master, codes don't match" defect that was entirely a wrong-lookup false positive — the
correctly-named group matched the live design doc's codes perfectly. Always copy the group name
verbatim from the design doc's own delegation note (including any parenthesized qualifier) and match
it exactly against the 区分名称 sheet's group-header cells, rather than fuzzy/keyword-matching the
first similarly-named group found.

## Seeing the actual screen layout (画面設計書 Ⅰ.画面ﾚｲｱｳﾄ)

The `Ⅰ.画面ﾚｲｱｳﾄ` section's cells are **empty in every dump** — the layout is a floating picture, not
cell content, so a text dump can never show it. When a check needs the visual layout (item placement,
which group frame an item sits in, whether a control is drawn at all), get the picture instead:

- **Do not** try `Range.CopyPicture` + `ChartObjects().Add` + `Chart.Paste` + `Chart.Export`. It runs
  without error against an invisible Excel instance but writes ~200-byte blank PNGs, because the
  clipboard round-trip does not work headless. (`$ws.ChartObjects` also needs the parenthesised
  `$ws.ChartObjects()` form in PowerShell, or `.Add` reports "does not contain a method named 'Add'".)
- **Do** unzip the workbook and read `xl/media/` directly — no Excel at all:
  `[System.IO.Compression.ZipFile]::OpenRead($xlsx)`, then extract entries under `xl/media/`.
  The layouts are usually `.emf` (vector); convert each with
  `System.Drawing.Imaging.Metafile` → `Bitmap` → `Save(...,ImageFormat::Png)` at ~1.6x scale, then
  Read the PNG. `.png` entries in the same folder are typically pasted screenshots of the *old*
  (JAGUR-era) screen from the `画面ｲﾒｰｼﾞ*` sheet, often carrying review annotations — useful history,
  but not the current design; the `.emf` files are the current layout.
- Map picture → sheet by listing `$ws.Shapes` per sheet (`Type=13` is a picture) and comparing
  `Top`/`Width`/`Height` against the media entries' sizes; a 画面設計書 sheet typically holds one
  picture for the main screen plus one per 明細の続き strip.

`pdftoppm` is not installed, so exporting a sheet to PDF and reading its pages does not work either.

## Workbook template notes

- Read the dumped `.txt` files with the Read tool (not `cat`/Bash) — the `[row,col]=value`
  format is compact enough to scan quickly and lets you cite exact cells back to the user.
- Table/column-layout workbooks in `07_データベース・ファイル設計書(仮)`-style folders follow a
  fixed template: a **"ﾃｰﾌﾞﾙﾚｲｱｳﾄ"** sheet with the table ID/name at row 6
  (`[6,1]=<TableID>` / `[6,6]=<Table name>`), header at row 7
  (`No. | 項目名 | 項目ID | 属性 | 桁数 | DB桁 | I01... | notnull | 備考`), and one data row per
  column from row 8 onward — `[r,3]` is the Japanese column name (項目名), `[r,12]` the physical
  column id, `[r,19]`/`[r,23]` the DB type/length.
- Program/screen design workbooks (機能定義書系) follow a fixed sheet set: 表紙 → 機能定義書 →
  画面設計書 → 詳細設計書 / ﾁｪｯｸ処理設計書 → 更新条件表(<TableID>) — section headers use Roman
  numerals (Ⅰ．Ⅱ．Ⅲ...) which are stable anchors to grep/read for across programs.
