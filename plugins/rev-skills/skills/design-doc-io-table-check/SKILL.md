---
name: design-doc-io-table-check
description: Check whether a program's 機能定義書「Ⅲ．入出力定義」CRUD list is complete and accurate in both directions: every table used in 画面設計書/更新条件表/帳票設計書/ﾌｧｲﾙ入出力仕様書 (including via 07.共通項目取得 delegation) is declared with the right C/R/U/D flags, and every declared table is really used and has a DB一覧 entry and a ﾃｰﾌﾞﾙﾚｲｱｳﾄ file. The highest-yield check in a program REV (delegation tracing, PH2/PH3 layout precedence, exact-ID matching, struck-block exclusion), split from `design-doc-internal-consistency`. Use when the user asks to check 入出力定義表/CRUD一覧 completeness, or whether declared tables match what the doc actually uses. For a full REV without named checks the entry point is `rev-program-review`; run this standalone only when selected there or asked for by name.
---

# design-doc-io-table-check

Checks one program's 機能定義書「Ⅲ．入出力定義」table for completeness in both directions, grounded
in this project's own checklist: `01_Doc/99.共通資料/設計書記述ルール/05.設計書記述ルール_チェックリスト.xlsx`
(sheet "ﾁｪｯｸﾘｽﾄ", items 3-2/3-3) — see the frontmatter `description` above for exactly what "complete
in both directions" covers and why this was split out of `design-doc-internal-consistency`.

## Environment

**Read `_shared/agent-guide.md` first** — scope selection, environment, the live dump, reporting
conventions. Read the DB一覧 **live** (`_shared/scripts/live_dump.py`; it has struck = retired rows) and
the ﾃｰﾌﾞﾙﾚｲｱｳﾄ files live too (`live_dump.py … --sheets "^ﾃｰﾌﾞﾙﾚｲｱｳﾄ$"`, as in `xlsx-db-column-check`
step 3 — never the COM cache "Batch variant", which keeps struck text). Route the DB一覧 by the table
ID's JOBコード (characters 2-4): `XJC`/`SJC` → `06-06_DB一覧_工程管理`, `XJA`/`SJA` → `_基準情報`,
`XJB`/`SJB` → `_受注出荷`, `XJD`/`SJD` → `_品質管理`, `XJZ`/`SJZ` → `_共通` (all in
`01_Doc\06_システム設計書（一覧、管理台帳）`; prefixes confirmed against each file's live ID column,
2026-10-01). A JA table may also sit on 工程管理's `作成状況一覧` (98 `XJA` + 72 `SJA` IDs there) — either
registration counts. A JOBコード outside this table (`XAC`: `DXACM001`-`003` are on `06-06_DB一覧_工程管理`
`DB一覧` rows 248-250) → search all five DB一覧 before calling it unregistered. **Resolve layouts by `xlsx-db-column-check` step 3 alone** — its rule 4 is the
single layout-scope rule: `VIEW\`, `STEP1暫定テーブル\`, `90_JAGURﾃｰﾌﾞﾙﾚｲｱｳﾄ(<date>時点)\` and `10_共通WG`
mirrors are never a fallback, and a table found only there is "layout not present" (name the old copy).
The JAGUR folder alone holds ~30 of `PXJCO125`'s tables, some under old names
(`TXJCD407_ﾁｪｯｸｼｰﾄ実績.xlsx`), which is why no fallback is allowed. Delegated blocks: see
"Delegation depth" below. Anything visible in the dump is live evidence; removed content lives in
`_DELETED_DIGEST.txt`, which this check needs only to explain why a declared table has no live usage
left. Sources: 機能定義書/画面設計書/ﾁｪｯｸ処理設計書/更新条件表/帳票設計書/ﾌｧｲﾙ出力仕様書, and
`ﾌｧｲﾙ入出力仕様書(<ID>)` / `ﾌｧｲﾙ入力仕様書(<ID>)` — their Ⅱ `入出力先` names the table a file is loaded into
(`PSJCO403` `ﾌｧｲﾙ入出力仕様書(FSJC018)` `[15,18]=TSJCD401:社内加工予定金額`, a 洗替 target per Ⅳ B-⑤): that
is usage (C/D for 洗替, C/U for 登録/更新 — read Ⅳ's verb), and the file ID itself is an `F…` row.
Also 画面設計書 Ⅲ 検索条件一覧's `ﾒｲﾝﾃｰﾌﾞﾙ` column (an R; `PSJCO403` `GSJC403A` `[95,31]`-`[107,31]`; `[100,31]`
writes `SJCD401:社内加工予定金額` — TSJCD401 without its leading `T`, itself worth a 低 note).
`【更新条件表(<ID>) <section>】` pointers are resolved by `design-doc-internal-consistency` check 12, not here.

## Procedure

**Quick reference — the checks in order (each detailed below with confirmed real-world examples):**

1. 機能定義書「Ⅲ．入出力定義」often has TWO table-ID sources (a live CRUD summary table + a separate
   INPUT/OUTPUT breakdown table right below it). The breakdown table is frequently dead boilerplate,
   in which case the live dump won't contain it at all — so treat whatever the dump *does* show as
   the authoritative declaration set, and never resurrect IDs from `_DELETED_DIGEST.txt` as
   declarations.
2. Collect every table ID/alias from: Ⅲ．入出力定義 itself, 画面設計書's 参照ｴﾝﾃｨﾃｨ blocks (or the
   batch/report-program equivalents, **including a ﾌｧｲﾙ出力仕様書's own inline 参照ｴﾝﾃｨﾃｨ blocks in
   Ⅰ．ﾌｧｲﾙ出力条件** — 6 of the 工程管理 PHASE1-3 file specs read tables directly there rather than
   delegating to a 画面設計書 block, and each such table owes the program an R), AND every 更新条件表 sheet's *body* (not just its header) —
   the 取得内容/取得条件 column often cites other tables by Japanese name only, never by ID, so
   search by name too before calling a table unused. **Include the 更新条件表 side blocks from col 54**,
   under any of their labels — `取得ﾃｰﾌﾞﾙ名`, `参照ﾃｰﾌﾞﾙ名` (with `参照条件`) and inline `参照ｴﾝﾃｨﾃｨ`; compare
   labels with whitespace stripped. Each is an R on the table it names. The two 高 findings on
   `PXJCO125` (2026-10-01, both "R missing") came from here, and a literal scan missed both:
   `TXJCM012` is read **only** through `参照ﾃｰﾌﾞﾙ名` blocks (`更新条件表(TXJAM068)` `[8,54]`/`[24,54]`/`[97,54]`),
   `TXJCD404` through `参照ﾃｰﾌﾞﾙ名` (`[39,54]`/`[54,54]`/`[113,54]`) and the `TXJCA404` history INSERT's
   `取得ﾃｰﾌﾞﾙ名` (`更新条件表(TXJCA404)` `[8,54]`/`[21,54]`).
3. Follow every `※<共通設計書名>.<項目> 参照` delegation line to its actual target sheet (check both
   the WG-specific sheet and the shared sheet in the common-design workbook) — but a table reached
   **only** through a delegated block is 要確認, never 高 (see "Delegation depth").
4. Deprecated blocks are already gone from the shared live dump — **do not run your own
   struck-through/grayed-out scan.** A 参照ｴﾝﾃｨﾃｨ block visible in the dump is live evidence; a
   deprecated one is simply absent, in both diff directions.
5. Require an **exact** ID string match throughout (never prefix/substring) — suffix variants
   (`WF`, numeric suffixes like `_31`) are different tables from their base ID, and this applies to
   locating the DB layout file too, not just the diff itself. **Before matching, strip a developer
   schema prefix** — `HAYA_TSJCD403` in `PSJCO403` `画面設計書(GSJC403A)`'s col-54 SQL `[295,54]` is
   `TSJCD403` — and note the prefix once (低). Only a leading `<NAME>_` before a table ID; suffixes are
   never stripped.
6. Diff declared vs. used in both directions, then separately verify each C/R/U/D letter against
   what the live references actually do (a table can be "used" without every declared letter being
   backed by a real operation, or vice versa). Also confirm DB一覧 registration and テーブルレイアウト
   existence — search the correct top-level `<WG番号>_<WG名>WG\07_データベース・ファイル設計書(仮)`
   folder (PH3 > PH2 > top-level precedence), plus `01_Doc\07_データベース・ファイル設計書\` for
   基準情報(JA)-prefix common tables, before ever reporting a layout as missing. A C/U/D letter with
   **no 更新条件表 sheet behind it** is not skipped — see "No 更新条件表 for a written table" below.

**Detailed rules and confirmed examples:**

**機能定義書「Ⅲ．入出力定義」routinely contains TWO separate table-ID sources, not one — a CRUD
summary table (columns/rows listing each table ID with its C/R/U/D flags) AND, directly beneath
it, a separate "1．INPUT"/"2．OUTPUT" breakdown table (per-table rows with No./ID/名称/用途
columns) that restates largely the same tables in more prose-like form.** This second, breakdown
table is a well-known trouble spot in this project: it is frequently left ENTIRELY red+struck-through
top to bottom (including its own header row) as leftover, never-cleaned-up boilerplate from the
source workbook it was copied from — while the CRUD summary table right above it in the same
sheet is fully live. Confirmed for real on `PXJCO134_流動停止解除.xlsx`: the OUTPUT breakdown
table at rows 58-64 was entirely red+struck (including a row citing `TXJCM502`/`TXJCA301`, whose
apparent "duplicate No.5" numbering was flagged as a defect by a review that hadn't checked this
block's formatting — a false finding, since the whole block is dead). The same shape was also the
root of a separate finding on `PXJCB102_製造ｵｰﾀﾞｰ完了処理.xlsx`, and again on
`SXJCB147_処置指示発行(ｻﾌﾞﾌﾟﾛ).xlsx` (rows 63-86, a fully struck "1.INPUT"/"2.OUTPUT" pair citing a
dozen tables that appear nowhere else live).

**With a live dump this resolves itself: a dead breakdown table is simply not in the `.txt`.** So
the practical rule is now the reverse of the old one — if 機能定義書 shows only one table-ID block
where you expected two, that is the expected outcome, not a dump failure or a missing section. Do
not go looking for the breakdown table in `_DELETED_DIGEST.txt` and reinstate its IDs as
declarations, and do not report numbering/content oddities inside it; a one-line note that the
breakdown table is dead boilerplate is all it warrants.

Collect every table ID mentioned in "Ⅲ．入出力定義" (機能定義書 sheet). Separately collect every
table ID/alias used in 画面設計書's "参照ｴﾝﾃｨﾃｨ" blocks (or, for a batch program with no
画面設計書, the equivalent inline 参照ｴﾝﾃｨﾃｨ blocks inside 機能定義書's own "Ⅳ．機能処理概要", or a
帳票設計書's own get-item blocks) and in every 更新条件表(<TableID>) sheet's header. **Also read
every 更新条件表 sheet's body**, not just its header — the "取得内容"/"取得条件" column (the source
of each column being set) routinely cites *other* tables that feed values into the one being
updated, and it does this **by the source table's Japanese name only, never by table ID** (e.g.
"④ﾜｰｸﾌﾛｰ承認ﾃﾞｰﾀ(No1)" / "ﾜｰｸﾌﾛｰ承認ﾃﾞｰﾀ.会社ｺｰﾄﾞ" — no "TSJAM722" string appears anywhere in that
sheet). A plain text search for a table ID string will silently miss this kind of reference. Before
concluding a table in the I/O list is "declared but unused," don't stop at an ID-string search —
also search every 更新条件表/画面設計書/帳票設計書 sheet for that table's Japanese name (from
Ⅲ．入出力定義's own 名称 column), since that's frequently the only place it's actually written.
This happened for real: `TSJAM722`(ﾜｰｸﾌﾛｰ承認ﾃﾞｰﾀ) was missed this way — an ID-only search across
the whole workbook came back empty, but the table was genuinely referenced (and should carry an R
flag) once searched by name. **Keep the name search tractable**: search only the 名称 of
declared-but-not-yet-found tables, and only in cells of the 取得ﾃｰﾌﾞﾙ名 / 参照ｴﾝﾃｨﾃｨ / 取得内容 columns — a
full-workbook regex over ~600 names timed out at 600 s on the 9 MB `PXJCO161`.

**Also follow every "※<共通設計書名>.<項目> 参照" delegation line to its actual target sheet** —
画面設計書/機能定義書 routinely delegate a whole get-item block to a shared common-design workbook
(most often `01_Doc/04_共通設計/07.共通項目取得.xlsx`, e.g. "※共通項目取得.グループ名 参照")
instead of writing the 参照ｴﾝﾃｨﾃｨ inline. Don't stop at "this program just says 'see common doc'" —
open the target and collect the tables it reads; how to grade them is "Delegation depth" below. Open the target workbook, find the
matching named section (search both the WG-specific sheet, e.g. `共通項目取得(工程管理)`, and the
shared `共通項目取得` sheet — items get migrated from the WG-specific sheet to the shared one over
time, noted in a revision comment like "改訂履歴No94" when it happens, so the current live copy
may not be where you'd first expect), and pull its 参照ｴﾝﾃｨﾃｨ table(s) as if they appeared inline
in this program's own doc. This happened for real: `TSJAM726`(ﾜｰｸﾌﾛｰ承認ｸﾞﾙｰﾌﾟ) was missed because
the design doc only wrote "※共通項目取得.ｸﾞﾙｰﾌﾟ名 参照" with no table ID or name anywhere in the
program's own workbook — the actual `TSJAM726` reference only existed inside
`07.共通項目取得.xlsx`'s "(37)ｸﾞﾙｰﾌﾟ名" section, which nothing in the program's own workbook
would surface without deliberately following the delegation. Likewise, a batch program's argument
set may not obviously reach a delegated block's own downstream branches (e.g. a delegated
"移動ﾛｯﾄ(最新)" block that itself references a table via an alias fed by a *different* delegated
block) — if the program-level docs alone can't settle whether that path is actually taken, report
it as a judgment call for the designer rather than asserting it either way.

**Delegation depth — one rule.** Two delegation targets occur:
`※共通項目取得.<項目> 参照` → `01_Doc\04_共通設計\07.共通項目取得.xlsx`, and
`※ｼｽﾃﾑ共通設計(書).工程管理共通ﾙｰﾙ.<rule> 参照` (both spellings occur) → sheet `工程管理共通ﾙｰﾙ` of
`01_Doc\04_共通設計\05.ｼｽﾃﾑ共通設計書.xlsx` (dump just that sheet: `--sheets "^工程管理共通ﾙｰﾙ$"`). On
`PXJCO125` the latter appears in ﾁｪｯｸ処理設計書 `[35,59]`/`[37,59]`, `更新条件表(TXJCM003)` `[136,2]` and
`(TSJCD302)` `[9,16]`. A third: `ｼｽﾃﾑ共通設計書(<sheet>)` / `XJC_ｼｽﾃﾑ共通設計書「<sheet>」` naming a sheet 05 does
not have (`着手完了判断`, `実績表項目設定`, `QA判定`, `ﾛｯﾄ停止ﾁｪｯｸ`, `加工期限ﾁｪｯｸ`) →
`01_Doc\08_機能定義書\11_工程管理\XJC_ｼｽﾃﾑ共通設計書.xlsx`, same level-1 rule below.
- **Level 1** is the named section itself **plus its in-section sub-blocks** (`(17-1)`, `(17-2)` under
  `(17)移動ﾛｯﾄ(最新)`) — those are the section's own usage. A further `共通項目取得.X 参照` / `工程管理共通ﾙｰﾙ.X
  参照` written inside the section is level 2: note it, don't chase it.
- **A table read only at level 1 — never in the program's own blocks — is reported as 要確認, never
  高**, with the delegation path cited (`TSJCD504` via `工程管理共通ﾙｰﾙ.製品輸送中ﾁｪｯｸ`, `TSJAM999` via
  `(17-1)実績管理部門GRP取得`, `TXJCA003`'s R via `製造ｵｰﾀﾞｰ今回減算数取得`). Whether such tables belong in
  Ⅲ．入出力定義 at all is an **open decision for the user** — do not decide it either way, and do not
  report a delegated-only table as "declared but unused" either.
- **Declared ID ≠ the table the delegated block reads** — Ⅲ declares `VXJAM005` 加工GRPﾏｽﾀ (`PXJCO161`
  `[85,5]`) but the `共通項目取得.加工GRP` block it stands for reads `VSJAM003` (`(21)加工GRP`, changed by 改訂履歴No.61;
read the target live — the old VXJAM005 block sits just above it): 中/要確認, "declared ID differs
  from the delegated table", not a pair of unused/undeclared findings.
- The `TSJAM726` miss above is why the delegation is still followed: the finding exists, only its
  grade waits on that decision.

**A 参照ｴﾝﾃｨﾃｨ block that is present in the live dump is, by construction, live evidence — the dump
script already dropped the deprecated ones.** This replaces what used to be a mandatory per-block
formatting scan, and it removes the failure mode that made that scan mandatory in the first place:
a real review reported six false "used but not declared" findings on a single large 画面設計書
because a red-and-struck-through `処理区分="..."の場合` branch (and everything under it) was cited as
live evidence, and the same mistake recurred on `SXJCB147_処置指示発行(ｻﾌﾞﾌﾟﾛ).xlsx`, sheet
`帳票設計書(RXJC040)`, on a plain ~10-row block (rows 46-55, "(3)WF情報取得", 参照ｴﾝﾃｨﾃｨ citing
`TXJCM137WF`) — flagged as "`TXJCM137WF` used in 帳票設計書 but absent from 機能定義書's CRUD table"
when the whole block was deprecated and there was no live reference to flag at all. Neither can
happen from a live dump: those blocks aren't in it.

What still needs your attention is the **opposite** direction. A table declared in Ⅲ．入出力定義
whose only usage was in a now-deleted block will look "declared but unused" — which is a real
finding, but the useful report says *why*. Check `_DELETED_DIGEST.txt` for that table before writing
it up, so you can distinguish "the declaration is stale, delete it too" from "the usage was removed
by mistake."

**When comparing table IDs between the I/O table and actual usage, always require an exact
string match — never treat one ID as a match for another just because one is a prefix/substring
of the other.** This project has numeric-suffixed view-ID variants (e.g. `VXJCM004` vs
`VXJCM004_31` vs `VXJCM004_31_ALL`) where the *unsuffixed* base ID commonly has no design file or
real existence of its own at all — only a specific suffixed variant is actually implemented
(mirrors the already-known `WF`-suffix case, e.g. `TXJAM061` vs `TXJAM061WF`, being separate
workbooks). This happened for real: `PXJCO152_処置指示発行.xlsx`'s 機能定義書「Ⅲ．入出力定義」
row 50 (No.12) declares the table ID as bare `VXJCM004`, but every other place this view is
actually used in the same workbook — `更新条件表(TXJCM006)`'s header `[8,68]` and
`ﾁｪｯｸ処理設計書(GXJC152A)`'s 参照ｴﾝﾃｨﾃｨ blocks at `[224,14]`/`[648,14]` — correctly cites the full
`VXJCM004_31`. A substring-based match would call this "consistent" (bare `VXJCM004` is literally
contained in `VXJCM004_31`) and miss that the I/O table's own ID is truncated/incomplete; there is
no design file for a bare `VXJCM004` anywhere in the project, only `VXJCM004_31`-suffixed ones.
Flag any case where an I/O table's declared ID is only a prefix of the ID actually used elsewhere
in the workbook (or vice versa) as its own defect — "declared ID is missing a suffix present in
every live usage" — rather than silently accepting it as the same table. This same exact-match
rule applies to locating the DB layout file below: search for the *exact* full ID (suffix
included), not a fuzzy "contains this substring" search, since a fuzzy search over
`VXJCM004_31_...xlsx`/`VXJCM004_31_ALL_...xlsx` filenames will "find" a file and wrongly appear to
confirm the bare, unsuffixed ID as valid.

Diff the two sets:
- Table used in 画面設計書/更新条件表/帳票設計書 but **absent** from Ⅲ．入出力定義 → flag ("used
  but not declared in the I/O table").
- Table listed in Ⅲ．入出力定義 with a C/R/U/D flag but **never actually referenced** anywhere
  in 画面設計書/更新条件表/帳票設計書 → flag ("declared but unused").
- **For every table you just confirmed IS referenced somewhere (i.e. it passed the previous
  bullet), don't stop there — separately verify each flagged letter (C/R/U/D) against what kind
  of operation the live references actually perform**, table by table. This is easy to skip once
  you've satisfied yourself a table "is used," but "is used" and "is used the way the flags say"
  are different questions — a table can have a live SELECT-style 参照ｴﾝﾃｨﾃｨ block (a real R) while
  its I/O row only carries a C flag from a 更新条件表 sheet, and that gap won't surface unless you
  check flag-by-flag. A real review missed exactly this: `TXJAM025`/`TXJAM026` each had a live,
  unstruck 参照ｴﾝﾃｨﾃｨ reference fetching columns from them (a genuine R), but their I/O rows only
  had a C flag (from their 更新条件表 sheets) — the review confirmed "used, not unused" and moved
  on without checking that the R itself was missing from the flags. Flag any letter present in
  actual usage but absent from the row, and vice versa (a flagged letter with no matching
  operation anywhere). Read a 更新条件表 block's verb with `^[【\s]*(INSERT|UPDATE|DELETE|MERGE)` — labels
  are decorated (`【DELETE】※1` `更新条件表(TXJCM007)` `[16,16]`, `INSERT(新規入力の場合)` `(TXJCD102)` `[91,16]`,
  `DELETE1`/`DELETE2`). Prose that tests existence before writing (`…存在しない場合(未登録の場合)、…`
  `TXJCD318`, `存在する場合UPDATE、存在しない場合INSERT` `TSJCD330`) is an implicit read: R expected, but
  grade a missing R 低/要確認, not 高.
- **No 更新条件表 for a written table.** When Ⅲ marks C/U/D (`○`/`●`) for a table that has no
  `更新条件表(<ID>)` sheet, don't skip the letter check and don't call the letters unbacked: judge them
  from Ⅳ．機能処理概要's wording (登録 → C, 更新 → U, 削除 → D, 洗替 → D+C) and mark each verdict 要確認.
  Then report the missing sheets **once**, as one finding listing every such table — this check owns
  it; `xlsx-db-column-check` and `update-condition-completeness` defer to it. `PSJCO403` (2026-10-06)
  has no 更新条件表 sheet at all while Ⅲ `[106,5]`-`[111,5]` marks C/U/D on TXJAM100, TSJCD400-402,
  TSJCD404 and TXJCM008, and Ⅳ B-③/B-⑤ point at a `【更新条件表(TSJCD404) 洗替登録】` `[170,30]`/`[180,30]`
  that does not exist — cite such dangling pointers in the same finding.
- **An `F…` ID in Ⅲ．入出力定義 is a file, not a table** (`FXJA020` 取引先ﾏｽﾀﾌｧｲﾙ, `FSJC051`
  損金一覧ﾃﾞｰﾀ). It has no DB一覧 row and no ﾃｰﾌﾞﾙﾚｲｱｳﾄ by design, so never report it as
  unregistered or layout-less. Its registration (`06-04_ﾌｧｲﾙ一覧_<WG>.xlsm`) and its agreement with the
  `ﾌｧｲﾙ出力仕様書(<ID>)` sheet are `file-output-spec-check` F1/F2. What does stay here is the plain
  both-direction diff: a visible `ﾌｧｲﾙ出力仕様書(<ID>)` sheet whose ID has no 入出力定義 row is "used but
  not declared", exactly like a table.
- Also check whether each table ID appears in the DB一覧 and whether its ﾃｰﾌﾞﾙﾚｲｱｳﾄ workbook exists.
  **The DB一覧 is per-WG** — `01_Doc/06_システム設計書（一覧、管理台帳）/06-06_DB一覧_<WG名>.xlsx`
  (e.g. `06-06_DB一覧_工程管理.xlsx`; route by JOBコード as in Environment); the `_共通` copy holds only
  the shared tables, so checking it for a WG's own table reports a false "unregistered". Note the WG file carries **two overlapping
  table lists on two visible sheets with different header labels** — `作成状況一覧` (header row 1,
  `ﾃｰﾌﾞﾙID`/`ﾃｰﾌﾞﾙ名`) and `DB一覧` (header row 35, `ID`/`名称`) — and neither is a superset of the
  other. Check both before reporting a table as unregistered. Measured 2026-10-01 on a live dump
  (source length 401,141 / mtime `2026-09-08T09:55:22Z`; distinct ID-shaped cells per sheet): 353 live
  IDs on `作成状況一覧` and 211 on `DB一覧`, sharing 180. Re-measured **2026-10-06** (length 399,384 /
  mtime 2026-10-03): `作成状況一覧` is 355 rows / 353 IDs, `DB一覧` 212 IDs (dead=28 partial=19). Struck
  `DB一覧` rows are retired — the live dump already drops them; never count them as registered. These
  figures go stale within days: re-measure rather than trust them. The third visible sheet (`改訂履歴`) is not a table list at all.
  When resolving a table to its layout file, confirm the ID in the **`ﾃｰﾌﾞﾙﾚｲｱｳﾄ` sheet's** `A6`
  cell (under the `ﾃｰﾌﾞﾙID` label in `A5`): a `<ID>_*.xlsx` glob also matches suffixed *other* tables
  (`TXJCM003_B`, `TXJAM008_IN`), and accepting one silently checks the wrong table's column list.
  Take `A6` from that sheet specifically — the `旧ﾃｰﾌﾞﾙﾚｲｱｳﾄ` sheets in the same workbook hold the
  legacy ID there (`FDMBM03`, `FDCJM03` in `TXJCM003_製造ｵｰﾀﾞｰ.xlsx`), so reading the wrong sheet
  makes the correct file look like a mismatch. Select that sheet by exact name (`^ﾃｰﾌﾞﾙﾚｲｱｳﾄ$`): dated
  snapshots such as `ﾃｰﾌﾞﾙﾚｲｱｳﾄ_20251002時点` carry the same `A6` and would pass the check. Two files
  with the same `A6` in one phase folder (PH3 `TSJCD401_社内加工(予定金額).xlsx` and
  `TSJCD401_社内加工予定金額ﾃﾞｰﾀ(IS).xlsx`, different columns) are settled by matching `F6` against the
  DB一覧 名称 — `xlsx-db-column-check` step 3 rule 2 — and the duplicate is reported once as 要確認.
  **The DB design folder for a WG is a TOP-LEVEL project folder named
  `<WG番号>_<WG名>WG\07_データベース・ファイル設計書(仮)` (e.g. `11_工程管理WG\07_データベース・
  ファイル設計書(仮)`) — a sibling of `01_Doc`, NOT nested inside it**, even though most other
  design-doc types live under `01_Doc\...`. A real review once searched under the wrong,
  non-existent path `01_Doc\07_データベース・ファイル設計書(仮)` and wrongly reported 5 real
  tables (TSJAM036, TSJAM037, TSJAM081, TSJAM999, VSJCM137) as having no layout file at all — they
  all existed under the correct top-level WG folder (one in a `PH3` subfolder). Search recursively
  under the correct top-level folder (including any `PH2`/`PH3` subfolders, preferring PH3 > PH2 >
  top-level if a table appears in more than one) before ever reporting a table's layout as missing.
  **基準情報(JA)-prefix common/shared tables have instead been found under
  `01_Doc\07_データベース・ファイル設計書\`** (a different location than the WG-top-level pattern)
  — check there too before concluding a table's layout is genuinely absent. If it's genuinely not
  found anywhere, flag it (note: a table living in a *different* WG's folder is not itself a
  defect, just note it so the user can confirm ownership).

## Reporting

The output goes to the designer who owns the doc, not into your own working notes — report only
what they'd actually need to act on.

For each real finding: state the defect in one or two sentences, cite the exact sheet name and cell
coordinates (and the master file it was checked against, e.g. 06-06_DB一覧, a specific ﾃｰﾌﾞﾙﾚｲｱｳﾄ
workbook path) so it's actionable, and give a concrete suggested fix when one is obvious (e.g. "add a
row for table Z with an R flag", "add the missing R flag to TXJAM026's existing row"). Group findings
under a heading only when several findings share one; don't force every finding into a rigid
structure if it doesn't fit cleanly.

Order by impact, most important first — a genuinely missing/undeclared table or a flag that doesn't
match actual usage belongs well before a minor naming nit. When a finding depends on a judgment call
rather than a clear rule (a cross-WG reference you couldn't fully verify, a delegation path you
couldn't confirm the program actually takes, a pattern that could be deliberate rather than an
omission), say so explicitly and mark it as needing the designer's confirmation — don't present it
with the same confidence as a confirmed defect.

Omit entirely, unless the user explicitly asked for a full pass-by-pass audit:
- "確認済み・問題なし" lines for tables that checked out clean
- tallies of how many tables were checked
- narration of your own process (which files you dumped, how you sampled, which scratch path you
  used)

The one exception worth keeping brief: if a whole section couldn't be verified at all (a DB design
folder was missing, a table's layout file couldn't be found anywhere), a one-line note is fine —
that's a real gap in review coverage the designer should know about, not a process detail.
