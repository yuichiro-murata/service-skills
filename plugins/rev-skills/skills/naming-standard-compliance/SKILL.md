---
name: naming-standard-compliance
description: Check that program/screen/table/file/report/zoom/message IDs and design-doc headers in this project follow the official ID-numbering rules and doc-writing checklist (05.システム共通設計書「各ID採番」, 05.設計書記述ルール_チェックリスト). Use when the user asks to check ID命名規則/採番ルール compliance, or wants a broad "is this design doc written correctly" pass distinct from cross-reference or DB-column checks. For a full REV without named checks the entry point is `rev-program-review`; run this standalone only when selected there or asked for by name.
---

# naming-standard-compliance

Checks that IDs used in design docs (and the docs' own header/structure) follow this project's
documented rules — see the frontmatter `description` above for which sibling skill covers
cross-references, I/O table completeness, DB-column existence, or typo proofreading instead.

Grounded in:

- `01_Doc/04_共通設計/05.ｼｽﾃﾑ共通設計書.xlsx`, sheet **"各ID採番"** — the numbering rule for each ID
  type (read it fresh each run; do not rely on a paraphrase from memory, since it's the single
  source of truth and this skill's job is to catch drift from it).
- `01_Doc/99.共通資料/設計書記述ルール/05.設計書記述ルール_チェックリスト.xlsx`, sheet
  **"ﾁｪｯｸﾘｽﾄ"** — items 1-x/2-x (general/表紙 formatting rules) plus the header-info check (1-6:
  機能ID・システムID・作成日 etc. must be correct and consistent across every sheet of one
  workbook).

## Environment

**Read `_shared/agent-guide.md` first** — scope selection, environment, the live dump, reporting
conventions. Read `05.ｼｽﾃﾑ共通設計書` and the `06-01_機能一覧_*` files **live** with
`_shared/scripts/live_dump.py <master> <out_dir> --sheets <regex>` (e.g. `--sheets 各ID採番`), one
`out_dir` per master — each run rewrites `_DELETED_DIGEST.txt`. Deleted IDs are already gone from the dump — every ID string you can see is a live one to
validate against the numbering rule.

Two consequences worth holding onto. First, a partially-struck cell arrives as its live remainder,
so validate exactly the string you see and never reconstruct the raw one from
`_DELETED_DIGEST.txt` — a raw-text check once reported `"GSXJC205A"` as a malformed screen ID when
only the `X` was struck and the live text `GSJC205A` was correct. Second, an ID that appears only in
the digest is retired, not a numbering violation: report neither its format nor its absence.

## ID rules (from "各ID採番" — re-verify against the live sheet, this is a summary)

| ID種別 | 構成 | 例 |
|---|---|---|
| プログラムID | `[P=プログラム\|S=サブプログラム]` + `<3桁JOBコード>` + `[O=オンライン\|B=バッチ]` + `<3桁連番(0埋め)>` | `PSJCO315`, `PSJCB315` |
| 画面ID | `G` + `<3桁JOBコード>`(プログラムIDと同一) + `<3桁連番>`(プログラムIDと同一) + `<プログラム内連番:A,B,C...>` | `GSJC315A` |
| テーブルID | `[T=テーブル\|V=ビュー]` + `<3桁JOBコード>` + `[M=マスタ\|D=データ\|V=トラン\|A=累積\|P=パラメータ]` + `<3桁連番(0埋め)>` | `TXJAM025`, `TSJCV501`, `VSJAM060` |
| ファイルID | `F` + `<3桁JOBコード>` + `<3桁連番(0埋め)>` | `FSJA004` |
| 帳票ID | `R` + `<3桁JOBコード>` + `<3桁連番(0埋め)>` | — |
| ズームID | `Z` + `<基準情報の固定JOBコード>` + `<3桁連番(0埋め)>` | `ZXJA047` / `ZSJA047` — the rule sheet's literal text only shows `XJA`, but in practice **both `XJA` and `SJA` are valid**, following the same 共通(X)/個社固有(S) pairing used by every other ID type's JOBコード in this project (e.g. `XJC`/`SJC` for プログラムID). Confirmed by the user directly — this is not a rule violation, don't flag `SJA` zoom IDs at all. Only flag a zoom ID whose basic-info segment is neither `XJA` nor `SJA`. |
| メッセージID | `<3桁主管理JOBコード>` + `-` + `[0=バインド変数無し\|1=バインド変数有り]` + `<5桁連番(0埋め)>` | — |
| 区分コード | `<3桁主管理JOBコード>` + `<5桁連番(0埋め)>` | — |
| 共通コード | `<3桁JOBコード>` + `<キー1:2桁>` + `<キー2:90桁>` + `<キー3:150桁>` + ボディ(30項目×各1000桁) | — |

Note: **画面項目ID** (e.g. `XJC0036`, `SJC0662` seen inside 画面設計書「Ⅴ．画面項目定義」) is a
*different* ID space not defined in "各ID採番" at all — its correctness is "is it registered in the
`82.画面項目辞書_*.xlsx` its own prefix routes to" (`XJC`/`SJC` → `_工程管理`, `XJZ`/`SJZ` → `_共通`,
and so on; the routing table is in `_shared/agent-guide.md`), which is
`design-doc-internal-consistency`'s job, not a fixed regex to validate here.

## Procedure

1. Given a target file, WG folder, or the whole project, enumerate design-doc workbooks in scope
   (`表紙`/`機能定義書`/`画面設計書`/`更新条件表` sheets carry the IDs). Mock-up and reference sheets
   (`実績入力_B画面`…, `(参考)…`, `【JAGUR】…` in PXJCO161) are out of scope for the ID rules: they hold
   sample data and pasted legacy screens, not IDs this program defines. **Strip trailing spaces from a
   sheet name before parsing its `(<ID>)`** (`PXJCO161` `更新条件表(TSJCD330) `, `更新条件表(TXJCM057) `);
   a sheet name that changes when stripped is one 低 line per workbook. Use the dump
   `rev-program-review` handed you; standalone, run `_shared/scripts/live_dump.py` on each workbook.
   The program registry is three sheets in two files (verified for 工程管理): in
   `06-01_機能一覧_工程管理.xlsx`, the WG sheet `工程管理` and `ｻﾌﾞﾌﾟﾛ一覧` (`--sheets "^(工程管理|ｻﾌﾞﾌﾟﾛ一覧)$"`);
   in `06-01_機能一覧_工程管理(付属機能).xlsx`, the sheet `付属機能` (`--sheets "^付属機能$"`), which also
   lists sub-programs (e.g. `SXJCB184`). An ID is unregistered only when all three lack it. The
   neighbouring `ｻﾌﾞﾌﾟﾛ利用機能一覧(工程管理)` is a caller-by-sub-program usage matrix, not the registry.
   Other WGs name their sheets differently (`品質管理`, `基準情報`'s `機能一覧`) — list the sheet names first.
   The orchestrator may also pass `NON-SCREEN-ID` lines from `build_screen_list.py` (`<value>  <sheet>
   <file>`): a 画面設計書 header whose 画面ID cell holds a program ID (`PXJDO601` in 画面設計書(GXJD601A)).
   For a file in scope, report it as a step-4 header-ID error (expected: the sheet name's `(ID)`); ignore
   lines for other files.
2. Extract every ID of each type from its home location:
   - プログラムID/画面ID/ファイルID from the 表紙 and the header rows of each sheet, matched by label
     (`ﾌﾟﾛｸﾞﾗﾑID`/`画面ID`/`ﾌｧｲﾙID`), not by row — see step 4 for which labels each template has.
   - テーブルID, in a program workbook (which has no ﾃｰﾌﾞﾙﾚｲｱｳﾄ sheet): 機能定義書 Ⅲ．入出力定義 column 5
     under the `ID` header (find the row by that label — PXJCO125 has it at `[35,5]`, one example), each `更新条件表(<ID>)` sheet
     name and its 更新ﾃｰﾌﾞﾙ cell `[8,16]` (`TXJCD203:個別不良項目明細` — the ID is the part before `:`),
     and 参照ｴﾝﾃｨﾃｨ cells. Read a ﾃｰﾌﾞﾙﾚｲｱｳﾄ workbook's own header only when that is the review target.
   - ファイルID/帳票ID/ズームID from wherever the doc set defines them (ﾌｧｲﾙ出力仕様書/帳票設計書/
     ズーム設計書 headers, or references to them inside 画面設計書 event descriptions).
3. Parse each extracted ID against its rule's regex shape above. Flag:
   - IDs that don't parse at all (wrong length, unexpected letter in a fixed-value slot).
   - A JOBコード/sequence mismatch **between the プログラムID and its own 画面ID** (プログラムID says `SJC`
     but 画面ID says `SJA` for the paired screen — the rule defines them to share JOBコード and sequence),
     and an `X`/`S` slot mismatch inside that pair (`PXJCO…` with `GSJC…`). Do not compare against the
     JOBコードs of other IDs in the doc: tables of other WGs (`TXJAM…`, `VSJAM…`, `TSJC…`, `TXJZM…`) are
     legitimately referenced and gave ~25 false hits on PXJCO161.
   - An ID of a foreign system (`DXACM001`-`003`, prefix `DXA…`, registered on `06-06_DB一覧_工程管理`)
     does not follow `[T|V]…` by design — 要確認 at most, not a violation.
   - Object-type-letter mismatches for テーブルID (e.g. ID starts with `V` but the workbook's own
     ﾃｰﾌﾞﾙ名/構造 looks like a physical table, not a view, or vice versa).

   **Suffixed テーブルIDs** (`VXJCM004_31`, `TXJCM007_B`) are not violations when the base part parses
   and a layout exists for the full ID. Resolve the layout by `xlsx-db-column-check` step 3 (exact `A6`
   on the `^ﾃｰﾌﾞﾙﾚｲｱｳﾄ$` sheet, PH3 > PH2 > top-level, the flat `01_Doc\07_…` copy at its own level
   only), not a filename glob: `<ID>_*.xlsx` also matches suffixed *other* tables, and a recursive
   search offers the superseded `90_JAGURﾃｰﾌﾞﾙﾚｲｱｳﾄ(<date>時点)` snapshots. Both examples resolve in
   `11_工程管理WG\07_データベース・ファイル設計書(仮)` (`TXJCM007_B` in `PH2`). Flag a suffixed ID only
   when no layout or registry entry exists for it.
4. Header-info check (checklist 1-6): within one workbook, confirm the **workbook-level** header
   values (システムID, 計画書No, 機能ID, ﾌﾟﾛｸﾞﾗﾑID and its name, 作成日) are identical across every sheet's
   header block (機能定義書, 画面設計書, ﾁｪｯｸ処理設計書, 更新条件表, ﾌｧｲﾙ出力/入出力仕様書 all repeat this
   header). **Label sets differ by template**: 機能定義書/更新条件表 carry 機能ID (row 3) and ﾌﾟﾛｸﾞﾗﾑID
   (row 4); 画面設計書/ﾁｪｯｸ処理設計書 carry ﾌﾟﾛｸﾞﾗﾑID (row 3) and 画面ID (row 4), no 機能ID (verified on
   PXJCO125); ﾌｧｲﾙ出力/入出力仕様書 carry ﾌﾟﾛｸﾞﾗﾑID (row 3) and ﾌｧｲﾙID (row 4). Pair cells by label and
   compare only labels both sheets have. **画面ID and ﾌｧｲﾙID are per-sheet, not workbook-wide**: each must
   equal the `(<ID>)` in its own sheet name, and the name beside it (col 15) must be that screen's/file's
   name — so in a multi-screen or multi-file program they *should* differ between sheets. The real defect
   is the opposite: on PSJCO403, 画面設計書(GSJC403B/C/D) and ﾌｧｲﾙ(入)出力仕様書(FSJC018/019/020) all carry
   `GSJC403A` in `[4,10]`/`[4,62]` and `社内加工ﾃﾞｰﾀ作成` in `[4,15]`/`[4,67]` — copies of sheet A never
   updated; check the names as well as the IDs (高; the validation
   run found only one other of PHASE3's 104 画面設計書 sheets with this, in PSJCO309). Report them as one
   finding grouped by pattern. A row-4 **label** that differs between the left and right copies belongs
   in the same finding (FSJC018: `[4,5]`=`画面ID`, `[4,57]`=`ﾌｧｲﾙID`). Flag any sheet whose workbook-level header
   contradicts the others in the same file. Exceptions: **作成日 may legitimately be
   later on a sheet added in a later revision** (check 表紙 Ⅲ．改訂履歴 for that sheet's addition
   before reporting); **表紙 has no header block** (its IDs sit in the body); a template with a
   single header copy (e.g. `ｽﾍﾟｰｼﾝｸﾞﾁｬｰﾄ`) skips the left/right comparison below. An **更新日 older
   than a dated revision note inside the same sheet** (`2026/6/10 福浦 修正` on a sheet whose header
   更新日 is earlier) is a 低 finding (header bookkeeping — 更新日 diverges in 43 of 99 sibling
   workbooks). Ignore notes dated on or before the header 作成日: those are the creation memos, and
   a `-` 更新日 is then correct (11 false positives on PXJCO125, e.g. 更新条件表(TXJCD203), whose notes
   are all 2025/7/2 = 作成日 45840). **Deliberately skip `詳細設計書` sheets
   here** beyond their header rows (`[1-4,*]`) —  across every program reviewed with this skill's sibling
   `design-doc-internal-consistency` so far, 詳細設計書 has turned out to be a near-empty
   header-only sheet, so checking its header costs a full-sheet dump for essentially no chance of
   catching a real mismatch. This is a deliberate, documented coverage trade for token savings, not
   a silent gap — if the user specifically asks about 詳細設計書 header consistency, check it then.
   `live_dump.py` skips these sheets by default, so read rows 1-4 with
   `live_dump.py <workbook> <out_dir> --sheets "^詳細設計" --include-detail` (only the header rows are needed)
   or with openpyxl read-only.

   **機能ID against the registry**: the header 機能ID must equal the registry's 機能ID (col 7, header
   `機能ID` on row 6) on the program's row. A wrong value repeated on every copy is invisible to the
   left/right and cross-sheet diffs: `PSJCO403` reads `SJC030` on every sheet while 付属機能 row 14 has
   `SJC040` (中 — name the registry value).

   **Also check WITHIN each individual sheet, not just across sheets**: every header block in this
   project's template is physically duplicated side-by-side on one row (e.g. row 1's `2C` label
   with the header starting around column 5, and a second `2D`-labeled copy of the same header
   starting around column 57 — the same pairing exists for `3C`/`3D` on 画面設計書/ﾁｪｯｸ処理設計書/
   更新条件表/etc.). These two copies are meant to be identical but can drift independently of the
   cross-sheet check above. Example of the shape (a `-` copy, so it now counts as drift — see the next
   paragraph) from `PXJCB102_製造ｵｰﾀﾞｰ完了処理.xlsx`: both
   `機能定義書(PXJCB102)` and `更新条件表(TXJCM003)` showed 作成日=`46197` in the left (`2C`/`3C`)
   copy but 作成日=`-` (blank) in the right (`2D`/`3D`) copy of the very same header row — while a
   third sheet in the same workbook, `ﾁｪｯｸ処理設計書(PXJCB102)`, correctly showed `46197` on both
   sides. A cross-sheet-only check would have missed this, since it only compares one canonical
   value per sheet and never looks at whether a sheet agrees with itself. Extract both copies from
   every sheet's header rows (1-4; left labels at columns 5/29/41, right at 57/81/93, each value five
   columns to the right of its label), pair them label by label, and diff them in addition to the
   cross-sheet diff.

   **Report a left/right difference only when the copies contradict** — two different non-`-`
   values for one label. A `-` (or empty) copy beside a filled one is template-level drift, not a
   per-sheet defect: on PHASE2, 209 of 424 sheets have differing C/D 更新日 and ~182 sheet pairs show
   the `-` pattern, and the orchestrator's Step 3.5 rule 2 withdraws such findings. Summarise it as
   **one 低 line per workbook** (PSJCO403: FSJC019/020 right-copy 作成日/作成者 `-`). List the
   contradicting sheets of one workbook in a single finding grouped by pattern, not one per sheet.
   This check owns both the contradictions and the drift line; `design-doc-formatting-consistency`
   reports them only when this check is not in the run. A contradiction in **更新日** (or 更新者) is 低
   bookkeeping; say which copy is authoritative by comparing both dates with 表紙 Ⅲ．改訂履歴 and the
   sheet's own dated notes — it is not always the right copy that lags (`PXJCO161` `更新条件表(TXJCD104)`:
   left `[3,46]` 46227, right `[3,98]` 46239 — the left is stale). Contradictions in 機能ID, ﾌﾟﾛｸﾞﾗﾑID or
   画面ID keep the severity given above.

5. **表紙's "Ⅱ．設計書構成" list vs the sheets actually present** (checklist item near 2-1) — run
   this as a **standard, always-on check**, not an optional one: compare every row's Customer/
   Developer marker (`●`/`-`) against whether a matching sheet genuinely exists in the workbook
   *and has real content* (not just an empty template). This has caught a real defect in every
   program reviewed with it so far that had extra doc types beyond the bare minimum: on
   `PXJCB134_流動停止(ﾊﾞｯﾁ).xlsx`, 表紙 marked both `ﾁｪｯｸ処理設計書`(No.4) and `詳細設計`(No.9) as
   `-`/`-` (not applicable) despite both `ﾁｪｯｸ処理設計書(PXJCB134)` (populated with real check
   items) and two detail-design sheets existing — and 表紙's own 改訂履歴 No.4 even said "詳細設計書
   を作成" (created the detail-design doc), directly contradicting the `-` marker. On
   `PXJCB102_製造ｵｰﾀﾞｰ完了処理.xlsx`, the same thing happened for `詳細設計`(No.9) alone. Given this
   hit rate, don't treat this as a "only if asked for a full pass" item — check it every time, the
   same as the header-info check above.

   **Two parts of 表紙 are template boilerplate and are NOT findings — do not report them.** Both
   were raised as findings on the `PXJCO124_ﾛｯﾄ振向け` REV and both were withdrawn after measuring
   the 28 sibling workbooks in `01_Doc/08_機能定義書/11_工程管理/PHASE3`:

   - **The "その他設計書" column of Ⅱ．設計書構成 (column 34, rows 24-33) is empty in every
     workbook** — 0 of 28 had any entry, while the No. column beside it (column 32) is always
     pre-numbered 1-10. It is an unused template column, so "sheet X appears nowhere in the
     構成表" is not a defect just because X is only explainable by that column. The Customer/
     Developer marker check in the paragraph above is the real check here; the その他 column is not.
   - **A trailing 改訂履歴 row carrying only the next sequence number, with 改訂日/改訂者/計画書No/
     改訂内容 all blank, is the project's convention** — 27 of 28 workbooks had one (the lone
     exception was `PSJCO205_返品処置指示発行`). Rows are pre-numbered for the next revision. Only
     report a 改訂履歴 row as incomplete when a row **other than the last** is partially filled.

   Two Customer/Developer markers are convention too (all 61 PHASE3 workbooks, measured 2026-10-06):
   - **`画面遷移図` `-`/`-`** in 61 of 61, and none has a 画面遷移図 sheet (`ﾀﾌﾞ遷移説明`,
     `《参考》画面項目遷移` are reference sheets, not one). Not a finding.
   - **`詳細設計` Developer `●` with a header-only `詳細設計書(<ID>)` sheet** — 56 of 61 mark Developer
     `●`; in 25 of them the sheet is empty below row 4 (PSJCO403 among them) and most others hold only an
     API list or `drawio参照`. Read this `●` as "the doc type applies", not "already written" — "has real
     content" does not apply to this row. The real defect stays the reverse one above: a populated sheet marked `-`.

   When a 表紙 observation looks like an omission, measure it across the sibling workbooks in the
   same PHASE folder before reporting it. Both false positives above would have been caught by one
   openpyxl pass over the folder, and the same reasoning applies to any other "this cell is blank"
   observation on a shared template.

6. 表紙 checks that remain lighter-weight/manual-leaning (do these only if asked for a full pass —
   a REV that runs all checks counts as one — they require reading prose or a slower cross-file lookup): 表紙's 画面ID/画面名 matches the
   program's WG-specific 機能一覧 workbook under
   `01_Doc/06_システム設計書（一覧、管理台帳）/06-01_機能一覧_<WG名>.xlsx` (e.g.
   `06-01_機能一覧_工程管理.xlsx` for a 工程管理 program, `06-01_機能一覧_品質管理.xlsx` for
   品質管理, etc.; 受注出荷 is the exception, below) — **not** `06-01_機能一覧_共通.xlsx`. That "共通" workbook only lists the
   cross-cutting `XJZ`-prefixed common functions (confirmed by dumping it: its one data sheet is
   named `共通`, ~52 rows, all `機能分類=共通(XJZ)`) and will not contain a program from any other
   WG at all — a program ID/機能ID lookup there for e.g. a `SJC`-prefixed 工程管理 program simply
   comes back "not found", which is a wrong-file miss, not a real "画面ID missing from 機能一覧"
   finding. Pick the 機能一覧 file whose WG suffix matches the target program's JOBコード domain
   (`XJC`/`SJC` → 工程管理, `XJB`/`SJB` → 受注出荷, etc.), and only fall back to `_共通.xlsx` when
   the program itself is genuinely an `XJZ` common one. For 工程管理, look in **both registry files from
   step 1**: the WG sheet `工程管理` and `06-01_機能一覧_工程管理(付属機能).xlsx` sheet `付属機能` — the
   付属機能 programs (e.g. the `SJC040` 社内加工 family, PSJCO403) are only in the latter. **受注出荷 is
   named differently**: `06_01_機能一覧_受注出荷.xlsx` (underscore); read the sheet `受注出荷_XJB混在` and
   skip `受注出荷_bkup`, `受注出荷_XJB混在_bkup` and `受注出荷_XJB混在_20240627bkup` (stale copies, all
   visible). In the registry sheet, the row for a given `PGMID` (col 34) also carries the screen's
   画面/帳票/ﾌｧｲﾙID (col 49, header `画面/帳票/ﾌｧｲﾙID`) and a 処理内容 label (col 53) whose first
   line is the screen name — that's the pair to diff against the workbook's own 画面ID/画面名, since
   this sheet has no separate dedicated 画面ID column of its own. 受注出荷 has them at cols 32/47/51 —
   locate all three by the row-6 headers. A program's second and later screens/files sit on
   continuation rows with `PGMID` blank — read the program row and all its continuation rows, down to
   the next row with a PGMID (付属機能: PSJCO403 on row 14, GSJC403B/C and FSJC018-020 on rows 15-19, and
   FSJC091 on row 20; row numbers are one snapshot). Strip the area prefix (`^[A-Z]\.`,
   e.g. `A.ﾛｯﾄ取消`; 受注出荷 uses `【…】`) before comparing; a trailing `画面` on the design-doc side (`ﾛｯﾄ取消画面`, 15 of 159
   XJC/SJC screens) is 要確認 at most. Also (still in this lighter
   category): new files have a 改訂履歴 entry marked 新規作成 (and 流用新規 files cite their 流用元
   計画書No + 改訂履歴 No).

7. **(moved)** The self-check writing rules (取得件数, ﾚｽﾎﾟﾝｽ表記, 画面名＋画面ID) are now
   `design-doc-writing-rules` W3-W5.

## Reporting

The output goes to the designer who owns the doc, not into your own working notes — report only
what they'd actually need to act on.

Group by ID type / checklist item, but only for types that actually produced a finding. For every
flagged ID, show the raw string, which rule it was checked against (quote the rule briefly), the
location it was found, and — when the fix is obvious — what it should be instead. Order by impact:
a JOBコード mismatch between paired IDs or a header-info contradiction (checklist 1-6) belongs before
a cosmetic 表紙 formatting nit. Distinguish clear rule violations from "the rule sheet and reality
disagree, worth asking a human" cases and mark the latter as needing the designer's judgment — don't
assert the design doc is wrong when it's equally possible the common-design rule sheet is the stale
one (though note the ズームID `XJA`/`SJA` case above is now a *resolved* example of this, not an
open one — don't re-flag it).

Omit entirely: IDs that parsed cleanly and matched (don't list "プログラムID: OK"), a tally of how
many IDs of each type were checked, and narration of your own process. The one exception worth a
one-line mention: an ID category you couldn't check at all because its home sheet/section was
missing or empty — that's a coverage gap worth flagging, not a process detail.
