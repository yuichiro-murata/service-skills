---
name: naming-standard-compliance
description: Check that program/screen/table/file/report/zoom/message IDs and design-doc headers in this project follow the official ID-numbering rules and doc-writing checklist (05.システム共通設計書「各ID採番」, 05.設計書記述ルール_チェックリスト). Use when the user asks to check ID命名規則/採番ルール compliance, or wants a broad "is this design doc written correctly" pass distinct from cross-reference or DB-column checks. When the user asks to REV a single design-doc workbook without naming which checks they want, the entry point is `rev-program-review`: it first asks the user, checkbox-style, which of the 9 single-program checks to run, then runs only those as one combined pass. Do not launch all nine yourself. Run this skill standalone only when it was one of the selected checks, or when the user asked for this check by name.
---

# naming-standard-compliance

Checks that IDs used in design docs (and the docs' own header/structure) follow this project's
documented rules — see the frontmatter `description` above for which sibling skill covers
cross-references, I/O table completeness, DB-column existence, or typo proofreading instead.

**Scope selection comes first.** When the user asks to REV a single program's design-doc workbook
without naming specific checks, `rev-program-review` is the entry point: it presents the 9
single-program checks as a checkbox list (`AskUserQuestion`, multiSelect), then runs only the
selected ones as one combined pass, dumping the workbook once up front and sharing the text with
every check (see `_shared/xlsx-excel-com-dump.md`'s "dump once, share the text" section). Do **not**
unconditionally launch all 9 yourself, and do not post a per-check status update — the combined
report is posted once, after every selected check has finished. This skill runs on its own when it
was one of the selected checks, or when the user asked for this check by name.

Grounded in:

- `01_Doc/04_共通設計/05.ｼｽﾃﾑ共通設計書.xlsx`, sheet **"各ID採番"** — the numbering rule for each ID
  type (read it fresh each run; do not rely on a paraphrase from memory, since it's the single
  source of truth and this skill's job is to catch drift from it).
- `01_Doc/99.共通資料/設計書記述ルール/05.設計書記述ルール_チェックリスト.xlsx`, sheet
  **"ﾁｪｯｸﾘｽﾄ"** — items 1-x/2-x (general/表紙 formatting rules) plus the header-info check (1-6:
  機能ID・システムID・作成日 etc. must be correct and consistent across every sheet of one
  workbook).

## Environment

**Where the shared docs live:** the `_shared/*.md` files ship **inside this plugin**, in the
`_shared/` folder next to this skill's own directory (`<plugin root>/skills/_shared/`) — **not** in
`~/.claude/skills/_shared/`, which does not exist on a normal install. Resolve every `_shared/...`
reference below against that folder; if it doesn't resolve, glob
`**/rev-skills/**/skills/_shared/<filename>` and read the hit. Do not skip it and improvise the
Excel COM dump instead: that doc carries the guard against attaching to — and then `Quit()`-ing —
the user's own live Excel session, the strikethrough-exclusion scan, and the reference-file cache.

Read `_shared/xlsx-excel-com-dump.md` first for how to dump `.xlsx` sheets to text
via PowerShell + Excel COM (the validated reader for these workbooks — Python/openpyxl is
installed but has never been validated against them, see that doc's opening note) — including its "Excluding
struck-through / grayed-out rows from review" section. **Deleted IDs are already gone from the dump,
so do NOT run a formatting scan of your own** — every ID string you can see is a live one to
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
and so on; the table is in `_shared/reference-index.md`), which is
`design-doc-internal-consistency`'s job, not a fixed regex to validate here.

## Procedure

1. Given a target file, WG folder, or the whole project, enumerate design-doc workbooks in scope
   (`表紙`/`機能定義書`/`画面設計書`/`ﾃｰﾌﾞﾙﾚｲｱｳﾄ` sheets carry the IDs — dump those sheets per the
   shared COM script).
2. Extract every ID of each type from its home location:
   - プログラムID/画面ID from the 表紙 and 機能定義書/画面設計書 header rows (row 3-4 area, labeled
     `ﾌﾟﾛｸﾞﾗﾑID`/`画面ID`).
   - テーブルID from ﾃｰﾌﾞﾙﾚｲｱｳﾄ row 6.
   - ファイルID/帳票ID/ズームID from wherever the doc set defines them (ﾌｧｲﾙ出力仕様書/帳票設計書/
     ズーム設計書 headers, or references to them inside 画面設計書 event descriptions).
3. Parse each extracted ID against its rule's regex shape above. Flag:
   - IDs that don't parse at all (wrong length, unexpected letter in a fixed-value slot).
   - IDs where the JOBコード segment doesn't match the JOBコード actually used elsewhere in the same
     doc (e.g. プログラムID says `SJC` but 画面ID says `SJA` for the supposedly-paired screen —
     these two IDs are defined by the rule to share the same JOBコード and sequence).
   - Object-type-letter mismatches for テーブルID (e.g. ID starts with `V` but the workbook's own
     ﾃｰﾌﾞﾙ名/構造 looks like a physical table, not a view, or vice versa).
4. Header-info check (checklist 1-6): within one workbook, confirm 機能ID/画面ID/システムID/
   計画書No/作成日 are identical across every sheet's header block (表紙, 機能定義書, 画面設計書,
   ﾁｪｯｸ処理設計書, 更新条件表 all repeat this header — they should never disagree). Flag any sheet
   whose header contradicts the others in the same file. **Deliberately skip `詳細設計書` sheets
   here** (and don't dump them at all) — across every program reviewed with this skill's sibling
   `design-doc-internal-consistency` so far, 詳細設計書 has turned out to be a near-empty
   header-only sheet, so checking its header costs a full-sheet dump for essentially no chance of
   catching a real mismatch. This is a deliberate, documented coverage trade for token savings, not
   a silent gap — if the user specifically asks about 詳細設計書 header consistency, check it then.

   **Also check WITHIN each individual sheet, not just across sheets**: every header block in this
   project's template is physically duplicated side-by-side on one row (e.g. row 1's `2C` label
   with the header starting around column 5, and a second `2D`-labeled copy of the same header
   starting around column 57 — the same pairing exists for `3C`/`3D` on 画面設計書/ﾁｪｯｸ処理設計書/
   更新条件表/etc.). These two copies are meant to be identical but can drift independently of the
   cross-sheet check above. This happened for real on `PXJCB102_製造ｵｰﾀﾞｰ完了処理.xlsx`: both
   `機能定義書(PXJCB102)` and `更新条件表(TXJCM003)` showed 作成日=`46197` in the left (`2C`/`3C`)
   copy but 作成日=`-` (blank) in the right (`2D`/`3D`) copy of the very same header row — while a
   third sheet in the same workbook, `ﾁｪｯｸ処理設計書(PXJCB102)`, correctly showed `46197` on both
   sides. A cross-sheet-only check would have missed this, since it only compares one canonical
   value per sheet and never looks at whether a sheet agrees with itself. Extract both copies from
   every sheet's header row and diff them against each other in addition to the cross-sheet diff.

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

   When a 表紙 observation looks like an omission, measure it across the sibling workbooks in the
   same PHASE folder before reporting it. Both false positives above would have been caught by one
   openpyxl pass over the folder, and the same reasoning applies to any other "this cell is blank"
   observation on a shared template.

6. 表紙 checks that remain lighter-weight/manual-leaning (do these only if asked for a full pass,
   they require reading prose or a slower cross-file lookup): 表紙's 画面ID/画面名 matches the
   program's WG-specific 機能一覧 workbook under
   `01_Doc/06_システム設計書（一覧、管理台帳）/06-01_機能一覧_<WG名>.xlsx` (e.g.
   `06-01_機能一覧_工程管理.xlsx` for a 工程管理 program, `06-01_機能一覧_品質管理.xlsx` for
   品質管理, etc.) — **not** `06-01_機能一覧_共通.xlsx`. That "共通" workbook only lists the
   cross-cutting `XJZ`-prefixed common functions (confirmed by dumping it: its one data sheet is
   named `共通`, ~52 rows, all `機能分類=共通(XJZ)`) and will not contain a program from any other
   WG at all — a program ID/機能ID lookup there for e.g. a `SJC`-prefixed 工程管理 program simply
   comes back "not found", which is a wrong-file miss, not a real "画面ID missing from 機能一覧"
   finding. Pick the 機能一覧 file whose WG suffix matches the target program's JOBコード domain
   (`XJC`/`SJC` → 工程管理, `XJB`/`SJB` → 受注出荷, etc.), and only fall back to `_共通.xlsx` when
   the program itself is genuinely an `XJZ` common one. In the WG-specific sheet (sheet name matches
   the WG, e.g. `工程管理`), the row for a given `PGMID` (around col 34) also carries the screen's
   画面/帳票/ﾌｧｲﾙID (col 49, header `画面/帳票/ﾌｧｲﾙID`) and a 処理内容 label (col 53) whose first
   line is the screen name — that's the pair to diff against the workbook's own 画面ID/画面名, since
   this sheet has no separate dedicated 画面ID column of its own. Also (still in this lighter
   category): new files have a 改訂履歴 entry marked 新規作成 (and 流用新規 files cite their 流用元
   計画書No + 改訂履歴 No).

7. **記述ルール from the self-check workbook** (`01_Doc\08_機能定義書\11_工程管理\XXXXX000_ｾﾙﾌﾁｪｯｸ用.xlsx`,
   call-outs on its sample sheets). Always-on, like steps 4-5. Calibrated on 2026-10-01 over 120
   工程管理 `PHASE1`-`PHASE3` workbooks; counts and examples are one snapshot — re-derive every verdict.

   **7a — 画面表示仕様の各ﾌﾞﾛｯｸ末行に取得件数.** Call-out on 画面設計書 Ⅲ: "取得件数を各末行に記載する".
   Every block of 画面設計書 Ⅲ．画面表示仕様 that starts with `参照ｴﾝﾃｨﾃｨ` **and has a `取得項目`
   table** must close with a `取得件数` row (`[27,4]=取得件数 [27,9]=1件 [27,15]=取得できない場合 [27,24]=-`).
   Don't fix the column: a nested child block starts further right and so does its 取得件数
   (`PSJCO503` `GSJC503A` `[156,6]` … `[172,6]`). Scope is 画面設計書 only — a ﾁｪｯｸ処理設計書 Ⅱ block with
   参照ｴﾝﾃｨﾃｨ is not covered by this call-out.
   - Exempt a block with no 取得項目 — a sort-only sub-block under a 集計単位 variant (`PXJCO164`
     `GXJC164A` `①集計単位="管理No"` … 参照ｴﾝﾃｨﾃｨ/集計条件/ｿｰﾄ順 only) or a `UNION ALL` of two earlier
     blocks (`PSJCO505` `GSJC505A` `[1044]`). 14 such blocks in the corpus.
   - Measured: 1,007 of 1,058 blocks comply. **Missing — 低** (7: `PSJAO704` `GSJA704A` `[213]`,
     `PSJCO503` `GSJC503A` `[432]`, `PXJCO128` `GXJC128B` `[124]` …).
   - **Legacy shape** — the count written as `想定件数 0件以上` on the 取得項目 header row instead of a
     closing 取得件数 row: 44 blocks in 5 sheets (`PSJCO602` `GSJC602A` alone has 16). Report once per
     sheet as 低 ("旧書式。末行に取得件数を記載"). A block with **neither** shape on a legacy sheet
     (`GSJC602A` `[264,4]`, `[462,4]`) is the plain "missing" finding, listed separately.
   - A 取得件数 row present but with **no 件数** (`取得できない場合 -` only; 43 rows, against 551 `1件` and
     409 `複数件`) — 低: the reader can't tell whether the fetch is single- or multi-row (レビュー観点 No.25).

   **7b — ﾚｽﾎﾟﾝｽ定義の表記.** Call-out on 機能定義書: "「ｼｽﾃﾑ共通設計書.ﾚｽﾎﾟﾝｽ 参照」に統一". Read the body
   of 機能定義書 `Ⅶ．ﾚｽﾎﾟﾝｽ定義`. Measured over 109 sheets: 97 read exactly `※ｼｽﾃﾑ共通設計書.ﾚｽﾎﾟﾝｽ 参照`.
   - `※呼び出し元で記載` / `※呼出元で記載` (9) — every one is a **サブプロ/バッチ** (`S…B…`), whose response is
     the caller's; correct, don't flag.
   - `※06-09_ﾚｽﾎﾟﾝｽ一覧_工程管理を参照` (`PSJCO602`, `PSJCO604`) — 要確認: either the program has its own
     responses (then `design-doc-internal-consistency` check 2 verifies 06-09) or it should use the
     standard wording.
   - Any other wording of the standard reference (missing `※`, `ｼｽﾃﾑ共設計書`, extra text) — 低, give the
     exact standard string.

   **7c — 画面名の後ろに画面ID.** Call-out on 機能定義書: "画面名を記載する場合は後ろに画面IDを記載すること".
   In 機能定義書 `Ⅳ．機能処理概要` and 画面設計書 `Ⅳ．画面項目ｲﾍﾞﾝﾄ詳細`, a known screen name followed by
   `画面` (`取引先詳細画面`) in a sentence that transitions to, opens, returns to or shows it
   (`遷移`/`起動`/`戻る`/`表示する`/`開く`/`呼出`) must be followed by `(<画面ID>)`
   (`取引先詳細画面(GXJA802B)に遷移する`). Build the screen-name list from every 画面設計書's 3C header
   `[4,15]` (screen name) / `[4,10]` (画面ID) in the folder, not just this workbook — the name you meet is
   often another program's screen. Normalise the header name first: some end in `画面` (`受入入力画面`),
   most don't, so match `name` + optional `画面`. The program's **own** screens count too
   (`PSJCO602` `[52,4]` `棚卸用仕掛在庫ﾃﾞｰﾀ作成画面を表示する` → `(GSJC602A)`).
   **The screen must be the object of the verb** — `…画面に遷移する` / `…画面に戻る` / `…画面を起動する` /
   `…画面を表示する` / `…画面を開く`. `…画面に表示する` / `…画面で選択された…` put data *on* a screen and
   are not this rule (`PSJCO602` `[145,4]` `…集計し、棚卸用仕掛在庫ﾃﾞｰﾀ作成画面に表示する。`). Don't scan
   other sections: overview prose (Ⅰ．機能概要) and generic words (`画面ｸﾘｱ`, `呼出元画面`) are not this rule.
   `ﾀﾞｳﾝﾛｰﾄﾞ画面` is excluded although `GSJA702A`'s header name is `ﾀﾞｳﾝﾛｰﾄﾞ`: the common download screen
   is cited by program ID in the house wording (`【共通】ﾀﾞｳﾝﾛｰﾄﾞ(PSJAO702)を起動する`), and the bare
   `ﾀﾞｳﾝﾛｰﾄﾞ画面の戻り値…` phrases are fixed template sentences. Several names map to more than one ID
   (`社内加工単価登録` GSJC401A/GSJC402A, `工順比較` GSJA401G/GXJC124B, `工程詳細` GSJA401C/GXJA243B) — give
   the candidates, don't guess.
   Measured in that scope: 405 mentions carry the ID, **310 don't, across 110 sheets** — a widespread
   omission, so report it **once per sheet, 低**, listing the cells and the ID to append
   (`PSJAO704` `機能定義書` `[70,4]` `承認ｸﾞﾙｰﾌﾟ一覧画面に遷移する` → `承認ｸﾞﾙｰﾌﾟ一覧画面(GSJA704G)`). When the
   name matches more than one screen ID, give the candidates rather than guessing.

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
