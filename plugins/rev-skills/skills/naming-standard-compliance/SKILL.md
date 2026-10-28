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
  機能ID・システムID etc. must be correct and consistent across every sheet of one
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
   `rev-program-review` handed you; standalone, run `_shared/scripts/live_dump.py <workbook> <out> --hidden` on each
   workbook (step 5 must see hidden sheets to judge them absent).
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
   values (システムID, 計画書No, 機能ID, ﾌﾟﾛｸﾞﾗﾑID and its name) are identical across every sheet's
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
   contradicts the others in the same file. Exceptions: **表紙 has no header block** (its IDs sit in the
   body); a template with a single header copy (e.g. `ｽﾍﾟｰｼﾝｸﾞﾁｬｰﾄ`) skips the left/right comparison below.
   **Not reviewed (designers' answer 2026-10-02 No.48/50/73-76/102, adopted 2026-10-19):** 表紙 Ⅲ．改訂履歴, the revision memos right of the form (wording and dates), and every sheet header's 作成日/更新日 are left to the designer's own check — do not report their values, their left/right copies, their order, or rows left incomplete. (A struck row with no revision memo at all is a different matter, not covered by this skill yet.) **Deliberately skip `詳細設計書` sheets
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
   cross-sheet check above (the IDs and names only — 作成日/更新日 are not reviewed, see above). Extract both copies from
   every sheet's header rows (1-4; left labels at columns 5/29/41, right at 57/81/93, each value five
   columns to the right of its label), pair them label by label, and diff them in addition to the
   cross-sheet diff.

   **Report a left/right difference only when the copies contradict** — two different non-`-`
   values for one label. A `-` (or empty) copy beside a filled one is template-level drift, not a
   per-sheet defect (~182 PHASE2 sheet pairs show the `-` pattern — counted while header dates were still
   compared; re-count 作成者/更新者 before quoting a figure), and the orchestrator's Step 3.5 rule 2 withdraws such findings. Summarise it as
   **one 低 line per workbook** (PSJCO403: FSJC019/020 right-copy 作成者 `-`). List the
   contradicting sheets of one workbook in a single finding grouped by pattern, not one per sheet.
   This check owns both the contradictions and the drift line; `design-doc-formatting-consistency`
   reports them only when this check is not in the run. A contradiction in 作成者/更新者 is 低 bookkeeping;
   作成日/更新日 are not compared at all. Contradictions in 機能ID, ﾌﾟﾛｸﾞﾗﾑID or 画面ID keep the severity given
   above.

   **Header IDs against the workbook itself** (my-review-skills 〔PH2-49・82・101・423・467・741・870・953〕, adopted
   2026-10-28; measured over the 121 工程管理 books: 95 cells in 16 books). The cross-sheet diff above cannot see a
   value copied to *every* sheet, so also compare each sheet's ﾌﾟﾛｸﾞﾗﾑID (`[3,10]`/`[3,62]`; row 4 on
   機能定義書/更新条件表) with the PID in the workbook's file name, and the three-digit number in a
   画面設計書/ﾁｪｯｸ処理設計書 sheet name's `(G…)` with the PID's. One finding per workbook grouped by pattern; when the cross-sheet diff above fires on the same cells (half the sheets
   copied, half correct — `PSJCO404`), this paragraph's line absorbs it, not two findings:
   - **中 — another program's ID**: `PSJCO404` carries `PSJCO403` (and its name `社内加工ﾃﾞｰﾀ作成(めっきIS以外)`) in `[3,10]`/`[3,62]` of 12 of its 24
     sheets (the `NEW_…` sheets, 画面/ﾁｪｯｸ `(GSJC403C/D)`, `ﾌｧｲﾙ入出力仕様書(FSJC018)`) and names six sheets `GSJC403A-D`
     while its `更新条件表(TSJCD400)` `[9,16]` says `①画面(GSJC404D)` — the number says which side is wrong; the sheet's
     letter (`(GSJC403C)` holding the D screen) is `design-doc-internal-consistency`'s; `PXJCO129` `更新条件表(TSJCA006)` `[4,62]` `PXJCB101`; `PSJCO604` `ﾁｪｯｸ処理設計書(GSJC604D)`
     `[4,10]`/`[4,62]` 画面ID `GXJA253B` (copied from PXJAO253); a typo of the own ID (`PSJC207` on `PSJCO207`'s
     画面/ﾁｪｯｸ `[3,10]`/`[3,62]`); `#REF!` (`PSJCO606` `ﾁｪｯｸ処理設計書(GSJC606A)` `[3,10]`/`[3,62]`/`[4,10]`/`[4,62]`, the
     names `[3,15]`/`[4,15]` too).
   - **低 — the wrong kind of ID under the label**: a table ID under ﾌﾟﾛｸﾞﾗﾑID in both copies (`PSJAO704` twelve 更新条件表
     `[4,10]`/`[4,62]` `TSJAM721`…), a screen ID (`PSJCO206` `[4,62]` `GSJC206A`), a 機能ID where the sheet's own label says
     ﾌﾟﾛｸﾞﾗﾑID (`SSJCB605` `[3,10]`/`[3,62]` `SJC060`, `PSJCO803` `FSJC080` `SJC080`, `PSJCO403` `FSJC018-020` `[3,62]` `SJC040`).
     When only the **right copy's labels** have drifted (`PSJCO602` six 更新条件表: `[3,5]`=機能ID/`[3,10]`=`SJC060` is right,
     `[3,57]`=ﾌﾟﾛｸﾞﾗﾑID/`[4,57]`=画面ID are the wrong labels), that is the left/right label mismatch above, not a wrong ID.
   - **低 — a 更新条件表 更新概要 `①画面(<ID>)` naming a screen or file the workbook has no sheet for** (a copy-source
     leftover; 8 cells). Not this bullet: an ID whose number *is* the PID's (`GSJC404D` above — evidence for the 中 line,
     the sheet name being wrong), and an ID of the wrong shape (`GSJCO404B`, 9 characters, `PSJCO404` `更新条件表(TSJCD401)`
     `[41,16]`/`[44,16]` — step 3's malformed-ID finding). Examples: `PXJCO125` `TSJCD101`/`TSJCD014`/`TXJAM068` `[9,16]` (`[71,16]`) `GSJC201B`, `PSJAO401` `TSJAA025`
     `[52,16]` `GSJA241B,GSJA241F`, `PSJCO403` `TSJCD401` `[9,16]` `FSJC025`. `update-condition-completeness` C9 checks
     the 更新概要's shape, not this.
   - **Not findings:** a ﾌｧｲﾙ出力仕様書's ﾌﾟﾛｸﾞﾗﾑID that is the sub-system's ID rather than the program's (my-review-skills
     excludes it too); a PID missing from the 機能一覧 registry is step 6's 未登録 line, not this paragraph's; other programs'
     G/F IDs in the body (679 cells in 71 books) — 遷移/起動/参照 (322), ﾜｰｸﾌﾛｰ screen IDs set as values in 更新条件表
     (`"GSJA704B"`, ~60), `①GSJA702A_ﾀﾞｳﾝﾛｰﾄﾞ` (15 books), TXJAM100 keys (0 foreign); `PSJAO502`'s `GSJA501DE` (a
     sheet serving two screens — 要確認 at most).
   - **ｼｽﾃﾑ名**: the 表紙 `[6,6]` and every header ｼｽﾃﾑ名 (`[2,34]`/`[2,86]`) read `【UNICORN】川内生産管理ｼｽﾃﾑ再開発（工程管理）`
     in 121 of 121 books (2,528 header cells; call-out 表紙 列47/行1 「再開発」). Any other value is one 低 line.

5. **表紙's "Ⅱ．設計書構成" list vs the sheets actually present** (checklist item near 2-1) — run
   this as a **standard, always-on check**, not an optional one: compare every row's **Customer** marker (`●`/`-`) against whether a matching sheet genuinely
   exists in the workbook *and has real content* (not just an empty template), and every row's
   **Developer** marker against whether those sheets carry **developer-side content** (designers'
   answer 2026-10-10 No.92, adopted 2026-10-19):
   - **Developer-side area** = columns BA-CZ (53-104), rows 5 down, on each sheet whose `BA1` holds
     the `nD` marker (`2D` on 機能定義書, `3D` elsewhere). Rows 1-4 there are the developer copy of the
     header and never count; column DA (105) onward is the revision-memo margin, not this area.
   - **Not developer content:** the 画面項目ID column (BB, col 54, on 画面設計書 Ⅳ/Ⅴ, ﾌｧｲﾙ出力/入力仕様書
     and 帳票設計書 Ⅱ — its header, `-`, IDs, placeholders `XXXXXXX`/`SJC9xxxx`, and cells holding only
     labels (half- or full-width colon) and IDs separated by line breaks or `/`, such as `処理前：⏎XJZ7006⏎処理後：⏎XJC7009`); the column headers the template pre-prints (ﾁｪｯｸ処理設計書 row 8
     `ｴﾗｰｺｰﾄﾞ`/`ﾁｪｯｸ詳細`/`ﾁｪｯｸ補足`); a copied section title (`Ⅰ．画面ﾚｲｱｳﾄ`); dated change notes
     (`2025/2/3 藤崎 開発指摘対応：…`).
   - **Developer content:** anything else live there — 画面設計書 Ⅴ physical names in BF/BG
     (`sagyobaCd`, `gxjco164aList1G1`; the developer's area, answer 2026-10-02 No.6・7), 更新条件表
     取得ﾃｰﾌﾞﾙ名/取得条件 blocks (they always sit in BA-CZ), ﾁｪｯｸ処理設計書 ｴﾗｰｺｰﾄﾞ/ﾁｪｯｸ詳細/ﾁｪｯｸ補足 values,
     帳票設計書 横幅超過/水平揃え/垂直揃え values, SQL, notes.
   - **Expected** (per doc type — `●` when **any** sheet of that type has content: `PSJCO203` has 10 of 17
     更新条件表 with a 取得 block): content → `●`; no content, including a sheet whose BA-CZ holds only 画面項目IDs →
     `-`. Report `●` without content as 低 「Developer を「-」に」 (更新条件表 with no 取得 block,
     機能定義書, ID-only 画面設計書/ﾌｧｲﾙ出力仕様書, `PSJCO403` ﾁｪｯｸ処理設計書 with only the row-8 headers), and
     `-` with content as 低 「Developer を「●」に」, citing up to three cells (`PXJCO131`
     `画面設計書(GXJC131A)` `[1636,58]` `sagyobaCd`; `PSJCO203` `更新条件表(TXJCA003)` `[8,54]` 取得ﾃｰﾌﾞﾙ名).
     When the only content is one to three stray memo cells (`固定で大丈夫`, `ﾎﾞﾀﾝは9000番台`, an image
     caption), report 要確認: move the memo out of BA-CZ or set `●`.
   - **Not a finding:** Developer `-` on a sheet whose BA-CZ holds only 画面項目IDs (ﾌｧｲﾙ出力仕様書 BB,
     画面設計書 Ⅳ/Ⅴ BB) — No.92.
   - Measured 2026-10-19 over PHASE1-3 (121 表紙): 349 rows agree, 59 `●` without content (38 of them
     ID-only), 54 `-` with content (11 of them memo-only). 詳細設計, 画面遷移図 and ｽﾍﾟｰｼﾝｸﾞﾁｬｰﾄ follow the
     convention bullets below, not this test.
   **A hidden sheet does not count as present**
   (the dump now includes hidden sheets — `_HIDDEN.txt` lists them): a doc type whose only sheet is
   hidden is judged as absent; the 【非表示】 line for that sheet belongs to its owning check, not here
   (`agent-guide.md`, "Hidden sheets, rows and columns"). This has caught a real defect in every
   program reviewed with it so far that had extra doc types beyond the bare minimum: on
   `PXJCB134_流動停止(ﾊﾞｯﾁ).xlsx`, 表紙 marked both `ﾁｪｯｸ処理設計書`(No.4) and `詳細設計`(No.9) as
   `-`/`-` (not applicable) despite both `ﾁｪｯｸ処理設計書(PXJCB134)` (populated with real check
   items) and two detail-design sheets existing — and 表紙's own 改訂履歴 No.4 even said "詳細設計書
   を作成" (created the detail-design doc), directly contradicting the `-` marker (改訂履歴 used here only as
   evidence — it is never itself reviewed). Given this
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
   - **表紙 Ⅲ．改訂履歴 is not reviewed at all** (no incomplete rows, no numbering, no 新規作成 entry) — see the
     rule in step 4.

   Three Customer/Developer markers are convention too (61 PHASE3 workbooks, 2026-10-06; ｽﾍﾟｰｼﾝｸﾞﾁｬｰﾄ over
   PHASE1-3, 2026-10-19):
   - **`画面遷移図` `-`/`-`** in 61 of 61, and none has a 画面遷移図 sheet (`ﾀﾌﾞ遷移説明`,
     `《参考》画面項目遷移` are reference sheets, not one). Not a finding.
   - **`詳細設計` Developer `●` with a header-only `詳細設計書(<ID>)` sheet** — 56 of 61 mark Developer
     `●`; in 25 of them the sheet is empty below row 4 (PSJCO403 among them) and most others hold only an
     API list or `drawio参照`. Read this `●` as "the doc type applies", not "already written" — "has real
     content" does not apply to this row. The real defect stays the reverse one above: a populated sheet marked `-`.
     The BA-CZ test above does not apply: 詳細設計書 is `3D` on both halves (116 of 124 sheets), i.e.
     wholly the developer's document. PHASE1-3: 115 `●` (24 header-only); the two populated sheets
     marked `-`/`-` are `PXJCO130` and `SXJCB178`.
   - **`ｽﾍﾟｰｼﾝｸﾞﾁｬｰﾄ` Developer `-`** — 121 of 121 (PHASE1-3, 2026-10-19). The sheet is one wide
     customer page (`3C` only, 計画書No at col 82/108) with no `nD` developer half, so columns 53-104
     are the chart itself; never apply the BA-CZ test to it.

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
   XJC/SJC screens) is 要確認 at most.

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
