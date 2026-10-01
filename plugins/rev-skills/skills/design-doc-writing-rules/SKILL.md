---
name: design-doc-writing-rules
description: Check a program's design-doc workbook against the self-check workbook's (XXXXX000_ｾﾙﾌﾁｪｯｸ用.xlsx) writing rules — Ⅴ．画面項目定義 per-item rules (upper-case conversion note on code inputs, button IDs in the 9xxx band, 選択 = □ without ID, No = ZZ9/1から連番/XJZ0425, TextBox/TextArea always carry 桁数 and 入力可文字種), 項番 gaps/duplicates (No. columns, ﾁｪｯｸ処理 numbers, circled 機能処理概要 sections), 取得件数 at the end of every 画面表示仕様 block, the standard ﾚｽﾎﾟﾝｽ wording, and screen names followed by their 画面ID. Rules about how a doc is written, not cross-references (design-doc-internal-consistency) or ID numbering (naming-standard-compliance). Use when the user asks for 記述ルール/セルフチェック観点 compliance. For a full REV without named checks the entry point is `rev-program-review`; run this standalone only when selected there or asked for by name.
---

# design-doc-writing-rules

Checks the writing rules that the self-check workbook
`01_Doc\08_機能定義書\11_工程管理\XXXXX000_ｾﾙﾌﾁｪｯｸ用.xlsx` (sheet `レビュー観点` and the call-outs on its
sample sheets) adds on top of the 05 checklist. These rules look at one row or one list at a time —
they ask "is this written the way the project writes it", not "does X here agree with Y there"
(`design-doc-internal-consistency`) and not "does this ID follow 各ID採番" (`naming-standard-compliance`).
They were split out of those two skills so each agent stays focused; the rule text below is moved
verbatim, and every threshold was calibrated on the 工程管理 `PHASE1`-`PHASE3` corpus as stated in each
section.

## Environment

**Read `_shared/agent-guide.md` first** — scope selection, environment, the live dump (W2 needs
`_DELETED_DIGEST.txt`), reporting conventions. W5 uses the folder-wide screen-name list: use the one your prompt names, and build it yourself
when none was handed to you (standalone, or the orchestrator didn't prebuild it).

## Checks

### W1 — Ⅴ．画面項目定義 per-item rules (formerly design-doc-internal-consistency check 8)


These come from `01_Doc\08_機能定義書\11_工程管理\XXXXX000_ｾﾙﾌﾁｪｯｸ用.xlsx` (sheet `レビュー観点`
No.43/44/48/49 and the call-outs on its `画面設計書(GXJA802A)` sample), not from the 05 checklist.
They look at one row at a time, so they are cheap — but every threshold below was
calibrated on 2026-09-30 against 200 `画面設計書(*)` sheets / 10,321 Ⅴ rows in the 工程管理
`PHASE1`-`PHASE3` workbooks, because the literal rule text over-fires on this corpus in each case.
Re-derive every verdict from the dump; the counts and examples are one snapshot.

**Parse Ⅴ by header label, per area.** Ⅴ is split into areas (`共通`, `G1)検索条件領域`, …) whose
titles sit in col 3; each area has its own header row. Locate columns by normalised label
(strip whitespace/newlines — the headers read `入力可⏎文字種` and `説明/⏎表示時参照ﾃｰﾌﾞﾙ`): `画面項目名`,
`表示`, `属性`, `TAB`, `桁数`, `表示形式`, `入力可文字種` (some sheets: `入力可`), `必須`, `初期値`, `説明`,
`画面項目ID` (commonly cols 5/11/17/21/23/25/30/33/35/43/54 — resolve by label, never by these).
Item rows are those whose No. (col 3) is a number. Strip a leading full-width space before
comparing an ID (the self-check sample itself has `　XJA0130`).

**W1a — 小文字→大文字変換の記載 (レビュー観点 No.43).** For an **input** item (`TextBox`, `TextArea`,
`ComboBox` — never `Label`/`ReadOnly`/`Hidden`/`LinkLabel`) whose 画面項目名 is one of
`管理No` / `在庫No` / `作業者ｺｰﾄﾞ` / `作業場ｺｰﾄﾞ` / `品目ｺｰﾄﾞ` / `工程ｺｰﾄﾞ` / `資源ｺｰﾄﾞ` / `ﾃｰﾌﾟﾛｯﾄNo` /
`KC品名` — after stripping a `(FROM)`/`(TO)`, a trailing digit, `(1-N)` — the 説明 column must say the
input is upper-cased. The house wording is `値入力時：小文字ｱﾙﾌｧﾍﾞｯﾄを大文字に変換` (329 of 366
occurrences); `ﾛｽﾄﾌｫｰｶｽ時：…`, a half-width colon or a `・`/`①` prefix are the same rule — match on
`大文字`. Measured: 207 of 227 exact-name TextBoxes comply; the 20 that don't are findings, and half
of them are `KC品名` search fields. A **compound** name ending in one of these
(`送付元作業場ｺｰﾄﾞ`, `開始工程ｺｰﾄﾞ`, `取消作業者ｺｰﾄﾞ`) complies 100 : 19 — report the 19 as 要確認 only,
because a prefix can make it a different entity (`生管理No`, `ｲﾝｸ管理No`, `部材管理No` are not
管理No). An upper-case note on an item *outside* the list (`加工手順ｺｰﾄﾞ`, `製造ﾛｯﾄNo`) is fine.
**Also compare within the workbook:** an input item that carries the note on one screen and not on
another is a gap even when its name is outside the list or drifts — match by 画面項目ID when both
have one, else by name with a trailing `ｺｰﾄﾞ` removed. Confirmed on `PSJCO805`: `出庫先作業場ｺｰﾄﾞ`
has the note on `GSJC805A`/`E`/`F`, while the same input named `出庫先作業場` on `GSJC805B` `[123]`
and `GSJC805C` `[385]` has none — the literal list cannot see it. Report these as 要確認.

**W1b — ﾎﾞﾀﾝの画面項目ID (No.44).** The self-check says "共通のﾎﾞﾀﾝは9xxx番台" and a call-out says
"ﾎﾞﾀﾝは全て共通のXJZ9xxx(SJZ9xxx)を利用/採番". **Check the `9xxx` band, not the `XJZ`/`SJZ` prefix.**
Measured on non-zoom `Button` rows in `共通` areas: 721 use `XJZ9xxx`/`SJZ9xxx`, but 149 use a
WG-prefixed `9xxx` (`SJA9018` 基本情報ﾎﾞﾀﾝ, `SJC9004` ﾜｰｸﾌﾛｰﾎﾞﾀﾝ) — program-specific buttons numbered
in their own WG per the same workbook's other call-out "新規採番はSJC/SJAで採番する". Flagging
those would be ~150 false findings. Flag, for a non-zoom Button in a `共通` area:
- ID blank or `-` (59 blank — e.g. `PSJCO805` `GSJC805A` has no ID on any of its 12 共通 buttons);
- the placeholder `XXXXXXX` (`PSJCO402`, `PSJCO403`);
- a malformed prefix (`PSJAO704` `GSJA704I` 行削除ﾎﾞﾀﾝ `XZJ9031` — `XZJ` for `XJZ`);
- an ID outside `9xxx` (14 — `XJC0323` 削除ﾎﾞﾀﾝ, `SJA0349` 削除ﾎﾞﾀﾝ, `XJZ0079` 返却ﾎﾞﾀﾝ). When a
  standard button (検索/画面ｸﾘｱ/削除/戻る/更新…) has a non-`9xxx` ID, name the `XJZ9xxx` the other
  sheets use for the same button as the fix. A cell holding two IDs (`SJA0873⏎SJA9048`) passes if
  one of them is `9xxx`.
Two checks the band test cannot catch, both worth a line:
- **the same ID on two different buttons of one sheet** — `PSJAO205` `GSJA205D` gives 処理状況ﾎﾞﾀﾝ
  `[86,54]` the ID `XJZ9009`, which is its own 戻るﾎﾞﾀﾝ's `[85,54]`; a duplicate ID means one label
  is wrong at runtime;
- **a standard button whose ID differs from the one the rest of the corpus uses** for that button
  name (処理状況 = `XJZ9017` on 5 other sheets). Build the name → most-used-ID table from the other
  screens you have and report a deviation as 要確認, naming the usual ID.
**When a sheet's whole 画面項目ID column is empty** (`PSJCO805`: all six screens, every row), do
not list every button, No and input separately — report "画面項目ID 未採番" once per sheet, with the
standard IDs you can already suggest (検索 `XJZ9002`, 画面ｸﾘｱ `XJZ9003`, ﾀﾞｳﾝﾛｰﾄﾞ `XJZ9004`, 新規登録
`XJZ9028`, 削除 `XJZ9008`, 戻る `XJZ9009`, No `XJZ0425`) and which buttons need a new number.
Outside `共通`, zoom buttons (`表示` = `Z`, or a `…ｽﾞｰﾑ` name — 460 rows, all `-`), calendar buttons
(`ｶﾚﾝﾀﾞｰ(…)`) and tab controls legitimately carry `-`: report only a placeholder or a non-`9xxx` ID
there, never a `-` or a blank (blank is as common as `-` on zoom/calendar buttons).
"Most-used ID" comparisons: use a prebuilt button-ID table if your prompt names one; otherwise compare
within the workbook and against the standard IDs listed here (検索 `XJZ9002`, 画面ｸﾘｱ `XJZ9003`, ﾀﾞｳﾝﾛｰﾄﾞ
`XJZ9004`, 更新 `XJZ9007`, 削除 `XJZ9008`, 戻る `XJZ9009`, 新規登録 `XJZ9028`, 処理状況 `XJZ9017`), and
leave other buttons' correct number to the designer.

**W1c — 選択 (No.48).** ("明細-type area" below = any area of repeating rows — titles with
`明細`/`一覧`/`履歴`/`通知先`…, or any area that has a `No` or `選択` item.) A `選択` item in a 明細-type area should be `表示` = `□` (`CheckBox`) with
画面項目ID `-`. Measured: 72 comply. Flag `表示` `-`/blank on a CheckBox 選択 (6 — e.g.
`PSJCO302` `GSJC302A` `[567]`; `PSJCO502` `GSJC502B` has two with every column blank), and a
画面項目ID on 選択 (10 — 6 of them `XJC8032`, so the designer may say it is a deliberate convention;
report it low and let them decide). A `RadioButton` 選択 (single
selection, `表示` `○` or `-`) is a different control: don't apply the `□` rule to it. A
`選択(ﾍｯﾀﾞｰ)` (the select-all checkbox in the list header) follows the same rule as `選択`.

**W1d — No (No.49).** A `No` / `No.` item in a 明細 area should be `Label`, 表示形式 `ZZ9`, 初期値
`1から連番`, 画面項目ID `XJZ0425`. Measured over 159 No rows: 51 fully comply, and the literal rule
over-fires on all three fields:
- **表示形式**: `-` in 74 rows, nearly as common as `ZZ9`. Still a finding (the rule is explicit and
  the value is needed for the display), but group it: one low-severity line per sheet, not per row.
  `ZZZ9` (a list allowed past 999 rows) is not a finding when the area note says so.
- **初期値**: `1からの連番`, and a value sourced from data (`(5).表示順`, `新規:1から連番⏎変更:(6)…`)
  are correct. Flag only `-`/blank, or `右記参照` with nothing to the right.
- **画面項目ID**: `XJA0042` is the 基準情報 dictionary's own No item and appears across 基準情報
  screens — report it as 要確認 (XJZ0425 is the common one), not as a defect. A different ID, or
  none, is a finding.

**W1e — TextBox/TextArea の桁数・入力可文字種 (call-out "属性がTextBox、TextAreaの場合は必ず
入力可能文字種、桁数を記載する").** Every `TextBox`/`TextArea` row needs a number in 桁数 and a value
in 入力可文字種. Measured: 1,555 of 1,688 comply. **Do not exempt date/time types**: `年月日(8桁)`
has 213 rows with 桁数 `8` and exactly 1 with `-`, so a `-` beside `年月日(8桁)` / `時間(4桁)` /
`数値(整数)正数` is the same omission as anywhere else. Report 桁数 `-` and 入力可文字種 `-`/blank
separately; a row missing both is one finding. `ﾊﾟｽﾜｰﾄﾞ` is a 入力可文字種 value and still needs 桁数.
A placeholder in 桁数 (`?`, `XX` — `PSJCO805` has six `?`) is the same as `-`: a length nobody has
decided yet.
This is stricter than, and complementary to, `xlsx-db-column-check` step 5 (which reports a `-` only
when another screen gives the same item a number): report here regardless, and don't repeat that
skill's DB-length comparison.

### W2 — 項番の飛び・重複 (No.6) (formerly design-doc-internal-consistency check 10)

Every numbered list in the four sheet types — a column (≤ col 8) whose
header cell is `No.`/`No`/`項番`, followed by integer cells — must run 1, 2, 3 … Also the circled
section numbers of 機能定義書 Ⅳ (`A-①`, `A-②` … per letter). **ﾁｪｯｸ処理設計書 is the exception to the
header rule**: its number column is col 1 under the header `ﾁｪｯｸ項目` (`[8,1]`), and each
`【…押下時】` heading starts a new list — scan it explicitly or its defects are missed (`PXJCO129`
`ﾁｪｯｸ処理設計書(GXJC129A)` `[16,1]`-`[18,1]` = 5, 6, 7 after 1-6; `GXJC129B` `[96,1]`/`[98,1]` both 84).
**End a list only on a heading in col ≤ 3** (an area title, `Ⅳ．…`, `【…】`) — the Ⅴ 説明 column (col 43)
routinely holds `Ⅳ．画面項目ｲﾍﾞﾝﾄ詳細 参照`, and treating that as a section break hides duplicates
(`GSJA704J` `[255,3]`). **Judge the sequence on live numbers only**; a struck row that still shows its
old number (`ﾁｪｯｸ処理設計書(GSJA704B)`: live 1, struck 2 and 3, live 2) is not part of the sequence and
must not produce a "backwards" finding — consult the struck numbers only to explain a gap.
**Also include the numbered sub-tables of 画面設計書 Ⅲ．画面表示仕様** (取得項目 / 検索条件 / 結合条件 / ｿｰﾄ順 /
集約条件, numbers in col 4 with no `No.` header): each list starts after its own label row in col 4 and ends at
the next such label or block. These were not part of the corpus calibration, but they hold real defects
(`PSJAO704` `GSJA704B`: ｿｰﾄ順 `[179,4]` 2→2, 検索条件 `[224,4]` a live row with no number where the struck
`4` was, `[454,4]` 3→5).

Normalise before flagging — measured, the naive check fires ~150 times on 94,445 numbered cells and
most of it is structure:
- **A drop back to 1 is a new list**, not an error (ﾁｪｯｸ処理 restarts per `【…押下時】` event; 39 such).
- **A gap explained by a struck row is the project's convention** — deleted rows keep their number in
  strikethrough (26 gaps, e.g. `PSJAO241` `更新条件表(TXJAM002)` 45→47 with a struck `46`). Before
  reporting a gap, look in `_DELETED_DIGEST.txt` for **deleted content on the rows between** the two
  live numbers — the COM dump's digest omits a bare struck number as filler, so don't wait for the
  number itself to appear. A row whose content is struck but whose number is **not** (`GSJA704B`
  【引戻ﾎﾞﾀﾝ押下時】 `[26,1]`-`[28,1]` live 2/3/4 over struck checks) is an empty live item — 要確認.
- **A constant step of 2 over three or more items** (`PXJCO906` `GXJC906A` 1, 3, 5, 7 in col 6) is a
  two-row item layout, not missing numbers.
- A sub-number (`A-②-1`, `A-②(1)`) belongs to its own series; don't fold it into the circled list.
What survives is real: **duplicate** (16 — `PSJAO704` `GSJA704J` `[255,3]` 9→9, `PSJCOA09`
`GSJCA09C` four `3`s in a row), **gap** (33 — `PSJAO704` `GSJA704F` `[174,3]` 4→6), **typo / backwards**
(16 — `PXJAO701` `GXJA701A` `[625,3]` 11→**123**→13, i.e. `12` mistyped; `GSJA704F` is in Ⅳ．画面項目ｲﾍﾞﾝﾄ詳細,
not Ⅴ — say which section), **list not starting at 1** (4),
and in 機能処理概要 (11 lists) — a skipped circled number (`PXJCO152` `A`: ①②⑥⑦⑧), a duplicate
(`PSJCO402` `A`/`B`: ④④), and a section filed under the wrong letter (`PSJAO704`: inside the `E` block,
`[143,3]` reads `F-⑥.戻る処理`, so E jumps ⑤→⑦ — the `E-⑥` was written as `F-⑥`; F itself is fine). Severity 低, except a duplicate/misfiled
circled section that another doc cites by number (`A-⑥ 参照`), which is 中 — the citation now points at
the wrong section. Some lists interleave two sequences in one column (`PXJAO243`
`更新条件表(TXJAM023)` col 2: 22, 25, 23, 29, 24, 35 …); report it once as "番号が二系統混在" rather than
as a cascade of gaps and backward steps.

### W3-W5 — 記述ルール from the self-check workbook (formerly naming-standard-compliance step 7)

(`01_Doc\08_機能定義書\11_工程管理\XXXXX000_ｾﾙﾌﾁｪｯｸ用.xlsx`,
call-outs on its sample sheets). Calibrated on 2026-10-01 over 120
工程管理 `PHASE1`-`PHASE3` workbooks; counts and examples are one snapshot — re-derive every verdict.

**W3 — 画面表示仕様の各ﾌﾞﾛｯｸ末行に取得件数.** Call-out on 画面設計書 Ⅲ: "取得件数を各末行に記載する".
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

**W4 — ﾚｽﾎﾟﾝｽ定義の表記.** Call-out on 機能定義書: "「ｼｽﾃﾑ共通設計書.ﾚｽﾎﾟﾝｽ 参照」に統一". Read the body
of 機能定義書 `Ⅶ．ﾚｽﾎﾟﾝｽ定義`. Measured over 109 sheets: 97 read exactly `※ｼｽﾃﾑ共通設計書.ﾚｽﾎﾟﾝｽ 参照`.
- `※呼び出し元で記載` / `※呼出元で記載` (9) — every one is a **サブプロ/バッチ** (`S…B…`), whose response is
  the caller's; correct, don't flag.
- `※06-09_ﾚｽﾎﾟﾝｽ一覧_工程管理を参照` (`PSJCO602`, `PSJCO604`) — 要確認: either the program has its own
  responses (then `design-doc-internal-consistency` check 2 verifies 06-09) or it should use the
  standard wording.
- Any other wording of the standard reference (missing `※`, `ｼｽﾃﾑ共設計書`, extra text) — 低, give the
  exact standard string.

**W5 — 画面名の後ろに画面ID.** Call-out on 機能定義書: "画面名を記載する場合は後ろに画面IDを記載すること".
In 機能定義書 `Ⅳ．機能処理概要` and 画面設計書 `Ⅳ．画面項目ｲﾍﾞﾝﾄ詳細`, a known screen name followed by
`画面` (`取引先詳細画面`) in a sentence that transitions to, opens, returns to or shows it
(`遷移`/`起動`/`戻る`/`表示する`/`開く`/`呼出`) must be followed by `(<画面ID>)`
(`取引先詳細画面(GXJA802B)に遷移する`). Build the screen-name list from every 画面設計書's 3C header
`[4,15]` (screen name) / `[4,10]` (画面ID) in the folder, not just this workbook — the name you meet is
often another program's screen (use the prebuilt list your prompt names; build it yourself when
none was handed to you). Normalise the header name first: some end in `画面` (`受入入力画面`),
most don't, so match `name` + optional `画面`. The program's **own** screens count too
(`PSJCO602` `[52,4]` `棚卸用仕掛在庫ﾃﾞｰﾀ作成画面を表示する` → `(GSJC602A)`).
**The screen must be the object of the verb** — `…画面に遷移する` / `…画面に戻る` / `…画面を起動する` /
`…画面を表示する` / `…画面を開く` / `…画面に起動する` / `…画面へ遷移する`. Normalise full-/half-width
parentheses before matching (`承認ｸﾞﾙｰﾌﾟ（ﾕｰｻﾞｰ)登録画面` = `承認ｸﾞﾙｰﾌﾟ(ﾕｰｻﾞｰ)登録画面`). An abbreviated name
(`承認ﾌﾛｰ一覧画面` for `ﾜｰｸﾌﾛｰ承認ﾌﾛｰ一覧`) or a dropped `画面` (`…登録に遷移`) is 要確認 with the
candidate ID, not a W5 finding. `…画面に表示する` / `…画面で選択された…` put data *on* a screen and
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

Group by check (W1-W5), then by sheet. Every rule here is low severity on its own unless the check
says otherwise (a duplicate/misfiled circled section that another doc cites by number is 中; a
placeholder ID such as `XXXXXXX` is 中). When one deviation repeats across a whole sheet, report it
once per sheet with the cells listed, as each section says — these rules fire widely and a row-by-row
list buries the useful findings.
