---
name: design-doc-writing-rules
description: Check a design-doc workbook against the self-check workbook's (XXXXX000_ｾﾙﾌﾁｪｯｸ用.xlsx) writing rules: Ⅴ per-item rules (upper-case note on code inputs, button IDs (共通 XJZ/SJZ 9xxx first, else the WG's 9xxx), 選択 = □, No = ZZ9/1から連番/XJZ0425, TextBox/TextArea carry 桁数 and 入力可文字種, 入力可文字種 consistent across screens, zoom buttons 表示 Z + Ⅳ参照), 項番 gaps/duplicates, 取得件数 at the end of each 画面表示仕様 block, the standard ﾚｽﾎﾟﾝｽ wording, screen names followed by their 画面ID, and 画面表示仕様 search conditions (LIKE needs a match mode, 名称 searches 部分一致) and join order (INNER before LEFT). How a doc is written, not cross-references (design-doc-internal-consistency) or ID numbering (naming-standard-compliance). Use when the user asks for 記述ルール/セルフチェック観点 compliance. For a full REV without named checks the entry point is `rev-program-review`; run this standalone only when selected there or asked for by name. Also W7 (zoom names/IDs vs 03.ｽﾞｰﾑ管理台帳) and W8 (other programs' names = their 機能定義書 file names).
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
`_DELETED_DIGEST.txt`), reporting conventions. W5 uses the all-WG screen-name list: use the one your
prompt names, and build it yourself (see W5) when none was handed to you. W7 and W8 use the zoom list and
the program-name list the same way. Hidden rows and columns of a live sheet (`_HIDDEN.txt`) are read like the
rest: a W1-W8 finding there is reported with 【非表示】 (this check owns hidden Ⅴ rows, `agent-guide.md`); a
whole hidden leftover sheet is not this check's.

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
comparing an ID (the self-check sample itself has `　XJA0130`). Re-resolve the columns at every header
row — they shift by one within a sheet (`PXJCO161` `GXJC161B` `[2492]`: 属性 col 18, 説明 col 43).
**Ⅴ ends at the `Ⅵ．` heading** in col 2 (`Ⅵ.` with a half-width dot occurs too, as does `Ⅴ.`):
Ⅵ．画面項目制御・出力仕様 repeats Ⅴ's area titles and numbered rows (No. col 3, 画面項目 col 5), so a
parser that keeps the last Ⅴ header past `Ⅵ．` reads its ○/× grid as Ⅴ items and gives false W1b/W1e hits.

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

**W1b — ﾎﾞﾀﾝの画面項目ID (No.44, No.35; rule fixed 2026-10-19).** The self-check says "共通のﾎﾞﾀﾝは9xxx番台" and
a call-out says "ﾎﾞﾀﾝは全て共通のXJZ9xxx(SJZ9xxx)を利用/採番"; the designers' answer No.35 allows a WG-prefixed
`9xxx`, and the user's decision of 2026-10-19 (my-review-skills 〔ユーザー指摘 20261006 PXJCO124 BB992〕〔回答 20261006〕)
narrows it: **whatever the program's subsystem (`XJC`/`SJC`, `XJA`/`SJA`), a button first takes the `共通`
(`XJZ`/`SJZ`) `9xxx` ID of the same name; only when `共通` has no button of that name does it take its own WG's
`9xxx`** ("新規採番はSJC/SJAで採番する" — `SJA9018` 基本情報ﾎﾞﾀﾝ stays correct). So the band test (`9xxx`, not the
prefix) is the first filter, and a `共通` lookup decides the rest.
**The 共通 button list.** Use the `82.画面項目辞書_共通` index (`_shared/reference-index.md`; the orchestrator hands
its path over — standalone, build it or read the file live with `scripts/xlsx_zip.py`). Keep the `XJZ9xxx`/
`SJZ9xxx` rows whose name is non-empty: 221 IDs / 215 names on 2026-10-19 (`XJZ` 147 live — `XJZ9124`-`9135` are
struck, retired; `SJZ` 74 — `SJZ9054`/`9055` and `9077`-`9147` are blank reserved rows). Six names occur twice:
行削除 `XJZ9031`/`XJZ9142`, 一括承認 `XJZ9102`/`XJZ9159`, 検査表出力設定 `XJZ9055`/`XJZ9115`, 処理状況
`XJZ9017`/`SJZ9052`, 反映 `XJZ9065`/`SJZ9075`, ｱｯﾌﾟﾛｰﾄﾞ `SJZ9032`/`XJZ9092`. Find duplicates yourself — the
dictionary's own `重複` column (col 2, a `COUNTIF` over the name column) also marks clashes with non-button items
(完了 `XJZ9139` × because of item `XJZ0033`) and never compares the `XJZ` sheet with the `SJZ` sheet.
**Match on `表示`, not 画面項目名** — the dictionary holds the caption (as in internal-consistency check 3), and the
corpus follows it: 帳票ﾃﾝﾌﾟﾚｰﾄ削除ﾎﾞﾀﾝ / 表示 `削除` is `XJZ9008` (20 rows, `PSJAO204` `GSJA204B`, `PSJAO705`
`GSJA705B`), ﾌｧｲﾙ選択ﾎﾞﾀﾝ / `ﾌｧｲﾙの選択` is `XJZ9064`, 上移動ﾎﾞﾀﾝ / `▲` is `SJZ9062`. Fall back to 画面項目名 only
when `表示` is `-`, blank or only `※n`; a glyph caption that is not in 共通 is not a fallback case (`PSJCO204`
`GSJC204A` `[807]` 解除ﾎﾞﾀﾝ / `⊟` is not `SJZ9015` 解除). Normalise both sides: NFKC, drop all whitespace
(`SJC9044` `実　行`), a trailing key hint `(ENT)`/`(R)`/`(F9)`…, `※n`, series markers `(1-N)`/`1～15`, and a
trailing `ﾎﾞﾀﾝ` (the dictionary has `自動検索ONﾎﾞﾀﾝ`). **Exact match only** — 160 of the 311 own-band buttons with
no 共通 match *contain* a 共通 name and are different buttons (`再印刷`, `一括更新`, `ｸﾞﾚｰ行削除`, `流動停止解除`,
`ﾌｧｲﾙ選択` vs `ﾌｧｲﾙの選択` on the 名 side). **Never match a `表示` of `-`** to `SJZ9065` `-`: it means "no caption"
(18 false hits — right-click 切り取り/貼り付け on `GSJA401B`/`E`/`F`, 連番1-4 on `GSJC309A`) unless the cell says
so (`PSJAO702` `GSJA702A` `[215]` `-⏎※ﾏｲﾅｽﾎﾞﾀﾝ`). A nested 画面項目名 sits one or two columns right of the header
(`GXJC161B` `[2446]` col 6, `GSJC308A` `[1488]` col 7): take the first non-empty cell between the 画面項目名 and
表示 columns.
**Judge each non-zoom Button, in any area** (measured 2026-10-19 on 1,472 such rows in 122 books / 215 sheets):
- 共通 has the name and the ID is one of its IDs → pass (909 of the 921 rows citing `XJZ`/`SJZ`; a doc citing
  the higher of a duplicated pair passes — 処理状況 `SJZ9052` ×2, ｱｯﾌﾟﾛｰﾄﾞ `XJZ9092` ×2).
- **共通 has the name and the ID is an own-WG `9xxx` → finding, 低, 確定** (not 要確認): 「共通に同名のﾎﾞﾀﾝ
  `<ID>` があるため共通の画面項目IDを使用」. The fix is the 共通 ID; for a duplicated name, the lower number
  (numeric part, `XJZ` before `SJZ` on a tie) — still 確定. 15 rows on 11 sheets: 行削除 `XJC9083` →
  `XJZ9031` (`PSJCO309` `GSJC309A` `[1324,54]`, `PSJCO308` `GSJC308A` `[1460,54]`), 行追加 `XJC9082` →
  `XJZ9030` (same sheets `[1323,54]`, `[1459,54]`), ﾜｰｸﾌﾛｰ `SJC9004` → `SJZ9067` (`GSJA203A` `[415]`,
  `GSJA241A` `[557]`, `GSJA401A` `[312]`, `GSJC402A` `[116]`, `GSJC402B` `[140]`), 実行 `SJC9044` → `XJZ9016`
  (`GSJC509A` `[333]`), 電子配信 `XJC9027` → `XJZ9093` (`GXJC161B` `[2027]`), 設定 `XJC9009` → `XJZ9071`
  (`GXJC906A` `[474]`), 完了 `XJC9042` → `XJZ9139` in 実績表項目領域 (`GXJC130C` `[303]`, `GXJC161B` `[2446]`,
  `GSJC308A` `[1488]` — four other sheets' 実績表項目領域 already use `XJZ9139`).
- 共通 has the name and the ID is a *different* `XJZ`/`SJZ` ID → 確定 too, naming the matching ID: `GSJC402A`
  `[112]` / `GSJC402B` `[136]` 画面ｸﾘｱ `XJZ9030` (行追加) → `XJZ9003`; `GSJC604B` `[797]` 検索 `XJZ9007` →
  `XJZ9002`; `GSJC606A` `[294]` 更新 `XJZ9013` → `XJZ9007`; `GSJA205D` `[86]` 処理状況 `XJZ9009` → `XJZ9017`.
- 共通 has no button of that name → the WG's own `9xxx` is correct (No.35; 311 rows — `SJA9018` 基本情報,
  `SJC9012` 再印刷, `SJA9023` 変更履歴). Only these buttons go through the most-used-ID comparison below.
- ID blank, `-`, placeholder, malformed or outside `9xxx` → the bullets below; whenever 共通 has the name, the
  fix is that 共通 ID (50 blank/placeholder and 24 non-`9xxx` rows have one).
Glyph captions are real 共通 names (`+` `SJZ9064`, `-` `SJZ9065`, `▲`/`▼` `SJZ9062`/`9063`, `×` `SJZ9070`, `<`/`>`
`XJZ9156`/`9157`) and are matched when the caption is that glyph; `Z` `SJZ9066` is the zoom button and stays out
of W1b with the other zoom/calendar buttons, as below. Several buttons with the same caption on one sheet
(`GSJC801B` ｶﾗｰ選択 ×8 → `SJZ9056`) share the ID legitimately — not a duplicate-ID finding.
Flag, for a non-zoom Button in a `共通` area:
- ID blank or `-` (59 blank — e.g. `PSJCO805` `GSJC805A` has no ID on any of its 12 共通 buttons);
- the placeholder `XXXXXXX` (`PSJCO402`, `PSJCO403`);
- a malformed prefix (`PSJAO704` `GSJA704I` 行削除ﾎﾞﾀﾝ `XZJ9031` — `XZJ` for `XJZ`);
- an ID outside `9xxx` (14 — `XJC0323` 削除ﾎﾞﾀﾝ, `SJA0349` 削除ﾎﾞﾀﾝ, `XJZ0079` 返却ﾎﾞﾀﾝ). Name the 共通 `9xxx` of
  the same name as the fix (削除 `XJZ9008`); without one, the `9xxx` the other sheets use for that button. A
  cell holding two IDs (`SJA0873⏎SJA9048`) passes if one of them is `9xxx`.
When `表示` is `-` and the match came from 画面項目名 (`PSJCO308` `GSJC308A` 行追加 `[1459,54]` / 行削除 `[1460,54]`),
the 共通 fix stays 確定, but tell the designer to confirm in Ⅰ/the 画面ｲﾒｰｼﾞ sheet that it is not an icon button
(`+` `SJZ9064` / `-` `SJZ9065`), which would be the ID instead.
Two checks the band test cannot catch, both worth a line:
- **the same ID on two different buttons of one sheet** (not a button repeated on every detail row —
  a per-row `削除ﾎﾞﾀﾝ` sharing one ID is one control) — `PSJAO205` `GSJA205D` gives 処理状況ﾎﾞﾀﾝ
  `[86,54]` the ID `XJZ9009`, which is its own 戻るﾎﾞﾀﾝ's `[85,54]`; a duplicate ID means one label
  is wrong at runtime;
- **a button with no 共通 name whose ID differs from the one the rest of the corpus uses** for that
  button name. Build the name → most-used-ID table from the other screens you have and report a
  deviation as 要確認, naming the usual ID. (A button whose name *is* in 共通 is judged against the
  dictionary above, 確定 — not here.)
**When a sheet's whole 画面項目ID column is empty** (`PSJCO805`: all six screens, every row), do
not list every button, No and input separately — report "画面項目ID 未採番" once per sheet, with the
standard IDs you can already suggest (検索 `XJZ9002`, 画面ｸﾘｱ `XJZ9003`, ﾀﾞｳﾝﾛｰﾄﾞ `XJZ9004`, 新規登録
`XJZ9028`, 削除 `XJZ9008`, 戻る `XJZ9009`, 実行 `XJZ9016`, No `XJZ0425`) and which buttons need a new number.
Outside `共通`, zoom buttons (`表示` = `Z`, or a `…ｽﾞｰﾑ` name — 460 rows, all `-`), calendar buttons
(`ｶﾚﾝﾀﾞｰ(…)`) and tab controls legitimately carry `-`: report only a placeholder or a non-`9xxx` ID
there, never a `-` or a blank (blank is as common as `-` on zoom/calendar buttons).
Outside `共通`, a non-`9xxx` ID on a program-specific button is 要確認, not a defect — name the `9xxx`
IDs the sheet's other buttons use. Area-toggle buttons (the button that carries its area's name and
opens/closes it) are mixed in practice: `PXJCO161` `GXJC161B` 実績表項目 `XJC0252` / 不良項目 `XJC0333` /
QA判定 `XJC0334` against 分割 `XJC9087` / 保留分割 `SJC9053` / 資源/依頼ｺｰﾄﾞ `SJC9052` on the same sheet —
report them once per sheet as 要確認. A button whose ID is also a non-button item's ID on the same
sheet is a **duplicate-ID finding** (中), not a band question: `GXJC161B` 不良項目ﾎﾞﾀﾝ `[2493,54]` and the
Label 不良項目(1-N) `[2494,54]` are both `XJC0333` — the button needs its own number.
"Most-used ID" comparisons (buttons with no 共通 name only): use a prebuilt button-ID table if your prompt
names one; otherwise compare within the workbook. For a button whose name is in 共通 the dictionary is the
authority; if no 共通 index was handed over and you cannot read the file, fall back to the standard IDs
listed here (検索 `XJZ9002`, 画面ｸﾘｱ `XJZ9003`, ﾀﾞｳﾝﾛｰﾄﾞ `XJZ9004`, 更新 `XJZ9007`, 削除 `XJZ9008`, 戻る
`XJZ9009`, 実行 `XJZ9016`, 新規登録 `XJZ9028`, 処理状況 `XJZ9017`, 行追加 `XJZ9030`, 行削除 `XJZ9031` — all
confirmed in `82.画面項目辞書_共通` `XJZ(共通_JAGUR)` on 2026-10-19), say the 共通 lookup was skipped, and leave
other buttons' correct number to the designer.

**W1c — 選択 (No.48).** ("明細-type area" below = any area of repeating rows — titles with
`明細`/`一覧`/`履歴`/`通知先`…, or any area that has a `No` or `選択` item.) A `選択` item in a 明細-type area should be `表示` = `□` (`CheckBox`) with
画面項目ID `-`. Measured: 72 comply. Flag `表示` `-`/blank on a CheckBox 選択 (6 — e.g.
`PSJCO302` `GSJC302A` `[567]`; `PSJCO502` `GSJC502B` has two with every column blank), and a
画面項目ID on 選択 (10 — 6 of them `XJC8032`, so the designer may say it is a deliberate convention;
report it low and let them decide). A `RadioButton` 選択 (single
selection, `表示` `○` or `-`) is a different control: don't apply the `□` rule to it. A
A Ⅴ `選択(ﾍｯﾀﾞｰ)` row (the select-all checkbox in the list header) follows the same rule as `選択`. The
name difference between a Ⅳ `選択(ﾍｯﾀﾞｰ)` event and the area's Ⅴ `選択` row (when that Ⅴ area has no
`選択(ﾍｯﾀﾞｰ)` row) is the intended way to tell the header event from the row event — never report the
name difference or the Ⅳ name as absent from Ⅴ (`design-doc-internal-consistency` check 6).

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
**A row whose 属性 is `右記参照` is not skipped**: its attribute is given per case in 初期値/説明
(`PXJCO161` `GXJC161D` `[291]` 管理No: `属性:Label` for the first row in `[291,35]`, an input otherwise —
it has 桁数 15, TAB 2 and the upper-case note). If any case is `TextBox`/`TextArea`, apply W1e/W1f to
the row; if no case states an attribute, report 要確認. The converse — a `Label` row with TAB, 必須 or
入力可文字種 filled (`PSJCO403` `GSJC403A` `[1405]` 単価: TAB 24, 必須 ○, `数値(小数)`, 桁数 `-`, and the
check sheet has a 【単価変更時】 event) — is a finding (要確認): either the 属性 is wrong (an input) or the
input columns are leftovers; name which the rest of the doc implies.
This is stricter than, and complementary to, `xlsx-db-column-check` step 5 (which reports a `-` only
when another screen gives the same item a number): report here regardless, and don't repeat that
skill's DB-length comparison.

**W1f — 入力可文字種の画面間統一 (レビュー観点 No.42 「入力可文字種は他画面の同じ項目と同じ文字種か」).**
`xlsx-db-column-check` step 5 compares 桁数 across screens; this is the same comparison for
入力可文字種. Scope: `TextBox`/`TextArea` rows whose 文字種 is a real value (`-`/blank is W1e; a `※n`
footnote cannot be compared). Match items by **normalised 画面項目名** (strip whitespace, `(FROM)`/`(TO)`,
`(1-N)`, a trailing digit), never by 画面項目ID alone — `PSJCO403` `GSJC403A` Ⅴ's 作業工程GRP row carries 製造ﾛｯﾄNo's
`XJC0028`, which is a wrong ID (`design-doc-internal-consistency` check 3), not a 文字種 deviation.
Include `右記参照` rows as W1e resolves them. **Vocabulary**: a value outside the house forms
(`文字列(半英数)`, `文字列(半英数記号)`, `文字列(全角)`, `数値(整数)正数`, `年月日(8桁)` …) — `半角英数` with no
`文字列(…)` (`PXJCO161` `GXJC161D` `[291,30]`), `数値(少数)` — is one 低 line per workbook. `文字列(半数)` is a
house form (18 uses in 10 PHASE workbooks, no alternative spelling; designers cite the self-check sheet),
not a vocabulary outlier — compare it only in the cross-screen check below
listing the cells, and still compared below by its evident meaning.
Two comparisons:
- **Against other programs** — report a row whose 文字種 differs from the value other programs use
  for the same item **when that value is dominant**: ≥ 5 rows from ≥ 3 distinct other programs, and
  ≥ 80 % of those rows. Use the prebuilt table your prompt names, or standalone build it and list the
  candidates (scripts under `<plugin>\skills\_shared\scripts\`):
  `python build_kind_table.py <...\01_Doc\08_機能定義書\11_工程管理> kind.tsv` (~45 s, up to 3 min on a busy machine; the script
  itself limits a folder with `PHASE*` sub-folders to the workbooks directly inside them), then
  `python kind_consistency.py kind.tsv <ProgramID>` — it prints `[row,col]` and groups the rows of one
  item with one value. **Report one finding per item and value**, listing every cell (`PSJCO805`
  ｲﾝｸ品名 on `GSJC805A`/`D`/`E`/`F` is one line). Confirm each cell in the live dump before reporting.
- **Within the program** — the same item on two of its own screens with different 文字種 (the script's
  part (b)). Report as 要確認; a date granularity difference can be intended (`PSJCO805` 回答納期 is
  `年月(6桁)` on `GSJC805C` `[412]` and `年月日(8桁)` on four other screens).
Calibrated 2026-10-02 on 1,867 TextBox/TextArea rows (live text) in the 122 工程管理 `PHASE1`-`PHASE3` workbooks:
of 231 items used on ≥ 2 sheets, 202 are consistent, and the dominant-value rule yields **13 grouped
findings (17 cells) in 8 programs**, plus 6 within-program pairs. The examples below are that 2026-10-02 snapshot,
not current state (`PSJCO204` was edited 2026-10-05 and no longer reproduces) — re-run the script — e.g. `PSJCO204` `GSJC204A` `[728,30]` 品目ｺｰﾄﾞ `文字列(半英数)` against `文字列(半英数記号)` in
54 of 54 rows (a code with a symbol could not be entered; the same sheet's ﾃｰﾌﾟﾛｯﾄNo `[744,30]` and
製造ﾛｯﾄNo `[745,30]` have the same narrowing), `PSJCO305` `GSJC305A` `[407,31]`/`[409]`
資源ｺｰﾄﾞ1/2 `文字列(全角)` against `文字列(半英数)` 18/18, `PSJAO404` `GSJA404B` `[281]` 工程ｺｰﾄﾞ
`文字列(半数)` against `文字列(半英数)` 32/32, `PSJCO204` `[854]` 識別ｶｰﾄﾞ発行枚数 `文字列(半英数)` where
every other program has `数値(整数)正数`. Items that are mixed across the corpus have no dominant value
and are **not** reported — `KC品名` (`文字列` 16 / `半英数記号` 11 / `半英数` 4), `処置指示No` 7:5,
`KC図番`, `資源GRP`. Severity: 中 when the screen's type is narrower than or incompatible with the
dominant one (`半英数` vs `半英数記号`, `全角` vs `半英数`, a string type where others use `数値`, and
`数値` where others use `文字列` — leading zeros and letters are lost); 要確認 otherwise (a broader or
merely different type — `文字列(半角)` vs `文字列(半英数記号)`, `数値(整数)` vs `数値(整数)正数`, where a
negative entry may be intended).
Note that the dominant value can contradict the call-out "ｺｰﾄﾞ/No/ｷｰは半角英数": 品目ｺｰﾄﾞ, 製造ﾛｯﾄNo and
ﾃｰﾌﾟﾛｯﾄNo are overwhelmingly `半英数記号`. Follow the corpus here and do not raise that call-out against them.

**W1g — ｽﾞｰﾑﾎﾞﾀﾝの表示・説明 (レビュー観点 No.38 「G1)検索条件領域（他Gも）のズームの表示列はZの記載があるか。
説明列に「Ⅳ．画面項目ｲﾍﾞﾝﾄ詳細 参照」の記載があるか。」).** A zoom button is a Ⅴ row, in any area, whose 属性
is `Button` (first line — `Button⏎※5` counts) or `その他` (基準情報 sheets write zooms that way:
`PXJAO802` `GXJA802A` `[196]`), and whose `表示` is `Z` **or** whose 画面項目名 contains `ｽﾞｰﾑ`/`ズーム`
— recognise it by either, because each finding below breaks one of the two (`PXJCO134` has `ｽﾞｰﾑ` names
with `表示` `-`; `PSJCOA09` `GSJCA09F` `[133]`/`[138]` `管理加工GRPｿﾞｰﾝ` is a zoom only by its `Z`).
Out of scope: `ｶﾚﾝﾀﾞｰ(…)` and other buttons; `ｽﾞｰﾑｲﾝ`/`ｽﾞｰﾑｱｳﾄ` (diagram magnification, `PSJAO704`
`GSJA704B` `[557]`/`[558]`); the `Label` rows `F9:ｽﾞｰﾑ` (共通 key guide, 13 rows) and `ｽﾞｰﾑ名`. Columns as in
the parse rule above: `表示` by its exact label (`表示形式` also starts with 表示), 説明 by prefix — it reads
`説明/⏎表示時参照ﾃｰﾌﾞﾙ` or `説明／⏎表示時参照ﾃｰﾌﾞﾙ` (full-width slash, `PXJCO129` `GXJC129B` `[614,43]`), and a
lookup on the half-width form alone found no 説明 column on 125 zoom rows. Two rules:
- **表示 = `Z`.** Anything else — `-`, blank, a caption — is a finding.
- **説明 refers to Ⅳ.** After removing whitespace (half/full-width), 説明 must match
  `Ⅳ[.．]?画面項目(ｲﾍﾞﾝﾄ|イベント)詳細.*参照`. All of these are compliant: `Ⅳ．画面項目ｲﾍﾞﾝﾄ詳細 参照`,
  `Ⅳ.画面項目ｲﾍﾞﾝﾄ詳細参照`, full-width `イベント` (the self-check sample's own `[321,43]`), a qualifier
  (`…詳細.G1)検索条件領域(No.2)参照`, `…詳細.取引先ｽﾞｰﾑ参照`), and trailing text (`…参照⏎※1`,
  `…参照行追加ﾎﾞﾀﾝで追加した行のみ、表示`). A bare `Ⅳ参照` would not match; it does not occur in the corpus, so
  if you meet one, report it as 低 wording rather than as a missing reference. `-`, blank, or a prose
  description of the zoom is a finding — the prose restates Ⅳ and drifts from it: `PXJCO163` `GXJC163B`
  `[1005]` says `画面.資源ｺｰﾄﾞをもとに資源ｽﾞｰﾑを起動` while Ⅳ `[745]` also passes 有効期限ﾌﾗｸﾞ.
Calibrated 2026-10-05 on the 122 工程管理 `PHASE1`-`PHASE3` workbooks (live text): **535 zoom buttons on 136
sheets** (4 `ｽﾞｰﾑｲﾝ`/`ｱｳﾄ` excluded); 524 have `表示` `Z` and 523 carry the Ⅳ reference (485 in the two
plain forms, 38 variants). Findings — **11 `表示` on 4 sheets, 12 `説明` on 5 sheets**:
- `表示`: `PXJCO134` `GXJC134A` — 8 of its 9 zooms `-` (`[506]` 処置指示Noｽﾞｰﾑ … `[534]` 停止工程ｽﾞｰﾑ; only
  `[508]` KC品名ｽﾞｰﾑ has `Z`);
  `PXJAO233` `GXJA233B` `[584]` 不良項目ｺｰﾄﾞｽﾞｰﾑ `-`; `PXJAO701` `GXJA701B` `[1349]` 資源GRPｽﾞｰﾑ blank;
  `PSJCO205` `GSJC205A` `[451]` ｵｰﾀﾞｰ工程属性(共通)ｽﾞｰﾑ `工順詳細` with ID `SJC9061` — a captioned
  button that opens a zoom, so 要確認 (it may be meant as a text button).
- `説明`: `PSJCO805` `GSJC805B` `[121]`/`[124]`/`[135]` and `GSJC805C` `[383]`/`[386]`/`[389]` blank;
  `PSJAO704` `GSJA704E` `[172]`/`[176]` `-`; `PSJAO501` `GSJA501B` `[477]` `-`; `PXJCO163` `GXJC163B`
  `[953]`/`[1005]`/`[1008]` prose.
Severity 低 (the event itself is specified in Ⅳ; this is how Ⅴ points at it), except 要確認 for a caption in
`表示`. **Group by sheet**: one line per sheet and rule, listing the rows (`PXJCO134` `GXJC134A` is one line
of 8). When a zoom has no Ⅳ event at all, that is a cross-reference gap (`design-doc-internal-consistency`)
— mention it, don't count it here. A misspelt name such as `ｿﾞｰﾝ` belongs to the typo check.

### W2 — 項番の飛び・重複 (No.6) (formerly design-doc-internal-consistency check 10)

Every numbered list in 表紙 and the four sheet types (機能定義書, 画面設計書, ﾁｪｯｸ処理設計書, 更新条件表) — a
column (≤ col 8) whose header cell is `No.`/`No`/`項番`, followed by integer cells — must run 1, 2, 3 … This
excludes 表紙 Ⅲ．改訂履歴, which is not reviewed (designers' answer 2026-10-02 No.48/50/73-76/102). Also the circled
section numbers of 機能定義書 Ⅳ (`A-①`, `A-②` … per letter), and **a run of circled step numbers inside
one cell** — split the cell on newlines and read the line-initial ①②③…; they must not repeat or skip
(`PXJCO125` `画面設計書(GXJC125A)` `[416,24]` ①②**②**). A `※1` or a mid-line circled reference is not a step.
Apply this in-cell rule to Ⅳ prose only (機能定義書 Ⅳ, 画面設計書 Ⅳ). Elsewhere a line-initial `①.` followed
by an item reference is a source reference, one per condition branch — `PXJCO161` `更新条件表(TXJCM008)`
`[140,18]` `①.G4).連結完了工程＝空白の場合、NULL⏎①.G4).連結完了工程≠空白の場合、…` cites 更新概要's source ①
twice and is not a ①① duplicate (~9 such false hits on that workbook).
The same applies to **parenthesised step numbers** `1)`, `2)`, `3-1)`, `3-2a)` at line start — inside one cell
or one per row down a `※n` event block of 画面設計書 Ⅳ (the block ends at the next `※n` heading). Each level
(`n)`, `n-m)`, `n-ma)`) is its own series under its parent: `PSJCO403` `GSJC403A` ※2 `[1176,4]`/`[1177,4]`
both `2)` with no `1)`, ※10 `[1254,6]`/`[1255,6]` both `3-2a)` (meant `3-2a`/`3-2b`). A `2-2)を繰り返して` reference
mid-line is not a step.
**ﾁｪｯｸ処理設計書 is the exception to the
list-break rule**: its number column is col 1 under the header `No` (`[8,1]`; `[8,3]` is `ﾁｪｯｸ項目`, and when
matching header labels normalise half-width `･`/`・` — `[8,14]` reads `ﾁｪｯｸ内容･経緯`), and **each `【…】`
heading in col ≤ 3 ends the current list and starts a new one** (`[9,3]` `【検索ﾎﾞﾀﾝ押下時】`, `[40,3]`
`【実行ﾎﾞﾀﾝ押下時】` on `PXJCO125`) even though it sits right of col 1 — the general rule below would never
fire there. **Except** a heading that repeats the event already open with trailing text starting
`上記` — `PXJCO161` `ﾁｪｯｸ処理設計書(GXJC161B)` `[267,3]` `【実行ﾎﾞﾀﾝ押下時】上記ﾁｪｯｸ処理で1件もｴﾗｰにならなかった場合、実行する`
inside the 【実行ﾎﾞﾀﾝ押下時】 of `[22,3]`, numbering 236 → 237 — is a sub-heading: the list continues.
A bare heading repeated verbatim further down is still a new
list — scan it explicitly or its defects are missed (`PXJCO129`
`ﾁｪｯｸ処理設計書(GXJC129A)` `[16,1]`-`[18,1]` = 5, 6, 7 after 1-6; `GXJC129B` `[96,1]`/`[98,1]` both 84).
Elsewhere, **end a list only on a heading at or left of the number column** (an area title, `Ⅳ．…`, `【…】`;
col ≤ 3 for the usual col-2 number column) — the Ⅴ 説明 column (col 43)
routinely holds `Ⅳ．画面項目ｲﾍﾞﾝﾄ詳細 参照`, and treating that as a section break hides duplicates
(`GSJA704J` `[255,3]`). A `【…の場合】` condition line inside a table is a branch caption, not a heading —
the list continues (`PXJCO161` `GXJC161F` `[58,5]` `【(3)で取得した内外作区分="1"(社外)の場合】` between 検索条件
8 and 9). Zero-padded values are data, not numbering (`GXJC161B` `[2146,6]`/`[2147,6]` `003`/`004` under a
`層No` header in a worked example).
**Judge the sequence on live numbers only**; a struck row that still shows its
old number (`ﾁｪｯｸ処理設計書(GSJA704B)`: live 1, struck 2 and 3, live 2) is not part of the sequence and
must not produce a "backwards" finding — consult the struck numbers to classify a gap (below).
**Also include the numbered sub-tables of 画面設計書 Ⅲ．画面表示仕様** (取得項目 / 検索条件 / 結合条件 / ｿｰﾄ順 /
集約条件, no `No.` header): the number column is **the column of that table's own label**, so detect it
per table — col 4 at top level, further right in a nested block (`PXJCO125` `GXJC125A` `[73,6]` 取得項目
→ `[75,6]` 1, `[76,6]` 検索条件 → `[77,6]` 1). Each list starts after its label row and ends at the next
label or block at that column or left of it. These were not part of the corpus calibration, but they hold real defects
(`PSJAO704` `GSJA704B`: ｿｰﾄ順 `[179,4]` 2→2, 検索条件 `[224,4]` a live row with no number where the struck
`4` was, `[454,4]` 3→5).

Normalise before flagging — measured, the naive check fires ~150 times on 94,445 numbered cells and
most of it is structure:
- **A drop back to 1 is a new list**, not an error (ﾁｪｯｸ処理 restarts per `【…押下時】` event; 39 such).
- **A gap left by a struck row is a finding — the rows after it must be renumbered** (user decision
  2026-10-19, my-review チェック観点 〔課題152-1085〕: 取り消し線で行を削除したら、その後ろの項番を振り直して
  いるか). A deleted row stays in place in strikethrough; the live rows below it must still run 1, 2, 3 ….
  Measured 2026-10-19 on the 122 工程管理 `PHASE1`-`PHASE3` books: of the struck runs between two live numbers,
  667 were followed by renumbered live rows (575 No. columns, 92 Ⅲ sub-tables; also 22 of 22 `A-①` sections
  and 30 in-cell ①② runs renumbered in place) and 61 were not (29 + 32, in 40 lists, 19 sheets, 16 books) —
  renumbering is the norm, the gap is the defect. Examples: `PSJAO241` `更新条件表(TXJAM002)` `[60,2]` struck
  `46` 代表製品ｺｰﾄﾞ, live 45→47 in all four update blocks (`TXJAM002WF` 47→49 likewise; ﾃｰﾌﾞﾙﾚｲｱｳﾄ itself was
  renumbered — 社内管理番号 is its No.46); `PSJCO308` `GSJC308A` Ⅴ `[1188,3]` struck 26 (col 59 note
  `No26の行を削除`) and `[1194,3]`-`[1195,3]` struck 32/33, live 25→27, 31→34; `PXJCO161` `機能定義書(PXJCO161)`
  `[443,7]`/`[451,7]` 20→22, 28→30; `PSJCO404` `GSJC404A` G2) `[1723,3]` 1→3; `PSJCO803` `GSJC803A` 検索条件
  `[91,4]`-`[93,4]` 13→17; `PSJAO204` `GSJA204B` 取得項目 `[106,4]`-`[124,4]` every even number struck.
  **How to decide**: read the list with its struck numbers in sheet order — `live_dump.py`'s digest keeps
  them; the COM digest drops a bare struck number as filler, so there look for **deleted content on the rows
  between** the two live numbers. Report when live `a` is followed by live `b > a+1` and the missing numbers
  belong to struck rows between them — or when a list's first live number is `k+1` with struck `1`..`k` above it
  (`PSJCO308` `ﾁｪｯｸ処理設計書(GSJC308A)` 【実行ﾎﾞﾀﾝ押下時】 `[107,1]`/`[108,1]` struck 1/2, live 3-47 →
  「No.3以降を2つ繰り上げ」). Numbers quoted inside revision memos (col 59, col ≥ 100) are not citations. Not a finding:
  - struck `k` followed by live `k` (renumbered, or the row replaced in place), or a struck row still showing
    its pre-renumbering number (`PSJAO203` `GSJA203B` `[1285,3]` struck 21 between live 19 and 20; 271 such) —
    the live sequence is continuous, never call it backwards;
  - a struck run at the end of a list (273), a struck header (the whole table is deleted), a superseded sheet;
  - in-place renumbering in a heading — old number struck, new one typed beside it (`PXJAO243` `[128,3]`
    `B-②③` live `B-③`; `PXJCO152` `[84,3]` live `A- ③`): take the number from the live text, allow
    whitespace after `A-`, and fold ➀-➉ (`PSJAO401` `[312,3]` `F-➅`) into ①-⑩;
  - in-cell ①② where the strike also removed a line break, so the next step now sits mid-line
    (`PSJAO241` `GSJA241B-規格情報` `[1105,20]` `…戻る。⑤…`, `PSJCO501` `GSJC501B` `[790,24]`) — read a circled
    step after `。` too; the corpus has no real in-cell strike gap;
  - a struck row whose only live text is a `-` placeholder anywhere, or text in col ≥ 54 only (画面項目ID / SQL)
    or in the revision memos (col 59, col ≥ 100) — the row is still deleted (`PSJCO308` `[1188,54]` `-`).
  Reported differently: the number struck but the row content live (`GSJA704B` 検索条件 `[224,4]` struck `4`
  over a live `B.承認状態 NOT IN …`) — the deletion is unfinished, 要確認; the content struck but the number
  live (`GSJA704B` 【引戻ﾎﾞﾀﾝ押下時】 `[26,1]`-`[28,1]` live 2/3/4 over struck checks) is an empty live item —
  要確認; a struck row whose content reappears live elsewhere in the list (`PXJCO130` `GXJC130B` `[905,5]`-
  `[909,5]`: 1, 2 製伝No, 2 出荷ﾛｯﾄNo, struck 3 製伝No, 4) is one 「行移動後の振り直し漏れ」, not a duplicate plus a gap.
  Severity 低, one line per list naming the first live number to change (`No.47以降を1つ繰り上げ`); the same gap
  in repeated blocks of one sheet (TXJAM002 ×4) is one line. 中 when the numbers are cited elsewhere
  (`No.5`, `(n).`, `A-⑥ 参照`) — say that the citations must be renumbered with the list. The Ⅲ block headings
  `(n)` are not part of W2 (their gaps — 29 left, 29 renumbered in the corpus — are cited by number all over the
  doc; check 12 resolves those citations).
- **Numbers that were never written up** — pre-numbered rows with every other column empty
  (`PSJCO403` `ﾁｪｯｸ処理設計書(GSJC403A)` `[12,1]`-`[39,1]`: all 13 events from 【検索ﾎﾞﾀﾝ押下時】 to
  【単価変更時】) — are not a numbering defect. They are a finding of their own: one line per sheet, 中,
  "ﾁｪｯｸ内容未記載" listing the events; an event with no check says so (`[10,14]` `ﾁｪｯｸ無し` under 【初期処理】).
- **A constant step of 2 over three or more items** (`PXJCO906` `GXJC906A` 1, 3, 5, 7 in col 6) is a
  two-row item layout, not missing numbers.
- A sub-number (`A-②-1`, `A-②(1)`) belongs to its own series; don't fold it into the circled list.
What survives is real: **duplicate** (16 — `PSJAO704` `GSJA704J` `[255,3]` 9→9, `PSJCOA09`
`GSJCA09C` four `3`s in a row), **gap** (33 — `PSJAO704` `GSJA704F` `[174,3]` 4→6), **typo / backwards**
(16 — `PXJAO701` `GXJA701A` `[625,3]` 11→**123**→13, i.e. `12` mistyped; `GSJA704F` is in Ⅳ．画面項目ｲﾍﾞﾝﾄ詳細,
not Ⅴ — say which section), **list not starting at 1** (4),
and in 機能処理概要 (11 lists) — a skipped circled number (read from the live text — `PXJCO152` `A`, once cited here as ①②⑥⑦⑧, is ①-⑧ live after in-place renumbering), a duplicate
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
(`PSJCO503` `GSJC503A` `[156,6]` … `[172,6]`). A block ends only at a label or block title at its
column or left of it — **not at an opening bracket**, which sits one column left of the No. (W6
anatomy): `PXJCO161` `GXJC161A` `[369,5]` `(` belongs to 検索条件 7 of the col-6 block whose 取得件数 is
`[373,6]`. Scope is 画面設計書 only — a ﾁｪｯｸ処理設計書 Ⅱ block with
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
- **An unfilled template Ⅲ** — aliases with no entity names, numbered 取得項目/検索条件 rows with nothing in
  them, the template `取得件数 … 取得できない場合 -` (`PSJCO403` `GSJC403B` `[76,4]`-`[107,4]`, `GSJC403C`/`D` `[36,4]`-`[67,4]`) — is not
  a 件数 finding: report "Ⅲ 未記載" once per sheet instead of a line per block. A single stray template
  block inside an otherwise filled Ⅲ (`PSJCO403` `GSJC403A` `[1103]`-`[1136]`: title `()**取得`, empty
  取得項目 and 結合条件 rows, `取得件数` with no 件数) is likewise one 「未記載ﾌﾞﾛｯｸ」 line, not a 件数 finding.

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
(`取引先詳細画面(GXJA802B)に遷移する`). The screen-name list covers every visible 画面設計書 under
`01_Doc\08_機能定義書` (all WGs), not just this workbook — the name you meet is often another program's
screen. Its header row is found by label: the row (3-5) whose col 5 reads `画面ID` gives the ID in
col 10 and the name in col 15 (row 4 is `ﾌﾟﾛｸﾞﾗﾑID` in 04_品質管理/05_受注出荷, so never read a fixed
`[4,10]`). Use the prebuilt list your prompt names; standalone, build it with
`python _shared/scripts/build_screen_list.py <...\01_Doc\08_機能定義書> <out.tsv>` (~9 min cold). Its
columns are `name`, `screen_id` (header), `sheet`, `workbook`, `sheet_id` (the ID in `画面設計書(<ID>)`).
**Key a screen by `sheet_id` when present**, not the header ID: copied sheets keep a stale header
(`PSJCO403` `GSJC403B`/`C`/`D` all carry `GSJC403A` in `[4,10]`), so keying by header hides the program's own
screens and W5 ends up proposing another program's ID (`GSJC404B`). A header ID ≠ sheet-name ID is
`naming-standard-compliance`'s finding, not W5's — but the header **name** beside a stale ID is stale
too (`GSJC403B`/`C`/`D` all read `社内加工ﾃﾞｰﾀ作成`), so for those sheets take the screen name from the
機能定義書 Ⅳ block heading of that letter (`PSJCO403` `[158,3]` `B.予定金額設定`, `[192,3]` `C.実績ﾌｧｲﾙ取込`,
`[206,3]` `D.予定金額設定取込`). Normalise the header name first: some end in `画面` (`受入入力画面`),
most don't, so match `name` + optional `画面`. The program's **own** screens count too
(`PSJCO602` `[52,4]` `棚卸用仕掛在庫ﾃﾞｰﾀ作成画面を表示する` → `(GSJC602A)`).
**The screen must be the object of the verb** — `…画面に遷移する` / `…画面に戻る` / `…画面を起動する` /
`…画面を表示する` / `…画面を開く` / `…画面に起動する` / `…画面へ遷移する`. Normalise full-/half-width
parentheses before matching (`承認ｸﾞﾙｰﾌﾟ（ﾕｰｻﾞｰ)登録画面` = `承認ｸﾞﾙｰﾌﾟ(ﾕｰｻﾞｰ)登録画面`). An abbreviated name
(`承認ﾌﾛｰ一覧画面` for `ﾜｰｸﾌﾛｰ承認ﾌﾛｰ一覧`) or a dropped `画面` (`…登録に遷移`) is 要確認 with the
candidate ID, not a W5 finding. `…画面に表示する` / `…画面で選択された…` put data *on* a screen and
are not this rule (`PSJCO602` `[145,4]` `…集計し、棚卸用仕掛在庫ﾃﾞｰﾀ作成画面に表示する。`). Don't scan
other sections: overview prose (Ⅰ．機能概要) and generic words (`画面ｸﾘｱ`, `呼出元画面`) are not this rule.
The 機能定義書 Ⅳ screen-block headings `A.<画面名>` / `B.<画面名>` (`PXJCO125` `[86,3]` `A.ﾛｯﾄ取消画面`)
are titles, not sentences — exempt them.
`ﾀﾞｳﾝﾛｰﾄﾞ画面` is excluded although `GSJA702A`'s header name is `ﾀﾞｳﾝﾛｰﾄﾞ`: the common download screen
is cited by program ID in the house wording (`【共通】ﾀﾞｳﾝﾛｰﾄﾞ(PSJAO702)を起動する`), and the bare
`ﾀﾞｳﾝﾛｰﾄﾞ画面の戻り値…` phrases are fixed template sentences. When a name matches both the program's own screen and another program's, prefer the own one.
`名称(ID)画面` (`取引先詳細(GXJA802B)画面に遷移`) carries the ID and passes; mention the house order
`名称画面(ID)` once at 低 at most. So does the prefix form `<画面ID>_名称画面` (`PXJCO161` `GXJC161B`
`[1825,24]` `GXJC161D_実績入力(積層入力)画面を起動する`, `GXJC161A` `[752,24]`
`PXJCO130_ﾛｯﾄﾄﾚｰｽ.GXJC130B_ﾛｯﾄﾄﾚｰｽ(作業実績)画面へ遷移する`). A **program** ID in the parentheses
(`機能定義書` `[213,4]` `流動停止解除画面(PXJCO134_流動停止解除)に遷移する`) is 要確認: give the candidate
screen IDs of that program from the screen list. Several names map to more than one ID
(`社内加工単価登録` GSJC401A/GSJC402A, `工順比較` GSJA401G/GXJC124B, `工程詳細` GSJA401C/GXJA243B) — give
the candidates, don't guess.
Measured in that scope: 405 mentions carry the ID, **310 don't, across 110 sheets** — a widespread
omission, so report it **once per sheet, 低**, listing the cells and the ID to append
(`PSJAO704` `機能定義書` `[70,4]` `承認ｸﾞﾙｰﾌﾟ一覧画面に遷移する` → `承認ｸﾞﾙｰﾌﾟ一覧画面(GSJA704G)`). When the
name matches more than one screen ID, give the candidates rather than guessing.

### W6 — 画面表示仕様の検索条件と結合 (No.21/22, call-out)

Sources: `レビュー観点` No.21 「表の検索条件の演算子（=、LIKE）があっているか」, No.22 「表の検索条件のコード系は
前方一致、品名の名称系は部分一致となっているか」, and the call-out on `画面設計書(GXJA802A)` `@r142c54`:
「結合条件に"INNER JOIN"と"LEFT JOIN"が混在する場合は"INNER JOIN"を上部に固めて記述し、以降に"LEFT JOIN"を記述すること」.
Scope: 画面設計書 `Ⅲ．画面表示仕様` only, live text only (a struck row is not a condition). Calibrated
2026-10-05 on the 93 工程管理 `PHASE1`-`PHASE3` workbooks that have a 画面設計書 (216 sheets, 1,223
検索条件 tables / 6,782 rows, 414 blocks with 結合条件) — re-derive every verdict from the dump.

**Anatomy.** A block opens with `参照ｴﾝﾃｨﾃｨ` (label col = the block's column, 4 at top level, 5/6 when
nested); the rows under it map an alias (col +8) to an entity (col +10) — `Y`=`画面`, `Z`=`ﾛｸﾞｲﾝ情報`,
others are tables or earlier queries (`(4)-②受入済取得`). A `検索条件` table has no header row: each
row is No. (label col), left side (label col +2, e.g. `A.取引先ｺｰﾄﾞ`), **operator**, right side, then
notes (`※1`, `OR`, `)`) further right. **An opening bracket sits one column left of the No.**
(`PSJCO309` `GSJC309B` `[132,3]` `((`, `PSJCO405` `GSJC405A` `[96,3]` `(`) — it does not end the table; read the row. Ignore the revision margin
right of the form (col ≥ 100, typically 105-116: `2024/2/9 平島 改訂履歴No.2 変更`, `2026/9/15 宮園
順番を変更`) — it is change history, not a note. **Find the operator column per table** as the column that holds
the comparison tokens (`=`, `LIKE`, `<>`, `>=`, `<=`, `IN`, `IS`, `IS NOT`); it is label col +20 in the
standard layout (`[83,4]` → `[83,24]`, right side `[83,28]`), but a nested block can shift it
(`PSJCO403` `GSJC403A` `[229,22]`), so never read a fixed column. A token may carry a leading space
(`' LIKE'`). House vocabulary, measured: `=` 5,476, `LIKE` 512 (always upper case in the tables),
`<>` 153, `<=`/`>=` 292 (ranges `Y.xxx(FROM)`/`(TO)`), `IN` 89, `NOT IN` 8, `IS [NOT] NULL` 23. **The
match mode is written as text after the screen item, not as wildcards**: `Y.取引先ｺｰﾄﾞ(前方一致)` (412
LIKE rows), `Y.取引先名(部分一致)` (70), occasionally `(完全一致)`; no 検索条件 cell writes `%` or
`'%'||…||'%'` (those appear only in pasted SQL samples to the right, e.g. `PSJCO403` col 54, which are
out of scope). An operator cell holding `※n` (`PXJCO128` `GXJC128B` `[315,24]` `※8`: the operator
depends on 停止区分) is defined by that footnote — judge the footnote, not the cell.
A `※n` on the row points to a footnote `※n.…` (or `※n …`) at or left of the block's label column, possibly one
shared by several blocks (`PSJCO403` `GSJC403A` `[534,5]` `※1 G1)の検索条件が指定された場合のみの条件` serves
col-6 blocks above it) — **it can sit far below**
(`PXJCO131` `GXJC131A` `[237]` → `[615,4]`), so resolve it to the next `※n` footnote line after the row,
not within a fixed window. Joins are written as their own tables, one per joined alias:
`結合条件(A LEFT JOIN B)` / `(A INNER JOIN B)` / `(A LEFT OUTER JOIN B)` / `結合条件 A INNER JOIN B` /
`(A,B LEFT JOIN C)` at the block's label column, each followed by numbered `=` rows; `(+)` notation is
not used.

**W6a — 検索条件の演算子・一致方法 (No.21/22).** "Screen search input" below means a row with the `画面`
alias (`Y.…`, `Y.G1).…`) on **either** side — `Y.G1).出荷日From <= A.出荷日` (`PSJCO403` `GSJC403A`
`[364,8]`/`[471,8]`) counts; "left side" below then means the non-`Y` side — **and** whose own note or resolved `※n` footnote says the
condition applies only when the value is entered (`値が入力されていた場合のみ検索条件に含む`,
`未入力の場合、検索条件から除外`, `…に入力がある場合のみ…`, `G1)の検索条件が指定された場合のみ…`). A `Y.` row
without such a note is usually a key lookup (self-check sample `GXJA802A` `[233]` `A.取引先ｺｰﾄﾞ = Y.取引先ｺｰﾄﾞ`, a
detail screen's own key, a `③品目ﾘｽﾄ取得` sub-query) where `=` is right — don't apply a/c/d to it.
Item type comes from the left side's column name: **名称** = ends in `名`/`名称`/`略式名`/`ｶﾅ`
(`作業者名`, `取引先名`, `選択肢名`); **品名-ID** = `KC品名`/`IS品名`; **code** = ends in `ｺｰﾄﾞ`/`No`/`GRP`/`ID`
(`品目ｺｰﾄﾞ`, `管理No`, `加工GRP`, `ﾜｰｸﾌﾛｰID`); `図番`/`型番` are neither.
- **a. `LIKE` with no match mode — 中.** A `LIKE` row (any right side) with no `前方一致`/`部分一致`/
  `後方一致`/`完全一致`, no `%`, no `先頭n桁` in the right side or its notes, and no footnote that states the
  match. Without a wildcard `LIKE` behaves as `=`, so either the operator or the missing `%` is wrong.
  Measured: 19 rows in 4 programs — `PSJCO404` `GSJC404A` `[126,24]` `A.ｷｰ3 LIKE "300"` (the 取得項目 reads
  `SUBSTR(A.ｷｰ3,4,2)`, so a prefix was meant: `"300%"`), `PSJCO403` `GSJC403A` `[163,26]` `A.ｷｰ3 LIKE
  ﾒﾆｭｰﾊﾟﾗﾒｰﾀ`, `PSJCO201` `GSJC201B` `[256,24]` `A.部門GRP LIKE Z.部門GRP` (every other block uses `=` for
  部門GRP), and `PSJCO309` `GSJC309B` `[134,24]`/`[135]`/`[138]`/`[139]` (again at `[278,25]`-`[283]`, and on
  `GSJC309C`) — `((A.KC品名 >= Y.KC品名(開始) AND <= Y.KC品名(終了)) OR A.KC品名 LIKE Y.KC品名(開始) OR
  … LIKE Y.KC品名(終了))`, same for 製造ﾛｯﾄNo: the LIKE branches only add anything with a `%`. Report one
  line per table, listing the rows. Not findings:
  a footnote that defines the match (`PXJCO131` `GXJC131A` `[237,24]` `LIKE Y.作業場ｺｰﾄﾞ ※1,※11` →
  `※1. 画面.作業場条件設定=ﾁｪｯｸ無しの場合は前方一致、ﾁｪｯｸ有りの場合は完全一致`), and `Y.品目ｺｰﾄﾞ(先頭6桁)`
  (`PSJAO401` `GSJA401B` `[590,24]`, a prefix of the input — 3 rows).
- **b. `=` with a match mode — 中.** `=` whose right side or note says `部分一致`/`前方一致`/`後方一致` or
  carries `%`: the operator contradicts the text. Measured: `PSJAO501` `GSJA501A` `[170,24]`
  `A.実績表項目ID = Y.実績表項目ID(前方一致)`. This takes precedence over d (one finding per row).
  **Exempt** a left side that is a function cutting the column to the input's span —
  `SUBSTR(A.製造ﾛｯﾄNo,TO_NUMBER(F.FROM),…) = Y.製造ﾛｯﾄNo(部分一致)` (`PXJCO134` `GXJC134A` `[140]`/`[141]`,
  `PXJCO128` `GXJC128B` `[323]`/`[324]`, `GXJC128C` `[110]`/`[111]` — 6 rows) is a correct `=`.
  **Exempt** too a right side whose name — `Y.` stripped, trailing spaces and `※n` dropped — is a live
  **画面項目名 in the same sheet's Ⅴ．画面項目定義**: the parenthesis is then part of the item's name (an
  input labelled after the 停止区分 `6:品目ｺｰﾄﾞ+製造ﾛｯﾄNo(部分一致)`), not the match mode of this row. The
  designer's answer (2026-10-06): `PXJCO128` `GXJC128A` `[135,24]`/`[136,24]` `A.製造ﾛｯﾄNo =
  Y.製造ﾛｯﾄNo(部分一致)` / `(部分一致)2` against Ⅴ `[477,5]`/`[478,5]` `製造ﾛｯﾄNo(部分一致)` / `…2` is not to be
  reported. Apply
  this exemption to b only: rows a and d still read the parenthesis as the match mode
  (`PSJCO304` `GSJC304A` `[138,24]`, since corrected to live `LIKE Y.製造ﾛｯﾄNo(部分一致)` — `=` and the `1`
  struck — where Ⅴ `[495,5]` is `製造ﾛｯﾄNo(部分一致)`, must not turn into an a finding; it stays the d
  要確認 it was before). `PSJAO501`'s `実績表項目ID(前方一致)` is not a
  Ⅴ name (Ⅴ has `実績表項目ID`), so it stays a finding.
- **c. 名称/品名 search input not 部分一致.** Screen search inputs of 名称 type use `LIKE …(部分一致)` in 26
  rows (13 programs) and never `=` — a `=` there is **中**; a `LIKE …(前方一致)` is **要確認** (1:
  `PXJCO131` `GXJC131A` `[771,24]` `A.枠名 LIKE Y.枠名(前方一致)`). For **品名-ID** (`KC品名`/`IS品名`) the house
  is 27 部分一致 (20 programs) against 7 `=` (3 programs: `PSJCO403` `GSJC403A` `[369,26]`/`[476,26]`,
  `PSJCO404` `GSJC404A` `[270,25]`/`[345,25]`/`[618,24]`, `PXJCO128` `GXJC128A` `[313,24]`, `GXJC128B` `[206,24]`) —
  `=` is **要確認**, citing No.22, because `KC品名` behaves like an identifier here (upper-cased, `半英数記号`;
  see W1a/W1f) and an exact match may be deliberate.
- **d. code search input with `LIKE …(部分一致)` — 要確認.** No.22 says 前方一致; the corpus has none
  outside `図番`/`型番`, which always use 部分一致 (`PSJAO203` `GSJA203A` `[201]` `KC図番`, `PSJAO241`
  `GSJA241A` `[282]`/`[285]`, `PXJAO701` `GXJA701A` `[250]` `ﾓﾃﾞﾙNo/型番` — not findings). Judge the mode
  given for this row only — a footnote that lists several cases (`停止区分"6:品目ｺｰﾄﾞ+製造ﾛｯﾄNo(部分一致)"`)
  is not a 部分一致 on 品目ｺｰﾄﾞ.
- **e. Not an SQL operator — 低.** `=>` (15 rows, 10 sheets, 6 programs — all `A.停止日時 => ｼｽﾃﾑ日時` in
  the `(A.停止日時 = NULL OR …)` pair copied from one template — its `= NULL` half is `xlsx-db-column-check`
  S1 (`IS NULL`); W6e reports only the `=>` cell and S1 only the `= NULL` cell, each naming the other: `PSJAO241` `GSJA241B-基本情報` `[330,24]`,
  `PSJAO401` `GSJA401B` `[473,24]`, `PSJCO401`, `PSJCO402`, `PXJAO701` `GXJA701B` `[275,24]`, `PXJCO101`
  `GXJC101C` ×6). Report once per sheet, fix `>=`. Same for `=<`, `==` or a full-width `＝`/`＜` if met.
  A blank operator with both sides filled occurs 0 times; if you meet one it is 中.
- **Not reported: a code search input with `=`.** The self-check says code = 前方一致, but the corpus
  splits 251 `LIKE …(前方一致)` (36 programs) against 88 `=` (19 programs) — codes picked from a zoom or
  a ComboBox (`層No`, `依頼元事業所ｺｰﾄﾞ`, `ｸﾞﾙｰﾌﾟID`) are compared exactly by design. Raising No.22 here
  would be ~90 false findings; mention it only if the designer's own screens mix both for the same item.

**W6b — 結合の順序 (call-out).** Within one block (from its `参照ｴﾝﾃｨﾃｨ` to the next `参照ｴﾝﾃｨﾃｨ` at the
same or a smaller column), read the `結合条件` labels **at the block's own column** in row order —
a nested child block's labels are its own sequence — and take the join type from each label
(`INNER`, `LEFT`/`LEFT OUTER`). **Finding — 低: an `INNER JOIN` label after any `LEFT JOIN` label**;
report the first offending label with the block's sequence (`I L I`). **Skip an INNER join onto a
LEFT-joined alias** (`PSJCO309` `GSJC309A` `[623,5]` `結合条件 C3 INNER JOIN D1` where `C3` came in by LEFT
JOIN; `[958,4]` `B INNER JOIN A2` — 13 such labels, all `PSJCO309`): it cannot move above the join it
depends on — but keep reading the block, a later INNER on the driving alias still counts.
Measured: 51 of 414 join blocks mix both types; 38 comply, **13 violate (9 sheets, 7 programs)** —
`PSJAO203` `GSJA203B` `[480,4]` `(A INNER JOIN F)` after four LEFT, `PSJCO201` `GSJC201B` `[286,4]`
(`I L I L L L L`), `PSJCO307` `GSJC307B` `[205,5]`, `PSJCO604` `GSJC604D` `[65,4]`, `PXJCO122` `GXJC122A`
`[368,4]`, `PSJCO302` `GSJC302B` `[77,4]` whose INNER join is conditional
(`【(2).内外作区分="1"(外注)の場合、結合】` — still a finding; say the condition can move with it), and 7 in
`PSJCO309`: three blocks repeated on `GSJC309A`/`B`/`C` — `[661,5]`/`[204,4]`/`[203,4]` `A INNER JOIN E1`
after the skipped chained ones, `[829,5]`/`[372,4]`/`[371,4]` `(A INNER JOIN D)` — plus `GSJC309A` `[962,4]`
`(A1 INNER JOIN D1)`. A block repeated on several screens is one finding with the cells listed.
Also exempt: `結合条件(CONNECT BY)` (2), and a bare `結合条件` followed
by `なし` (3). Worth a separate line when you see it: a label joining the driving alias to itself
(`PXJCO122` `GXJC122A` `[368,4]` `結合条件(A INNER JOIN A)` with `A.管理No = A.管理No`) — the right alias is
wrong (中, a `design-doc-internal-consistency` matter, but report it here if nothing else will).
Joins described only inside an entity name (`作業場ﾏｽﾀ(万一ﾏｽﾀとの紐付け切れた時の為LEFT_JOIN)`) or in pasted
SQL samples are out of scope.

### W7 — ズーム名・ズームIDが 03.ｽﾞｰﾑ管理台帳 と一致 (my-review-skills 「ズーム名・IDが台帳と違う」; 回答 20261010 No.54/55)

Source: 「Ⅳ・Ⅴ・Ⅵ に書くズームの名前・ズームIDが `03.ｽﾞｰﾑ管理台帳*.xlsx` の名称・ID と完全に一致しているか」 (adopted
2026-10-17). Calibrated on the 122 工程管理 `PHASE1`-`PHASE3` workbooks (live text); re-derive every verdict.

**The ledger.** The prebuilt zoom list your prompt names; standalone, build it with
`python _shared/scripts/build_zoom_list.py <...\01_Doc\04_共通設計> <out.tsv>` (~12 s, zip only; `--check` for freshness).
It reads the visible `ｽﾞｰﾑ一覧*` sheets of `03.ｽﾞｰﾑ管理台帳.xlsx` (共通: `ｽﾞｰﾑ一覧【XJx】`/`【SJx】`), `…_基準情報.xlsx` and
`…_工程管理.xlsx` (個別: `ｽﾞｰﾑ一覧【XJ】`/`【SJ】`) — never `90_JAGUR各管理台帳\`, never the hidden `旧ｽﾞｰﾑ管理台帳(共通)` /
`ｽﾞｰﾑ一覧（共通、個別）` / `WORK…` sheets — finding the header by its `ｽﾞｰﾑID` label, live rows only (103 IDs, 128 rows in
2026-10). Rules for the dictionary:
- Live rows only: `ZSJA025` was `承認ﾛｰﾙｽﾞｰﾑ` (struck, 共通 `【SJx】` `[27]`) and is now `処置指示Noｽﾞｰﾑ（ﾘｯﾁ）` in `_工程管理`.
- An ID may have several live rows (the 共通 ledger repeats a changed row below the original, `改訂履歴No.n` in col 61) and
  several ledgers (`ZXJA018` is in all three). Accept any live `ｽﾞｰﾑ名称` of the ID. `ZXJA040` has two: `品目ｽﾞｰﾑ` `[30]` and
  `品目ｺｰﾄﾞｽﾞｰﾑ` `[31]` (changed 2024/1/9; the corpus uses the latter 42 times).
- Normalise full-/half-width parentheses and whitespace, including a line break inside the name
  (`ﾃｰﾌﾞﾙ一覧ｽﾞｰﾑ(汎用ﾏｽﾀ)⏎(ZSJA020)`).
- `ｽﾞｰﾑ名称` is the authority, not `表示名`: the ﾘｯﾁ twins share a 表示名 with their XJ twin (`ZSJA028` `品目詳細ｽﾞｰﾑ(ﾘｯﾁ)` /
  表示名 `品目詳細ｽﾞｰﾑ` = `ZXJA005`'s 名称; also `ZSJA029`/`ZXJA029`, `ZSJA025`/`ZXJA035`, `ZSJA046`/`ZXJA030`).
- Some names do not end in `ｽﾞｰﾑ` (`ZSJA040` `返品情報検索`, `ZSJA041` `出荷情報検索(ﾛｯﾄﾄﾚｰｽ)`), so never extract a name by
  "…ｽﾞｰﾑ(".

**Where zoom IDs appear.** Only in Ⅳ — 画面設計書 `Ⅳ．画面項目ｲﾍﾞﾝﾄ詳細` and 機能定義書 `Ⅳ．機能処理概要`. Ⅴ and Ⅵ carry the
**button** name (`品目ｺｰﾄﾞｽﾞｰﾑ` opening `品目詳細ｽﾞｰﾑ(ﾘｯﾁ)`), which is not the zoom's name and is not this rule. Three shapes:
- **(A) inline `名称(ID)`** in 処理内容 — `PSJAO203` `GSJA203A` `[369,20]` `KC品名ｽﾞｰﾑ(ZXJA006)を起動する` (607 in 78 books,
  592 exact).
- **(B) the `※1.ｽﾞｰﾑ起動ﾊﾟﾗﾒｰﾀ` table** — heading `①KC品名ｽﾞｰﾑ`, then `ｽﾞｰﾑID` | `ZXJA006` or `"ZXJA006"` (`GSJA203A`
  `[386]`-`[389]`; 285 in 57 books). The heading is the button name (`②停止指示者ｽﾞｰﾑ` → `ZXJA031` 作業者ｽﾞｰﾑ), so **never
  compare a table heading with the ledger name**.
- **(C) a heading `ID:名称`** (`PSJCO308` `GSJC308A` `[1049,5]` `① ZXJA030:工程ｽﾞｰﾑ`) — compare like (A).

**Findings.**
- **a. ID not in the ledger — 中.** 0 in the corpus.
- **b. Table ID ≠ the Ⅳ event's ID — 中.** For each (B) table, find the Ⅳ row whose 画面項目 equals the heading and compare
  its inline (A) ID: 223 agree, **5 disagree** — `PSJCO402` `GSJC402B` `[123,13]`/`[127,13]` `"ZSJA043"` (ﾃｰﾌﾟﾛｯﾄNoｽﾞｰﾑ)
  while Ⅳ `[118,24]`/`[119,24]` launches `加工部門ｺｰﾄﾞｽﾞｰﾑ(ZSJA032)`; the same in `PSJCO404` `GSJC404A` `[1649,13]`/`[1653,13]`
  (hidden rows — `【非表示】`); `PSJAO241` `GSJA241B-基本情報` `[1551,11]` `①KC品名ｽﾞｰﾑ` → `ZSJA035`. With no matching Ⅳ row
  (57), report only when the heading's subject plainly contradicts the ID's ledger name: `PSJCO403` `GSJC403A`
  `[1540,5]`/`[1542,13]` and `PSJCO404` `GSJC404A` `[1585,13]` — `作業場ｺｰﾄﾞｽﾞｰﾑ` with 引数 `G1).作業場ｺｰﾄﾞ` but `"ZXJA031"`
  (作業者ｽﾞｰﾑ; Ⅳ uses `作業場ｽﾞｰﾑ(ZXJA029)`) — 中. Qualified button names are not that (`停止指示者ｽﾞｰﾑ`→作業者ｽﾞｰﾑ,
  `送付先作業場ｽﾞｰﾑ`→`作業場(納品用)ｽﾞｰﾑ`, `加工期限通知先選択ｽﾞｰﾑ`→ｸﾞﾙｰﾌﾟ選択ｽﾞｰﾑ).
- **c. (A)/(C) name ≠ the ledger 名称 of its ID.** Judge by **suffix match** on the normalised text before `(ID)` (balanced
  parentheses kept: `品目詳細ｽﾞｰﾑ(ﾘｯﾁ)(ZSJA028)`): (1) if one of the ID's own live 名称 is a suffix, it agrees — a qualifier
  or sentence in front is not reported (`[513,23]` reads 「…画面.管理Noをもとにｵｰﾀﾞｰ工程属性ｽﾞｰﾑ」, `PSJCO306` `[589,24]`
  「"1"(有効のみ)をもとに資源ｽﾞｰﾑ」); (2) else if another ID's 名称 is a suffix (take the longest), the written name is that
  名称 — **中**; (3) else the written name is the text after the last `をもとに`/`、`/`。`/space — **低**.
  **中** when the written name is the 名称 of **another** ID — name or ID is wrong:
  `PSJCO204` `GSJC204A` `[638,20]` `管理Noｽﾞｰﾑ(ZSJA042)` (ledger `管理Noｽﾞｰﾑ(不良ﾛｯﾄ統合用)`; `管理Noｽﾞｰﾑ` is `ZXJA012`);
  `PXJCO123` `GXJC123A` `[513,23]` `ｵｰﾀﾞｰ工程属性ｽﾞｰﾑ(ZXJA034)` (`ZXJA011`'s name; ledger `ｵｰﾀﾞｰ工程属性(共通)ｽﾞｰﾑ`). **低**
  otherwise — give the ledger name: `PSJCO205` `GSJC205A` `[385,24]`/`[387,24]` `返品情報検索ｽﾞｰﾑ(ZSJA040)`; `PXJCO164`
  `GXJC164A` `[591,20]` `汎用GRPｽﾞｰﾑ(ZSJA021)` (→ `汎用工程GRPｽﾞｰﾑ`); `PXJCO122` `GXJC122A` `[638,20]`
  `ｵｰﾀﾞｰ工程属性ｽﾞｰﾑ(共通)(ZXJA034)`; `PSJCO308` `GSJC308A` `[1038,24]` `共通ｽﾞｰﾑ(焼成管理No一覧(ZSJA048))`. A name that
  matches only the ID's 表示名, while being its XJ twin's 名称, is 低, "add `(ﾘｯﾁ)`": `PSJCO309` `GSJC309A` `[1227,24]`
  `作業場ｽﾞｰﾑ(ZSJA029)`, `PSJAO401` `GSJA401A` `[267,20]` `処置指示Noｽﾞｰﾑ(ZSJA025)`. A qualifier prefix with the ledger name
  intact (`PXJCO161` `GXJC161B` `[1833,24]` `次工程作業場ｽﾞｰﾑ(ZXJA029)`) is not reported.
- **d. Placeholder `XXXXｽﾞｰﾑ(XXXXXXX)` — 中** (the zoom has not been chosen; the sheet's call-outs often ask the question,
  `PSJCO802` 「依頼元部門のｽﾞｰﾑは？」): 8 cells in 3 programs — `PSJCO802` `GSJC802A` `[140,24]`/`GSJC802B` `[225,24]`,
  `PSJCO805` `GSJC805A`/`B`/`C`/`E`/`F`, `PSJCOA06` `GSJCA06B` `[226,24]`. One line per program.
- **Not findings:** using only some of the ledger's 起動ﾊﾟﾗﾒｰﾀ/戻り値 (回答 20261010 No.54/55: `ﾃｰﾌﾟﾛｯﾄNoｽﾞｰﾑ(ZSJA043)` given
  only ﾃｰﾌﾟﾛｯﾄNo is normal); a GX screen calling a ZS zoom or the reverse (414 GS→ZX, 33 GX→ZS in the corpus); `ｽﾞｰﾑｲﾝ`/`ｱｳﾄ`,
  `F9:ｽﾞｰﾑ`, `ｽﾞｰﾑ画面で選択された値`, the 表示 `Z`, and edit masks such as `ZZZ9` (the ID regex is `Z[A-Z]{3}\d{3}`). The
  LinkLabel rule for Z programs themselves concerns the ledger's own design sheets — out of scope here.
Report b and d one line per sheet/program, c one line per sheet listing the cells. A zoom with no Ⅳ event at all stays
`design-doc-internal-consistency`'s finding.

### W8 — 他機能の名前は呼び出し先の機能定義書ファイル名どおり (課題152-1041・1046・1047・1050, 課題165-168)

Source: my-review-skills 「他の機能（ｻﾌﾞﾌﾟﾛ・画面・バッチ）の名前を書くときは、呼び出し先の機能定義書のファイル名
（`<機能ID>_<機能名>.xlsx` の機能名。例: `ﾛｯﾄ分割処理(ｻﾌﾞﾌﾟﾛ)`）と完全に一致させているか」 — abbreviations such as
`ﾛｯﾄ分割処理ｻﾌﾞﾌﾟﾛ` or `処置指示書発行(ｻﾌﾞﾌﾟﾛ)` are findings, and once one is found every occurrence (I/O関連図, 機能処理概要,
other sheets) is checked (adopted 2026-10-17). Calibrated on the 122 工程管理 books and validated on `PXJCO123`
(reproduced) and `PSJCO306` (no finding, none false); re-derive every verdict.

**The program list.** `機能ID → 機能名` from the **file names** of every workbook under `01_Doc\08_機能定義書` (all WGs), in
`build_screen_list.py`'s scope — the list your prompt names; standalone,
`python _shared/scripts/build_program_list.py <...\01_Doc\08_機能定義書> <out.tsv>` (~1 s, opens no workbook; 287 programs
in 2026-10; `DUP-ID`/`UNPARSED` lines are not findings). Name = the file name after the first `_` (04_品質管理 carries a
branch: `PXJDO101_010_規格登録` → key `PXJDO101_010`). `【共通】` is part of the name (`【共通】ﾀﾞｳﾝﾛｰﾄﾞ(PSJAO702)` is exact).

**Scope.** Live text of every design sheet (hidden ones too, findings 【非表示】 — except leftover/withdrawn hidden
sheets, `agent-guide.md`) plus its **shape text** (I/O関連図 and other drawings: the `<a:t>` runs of
`xl/drawings/drawingN.xml` for the sheet, struck runs dropped — map the sheet to its drawing as in check 13). Join a shape's
runs and compare after removing line breaks, spaces and full/half-width differences in parentheses (`ﾛｯﾄ分割処理⏎(ｻﾌﾞﾌﾟﾛ)` is
exact). Excluded:
`表紙`, `改訂履歴`, `画面ｲﾒｰｼﾞ`, `詳細設計*`, and the revision margin (col ≥ 100 — not col 59: ﾁｪｯｸ処理設計書 keeps
ﾁｪｯｸ詳細 in col 59 and ﾁｪｯｸ補足 in 86, and 更新条件表 side blocks run to 86). The program's **own** name in its header/表紙 is
not this rule (`naming-standard-compliance`); its own Ⅴ．ﾌﾟﾛｸﾞﾗﾑ構成 row is, as e below.

**Reference shapes and findings.**
- **a. ID-anchored — `名称(ID)` / `【ID_名称】` / an ID cell followed by a name cell** (316 exact). Compare the adjacent name
  with the list name for that ID; prefix and suffix matches count as exact (`…をｾｯﾄしﾛｯﾄｶｰﾄﾞ発行(ｻﾌﾞﾌﾟﾛ)(SXJCB141)` — the name
  ends at the `(`; `【SXJCB122_ﾛｯﾄ分割処理(ｻﾌﾞﾌﾟﾛ)】参照` — it starts after the `_`). Strip a leading step number (`3-1)`,
  `(4)`, `②`) first. A different name with the right ID is **低** (the reader can still find the callee); group per sheet
  and callee, listing the cells. Measured: 28 cells in 10 programs — `PXJCO122` `[58,4]` `ﾛｯﾄ分割処理ｻﾌﾞﾌﾟﾛ(SXJCB122)`, the same
  in `PXJCO123` `[75,4]`/`[183,21]`/`[392,5]`, `PSJCO307` `[307,19]`, `PSJCO308` `[353,5]`, `PXJCO161` `[419,5]`/`[628,6]`/
  `[776,19]`; `ﾛｯﾄ統合処理ｻﾌﾞﾌﾟﾛ(SXJCB123)` ×6 in `PXJCO123`; `実績表項目ﾃﾞｰﾀ作成(処理)ｻﾌﾞﾌﾟﾛ(SXJCB178)` (file:
  `実績表ﾃﾞｰﾀ作成処理(ｻﾌﾞﾌﾟﾛ)`) in `PXJCO123` `[168,4]`/`[336,4]`/`[407,5]`/`[32,16]` and `PXJCO124` `[139,4]`/`[36,15]`, where a
  struck `ﾁｪｯｸｼｰﾄ` shows a half-finished rename — say so; `処置指示書発行(ｻﾌﾞﾌﾟﾛ)(SXJCB147)` in `PXJCO129` `GXJC129A` `[507,24]`,
  `PXJCO130` `GXJC130A` `[211,20]`, and `PXJCO123` `[30,16]` 「【SXJCB147_処置指示書発行】参照」 (no `(ｻﾌﾞﾌﾟﾛ)`); `ﾛｯﾄｶｰﾄﾞ発行ｻﾌﾞﾌﾟﾛ` `PSJAO401` `[146,5]`/`[314,5]`;
  `識別ｶｰﾄﾞ(共通)発行ｻﾌﾞﾌﾟﾛ` `PXJCO161` `[942,5]`; `PXJCB101` `[23,16]` `実績表ﾃﾞｰﾀ作成処理` (no `(ｻﾌﾞﾌﾟﾛ)`).
- **b. ID not in the list.** **中** when it is a typo or truncation — give the candidate found by name: `PSJCO201` `GSJC201B`
  `[435,24]` `識別ｶｰﾄﾞ(共通)発行(ｻﾌﾞﾌﾟﾛ)(SSJCB5072)` → `SSJCB507`; `PSJCO504` `[16,3]` `(PXJCO20)送品案内入力` → `PSJCO501`.
  **要確認** when no 機能定義書 exists yet: `SXJCB184` `QA判定実績登録(ｻﾌﾞﾌﾟﾛ)` in `PXJCO161` `[38,4]`/`[810,5]` and `PXJCO904`
  `[143,4]`/`GXJC904A` `[609,24]`.
- **c. No ID, but marked as a program — `Xｻﾌﾞﾌﾟﾛ`, `X(ｻﾌﾞﾌﾟﾛ)`, `X(ﾊﾞｯﾁ)`.** Normalise both sides by deleting `【共通】`,
  `(ｻﾌﾞﾌﾟﾛ)`, `ｻﾌﾞﾌﾟﾛ`, `(ﾊﾞｯﾁ[処理])`, `処理`, `ﾃﾞｰﾀ`, `書`, `項目`, `ﾏｽﾀ`. A text whose normalised form equals a listed name's
  while the text itself differs is **低**, with the exact name: `PSJCO201` `[94,17]`/`PSJCO203` `[85,17]` `ﾛｯﾄ作成ｻﾌﾞﾌﾟﾛ`;
  `PSJCO308` `[543,6]` `ﾛｯﾄ分割ｻﾌﾞﾌﾟﾛ`; `PXJCO161` 更新条件表 `TXJCD204` `[34,2]`/`[35,2]`, `TXJCM007` `[55,2]`/`[56,2]`/`[108,2]`/
  `[109,2]`, `TXJCM057`, `TXJCM058` (one line per sheet); `PXJCO123` `[284,4]` 「(ﾛｯﾄ統合ｻﾌﾞﾌﾟﾛで実行)」, `[309,8]`/`[333,8]`
  「(ﾛｯﾄ分割処理ｻﾌﾞﾌﾟﾛにて更新)」 (a mark inside a sentence counts); I/O関連図 `PXJCO123` shapes `実績表項目⏎作成処理⏎(ｻﾌﾞﾌﾟﾛ)`.
  Match a `ｻﾌﾞﾌﾟﾛ`-marked text only against `S…B` IDs and a `(ﾊﾞｯﾁ)`-marked one only against `…B` IDs, never against the
  program's own ID (`PXJCO123`'s own name normalises to `ﾛｯﾄ統合` too). A normalised form one character away from a listed
  name (a dropped or wrong kana, a missing ﾞ/ﾟ) is **低 (誤記)** with the exact name: `PXJCO123` `GXJC123A` `[491,23]`
  `ﾛｯﾄｶｰﾄ発行(ｻﾌﾞﾌﾟﾛ)` → `ﾛｯﾄｶｰﾄﾞ発行(ｻﾌﾞﾌﾟﾛ)`. Do **not**
  run this name-only matching on unmarked words: it fires on process and data words (`ﾛｯﾄ分割処理に関しては`,
  `処置指示発行ﾃﾞｰﾀﾏｽﾀ`, `流動停止`, `仕掛在庫照会ﾃﾞｰﾀ` — several hundred false hits).
- **d. The callee's own header disagrees with its file name — 要確認, reported here** (naming checks only the
  program under review): `PXJCB101`/`PXJCB102` `[4,15]` 「製造ｵｰﾀﾞｰ反映(ﾊﾞｯﾁ処理)」/「製造ｵｰﾀﾞｰ完了処理(ﾊﾞｯﾁ処理)」;
  callers copying it are the same line, not a separate finding.
- **Not this rule:** `…画面(PID)` (`着手入力画面(PXJCO904)` — W5's); an item list keyed by the callee (`PXJCO163` `GXJC163B`
  `[964,5]`-`[975,5]` `作業場ｺｰﾄﾞ(PXJCO161情報)`); generic words (`帳票出力ｻﾌﾞﾌﾟﾛ`, a bare `ｻﾌﾞﾌﾟﾛ`, `ﾊﾞｯﾁ処理(PSJCB903)`); a name
  used as a modifier inside a longer noun (`(23)処置指示発行ｻﾌﾞﾌﾟﾛ呼出処理無効ﾌﾗｸﾞ取得`).
- **e. The program's own name with another program's ID — 中** (a stale copy-source ID in Ⅴ．ﾌﾟﾛｸﾞﾗﾑ構成, reported
  here as 「Ⅴ．ﾌﾟﾛｸﾞﾗﾑ構成の機能IDが流用元のまま」): `PSJCO401` `[103,5]` `PSJAO401 | 社内加工単価登録`, `PSJCO405` `[168,5]`
  `PSJAO405`, `PXJCB131` `[145,5]` `PXJBO131`, `SXJCB123` `[161,5]` `SXJCO123`, `PSJCB807` sheet `機能定義書(PSJCB802)`
  `[139,5]`.

## Reporting

Group by check (W1-W8), then by sheet. Every rule here is low severity on its own unless the check
says otherwise (a duplicate/misfiled circled section that another doc cites by number is 中; a
placeholder ID such as `XXXXXXX` is 中). When one deviation repeats across a whole sheet, report it
once per sheet with the cells listed, as each section says — these rules fire widely and a row-by-row
list buries the useful findings.
