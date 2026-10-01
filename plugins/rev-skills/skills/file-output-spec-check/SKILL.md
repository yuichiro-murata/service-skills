---
name: file-output-spec-check
description: Review a program's ﾌｧｲﾙ出力仕様書(<ﾌｧｲﾙID>) sheets — the file/download output specification that no other REV skill reviews in its own right. Checks that the ﾌｧｲﾙID agrees everywhere it is written (sheet name, 3C/3D header, Ⅰ's 出力ｺｰﾄﾞ argument, 機能定義書 Ⅲ．入出力定義 and its 【ﾌｧｲﾙ出力仕様書(…)】 pointers), that the file is registered in 06-04_ﾌｧｲﾙ一覧_<WG>.xlsm under this program, that Ⅱ．ﾌｧｲﾙ出力仕様's fixed rows are filled in, that a 共通ﾀﾞｳﾝﾛｰﾄﾞ file carries its three standard annotations, that every ﾍｯﾀﾞｰ item has a 画面項目ID in the right number band and registered in 82.画面項目辞書, that the ﾍｯﾀﾞｰ and 明細 item lists agree, that every 明細 reference resolves to a real 画面設計書 block / 参照ﾃｰﾌﾞﾙ alias, that 編集方法 uses the common-design vocabulary and matches the value's format, that No is not output, and — when a ﾌｧｲﾙﾚｲｱｳﾄ workbook exists — that the spec agrees with it. Distinct from report-design-check (帳票設計書 only; explicitly does not read this sheet) and xlsx-db-column-check (column existence). Use when the user asks to REV ﾌｧｲﾙ出力仕様書/ﾀﾞｳﾝﾛｰﾄﾞﾌｧｲﾙ/CSV・TSV出力の仕様. For a full REV without named checks the entry point is `rev-program-review`; run this standalone only when selected there or asked for by name.
---

# file-output-spec-check

Reviews every visible `ﾌｧｲﾙ出力仕様書(<ﾌｧｲﾙID>)` sheet in one program's design-doc workbook. Until
this skill existed, nothing reviewed this sheet at all: `report-design-check` explicitly refuses it,
`xlsx-db-column-check` lists it as out of scope for 桁数, and the self-check workbook
`01_Doc\08_機能定義書\11_工程管理\XXXXX000_ｾﾙﾌﾁｪｯｸ用.xlsx` carries a whole block of review points for it
(レビュー観点 No.65-69 plus the call-outs on its `ﾌｧｲﾙ出力仕様書(FXJA002)` sample sheet). This skill
is the automatable part of that block.

**A program with no visible `ﾌｧｲﾙ出力仕様書(*)` sheet is not in scope.** Say so in one line and stop.
Hidden ones are out of scope too (the shared dump's default): across the project most hidden copies
are untouched template stubs named `ﾌｧｲﾙ出力仕様書(FXJDXXX)` / `ﾌｧｲﾙ出力仕様書()`, or withdrawn ones
renamed `削除)ﾌｧｲﾙ出力仕様書(…)`. Also out of scope: sheets that merely share the prefix but are not a
spec — `ﾌｧｲﾙ出力ｲﾒｰｼﾞ`, `ﾌｧｲﾙ出力のｲﾒｰｼﾞ→` (mock-ups) — and a `…_様式` companion sheet.

## Environment

**Read `_shared/agent-guide.md` first** — scope selection, environment, the live dump, reading the
画面項目辞書 index (F5), reporting conventions. Read the `06-04_ﾌｧｲﾙ一覧_*` registry and any ﾌｧｲﾙﾚｲｱｳﾄ
workbook **live** with `_shared/scripts/live_dump.py` (registries carry struck rows; the COM cache
keeps them).

## Sheet anatomy (measured, not assumed)

Measured on 2026-09-30 across 63 visible `ﾌｧｲﾙ出力仕様書(*)` sheets in 44 workbooks under
`01_Doc\08_機能定義書\11_工程管理\PHASE1`-`PHASE3` (plus `PXJCO130`'s three, which openpyxl cannot
open — see `agent-guide.md`'s Environment). Anchor on the **labels**; the column numbers are the common case, not a contract.
Every count quoted below is from that one snapshot of actively-edited workbooks — re-derive a verdict
from the current dump, never cite a count or an example here as a live finding.

```
[1,1]=3C [1,5]=ﾌｧｲﾙ出力仕様書 … [1,53]=3D [1,57]=ﾌｧｲﾙ出力仕様書        ← duplicated header (3C / 3D)
[3,5]=ﾌﾟﾛｸﾞﾗﾑID [3,10]=PSJCO901 [3,15]=損金一覧
[4,5]=ﾌｧｲﾙID    [4,10]=FSJC051  [4,15]=損金一覧ﾃﾞｰﾀ   [4,62]=FSJC051 [4,67]=損金一覧ﾃﾞｰﾀ
[6,2]=Ⅰ．ﾌｧｲﾙ出力条件
  [8,4]=画面設計書(GSJC901A) - (6)明細情報取得 参照 … ①    ← the 明細 source, named ① (or (n))
  [9,5]=※上記①のｿｰﾄ項目・出力順をﾀﾞｳﾝﾛｰﾄﾞ画面の引数に設定し、…   ← 共通ﾀﾞｳﾝﾛｰﾄﾞ sort note
  [11,4]=(1)ﾌｧｲﾙ出力先情報取得 / [12,4]=※05.ｼｽﾃﾑ共通設計書.ﾌｧｲﾙ出力制御 参照
  [14,4]=No [14,6]=引数 / [15,4]=1 [15,6]=出力ｺｰﾄﾞ [15,21]='FSJC051'   ← must equal the ﾌｧｲﾙID
[18,2]=Ⅱ．ﾌｧｲﾙ出力仕様   ← fixed label list in col 4, value in col 18 — see F3
[28,2]=Ⅲ．編集仕様
  [29,3]=1.ﾍｯﾀﾞｰ          ← section title in col 3 (also 1-1.ﾍｯﾀﾞｰ, 1.ｷｰ情報, (2)ﾍｯﾀﾞｰ in col 4 …)
    [30,5]=No. [30,7]=出力項目名 [30,17]=出力内容                       ← two-row header
    [31,17]=参照先 [31,23]=項目名・出力値 [31,45]=編集方法 [31,54]=画面項目ID
    [32,5]=1 [32,7]=損金区分 [32,17]=- [32,23]="損金区分" [32,45]=- [32,54]=SJC8128
  [77,3]=2.明細           ← same table without the 画面項目ID column
    [80,5]=1 [80,7]=損金区分 [80,17]=① [80,23]=損金区分 [80,45]=文字列 ※1
  [125,3]=3.ﾌｯﾀｰ          ← usually a single "編集無し" row
```

Parsing rules that matter:

- **The item table has a two-row header** (`No./出力項目名/出力内容`, then
  `参照先/項目名・出力値/編集方法/画面項目ID`). Anchor on the second row for the value columns.
- **`出力項目名` spans merged cells** — take the first cell (col 7). Merge repetition is not a
  finding here.
- **A cell value can contain a newline** (`\n収縮率 長辺` in `PSJAO203`'s FSJA008). Group tokens by the
  `[r,c]` inside them, never by physical line — see `agent-guide.md`'s "Parsing a dump line" — and strip
  leading/trailing whitespace from item names **only for comparison**, never when quoting.
- **Section titles vary**: `1.ﾍｯﾀﾞｰ`/`2.明細`/`3.ﾌｯﾀｰ` is the common case, but a multi-part file uses
  `1) 送品案内` → `1-1.ﾍｯﾀﾞｰ`/`1-2.明細`, and an Excel file with blocks uses `1.ｷｰ情報` →
  `(1)段落ﾀｲﾄﾙ`/`(2)ﾍｯﾀﾞｰ`/`(3)明細`. Pair a ﾍｯﾀﾞｰ table with the 明細 table **of the same part**, never
  across parts.
- **A table's end** is the first row that is neither a numbered item (`No.` column holds a number)
  nor its continuation line — typically a `※` note or the next section title. A `※` note row can
  itself carry `画面項目ID` / an ID in col 54 (`PXJCO128` `FXJC004` `[84,54]`/`[85,54]=SJC8478`, the
  multilingual literal "即停止" that note outputs). That is a legitimate ID for the note's literal —
  it does not make the 明細 table a ﾍｯﾀﾞｰ, and it is not a table row.

## Procedure

### 1. Get the dump, and enumerate the file specs

From the shared dump, list every visible `ﾌｧｲﾙ出力仕様書(<ID>)` sheet. You will also need, from the
same dump: the `機能定義書(*)` sheet (F1) and every `画面設計書(*)` sheet (F7).

### 2. Read the ﾌｧｲﾙ一覧 registry for each file ID's WG

Route by the **file ID's own JOBコード** (the three letters after `F`), exactly as
`report-design-check` routes 帳票ID:

| File ID prefix | Registry under `01_Doc\06_システム設計書（一覧、管理台帳）\` |
|---|---|
| `FXJC` / `FSJC` | `06-04_ﾌｧｲﾙ一覧_工程管理.xlsm` (sheet `ﾌｧｲﾙ一覧(工程管理)`) |
| `FXJA` / `FSJA` | `06-04_ﾌｧｲﾙ一覧_基準情報.xlsm` (sheet `ﾌｧｲﾙ一覧`) |
| `FXJZ` / `FSJZ` | `06-04_ﾌｧｲﾙ一覧_共通.xlsm` |
| `FXJD` / `FSJD` | `06-04_ﾌｧｲﾙ一覧_品質管理_.xlsm` (note the trailing `_`) |
| `FXJB` / `FSJB` | `06-04_ﾌｧｲﾙ一覧_受注出荷.xlsm` |

A program's files can span two prefixes (a 工程管理 program writing an `FSJA` file) — route per ID.
**Read the registry with strikethrough resolved** — openpyxl per `agent-guide.md`'s "Reading a
reference file yourself". Do not use the plain `xlsx-dumps` cache script
for this file: it writes `Value2` as-is, so a struck (withdrawn) row reads as registered and a
renamed-in-place name reads as old+new text concatenated. The registry's header row reads
`No. | ファイルID | ファイル名称 | 編成 | ＲＬ | ＢＦ | ＢＬ | 区分 | 備考 | 計画書No.` (row 4 on the
工程管理 file, data from row 5; locate it by the `ファイル名称` label after stripping whitespace — the cell
actually reads `フ　ァ　イ　ル　名　称` with full-width spaces). `備考` names
the owning function (`損金一覧`, or `付属機能/損金一覧` for an 付属機能 program), and `編成` is
`TSV` / `CSV` / `EXCEL`. Apply the same `LiveText` cascade as `report-design-check` step 2: a
partially-struck name is a rename in place, not a dead row.

### 3. Run the checks

**F1 — ﾌｧｲﾙIDの一貫性.** One file ID is written in up to six places. They must all agree, exactly
(no prefix/substring matching — `FSJC035` and `FSJC035A` are different files):

1. the sheet name's `(<ID>)`;
2. the 3C header `ﾌｧｲﾙID` (`[4,10]`) and the 3D copy (`[4,62]`) — the 3D copy is often `-`, which is
   `naming-standard-compliance`'s concern, not a mismatch here; only a *different* ID is;
3. Ⅰ．ﾌｧｲﾙ出力条件's `出力ｺｰﾄﾞ` argument literal (`'FSJC051'` / `"FSJA008"` — strip the quotes);
4. the 機能定義書 Ⅲ．入出力定義 row that declares the file (a row whose ID column holds an `F…` ID);
5. every `【ﾌｧｲﾙ出力仕様書(<ID>)】` pointer in 機能定義書 prose.

The 出力ｺｰﾄﾞ is the one that decides runtime behaviour — the common file-output routine looks up the
file name and output directory with it (`05.ｼｽﾃﾑ共通設計書` `ﾌｧｲﾙ出力制御` 1.) — so a sheet whose name
says one ID and whose 出力ｺｰﾄﾞ says another is a real defect, not cosmetics. The self-check workbook's
own sample is the textbook case: sheet `ﾌｧｲﾙ出力仕様書(FXJA002)`, header and 出力ｺｰﾄﾞ `FXJA020`,
入出力定義 `FXJA020`, and three 機能定義書 pointers that say `FXJA002`. Report which places carry which
ID, and name the one the registry holds as the likely correct value.

This is the highest-yield check here: 4 of the 57 sheets that pass an 出力ｺｰﾄﾞ carried a different ID
from their own header, and the shape is always the same — a sheet copied from a sibling and not
updated. `PSJCO309` `FSJC007` passes `'FSJC006'` (its neighbour); `PSJCO406` `FSJC029` passes
`'FSJC002'`, and its whole ﾍｯﾀﾞｰ list is `PSJCO301` `FSJC002`'s item-for-item, so say so — "this
sheet looks like an unedited copy of X" tells the designer more than a list of mismatches. Two more
(`PSJCO403` `FSJC019`/`FSJC020`) carry the **screen ID** `GSJC403A` in the ﾌｧｲﾙID header cell. When
the sheet name's parenthesis holds no file ID at all (`PSJCO206`'s `ﾌｧｲﾙ出力仕様書(LFﾓｰﾄﾞ)` /
`(CSVﾓｰﾄﾞ_0)` / `(CSVﾓｰﾄﾞ_1)`, header ID blank) the file has not been numbered — one lower-severity
line per sheet, not an F1 mismatch.

Also compare the ﾌｧｲﾙ名称 (`[4,15]`) with the registry's `ファイル名称` and the 入出力定義 row's name.

**F2 — ﾌｧｲﾙ一覧への登録.** Every live file ID must have a live registry row; that row's `備考` should
name this program's function (compare against the 3C `ﾌﾟﾛｸﾞﾗﾑ名` `[3,15]`, allowing the `付属機能/`
prefix and a `(ｻﾌﾞﾌﾟﾛ)`/`(ﾊﾞｯﾁ)` suffix). Report both directions: a spec with no registry row, and a
registry row whose 備考 names this program with no spec sheet (check for a hidden or `削除)` sheet
before calling it missing).

**Tell a stale name from a collision — they are different severities.** Measured: all 58 sheets
with a real file ID had a registry row, and 7 of those rows' 備考 disagreed with the program. Five
were a stale name for the same function (`ﾘｰﾄﾞﾀｲﾑ実績` vs `ﾘｰﾄﾞﾀｲﾑ実績照会`, `ﾛｯﾄ停止登録` vs
`ﾛｯﾄ停止指示登録`) — the registry was not updated after a rename; low severity, one line. But
`PSJCO604` `FSJC045`/`FSJC046`/`FSJC047` are registered to **other functions entirely**
(`付属機能/棚卸基本ﾃﾞｰﾀﾒﾝﾃﾅﾝｽ`, and `FSJC047` = `ｲﾝｸﾏｽﾀｰﾌｧｲﾙ` / `付属機能/ｲﾝｸﾏｽﾀﾒﾝﾃﾅﾝｽ`) — the program
reused IDs that belong to someone else's files, which collides at runtime because the 出力ｺｰﾄﾞ
lookup is by ID. Decide by the registry's `ファイル名称` as well as its 備考: both unrelated to this
sheet is a collision (high); either a near-match is a stale entry (low). Also compare `ファイル名称`
alone — 15 of 58 differed, nearly all by a trailing `ﾌｧｲﾙ`/`ﾃﾞｰﾀ` or a parenthesised qualifier;
report those as one grouped low-severity line per workbook, not one finding each. A 3C
`ﾌﾟﾛｸﾞﾗﾑ名` that is not this program's at all (`PSJCO803` `FSJC080`'s header says
`部材・治工具発注在庫管理`, the program is ﾃｰﾌﾟﾛｯﾄ管理) is a copied header — say so.

Compare the registry's `編成` (`TSV`/`CSV`/`EXCEL`) with Ⅱ `出力方法` **only for a non-共通ﾀﾞｳﾝﾛｰﾄﾞ file**
that states a fixed format. For a 共通ﾀﾞｳﾝﾛｰﾄﾞ file the user picks the format at run time, and the
registry's `TSV` is just the default — never a finding.

**F3 — Ⅱ．ﾌｧｲﾙ出力仕様の記入漏れ.** The section is a fixed label list in col 4 with the value in
col 18: `出力ﾌｧｲﾙ名` / `ﾌｧｲﾙ用途` / `出力方法` / `出力先` / `ﾚｺｰﾄﾞ長` / `文字ｺｰﾄﾞ` / `改行ｺｰﾄﾞ` / `0件出力`,
optionally `備考`. Accepted variants: `出力ﾌｧｲﾙ名(共通)` (10 sheets), `出力先(画面起動の場合)` (6), and an
extra `OUTPUT_CD` row on batch specs (7). Col 4 of this section also holds `※` notes and sub-headings
(`(1)ﾍｯﾀﾞｰ`, `(2)明細`, `参照先定義`) that are **not** labels — only the fixed labels above are
checked, or every note reads as a blank value.

Flag only a fixed label whose value is **genuinely empty**. A delegation (`【ｼｽﾃﾑ共通設計書　ﾌｧｲﾙ出力制御.基本方針】　参照`,
`機能定義書(PSJCOB07) Ⅳ．機能処理概要.A-③.ﾀﾞｳﾝﾛｰﾄﾞ処理 参照`) is a value, and `0件出力 = なし` is a
decision. Measured: zero genuinely empty fixed values in 63 sheets, so this check should be silent
on almost every run — a hit here is worth reading closely. A value can also start with a newline and
look empty on its first physical line; join the record before judging it (see the dump doc).

What does occur is a value **copied into the wrong row**: `PXJAO802` `FXJA020`'s `出力方法` reads
`ﾀﾞｳﾝﾛｰﾄﾞ画面の戻り値に設定された出力ﾌｧｲﾙ名` — the 出力ﾌｧｲﾙ名 sentence — where every other 共通ﾀﾞｳﾝﾛｰﾄﾞ
spec says `…出力形式`. Check that `出力方法` does not end in `出力ﾌｧｲﾙ名` and that `出力ﾌｧｲﾙ名` does not end
in `出力形式`.

**F4 — 共通ﾀﾞｳﾝﾛｰﾄﾞの注釈.** A file whose Ⅱ `出力ﾌｧｲﾙ名` reads `ﾀﾞｳﾝﾛｰﾄﾞ画面の戻り値に設定された出力ﾌｧｲﾙ名`
is produced through the common download screen (`PSJAO702_【共通】ﾀﾞｳﾝﾛｰﾄﾞ`), which lets the user
pick the format, the columns and the header row. Such a spec must say so in three places — the
self-check sample's call-outs require each one:

- Ⅰ: `※上記①のｿｰﾄ項目・出力順をﾀﾞｳﾝﾛｰﾄﾞ画面の引数に設定し、戻り値のｿｰﾄ項目・出力順でﾃﾞｰﾀ取得する。`
- Ⅲ ﾍｯﾀﾞｰ: `※.ﾀﾞｳﾝﾛｰﾄﾞ画面の戻り値のﾀｲﾄﾙ行の追加＝"1"(出力対象)の場合のみﾍｯﾀﾞｰを出力する。`
  (followed by the 出力項目ﾘｽﾄ sentence)
- Ⅲ 明細: `※.ﾀﾞｳﾝﾛｰﾄﾞ画面の戻り値に設定された出力項目のみ、出力項目ﾘｽﾄに設定された順番で出力する。`

Also, with 共通ﾀﾞｳﾝﾛｰﾄﾞ the `出力方法` should be `ﾀﾞｳﾝﾛｰﾄﾞ画面の戻り値に設定された出力形式`, not a
hard-coded `TSV`/`CSV` — the user picks it. Match on the distinctive phrases (`ｿｰﾄ項目・出力順`,
`ﾀｲﾄﾙ行の追加`, `出力項目ﾘｽﾄに設定された順番`), not the whole sentence: the numbering prefix and
full-width spacing vary.

Detect 共通ﾀﾞｳﾝﾛｰﾄﾞ from **either** `出力ﾌｧｲﾙ名` (or `出力ﾌｧｲﾙ名(共通)`) **or** `出力方法` containing
`ﾀﾞｳﾝﾛｰﾄﾞ画面の戻り値` — measured, 25 sheets said it in 出力ﾌｧｲﾙ名 but 31 said it in 出力方法, so keying
on one field misses files. Of the 共通ﾀﾞｳﾝﾛｰﾄﾞ sheets, 4-5 lacked the notes: `PSJCO505` `FSJC034`
and `PXJCO128` `FXJC004`/`FXJC005` lack all three; `PXJCO165` `FXJC013` lacks the ﾍｯﾀﾞｰ and 明細
notes. A multi-part file (`FSJC034`'s `1) 送品案内` / `2) 納品案内` / `3) 受入日報`) needs the ﾍｯﾀﾞｰ and
明細 notes per part, but one Ⅰ sort note covers the file.

**F5 — ﾍｯﾀﾞｰ項目の画面項目ID.** `05.ｼｽﾃﾑ共通設計書` `ﾌｧｲﾙ出力制御` 2.(3): the header row's labels are
fetched from the localization master for multilingual output, so **every ﾍｯﾀﾞｰ/段落ﾀｲﾄﾙ item needs a
画面項目ID**. Check, per ﾍｯﾀﾞｰ table:

- the table has a `画面項目ID` column at all, and no item in it is blank or `-` — except an item that
  is itself generated per row (`明細項目(1-N)`, whose label comes from the data; its `-` is correct);
- no ID is still the placeholder `XXXXXXX` — `PSJCO403` `FSJC019`/`FSJC020` and `PSJCO404`
  `FSJC023`/`FSJC024` have entire ﾍｯﾀﾞｰ columns of it (68 cells). An ID column holding a `※n`
  marker, `※` or `・・・` is an annotation, not an ID: don't report it as malformed, but check the note
  it points to exists;
- **number band by prefix**: `XJC`/`SJC` (工程管理) IDs on a file output are the `8xxx` block;
  `XJA`/`SJA` (基準情報) IDs are the ordinary `0xxx` dictionary IDs. Measured over every ﾍｯﾀﾞｰ ID in
  the 63 sheets: `XJC` 614 × `8xxx` / 29 × `0xxx`, `SJC` 546 / 11, but `XJA` 179 and `SJA` 355 — all
  `0xxx`, not one `8xxx`. So flag an `XJC`/`SJC` header ID outside `8000-8999`, and never flag an
  `XJA`/`SJA` one for not being `8xxx` — the self-check's レビュー観点 No.67 ("画面項目IDが8xxx番台")
  is a 工程管理 rule, and applying it to every prefix would bury the report in ~530 false findings.
  The 40 `XJC`/`SJC` `0xxx` IDs are the screen-item IDs reused on the file (`XJC0036` 品目ｺｰﾄﾞ where
  the file IDs use `XJC8049`). 37 of them sit in `PSJCO405`/`PSJCO406` (whole sheets built that way);
  the other three are lone outliers in otherwise-`8xxx` sheets (`PSJCO504` `FSJC033`, `PSJCO803`
  `FSJC080`, `PSJCOB07` `FSJC042`) — the higher-confidence shape. When the same 出力項目名 carries an
  `8xxx` ID on another file in the corpus or the dictionary, give it as the suggested fix;
- every ID is registered in the `82.画面項目辞書_*.xlsx` its prefix routes to, and the dictionary's
  画面項目名 matches the `出力項目名` — read the index per `_shared/reference-index.md`; split a cell
  holding several IDs on whitespace/newline.

明細 tables have no 画面項目ID column in the template. Its absence there is not a finding.

**F6 — ﾍｯﾀﾞｰと明細の対応.** Within one part, the ﾍｯﾀﾞｰ table's `出力項目名` list and the 明細 table's
list describe the same columns, so they must match **in the same order** — a header labelled for one
column over data from another is the silent kind of defect nobody notices until a user reads the
file. Report a set difference (an item in one list only) and, separately, an order difference.

A naive list diff flagged 18 of 54 pairs; most were not defects. **Normalise and skip first**, in
this order:

1. Strip `※n` / `※` markers and surrounding whitespace from each name (`実績表項目部 ※1` = `実績表項目部`).
2. Skip the pair when either side is the single row `編集無し` — that file has no header or no data
   section by design: an upload/template file (`PSJCO301` `FSJC002` 実績取込, `PSJCO405` `FSJC025`)
   has ﾍｯﾀﾞｰ only, and an interface file with physical column names (`PXJCO906` `FXJC027`-`029`,
   `ABKUBUN`/`BTR_NO`…) has 明細 only.
3. Skip generated pairs: a ﾍｯﾀﾞｰ `…名(1～N)` / `…項目(1～N)` against a 明細 `…値(1～N)` is one repeating
   column group (`不良項目名(1～N)` ↔ `不良項目値(1～N)`), not a mismatch.
4. Pair within one part only (`1-1.ﾍｯﾀﾞｰ` with `1-2.明細`, and each `(2)ﾍｯﾀﾞｰ` with the `(3)明細` of the
   same numbered block).

What survives is real, and it is exactly the kind of defect this check exists for:

- **Order swap** — `PSJCO502` `FSJC032`: ﾍｯﾀﾞｰ 4-7 read 送付先作業場ｺｰﾄﾞ/名 then 送付元作業場ｺｰﾄﾞ/名,
  明細 4-7 read 送付元 then 送付先. Every exported file labels the sender column as the recipient.
- **Name drift / typo** — `PSJCO504` `FSJC033` `製造ﾛｯﾄNo` vs `製造ﾛｯﾄNp`; `PSJCO604` `FSJC046`/`047`
  `KC品名` vs `品名`; `PXJAO701` `FXJA014` `事業所ｺｰﾄﾞ` vs `事業所`.

**F7 — 明細の参照先の解決.** The 明細 `参照先` column (col 17) names where each value comes from:

- `①` / `②` … → a symbol defined in Ⅰ．ﾌｧｲﾙ出力条件 by a line ending `… ①`
  (`画面設計書(GSJC901A) - (6)明細情報取得 参照 … ①`);
- `(n)` → **usually the block number of the delegated 画面設計書's Ⅲ．画面表示仕様, not of Ⅰ.** The common
  shape is Ⅰ `(1)ﾀﾞｳﾝﾛｰﾄﾞﾃﾞｰﾀ取得` → `※画面設計書(GSJC208A) Ⅲ.画面表示仕様 (7)明細ﾃﾞｰﾀ取得　参照`, and the
  明細 rows then say `(7)`. Measured on the 26 sheets whose 明細 uses `(n)`: reading `(n)` as Ⅰ's own
  block number would have left 18 of them "undefined" or pointing at a non-data block
  (`PXJCO128` `FXJC004`: Ⅰ `(1)` is `ﾌｧｲﾙ出力先情報取得`, the data is Ⅰ `(2)`, and `(1)` is the screen's
  `(1)ﾛｯﾄ停止指示`). Resolve in this order:
  (a) **Ⅰ itself numbers a block `(n)`** → `(n)` means that block. If it is a data block, resolve
  against it (`PSJCO505` `FSJC034` uses Ⅰ `(2)`-`(4)` this way). If it is the non-data
  `ﾌｧｲﾙ出力先情報取得` block (or another block with no 取得項目), that is a **finding**, even when the
  delegated screen happens to have a block `(n)` too. Confirmed on `PXJCO128` `FXJC004`/`FXJC005`:
  Ⅰ used to read `(1)…画面設計書(GXJC128A).Ⅲ.画面表示仕様(1) 参照`, the block was restructured on
  2026/6/9 into `(1)ﾌｧｲﾙ出力先情報取得` + `(2)ﾀﾞｳﾝﾛｰﾄﾞﾃﾞｰﾀ取得` (old line struck — see
  `_DELETED_DIGEST.txt`), and every 明細 `(1)` was left behind. Check the digest to say so; the fix is
  `(2)`. `PXJCO193` `FXJC018` has the same shape.
  (b) **Ⅰ has no block numbered `(n)`** (its blocks are unnumbered `・ﾌｧｲﾙ出力先情報取得` or use other
  numbers) → `(n)` is the delegated 画面設計書 Ⅲ block number; resolve it there (`PSJCO208`).
  Report a `(n)` that resolves in neither. Ⅰ's block titles sit in col 3 on some sheets and col 4 on
  others, so find them by pattern, not column;
- compound and external forms, all legitimate: `(2)-①` (sub-item ① of Ⅰ block `(2)`), `(1)、区分名称.… 参照`
  (block plus a 区分名称 lookup), `共通項目取得(工程管理).汎用工程GRP 参照` (delegation to
  `01_Doc\04_共通設計\07.共通項目取得.xlsx` — check the named item exists (sheet `共通項目取得` and the WG
  sheet such as `共通項目取得(工程管理)`; read it live per `agent-guide.md`));
- **an Ⅰ block with its own inline 取得項目** (`PSJCO502` `FSJC032`'s `② 納品案内明細情報取得`, beside
  a delegating `①`) → check the 明細 values against that block's own 取得項目 list, not a screen's;
- `A` / `B` … → an alias in Ⅰ's `参照ﾃｰﾌﾞﾙ` / `参照ｴﾝﾃｨﾃｨ` list;
- `-` → a literal (every ﾍｯﾀﾞｰ row) or a computed value.

Check that every symbol/block/alias used is defined in Ⅰ, and follow each delegation: the named
`画面設計書(<画面ID>)` sheet must exist in the workbook and its Ⅲ．画面表示仕様 must have the named
block (match `(n)` and the block title; titles drift, so a matching number with a different title is
a lower-confidence note, a missing number is a finding). Then check each 明細 `項目名・出力値` against
that block's `取得項目` list, **by name after stripping the decoration**: a trailing format
(`(YYYY/MM/DD)`, `(YYYY/MM/DD HH24:MI:SS)`), a `+':'+名称 ※…参照` 区分-name suffix, and quotes. A
printed value that the delegated block never fetches is an implementation hole — the same finding
`report-design-check` C4 raises for 帳票.

The delegation line is written several ways; recognise all of them before calling a 明細 source
undefined: `画面設計書(GXJA802A) - (1)明細ﾃﾞｰﾀ取得 参照 … ①`,
`※明細の取得内容は、画面設計書(GSJC505A)　Ⅲ.画面表示仕様 (4)①検索結果　参照`, and a bare
`参照ﾃｰﾌﾞﾙ` list whose aliases point at screen areas (`A | 製品仕様登録詳細.G1).画面情報` in
`PSJAO203` `FSJA008`) — the last kind resolves to a 画面 area, not a 画面設計書 Ⅲ block, so check the
area name exists in that screen's Ⅴ．画面項目定義 instead. A sheet ID in the pointer can carry a stray
space (`画面設計書(GSJC311A )`); trim it before looking the sheet up.

**F8 — 編集方法と出力値の整合、No の出力.** `05.ｼｽﾃﾑ共通設計書` `ﾌｧｲﾙ出力制御` 4. defines exactly four
編集方法 for data: `文字列` / `数値` / `日付` / `日付+時刻` (followed by a `※n` pointer to that section).
Flag, on 明細 rows:

- a value whose format annotation carries a time (`HH24:MI:SS`, `HH:MI`) with 編集方法 other than
  `日付+時刻`, and a date-only annotation with 編集方法 other than `日付`;
- a 編集方法 outside that vocabulary on a 明細 row (ﾍｯﾀﾞｰ rows legitimately read `-`, `編集無し` or
  `編集なし` — measured 599 / 658 / 38 — none of them is a finding);
- an item named `No` / `No.` in any 明細 or ﾍｯﾀﾞｰ table — the self-check call-out: "画面の明細にNoを
  追加していても、ﾌｧｲﾙ出力にはNoはいらない".

Do **not** flag a `日付` value that carries no `(YYYY/MM/DD)` annotation: the common design only says
a date "is passed as YYYY/MM/DD or YYYYMMDD" and does not require the annotation — 41 明細 date
values in the corpus have none, so flagging it would be pure noise.

Measured vocabulary on 明細 rows (after dropping the `※n` suffix): `文字列` 1011, `数値` 335, `日付` 64,
`日付+時刻` 19, `-` 168 (computed or literal values), plus the outliers worth a line: `日時` (2 — not a
defined 編集方法; the defined one is `日付+時刻`), `編集しない` (4), `"文字列"` in quotes (3), `・・・` (1).
Format/編集方法 contradictions that did occur: `PSJAO501` `FSJA012` rows 774/776/778 and `PXJCO165`
`FXJC013` rows 181/182 annotate a time but say `日付`; `PSJCO208` `FSJC090` gives `作業時刻` the
編集方法 `日付` (a name that says time, with no annotation either way — report as 要確認, since the
common design defines no time-only 編集方法).

`No` as an output item occurred in 10 places (`PSJAO241` `FSJA004`, `PSJCO204` `FSJC068`,
`PXJCO121` `FSJC001` in both ﾍｯﾀﾞｰ and 明細). Before reporting, look at the 明細 row's 参照先: a `No`
taken from the screen's list-row counter (`XJZ0425`, "1から連番") is the self-check's case and a
finding; a `No` that is a real DB column of the source block (a sequence number the user needs) is
not — say which it is.

**F9 — ﾌｧｲﾙﾚｲｱｳﾄとの一致（ﾚｲｱｳﾄがある場合のみ）.** Most download files have no layout workbook —
only interface files do. Look for `<ﾌｧｲﾙID>_*.xlsx` in, in this order:
`<WG番号>_<WG名>WG\07_データベース・ファイル設計書(仮)\ファイルレイアウト\` (glob
`*WG\07_データベース・ファイル設計書(仮)\ファイルレイアウト\<ﾌｧｲﾙID>_*.xlsx` — only `11_工程管理WG` has one today) and
`01_Doc\07_データベース・ファイル設計書\07_03_ファイルレイアウト\<nn>_<WG名>\`. Confirm the ID in the
layout sheet's `[6,1]` (under the `ﾌｧｲﾙID` label at `[5,1]`), exactly — the same suffix trap as table
layouts. No layout found is **not** a finding; say in one line which files had one.

A layout workbook holds one sheet per record type — `ﾌｧｲﾙﾚｲｱｳﾄ`, or `ﾌｧｲﾙﾚｲｱｳﾄ(ﾍｯﾀﾞｰ)` +
`ﾌｧｲﾙﾚｲｱｳﾄ(ﾃﾞｰﾀ)`, or named pairs like `品目情報(ﾍｯﾀﾞｰ)`/`品目情報(ﾃﾞｰﾀ)` (`FSJA004`). Row 6 gives
`形式`/`文字ｺｰﾄﾞ`/`改行ｺｰﾄﾞ`/`ﾚｺｰﾄﾞ長` (cols 29/35/41/47); items start at row 8 with `項目名` in col 3
and `TYPE` in col 15 (`X` = 文字, `9` = 数値), continuing in a right-hand column pair (col 27 No.,
col 29 項目名, col 41 TYPE) past row 59. Compare:

- item names, in order, against the matching spec table (`(ﾃﾞｰﾀ)` ↔ 明細, `(ﾍｯﾀﾞｰ)` ↔ ﾍｯﾀﾞｰ/ｷｰ情報);
- `TYPE` against 編集方法: `X` ↔ `文字列`/`日付`/`日付+時刻`, `9` ↔ `数値`;
- the layout's 形式/文字ｺｰﾄﾞ/改行ｺｰﾄﾞ against Ⅱ, only where Ⅱ states a value rather than delegating
  to `ｼｽﾃﾑ共通設計書 ﾌｧｲﾙ出力制御`.

Measured: only 8 layouts exist under the 工程管理WG folder (`FSJA004`, `FSJA008`, `FSJA013`,
`FSJA014`, `FSJA019`, `FSJC043`, `FSJC044`, `FSJC051`) and 9 under `07_03_ファイルレイアウト`, against 63
specs. Where both exist they disagree more often than not, and the layout is usually the older side:
`FSJC051`'s layout (作成 2024/1, 33 items, `No` first) against the spec's 42 items with no `No`;
`FSJA008`'s `(ﾍｯﾀﾞｰ)` sheet has 8 items while the spec's `1.ｷｰ情報` has 10 and a different order,
the spec having been revised through 2025/10. So report each layout disagreement as **要確認** with
both dates (layout `作成日`/`更新日` in its 3C header, spec `更新日` `[3,46]`), one grouped finding per
layout sheet listing the extra/missing/reordered items — not one finding per item.

**Out of scope here, handled elsewhere:**
- Whether the DB columns behind a delegated 画面設計書 block exist — `xlsx-db-column-check` (it now
  also harvests a ﾌｧｲﾙ出力仕様書's own inline 参照ｴﾝﾃｨﾃｨ blocks).
- Header 3C/3D agreement and 表紙 構成 markers — `naming-standard-compliance`.
- Prose typos in Ⅱ `ﾌｧｲﾙ用途` / `※` notes — `design-doc-typo-check`.
- Rich-client wording (`出力先はﾕｰｻﾞのﾀﾞｳﾝﾛｰﾄﾞﾌｫﾙﾀﾞとし、出力後ﾃﾞﾌｫﾙﾄｱﾌﾟﾘで開く` in 備考): the
  workbook does not say whether its screen is rich-client, so this can't be decided mechanically.
  Mention it only if 画面設計書 Ⅱ says `ｺﾝﾄﾛｰﾙﾎﾞｯｸｽ` (the rich-screen marker the self-check requires)
  and 備考 is missing.

### 4. Verify before reporting

Re-read the cells behind each candidate. The traps specific to this sheet:

- pairing a ﾍｯﾀﾞｰ with the 明細 of another part (F6);
- matching item names without stripping the newline/whitespace a merged cell carries, or without
  removing the value decoration (F6/F7/F9);
- flagging an `XJA`/`SJA` header ID for not being `8xxx` (F5);
- flagging a missing `(YYYY/MM/DD)` annotation (F8), or a missing ﾌｧｲﾙﾚｲｱｳﾄ (F9);
- treating a hidden template stub (`FXJDXXX`, `()`) as a live spec.

## Reporting

Write in the user's language, for the designer who owns the doc.

Group by ﾌｧｲﾙID. For each finding: cite the sheet and cell (`ﾌｧｲﾙ出力仕様書(FSJC051) [89,45]`) with a
clickable `[filename](relative/path)` link, and state the consequence in one clause (出力ｺｰﾄﾞが別
ﾌｧｲﾙを指しており出力先・ﾌｧｲﾙ名が取れない / ﾍｯﾀﾞｰが多言語化されない / ﾍｯﾀﾞｰと明細の列がずれる / 出力値が
取得されていない). Order by severity: F1 出力ｺｰﾄﾞ mismatches and F7 unfetched values first, then F5/F6,
then F2/F3/F4/F8, then F9. Mark F9 as 要確認 when the layout is older than the spec's 更新日 — the
layout may simply not have been maintained, and the designer decides which side is right.

Omit: files that checked out clean, tallies, and narration of what you dumped. One process fact is
worth a line — a registry or dictionary file that couldn't be located for a file's prefix, since that
leaves F2/F5 unchecked for it.
