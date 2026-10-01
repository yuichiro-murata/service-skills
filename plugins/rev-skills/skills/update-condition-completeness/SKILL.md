---
name: update-condition-completeness
description: Check a program's 更新条件表(<TableID>) sheets against the table's actual ﾃｰﾌﾞﾙﾚｲｱｳﾄ — every column of the table present in the right order, no notnull column left unset on INSERT, every primary-key column set on INSERT and carried as [KEY] on UPDATE/DELETE, and the house conventions for the 共通項目 block (登録者/登録日時 only at INSERT, 更新者/更新日時 on both, 排他ﾌﾗｸﾞ 1 / +1, 更新ﾌﾟﾛｸﾞﾗﾑID=画面ID) honored. Distinct from its siblings: `xlsx-db-column-check` asks whether a referenced column exists at all, `design-doc-io-table-check` asks whether the table is declared in the Ⅲ．入出力定義 CRUD list — this skill asks whether the UPDATE/INSERT/DELETE specification for an already-declared table is complete enough to code from. Use when the user asks to check 更新条件表 completeness, NOT NULL/主キー/共通項目の設定漏れ, 登録日時が更新されていないか, or ｼｽﾃﾑ日時の書式注記(YYYY/MM/DD HH24:MI:SS形式)漏れ. For a full REV without named checks the entry point is `rev-program-review`; run this standalone only when selected there or asked for by name.
---

# update-condition-completeness

Reviews every `更新条件表(<TableID>)` sheet in one program's design-doc workbook against the
authoritative `ﾃｰﾌﾞﾙﾚｲｱｳﾄ` of the table it updates. The question this check answers is **"could a
coder implement this INSERT/UPDATE/DELETE from the sheet alone, and would the row it writes satisfy
the table's own constraints?"** — not "does this column exist" (that is `xlsx-db-column-check`) and
not "is this table declared in the CRUD list" (that is `design-doc-io-table-check`).

This is the second-largest finding category in the project's own review history: the
第3フェーズ基本設計 指摘傾向レポート puts 更新条件表/DB at 84% of one designer's findings and 35% of
another's, behind only 処理フロー・I/O.

## Environment

**Read `_shared/agent-guide.md` first** — scope selection, environment, the live dump, folder
exclusions, reporting conventions. Read the ﾃｰﾌﾞﾙﾚｲｱｳﾄ files live with `live_dump.py` (step 2) — never
through the COM cross-session cache "Batch variant", which keeps struck text (PXJCO125's 35 layouts,
2026-10-01: `TXJCM006` `[81,40]` dead, `TXJCA317` `[23,40]`/`[34,40]` partial).

## Sheet anatomy (measured, not assumed)

Verified against all 13 `更新条件表(*)` sheets of
`01_Doc\08_機能定義書\11_工程管理\PHASE3\PXJCO192_処置指示登録.xlsx`. Confirm the shape holds on the
workbook in front of you before trusting the column numbers — the **labels** are the contract, the
column numbers are just the common case.

```
[8,2]=更新ﾃｰﾌﾞﾙ      [8,16]=TXJCM137WF:処置指示ﾏｽﾀWF      ← <TableID>:<table name>
[9,2]=更新概要        [9,16]=①画面(GXJC192A)、②ﾛｸﾞｲﾝ情報   ← defines the ①②… source symbols
[12,2]=更新条件       [12,16]=保存ﾎﾞﾀﾝ押下時 / 分岐条件      ← the trigger, often multi-line
[14,2]=No. [14,4]=項目名 [14,16]=INSERT [14,34]=UPDATE      ← the header row
[15,2]=1   [15,4]=会社ｺｰﾄﾞ [15,16]=② [15,18]=会社ｺｰﾄﾞ  [15,34]=② [15,36]=[KEY]会社ｺｰﾄﾞ
```

Rules that make the parse reliable:

- **The header row is the row with `No.` in col 2 AND `項目名` in col 4.** Do not anchor on row 14.
  Row 13 frequently carries a free-text flow note that happens to contain an operation word
  (`[13,16]=DELETE⇒INSERT`) and will be mistaken for the header by a looser rule.
- **One sheet can hold several blocks.** Half the sheets checked had two or three (e.g.
  `更新条件表(TXJCM137)` has header rows at 14, 68 and 122; `TXJAM026WF` at 14 and 164). Scan the
  whole sheet for header rows and treat each block independently — each has its own 更新ﾃｰﾌﾞﾙ,
  更新条件 and operation set. Stopping at the first block silently skips most of the specification.
- **Operation columns are read off the header row, not hardcoded.** Observed label columns: 16, 28,
  34, 40. Accept a header-row cell as an operation when it **starts with** an operation word —
  `re.match(r"^[\s　]*(?:\d+\.|【)?[\s　]*(INSERT|UPDATE|DELETE|MERGE)", text)` — and take its kinds from
  **every** operation word in it (`re.findall`). Keep the explicit full-width space `　`: Python 3's
  `\s` happens to cover U+3000, but an ASCII/`re.ASCII`/PowerShell `[ \t]` port does not, and silently
  drops `[14,16]=　INSERT` (`PXJCO125` `TXJCA205`/`TXJCA318`/`TXJCA407` — three whole blocks). **Never anchor the end of the label.** Measured on
  the raw text (struck and hidden-sheet labels included) of the 122 工程管理 PHASE1-3 workbooks
  (2026-10-01): of 1,444 header labels, an end-anchored
  `^(INSERT|…)[0-9０-９①-⑳]*$` silently dropped **98** — `UPDATE ※1`, `INSERT※1`, `DELETE　※1`,
  `INSERT① ※1`, `INSERT-1`, `INSERT_1`, `1.DELETE`, `【DELETE】※1`, `DELETE　(※1)`,
  `INSERT(新規入力の場合)`, `UPDATE(1)⏎(削除された…明細)` — each a whole block left unreviewed. The
  prefix match accepts all 1,443 real labels and rejects only a free-text note in the same row
  (`[115,74]=更新条件表(TXJCM006)のUPDATE①でｷｰにした枝番`). **Read the suffix**: a `※n` or a
  parenthetical says when the operation runs, which C3/C6 need. A column naming two operations
  (`DELETE　INSERT`, `UPDATEDELETE`, `UPDATE(存在すれば) INSERT`) gets both rule sets: C2 for its
  INSERT side and the `[KEY]` rule of C3 for its UPDATE/DELETE side — **judged on the live label**.
  Those three `PXJCO125` blocks read `DELETE　INSERT` raw, but the DELETE is struck (digest: `[14,16]
  PART`), so the live label `　INSERT` is a plain INSERT. Circled and ASCII digits both
  occur (`TXJAM026WF` `INSERT①`/`INSERT②`, `TXJAM027` `INSERT1`/`INSERT2`), and repeated plain labels
  (two bare `INSERT` columns) are real too. The start anchor matters: several sheets put an
  unrelated
  取得ﾃｰﾌﾞﾙ/検索条件 sub-block on the same rows further right (`[14,58]=ﾊﾟﾗﾒｰﾀﾏｽﾀ.ｷｰ1`,
  `[14,58]=共通ｺｰﾄﾞﾏｽﾀ.ｷｰ1`), and a "any non-empty cell to the right is an operation" rule turns
  those into phantom operations.
- **Each operation is two sub-columns: the label column `c` holds the source symbol (`①`, `②`, `-`),
  and `c+2` holds the value.** Confirmed for every observed label column (16→18, 28→30, 34→36,
  40→42).
- **"Not set" means the *value* column is `-` or empty — never the source column.** A column set
  from something other than the screen/login sources legitimately reads `[17,16]=-` with
  `[17,18]=ｼｽﾃﾑ日時`. Judging by the source column alone reports every system-supplied column as a
  missing value.
- **`[KEY]` prefixes the value** and marks the column as part of the WHERE clause for
  UPDATE/DELETE (`[15,36]=[KEY]会社ｺｰﾄﾞ`). A literal in a key position is normal
  (`[28,18]=[KEY]"000"`).
- **Struck-through blocks are already gone from the dump.** The shared dump resolves strikethrough
  live, so a withdrawn 更新条件表 block simply won't be there (in `PXJCO192`,
  `更新条件表(TXJCM137)` is a visible sheet whose entire content is struck out and marked
  `2026/6/10 福浦 削除`). Do not resurrect it from `_DELETED_DIGEST.txt` to review it. A table whose
  update spec was withdrawn but which is still declared in Ⅲ．入出力定義 is a real finding — but it
  belongs to `design-doc-io-table-check`, not here.
- **Hidden sheets are out of scope**, per the shared dump doc's default.
- **A dump record can span several physical lines.** The dump writes one line per sheet row, but a
  cell whose own value contains a newline (a multi-line 更新条件, a `※…`-annotated INSERT value)
  puts that newline straight into the file. Parse with `_shared/scripts/dump_cells.py` (`load()` groups by the
  `[r,c]=` tokens, so embedded newlines are harmless) — never line by line, or the row is truncated at the newline and every operation
  column after it reads as empty — which looks exactly like "未設定" and produces false C2/C3
  findings. Confirmed on `TXJCM137WF` row 26 (`ﾜｰｸﾌﾛｰID`), whose INSERT value carries a
  `※ｼｽﾃﾑ共通設計書…` continuation.

## Procedure

### 1. Get the workbook dump

If `rev-program-review` already dumped it, read those scratchpad files. Otherwise dump it live with `python _shared/scripts/live_dump.py <workbook> <out_dir>`
(openpyxl; same format and digest as the shared dump). Only `更新条件表(*)` sheets matter for this check, plus the
機能定義書 and 画面設計書 sheets for the trigger cross-check in step 5 C6.

### 2. Resolve each 更新ﾃｰﾌﾞﾙ to its layout file

The `更新ﾃｰﾌﾞﾙ` cell gives `<TableID>:<name>`. Resolve the ID to exactly one `ﾃｰﾌﾞﾙﾚｲｱｳﾄ` file using
**`xlsx-db-column-check`'s step 3 "Quick reference"** — read that section and follow it rather than
re-deriving it. The rules that bite hardest here: the search root is the top-level
`<WG番号>_<WG名>WG\07_データベース・ファイル設計書(仮)` (not nested under `01_Doc`), the ID must match
**exactly** (a `WF` suffix is a different table, and `更新条件表` sheets are full of `*WF` tables), the
`ﾃｰﾌﾞﾙﾚｲｱｳﾄ` sheet's `A6` cell must confirm the ID, and precedence is PH3 > PH2 > top-level.

Dump each one live, one out_dir per file — the same call as `xlsx-db-column-check` step 3, so a REV
running both reuses one set of dumps (the default prefix collides for `TXJCM007` vs `TXJCM007_B`):
`python _shared/scripts/live_dump.py <layout.xlsx> <out>/<TableID> --sheets "^ﾃｰﾌﾞﾙﾚｲｱｳﾄ$" --prefix <TableID>`
— anchor the regex: unanchored, it also matches `旧`/`JAG_`/`JAGUR_`/`CP2_` copies and dated snapshots
such as `ﾃｰﾌﾞﾙﾚｲｱｳﾄ_20251002時点`, which share the live sheet's `A6` (that skill's step 3 rule 2 lists
them). ~1–7 s each.

The layout sheet's header is at row 7 and the columns this check needs are:

| col | header | use |
|---|---|---|
| 1 | `No.` | column order |
| 3 | `項目名` | the name the 更新条件表 uses |
| 12 | `項目ID` | for reporting |
| 28-35 | `I01`…`I08` | index membership — the digit in the cell is that column's position within the index |
| 36 | `default` | a DB-side default, relevant to C2 |
| 38 | `notnull` | `Y` = NOT NULL |

- **The column list ends at the first empty 項目名 or at `＜ｲﾝﾃﾞｯｸｽ情報＞`.** Numbered blank rows often
  follow the last column (`TXJCD404` Nos 43-48, `TXJCM057` Nos 41-51, `TSJCA057` Nos 44-45); counted as
  columns they produced empty-named C1 "missing column" findings on 8 blocks.
- **`I01` is the primary key only when the `＜ｲﾝﾃﾞｯｸｽ情報＞` label says so** — `PRIMARY KEY I01`
  (`TXJCD404` `[57,1]`) or `ﾕﾆｰｸｷｰ　I01` (`TXJCM057` `[60,1]`). `ﾕﾆｰｸｷｰ　無し` / `ﾕﾆｰｸｲﾝﾃﾞｯｸｽ I01`
  (`TSJCA057` `[54,1]`/`[55,1]`, `TXJCA205` PH3 `[43,1]`/`[44,1]`) means **no key**: skip C3 for that table.
- **Phase.** PH3 > PH2 still picks the file, but when the layout is from a later phase than the
  workbook's own `PHASE` folder, a column that exists only in the later layout is 要確認
  ("PH3ﾚｲｱｳﾄで追加/変更 — 設計書の反映要否を確認"), not 中 (`xlsx-db-column-check` step 3 rule 3).

### 3. Calibrate the house convention from the workbook itself

Before flagging anything about the 共通項目 block, tabulate what **this workbook's own** sheets do for
each common column, per operation. The conventions below were measured on `PXJCO192`, and they held
across its sheets — but they are a project convention, not a written rule, so the workbook's own
dominant pattern wins over this table when the two disagree. Flag a deviation from the workbook's
dominant pattern; do not flag the pattern itself.

Observed standard for the leading common-column block
(`会社ｺｰﾄﾞ/登録者/登録日時/更新者/更新日時/更新ﾎｽﾄ名/更新ﾌﾟﾛｸﾞﾗﾑID/排他ﾌﾗｸﾞ/部門GRP/事業部ｺｰﾄﾞ/移行ｵﾌﾞｼﾞｪｸﾄID`):

| 項目名 | INSERT | UPDATE | DELETE |
|---|---|---|---|
| 会社ｺｰﾄﾞ | 設定 | `[KEY]` | `[KEY]` |
| 登録者 | 作業者ｺｰﾄﾞ | **`-`** | `-` |
| 登録日時 | ｼｽﾃﾑ日時 | **`-`** | `-` |
| 更新者 | 作業者ｺｰﾄﾞ | 作業者ｺｰﾄﾞ | `-` |
| 更新日時 | ｼｽﾃﾑ日時 | ｼｽﾃﾑ日時 | `-` |
| 更新ﾎｽﾄ名 | `-` | `-` | `-` |
| 更新ﾌﾟﾛｸﾞﾗﾑID | 画面ID | 画面ID | `-` |
| 排他ﾌﾗｸﾞ | 何らかの初期値 | 増分 | `-` |
| 部門GRP | 設定 | `[KEY]` | `[KEY]` |
| 事業部ｺｰﾄﾞ | 設定 | `-` | `-` |
| 移行ｵﾌﾞｼﾞｪｸﾄID | `-` | `-` | `-` |

WF tables carry four more after that block — `ﾜｰｸﾌﾛｰID`, `ｸﾞﾙｰﾌﾟID`, `ﾜｰｸﾌﾛｰ申請区分`, `元排他ﾌﾗｸﾞ` —
and the WF sheets in the same workbook are the calibration set for those.

**Do not hardcode `排他ﾌﾗｸﾞ = 1` on INSERT.** A first version of this table did, and running it over
`PXJCO192` produced four false findings in a row: `TSJCM139WF`, `TXJCM838WF` and both of
`TXJAM026WF`'s INSERT blocks set `ﾗﾝﾀﾞﾑ値`, and `TXJAM025WF` sets `※3` (a footnote reference), while
only `TXJCM137WF`/`TXJCM501` use the literal `1`. On this table the initial lock value is a design
choice per table, so the only INSERT-side defect worth reporting is `排他ﾌﾗｸﾞ` left **unset**. On
UPDATE the opposite holds: the value must express an increment (`+1`), because a literal that does
not advance the counter defeats optimistic locking.

### 4. Run the checks

**C1 — 項目リストが実テーブルと一致しているか.** The 更新条件表's `項目名` rows should be the layout's
`項目名` list, in the layout's `No.` order. Report: a layout column with **no row at all** in the
更新条件表 (the sheet was written against an older table definition — the coder has no instruction
for that column); a row whose 項目名 exists in **no** layout column; and a row order that diverges
from the layout (low confidence on its own — report only when it coincides with added/removed
columns, since it is then evidence the sheet was patched by hand rather than regenerated). A missing
non-key column in a **DELETE-only** block is 低 (the statement never writes it). Skip notes-only
sheets such as `更新条件表(その他)` (no `No.`/`項目名` header row).

**Pair the two directions before reporting.** When a "missing" layout name and an "extra" sheet name
are near-matches — one is a prefix or substring of the other, they are within **Damerau distance ≤ 2**
(an adjacent transposition counts as one edit: `UNICORN採番ﾌﾗｸﾞ` vs the layout's typo `UNICRON採番ﾌﾗｸﾞ`,
`PXJCO125` `TSJCM999` `[31,4]` / layout `[24,3]`, is distance 1 but plain Levenshtein 2), or the row's 項目ID
matches — they are one finding (a rename, or a key written only in part), not two, **and the pair
is then treated as one column for C2/C3** (an unpaired typo silently skips the NOT NULL check). `TSJCM139WF` is the live example: the layout's PK7 NOT NULL
column is `工程ｺｰﾄﾞ工程No` and the 更新条件表 row says `工程ｺｰﾄﾞ`. Reported as two lines it reads like
an unrelated deletion plus an unrelated addition; reported as one it says what the reviewer needs to
decide — either the doc carries a stale name, or the sheet is only setting the 工程ｺｰﾄﾞ half of a
concatenated key. (That the project really does concatenate this way is visible in the same
workbook: `TXJCM838WF`'s `指示工程ｺｰﾄﾞ` is set from `G6).指示工程 ※工程ｺｰﾄﾞ部分のみ`.)

This check earns its place: on `PXJCO192`'s `TXJCM501` it found three layout columns —
`停止工程GRP`, `製造ﾛｯﾄNo2`, `解除理由` — with no row in the 更新条件表 at all. Note this only
surfaces if step 2's PH3 > PH2 precedence was applied: `TXJCM501` has a layout file under **both**
phase folders, and the 更新条件表 matches the older one.

**C2 — INSERT で notnull 列が未設定.** For every operation whose kinds include `INSERT` (or
`MERGE`): every layout column with `notnull = Y` must have a non-`-`, non-empty value. An unset
NOT NULL column is a guaranteed runtime failure, so this is the highest-severity finding this check
produces. Two exemptions, both verifiable from the layout: the column has a `default` value (col 36),
or it is the target of a DB-side trigger/sequence noted in 備考 (col 40). State the exemption rather
than staying silent — "notnull だが default 設定あり" is useful to the reviewer.

**C3 — 主キーの扱い.** From `I01` (col 28), build the table's primary key in position order — only when
the `＜ｲﾝﾃﾞｯｸｽ情報＞` label makes `I01` a key (step 2); a table with `ﾕﾆｰｸｷｰ 無し` has no C3.
- INSERT: every PK column must be set. A missing one means the row can't be identified afterwards.
- Normalise `【KEY】` / `［KEY］` to `[KEY]` before matching (report the notation once at 低) —
  `PXJCO161` `TXJCD104` block 4 writes `【KEY】`, and a literal match calls all 7 PK columns unkeyed.
- An `I01` member whose `notnull` is blank cannot be in an Oracle PK (a ※ may say `PKから削除`,
  `TXJCD104` `[26,105]`); don't require it on INSERT or in `[KEY]`.
- UPDATE / DELETE: every PK column should carry `[KEY]`; a partial key updates or deletes a *range*.
  **A cascade delete is the normal case, not a finding.** Measured on `PXJCO161`/`163`/`125`
  (2026-10-01): ~45 standalone DELETEs keyed only on the parent key (会社/部門/管理No/枝番[/工程ｺｰﾄﾞ])
  and omitting child columns (SEQ, 作業者, 資源, ﾌｪｰｽﾞID…) — every one a 取消/削除 cascade, most with
  the 更新条件 saying so (`TXJCD102` `[12,16]` `①.G1).管理No、枝番を条件に削除する`). **A standalone DELETE
  whose 更新条件 names exactly its `[KEY]` columns (会社ｺｰﾄﾞ/部門GRP may go unnamed) is a documented
  range delete — not a finding, whichever PK columns it omits.** On `PXJCO125` (2026-10-01) a looser
  reading flagged 39 PK columns over 20 such DELETEs, all false. Without that text, still exempt it when
  the omitted columns are trailing/child PK columns and the trigger is a 取消/削除 button. **Report**
  (中) when the `[KEY]` set contradicts the columns the 更新条件 names (`TXJCM008` block@147: text says
  `管理No、枝番`, keys add `[172,18]=[KEY]G5).工程ｺｰﾄﾞ`), or when a leading/parent PK column is missing
  and nothing documents the range.
  For a partial-key **UPDATE**, report at 要確認 unless the 更新条件 states the range is intended
  (`次工程ｺｰﾄﾞで作業着手をUPDATE` explains it), quoting that text.
  **Exception — suppress it for the DELETE half of a DELETE⇒INSERT block.** Where a block pairs a
  DELETE with one or more INSERTs on the same table, deleting by a partial key is the whole point:
  the statement clears every child row for a parent key and the INSERTs rewrite them. Flagging it
  produced noise on two of `PXJCO192`'s sheets (`TXJCM838WF`, where the DELETE omits `SEQ` from a
  7-column PK, and `TXJAM026WF`, where it omits `工程ｺｰﾄﾞ`), and in both the paired INSERT sets the
  full key. Report a partial key only for a **standalone** DELETE, for an UPDATE, or when the paired
  INSERT does **not** set the full PK. **The exemption does not cover a PK column the paired INSERT
  sets to a fixed literal** (a discriminator such as `区分="1"`): leaving it out of the DELETE's
  `[KEY]` also wipes the rows of every *other* discriminator value, which the INSERT never rewrites
  (`PSJCO501` `TSJCD215` `[35,18]`). Report that one at 中, quoting the literal. **Also compare the
  DELETE's `[KEY]` values with the paired INSERT's values for the same columns**: `PXJCO161`
  `TXJCD407` deletes by `[114,18]=[KEY]G1).工程ｺｰﾄﾞ` but inserts `[114,30]=G5).ﾌｪｰｽﾞ工程ｺｰﾄﾞ` — if they
  differ, rows are re-inserted that were never deleted (PK violation). Report a mismatch at 中.
- **A column naming two operations** (`DELETE　INSERT`): apply the `[KEY]` rule to it only if some
  value in it carries `[KEY]`. Usually none does — the column holds the INSERT values and the
  DELETE's WHERE is in 更新条件; read it there, and report once (中) only if it is not stated. (The
  example formerly cited here, `PXJCO125` `TXJCA205` `[12,16]`, is not one: its DELETE is struck and the
  live label is `　INSERT`. No live example re-verified — 要確認.)
- **Repeated operation columns are independent statements by default** (`INSERT(更新ﾎﾞﾀﾝ押下時)` /
  `INSERT(削除ﾎﾞﾀﾝ押下時)`, `DELETE①`/`DELETE②`, `TXJCM007` DELETE1/DELETE2 each with its own key) —
  review each fully. Treat a 2nd+ column as **differential** only when a label/footnote says so, or it
  leaves `-` in common columns (会社ｺｰﾄﾞ/更新者/更新日時) that the first column sets; then don't read
  those `-` as "not set", and report the notation once per block at 低 if no footnote explains it.
- A `[KEY]` on a column that is **not** in `I01` is a filter (often a literal discriminator such as
  `[KEY]"1"(分割)`) — fold it into the partial-key judgement above; don't report it on its own.

**C4 — 共通項目の慣習違反.** Compare each block's common-column rows against the calibration from
step 3. The finding shape that actually shows up: **`登録者`/`登録日時` being set on UPDATE.** This is
not hypothetical — `更新条件表(TXJCM137WF)` in `PXJCO192` sets both on its UPDATE
(`[16,36]=作業者ｺｰﾄﾞ`, `[17,36]=ｼｽﾃﾑ日時(...)`) while every other UPDATE block in the same workbook,
including `TXJCM501`'s and `TXJCM137`'s, correctly leaves them `-`. Overwriting 登録日時 on every
update destroys the record's creation time. Also check: 排他ﾌﾗｸﾞ set to a non-incrementing literal
on UPDATE (the optimistic-lock counter never advances), 更新者/更新日時 left `-` on an UPDATE, and
**更新ﾌﾟﾛｸﾞﾗﾑID** set to anything but the workbook's dominant value (`PXJCO125` `TSJCA006` `[21,18]` uses
`ﾌﾟﾛｸﾞﾗﾑID` where the other 9 INSERT blocks use `画面ID` — 低, 要確認).

A full run of C1-C7 (measured before C8 existed) over `PXJCO192`'s seven live 更新条件表 sheets produced exactly three findings —
this one, the `TXJCM501` missing columns under C1, and the `TSJCM139WF` name mismatch — with no
C2 or C3 hits. That ratio is the target: this check is meant to be quiet on a clean sheet.

**C5 — 排他制御が更新と噛み合っているか.** If the table has an `排他ﾌﾗｸﾞ` column and the block has an
UPDATE, the optimistic-lock comparison should be specified somewhere, not just the `+1`. **In this
project it lives in ﾁｪｯｸ処理設計書** (`排他ﾁｪｯｸ`, message `XJZ-000020`) — accept that, or the 更新条件, or
the 機能定義書 overview; reading only the 更新条件 produced 42 false findings over three workbooks.
**Apply C5 only when the screen holds (or re-reads) that table's `排他ﾌﾗｸﾞ` for the rows the UPDATE
keys on** — typically a Hidden item in 画面項目定義 (`GXJC125A` `[525,5]=排他ﾌﾗｸﾞ`, 検索時 `(6)②.排他ﾌﾗｸﾞ`,
backing `TXJCM006` UPDATE①). A server-side UPDATE of rows the screen never loaded has nothing to
compare: on `PXJCO125` 8 such blocks (`TXJCM003`×3, `TXJCM004`×2, `TSJCD101`, `TXJAM068`×2) would all
have been false. `design-doc-internal-consistency`'s check 7 asks whether exclusive
control is mentioned *at all* for a program that updates; this check is the column-level follow-up —
report only what that check wouldn't already have said, and say plainly that it's the 更新条件表 side
of the same concern so the reviewer doesn't see it as two separate defects.

**C6 — 更新条件（トリガ）が書かれていて、画面イベントと対応しているか.** An empty `更新条件` cell is a
finding on its own. When it names a trigger (`保存ﾎﾞﾀﾝ押下時`, `一時保存ﾎﾞﾀﾝ押下時`), that event should
exist in the 画面設計書's event list for the screen the 更新概要 names. A trigger naming a button the
screen doesn't have is a stale copy-paste from another program's sheet.

**C7 — 参照元記号が定義されているか.** Every `①②③…` used in a value or source cell must be defined in
that block's own `更新概要` cell. A dangling `⑥` (used in `TXJCM501`'s UPDATE) means the reader can't
tell where the value comes from. Symbols are **block-scoped** — check against the block's own
更新概要, not the sheet's first one.

**C8 — 日時系の値に書式注記が付いているか.** Scope this check to `更新条件表(<TableID>)` sheets only.
A value cell that supplies a system timestamp must read `ｼｽﾃﾑ日時(YYYY/MM/DD HH24:MI:SS形式)` — the
bare `ｼｽﾃﾑ日時` with no format annotation is a finding. The annotation is what tells the coder which
`TO_CHAR`/`TO_DATE` format model to write; without it the column's precision (does it carry the time
part, or only the date?) is left to the implementer to guess, and two programs writing the same
column can diverge. Typical location is the 共通項目 block's `登録日時`/`更新日時` value cells
(`[17,18]`, `[17,36]` and their per-block equivalents), but apply it to **every** value cell in the
sheet whose content is `ｼｽﾃﾑ日時`, common column or not.

Unlike C4, this one is **not** calibrated against the workbook's dominant pattern: the annotated form
is always correct, so flag every bare `ｼｽﾃﾑ日時` even in a workbook where most sheets omit it. Match
on the value cell's full content — `ｼｽﾃﾑ日時` is half-width katakana (`ｼ ｽ ﾃ ﾑ`), and a cell already
carrying any parenthesized 形式 note is fine. A cell that merely *contains* `ｼｽﾃﾑ日時` inside a longer
sentence (a `※` footnote, a free-text 更新条件) is not a value cell — don't flag it.

Report it as 書式注記漏れ, one line per cell, and say the correct form explicitly so the designer can
paste it. This is **not** the same as `design-doc-internal-consistency`'s check 5: that one checks
letter-casing in 画面項目定義's 表示形式 column and explicitly does *not* touch 更新条件表's
`(YYYY/MM/DD HH24:MI:SS形式)` — the Oracle format model there is case-insensitive. Casing inside the
annotation is still never a finding; only its absence is.

Measured on `PXJCO192_処置指示登録.xlsx`, where all nine 更新条件表 sheets were scanned: 22 cells carry
the annotated form and 12 carry the bare one, so this is a live, common defect rather than a
hypothetical. The split runs almost exactly along the WF/non-WF line — every `*WF` sheet
(`TXJCM137WF`, `TSJCM139WF`, `TXJAM025WF`, `TXJAM026WF`, `TXJCM838WF`) is fully annotated, while the
plain sheets `TXJAM025` `[17,36]`/`[19,36]`, `TXJAM026` `[17,36]`/`[19,36]`, `TXJCM838`
`[17,36]`/`[19,36]` and `TXJCM501` `[17,18]` are bare. `TXJCM501` is the clearest case: its own
`[19,18]`/`[19,36]`/`[41,18]`/`[42,36]` are annotated and only `[17,18]` is not, so the sheet
contradicts itself.

**Run this check against the live dump, not a raw cell read.** Five more bare cells sit in
`更新条件表(TXJCM137)`, whose whole sheet is struck out and withdrawn — a raw `Value2` scan surfaces
them and a reviewer working from the shared dump will not see them at all. They are not findings.

**C9 — 更新概要の書式「①画面(画面ID)、②ﾛｸﾞｲﾝ情報」.** Call-out on the self-check workbook's
`更新条件表(TXJAM007)`: "①画面(画面ID)、②ﾛｸﾞｲﾝ情報 で記載すること。順番が逆や、ﾛｸﾞｲﾝ情報ではなくﾛｸﾞｲﾝは
ﾀﾞﾒです". Read each block's `更新概要` value (col 16, first line). Find the label by **prefix**: the
cell reads `更新概要⏎（仕様の背景なども記述）`, sometimes with a half-width `(`, so an exact match finds
nothing. Calibrated on 2026-10-01 over 1,053
blocks in 120 工程管理 PHASE1-3 workbooks; 634 are exactly `①画面(<画面ID>)、②ﾛｸﾞｲﾝ情報[、③…]`.
**Apply the rule only when the block's sources include the screen and the login info.** Many blocks
legitimately take values from elsewhere — `①承認処理API引数`, `①引戻し処理API引数`, `①ﾜｰｸﾌﾛｰからの引数`,
`①機能定義書(PSJCB301).Ⅳ．機能処理概要.(2)、②起動ﾊﾟﾗﾒｰﾀ` (batch), `①製品仕様ﾏｽﾀWF(削除対象)` — and are
not findings. Flag:
- **画面 without its ID** — `①画面、②ﾛｸﾞｲﾝ情報` (87 blocks, e.g. `PSJAO401` `更新条件表(TXJAM023)` `[9,16]`,
  most of `PSJCO503`): 低, name the 画面ID from the trigger/sheet context. A screen ID with a tab suffix
  (`画面(GSJA241B-工程情報)`) or two IDs (`画面(GXJC101A,GXJC101B)`) is fine;
- **order reversed** — `①ﾛｸﾞｲﾝ情報、②画面(GSJA241A)` (44 blocks, nearly all `PSJAO241`): 低. The ①②
  symbols are used in the value-source column, so when fixing the order the designer must renumber
  every source cell of that block too — say so in the finding;
- **`ﾛｸﾞｲﾝ` not followed by `情報`** — none in the current corpus, but the call-out names it explicitly;
  keep it.
Report one line per sheet when every block on it has the same deviation.

### 5. Verify before reporting

For each candidate finding, re-read the source cells in the dump. The three things that have
produced wrong findings on this shape of sheet:

- reading the **source** column instead of the value column (`c` instead of `c+2`) and reporting
  every system-supplied column as unset;
- anchoring on row 14 and missing the second and third blocks — or worse, comparing block 1's
  operations against block 2's rows;
- treating a right-hand 取得ﾃｰﾌﾞﾙ/検索条件 sub-block's header cells as operation columns.

## Reporting

Write in the user's language, for the designer who owns the doc.

Group by 更新条件表 sheet, and within a sheet by block when there is more than one. For each finding:
name the table and the column, cite the sheet and cell (`更新条件表(TXJCM137WF) [17,36]`) with a
clickable `[filename](relative/path)` link to the workbook, and state the consequence in one clause
(NOT NULL 違反で登録できない / 登録日時が更新で上書きされる / 主キーの一部しか指定されておらず複数行が
更新される). Order by severity: C2 and C3 partial-key first, then C4, then the rest.

Omit: sheets that checked out clean, a count of how many sheets/columns were compared, and any
narration of which files you dumped or resolved. The one process fact worth a line is a
更新ﾃｰﾌﾞﾙ whose layout file couldn't be resolved at all — that's a coverage gap, and say which table.
