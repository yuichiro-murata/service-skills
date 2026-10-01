# xlsx-db-column-check — TXJAM100 (検索条件保存ﾏｽﾀ) checks

Loaded by `xlsx-db-column-check` step 6 only when a 更新条件表 targets TXJAM100. Columns and dump
conventions are those of the main SKILL.md (steps 3 and 5).

### 6. Redundant code+name persistence in 検索条件保存マスタ-style generic tables

Some common tables in this project (confirmed for `TXJAM100`:検索条件保存ﾏｽﾀ, a shared table that
persists each user's last-used search filters across many different programs' screens) are
generic-column designs — their real ﾃｰﾌﾞﾙﾚｲｱｳﾄ has no per-field schema at all, just a long run of
polymorphic `項目1`...`項目N` columns (150 of them for TXJAM100), each holding whatever a given
program's 更新条件表 decides to put there. **The project's own design rule for this kind of table:
when a search field is a code that references a real master entity (品目ｺｰﾄﾞ, 作業場ｺｰﾄﾞ, 工程ｺｰﾄﾞ,
取引先ｺｰﾄﾞ, etc.), only the code itself should be persisted into the generic table — the
corresponding display name (KC品名, 作業場名, 工程名, 取引先名, etc.) must be re-fetched via a JOIN
against the real master table at read-time, never stored redundantly alongside the code.** A
classification/区分 code whose paired label comes from 共通ｺｰﾄﾞﾏｽﾀ is a different, correctly-handled
case — this project's own convention there is to persist the code alone and annotate it
"(区分値のみ)" in the value-source column, which is *not* a violation (it's the same discipline
applied correctly).

**Detection procedure** (run this specifically whenever a 更新条件表 targets `TXJAM100` or another
table you've confirmed follows the same generic-column/検索条件保存 convention): for every row whose
INSERT/UPDATE column shows a real, non-`-` source (i.e., it's actually being persisted), extract the
plain-text field label from the **value-source cell immediately right of that operation marker**
(e.g. `G1)作業場名` → `作業場名`; strip the leading `<画面エリア>)` prefix first). As step 5 notes,
there is **no column headed 取得内容 on a 更新条件表 sheet** — the source cell is unlabelled and is
located positionally from the `INSERT`/`UPDATE`/`DELETE` header (col 18 on TXJCM501; col 18 for
DELETE and col 36 for INSERT on TXJAM100). Classify each label as **code-type** (ends in `ｺｰﾄﾞ`/`CD`) or
**name-type** (ends in `名`/`名称`/`略式名`/`品名`, or is a known name field like `KC品名`). For every
code-type label, strip the code suffix to get its entity stem (e.g. `作業場ｺｰﾄﾞ` → `作業場`) and check
whether any name-type label in the same 更新条件表 shares that stem (`作業場` → `作業場名`) — or, for
the `品目`/`KC品名` case specifically, treat `KC品名` as `品目`'s paired name even though the surface
text doesn't share the stem literally (this pairing is specific to this project's terminology: `KC`
prefix names are this project's product-name field for a 品目ｺｰﾄﾞ). A code+name pair that are **both**
actually persisted (neither is `-`) is a *candidate*, not yet a finding — it must still clear the
control-type gate below. A code-type label whose paired name is annotated "(区分値のみ)" or has no
persisted name-type sibling at all is correctly designed — don't flag it.

**MANDATORY control-type gate — a persisted name is only a violation when the screen shows that name
as a Label.** Per explicit user direction, this gate decides the finding and there is no exception to
it. Before reporting any candidate pair, look the **name** item up by 画面項目名 in that screen's
`Ⅴ．画面項目定義` and read its 属性 (control-type) column — in the standard layout it is array column
18, with `画面項目名` at column 5 and 初期値 at column 36; resolve it by each Ⅴ area's own header-row
labels rather than trusting those positions:

- **属性 = `Label`** (or otherwise display-only) → the name is *derived from the code* via a master
  lookup and has no independent existence. Persisting it duplicates master data. **This is the
  finding.**
- **属性 = `TextBox`** (or any editable input: ComboBox a user picks, etc.) → the name is **its own
  independent search condition** that the user types, not a master-derived echo of the code. The
  program must persist it, because it is part of the filter the user entered and there is nothing to
  JOIN it back from — a partial/LIKE name search has no code to re-derive it. **Persisting it is
  correct. Do not flag it.**

The surface text is identical in both cases (`工程ｺｰﾄﾞ` + `工程名` persisted side by side), so a
name-pairing scan alone cannot tell a real defect from correct design. Skipping the gate turns every
screen that offers name-based text search into a page of false findings.

Confirmed for real on `PXJCO128_ﾛｯﾄ停止指示登録.xlsx`, where the gate flips the verdict on all four
candidates and the design turns out to be **entirely correct and internally consistent**:

| 画面項目 (`画面設計書(GXJC128A)` Ⅴ) | 属性 | TXJAM100 に保存? | 判定 |
|---|---|---|---|
| ﾛｯﾄ停止指示者名 `[483,5]` | TextBox | 保存 `[40,36]` | 正 — 独立した検索条件 |
| 工程名 `[486,5]` | TextBox | 保存 `[42,36]` | 正 — 独立した検索条件 |
| 停止者名 `[491,5]` | TextBox | 保存 `[45,36]` | 正 — 独立した検索条件 |
| 解除者名 `[500,5]` | TextBox | 保存 `[50,36]` | 正 — 独立した検索条件 |
| KC品名 `[487,5]` | TextBox | 保存 `[43,36]` | 正 — 品目ｺｰﾄﾞの派生名ではなく独立入力条件 |
| 停止工程GRP名 `[478,5]` | **Label** | 保存せず | 正 — ｺｰﾄﾞのみ保存 |
| 取引先名 `[507,5]` | **Label** | 保存せず | 正 — ※1復元表 `[538,17]` が `(2).取引先略式名` から再取得 |

The two Labels are exactly the ones left out of the persisted set, and `取引先名`'s restore rule
points at a delegated master lookup rather than a saved 項目N — the read-time-JOIN discipline this
rule is about, applied correctly. A review that reported the four TextBox names as redundant
persistence (as one did before this gate was written) is producing false findings on a correct
design.

Always re-check a candidate pair's 属性 before citing it, and cite the 属性 in the finding itself so
the designer can see the gate was applied.

**This gate is consistent with the self-check workbook — it is not a conflict.** `XXXXX000_ｾﾙﾌﾁｪｯｸ用.xlsx`
レビュー観点 No.47 says "TXJAM100_検索条件保存マスタを利用している場合、検索条件の項目が全てあるか。
名称も保存する。" Confirmed with the user (2026-10-01): **the 名称 there means a name the user enters as
a search condition (TextBox etc.), and does NOT include a Label name.** So No.47 and this gate say the
same thing — save every search condition the user enters, names included; never save a Label name
that is only a display of the code. Do not re-raise it as a skill-vs-checklist conflict.

#### 6b. The other half of No.47 — every search condition is saved

The gate above only catches *too much* being saved. No.47's first half — "検索条件の項目が全てあるか" —
is the opposite direction, a **missing** save, and nothing else in the REV set checks it. A condition the
user typed that is not saved comes back blank the next time the screen opens, while its neighbours are
restored: a quiet, user-visible defect.

For every 更新条件表 that targets `TXJAM100`:

1. **Find the screen.** Only screens named by a TXJAM100 block's 画面ID are in scope — a sibling
   screen of the same program that never saves to TXJAM100 is not.
    The `画面ID` row's value (`[KEY]"GXJC128A"`, `KEY:"GXJC131A"`, `"GSJC502A"`)
   names it. **Union every TXJAM100 block that carries the same 画面ID** before comparing —
   `PXJCO131` writes `GXJC131A` across three DELETE→INSERT blocks, and comparing block by block
   reports almost every item as missing.
2. **Collect the saved names**: the INSERT value cell of each `項目N` row whose source is not `-`.
   Normalise before matching — strip whitespace, a leading `G1)` / `(1).` area prefix, a trailing `※n`.
   A grouped name counts as saving each member, in **both** directions: `移動ﾛｯﾄ状態(作業前) ※1` saves
   the CheckBox `作業前` (group label outside, item inside — `PXJCO131` saves 28 status checkboxes this
   way), and `ｻﾝﾌﾟﾙ区分(通常)` / `(ｻﾝﾌﾟﾙ)` / `(先行)` save the CheckBox `ｻﾝﾌﾟﾙ区分` (item outside, its
   options inside — `PXJCO129`). So a saved name matches an item when they are equal, when the item
   appears inside the saved name's parentheses, or when the saved name minus a trailing `(…)` equals it.
   Keep the area prefix for one extra comparison before stripping it: a saved `G2)送付先作業場ｺｰﾄﾞ` for
   an item that lives in `G1)検索条件領域` (`PSJCO502`, where `G2)` is the 明細) is a low-severity
   prefix error worth a line.
3. **Collect the search conditions**: in that screen's Ⅴ．画面項目定義, the input items (`TextBox`,
   `ComboBox`, `CheckBox`, `RadioButton`, `TextArea`, `ListBox`) of every area whose title contains
   `検索条件`. Never Labels, Buttons, Accordions or zoom buttons — and per the decision above, a Label
   name is never expected to be saved.
4. **Exempt an item that is reset on purpose**: its 初期値 is derived from the system date
   (`ｼｽﾃﾑ日付-(2).日数`, `ｼｽﾃﾑ年月日(YYYY/MM/DD)-1ヶ月`). Not saving it is the design — it starts from
   today every time. Exempt the **pair**: when 発送日(From) is system-date-initialised, 発送日(To) is
   not saved either (`PSJCO501`/`PSJCO502`), and that is the same decision.
5. Report each remaining condition that no saved name covers.
6. **Cross-check against the restore table.** The screen's 初期値 note (`※1`/`※2` 復元表) maps each
   restored item to a slot — `KC品名 ← (13).項目19`. For every row of that table, the TXJAM100 INSERT
   value of `項目K` must be the same item. Also flag one item saved into **two** slots. Confirmed on
   `PXJCO129`: the restore table `[672,7]`/`[672,21]` reads 項目19 as KC品名, but 項目19 `[46,36]` was
   rewritten in place to `取引先ｺｰﾄﾞ` (struck `表示順`, live `取引先ｺｰﾄﾞ`), which is already in 項目22
   `[49,36]` — when the old KC品名 row `[44,36]` was repurposed, KC品名 dropped out. The screen then
   restores a 取引先ｺｰﾄﾞ into the KC品名 box. **A slot mismatch is 高**: it restores the wrong value,
   which is worse than a blank. The plain "not saved" finding for the same item merges into it.

Severity for the redundant-persistence gate itself (a Label name saved): **中** — redundant master
data, usually never read back (`PSJCO502`'s restore table re-fetches `(3).作業場名` from the master
and ignores the saved 項目4/項目6). Confirmed live on `PSJCO502`: TXJAM100 項目4 `[31,36]`
`送付先作業場名` and 項目6 `[33,36]` `送付元作業場名` are both Label on `GSJC502A` (`[428,17]`/`[431,17]`).

Scope note: step 3 keys on the area title containing `検索条件`. An input area named otherwise
(`PXJCO129` `G1)管理者ｺｰﾄﾞ指定領域`) is out of scope unless its items are themselves saved to TXJAM100
— then treat that area as a search-condition area too.

Measured on 2026-09-30 / 10-01 across the 34 工程管理 PHASE1-3 screens that save to TXJAM100: 439 of
469 search conditions are saved. Of the 30 left, grade by kind:
- **code/name inputs → 中** (6): `PXJCO129` `GXJC129A` KC品名 `[637]` (every other condition on that
  screen is saved — and it merges into step 6's slot-mismatch 高 below), `PSJAO502` `GSJA502A` 単位ｺｰﾄﾞ `[260]`, `PXJCO164` `GXJC164A` 作業者名 `[714]`,
  `PSJCO309` `GSJC309A` 作業場ｺｰﾄﾞ/作業者ｺｰﾄﾞ/焼成条件 (that screen saves only 3 items, so ask whether
  the omission is deliberate before calling it a defect);
- **date/time ranges and range-type numbers (案内書No(From)/(To)) → 要確認** (20): 67 date
  conditions are saved and 20 are not, so neither is "the rule"; when the item's 初期値 is `※n`, read
  the note — if it says the value is restored from 検索条件保存ﾏｽﾀ, the missing save is a definite 中;
- **CheckBox/RadioButton modifiers → 要確認** (4): `作業場条件設定`, `検索範囲`,
  `出荷選択済のみ表示` — option switches a designer may choose to reset.
Cite the screen row and the TXJAM100 sheet together, and say which neighbouring conditions *are*
saved, so the asymmetry is visible.

**`PSJCO304_着手ﾒｯｾｰｼﾞﾒﾝﾃﾅﾝｽ.xlsx` used to be this document's worked violation example, and it is
no longer one — the defect was remediated.** As written here, 更新条件表(TXJAM100) persisted
`作業場ｺｰﾄﾞ`+`作業場名`, `工程ｺｰﾄﾞ`+`工程名`, `取引先ｺｰﾄﾞ`+`取引先名` and `品目ｺｰﾄﾞ`+`KC品名` side by
side. The 2026/9/4 NCRN-8552 revision removed the three name columns from the persisted set; the
三つの名称 are now Label 属性 and are restored by read-time JOIN via the ※1 復元表, and the surviving
`KC品名` is a TextBox — an independent partial-match search condition, which the gate above passes as
correct. Re-verified on a REV of that workbook on 2026-09-14: **zero** redundant-persistence findings.

Take the general lesson, not just the correction: **a violation example named in a skill file is a
snapshot of one workbook at one moment, and these design docs are actively revised.** Never cite a
stored example as a live finding. Re-derive the verdict from the current dump every time, and if you
find an example here has been fixed, update this file (and bump the plugin `version`, or every
cached copy stays stale forever).

