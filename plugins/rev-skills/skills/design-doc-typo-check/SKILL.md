---
name: design-doc-typo-check
description: Proofread the free-text Japanese prose inside a program's design-doc workbook — feature/processing narratives, event descriptions, error-message text, footnotes, revision-history notes — for actual language mistakes (誤字/脱字/衍字/助詞の誤り/変換ミス), as distinct from its sibling skills' structural checks: cross-reference completeness (design-doc-internal-consistency), I/O table completeness (design-doc-io-table-check), DB column existence (xlsx-db-column-check), ID-numbering rules (naming-standard-compliance), and font/merge irregularities (design-doc-formatting-consistency). Use when the user asks to check 誤字脱字 or wants a language-proofreading pass on a design doc. For a full REV without named checks the entry point is `rev-program-review`; run this standalone only when selected there or asked for by name.
---

# design-doc-typo-check

Proofreads one program's design-doc workbook for genuine Japanese-language errors in its free-text
prose — wrong kanji, a missing character that breaks the sentence, a stray duplicated character, a
wrong particle that flips the meaning, a garbled leftover fragment from a copy-paste edit. This is
fundamentally different from every sibling REV skill: those are mechanical diffing/lookup tasks;
this one requires actually reading each sentence and understanding what it says. Do not try to
shortcut it with keyword search or pattern matching — read the prose.

This skill does **not** check cross-document terminology consistency, ID format, or font/merge
signals — see the frontmatter `description` above for which sibling skill covers each. A formatting
anomaly `design-doc-formatting-consistency` flags is often a good hint of *where* a copy-paste-era
typo might also be hiding, though, so it's worth cross-referencing that skill's findings if run in
the same pass.

## Environment

**Read `_shared/agent-guide.md` first** — scope selection, environment, the live dump, reporting
conventions. This check reads only the target workbook's dump: everything in the `.txt` is live
prose, and deleted-in-spirit text was never handed to you to proofread.

This matters more to proofreading than to any sibling check, because the artefact it removes looks
exactly like this skill's own quarry. A partially-struck cell arrives already reduced to its live
text, so a cell that would have read as a **garbled leftover fragment** in a raw dump now reads
correctly. Judge what you see; do not reconstruct the original from `_DELETED_DIGEST.txt` and
proofread that. Confirmed false findings from raw dumps: `"GSXJC205A"` reported as a misspelling of
`GSJC205A` (only the `X` was struck), and `注意事項備考` reported as a nonexistent term (only
`注意事項` was live).


## Scope: which cells are "prose" here

Structured tokens — IDs, table/column names, `<alias>.<column>` expressions, numeric codes, half-width
katakana terms that are this project's standard house style (`ﾃﾞｰﾀ`, `ﾛｯﾄ`, `ｵｰﾀﾞｰ`, `ｺｰﾄﾞ`, etc.) —
are NOT prose and are NOT this skill's concern; those belong to the other REV skills or are simply
normal project vocabulary. Only free-running Japanese sentences and phrases are in scope. In a typical
program workbook, that means:

- **表紙**: Ⅲ．改訂履歴 の 改訂内容 column (what each revision actually changed).
- **機能定義書**: Ⅰ．機能概要 (prose bullets), Ⅱ．機能要求事項, Ⅲ．入出力定義 の 用途 column
  (short prose), Ⅳ．機能処理概要 (the main narrative — numbered steps, ※ footnotes, and any prose
  woven into 検索条件/取得内容 descriptions), Ⅵ．前提条件, Ⅷ．その他特記事項.
- **画面設計書**: Ⅲ．画面表示仕様 の ※ footnotes, Ⅳ．画面項目ｲﾍﾞﾝﾄ詳細 (event-behavior
  descriptions — often long paragraphs), Ⅴ．画面項目定義 の 説明 column (col 43 — there is no 備考
  column), Ⅵ．画面項目制御・出力仕様 の 条件 prose.
- **ﾁｪｯｸ処理設計書**: the `ﾁｪｯｸ詳細` and `ﾁｪｯｸ補足` columns, error-message body text.
- **ﾌｧｲﾙ出力仕様書** (if present): 出力条件 / 編集内容 / 備考 prose.
- **更新条件表**: the 更新概要 prose at the top of the sheet, and ※-numbered footnotes (these are
  written in full sentences, unlike the 項目名/取得内容 columns which are mostly structured tokens).
  A structured token is still in scope for an **obvious** misspelling (`ﾛｸﾞｲﾝ画面.` with a stray
  period, `作場名` for `作業場名`) — skip only style and abbreviation questions there.
- **帳票設計書** (if present): 処理の流れ (numbered narrative steps), 備考 column.

## Procedure

1. Use the dump `rev-program-review` handed you; standalone, run
   `python _shared/scripts/live_dump.py <workbook> <out_dir>` (it skips 詳細設計書).
2. For each section listed above, read every cell's full text — sentence by sentence, not a keyword
   scan. Look for:
   - **誤字**: a wrong kanji/character that doesn't fit the intended word (a likely IME conversion
     slip, e.g. a homophone substituted for the correct word).
   - **脱字**: a missing character or word that makes the sentence ungrammatical or leaves an action
     without its object/verb.
   - **衍字**: an accidentally duplicated character or word (e.g. "確認するする").
   - **助詞の誤り**: a wrong particle that changes or breaks the meaning (e.g. "を" where "が" was
     needed, an action that reads backwards from what's clearly intended).
   - **Unmatched brackets/parentheses/quotes** within one cell.
   - **Garbled leftover fragments**: a word or clause that doesn't belong to the sentence it's sitting
     in — often the tell-tale sign of a copy-paste edit where the old text wasn't fully replaced.
3. Deprecated rows were already dropped at dump time — do not scan for them, and do not treat a
   sentence that reads oddly as evidence of a leftover you need to go verify in Excel.
4. Cross-check every candidate against context before reporting it as a real error:
   - Is this actually a deliberate project term or abbreviation, not a typo? (Check whether the same
     spelling appears consistently elsewhere in the project — if so, it's house style, not an error.)
   - Does a near-identical sentence exist elsewhere in the same doc (a sibling row in a repeating
     list, a parallel branch for a different condition)? Diffing against that twin is often the
     fastest way to confirm a real error versus a deliberate variation.
   - Could ambiguous phrasing be intentional shorthand common in this project's docs, rather than a
     mistake? If genuinely unsure, report it as a lower-confidence note rather than a confirmed error
     (see Reporting).

## Reporting

The output goes to the designer who owns the doc, not into your own working notes — report only
what they'd actually need to act on. Write in Japanese.

For each real finding: cite the exact sheet name and cell coordinate, quote the erroneous text
verbatim, state plainly what's wrong (誤字/脱字/衍字/助詞の誤り and which), and give the corrected
text when it's obvious. Order by impact: an error that changes or obscures the meaning of a processing
rule or condition belongs before a cosmetic misspelling that doesn't affect understanding. Distinguish
confirmed errors from "this reads oddly but might be intentional shorthand" judgment calls, and mark
the latter as needing the designer's confirmation rather than asserting it as a defect.

Omit entirely: a tally of how many cells/sections were proofread, "確認済み・問題なし" notes for
clean sections, and narration of your own process (which files you dumped, how you sampled). The one
exception worth a one-line mention: a whole prose section that couldn't be read at all (e.g. it was
entirely struck-through, or a sheet's used range exceeded the dump cap) — that's a real coverage gap
worth flagging, not a process detail.
