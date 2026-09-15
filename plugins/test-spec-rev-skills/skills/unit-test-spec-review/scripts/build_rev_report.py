# -*- coding: utf-8 -*-
"""単体テスト仕様書40件のREV結果を1つのExcelにまとめる。

unit-test-spec-review スキルのチェック1〜7を全ファイルに適用し、
指摘を「ファイル名でオートフィルタできる1枚の表」として出力する。
"""
import glob, os, re, datetime, collections, warnings
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

warnings.filterwarnings('ignore')
SRC = r"C:\Users\200451023\Downloads\単体テスト仕様2"
OUT = os.path.join(SRC, "REV指摘一覧_単体テスト仕様_V1.23.0.xlsx")
NOTRUN = (None, '', '-', '不可')


def asdate(v):
    return v.date() if isinstance(v, datetime.datetime) else v


def leaves_of(ws):
    return [r for r in range(34, ws.max_row + 1)
            if ws.cell(r, 1).value and str(ws.cell(r, 1).value).count('-') == 2]


findings = []   # dict rows


def add(fname, level, view, cell, body, note, kind):
    findings.append(dict(file=fname, level=level, view=view, cell=cell,
                         body=body, note=note, kind=kind))


files = sorted(glob.glob(os.path.join(SRC, '*.xlsx')))
files = [f for f in files if not os.path.basename(f).startswith('REV指摘一覧')]

# ---- 表記揺れの多数派を決めるための事前集計 ----
wording = collections.defaultdict(collections.Counter)
for path in files:
    ws = openpyxl.load_workbook(path, data_only=True)['ﾃｽﾄ仕様']
    for r in leaves_of(ws):
        no = str(ws.cell(r, 1).value)
        if no in ('S5-1-1', 'S1-1-6'):
            wording[no][str(ws.cell(r, 2).value).split('\n')[0]] += 1
major = {k: c.most_common(1)[0][0] for k, c in wording.items()}

meta = {}
for path in files:
    name = os.path.basename(path)
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb['ﾃｽﾄ仕様']
    manual = '手動移行_' in name
    m = re.search(r'(T[SX]J\w+?_\d{3})', name)
    tid = m.group(1) if m else ''
    lv = leaves_of(ws)
    judged = [r for r in lv if ws.cell(r, 16).value not in NOTRUN]
    ng = [r for r in lv if ws.cell(r, 16).value == 'NG']
    fixed = [r for r in ng if ws.cell(r, 18).value == 'OK']
    n3, q3 = asdate(ws['N3'].value), asdate(ws['Q3'].value)
    meta[name] = dict(tid=tid, manual=manual, leaves=len(lv), judged=len(judged),
                      ng=len(ng), e10=ws['E10'].value, n10=ws['N10'].value,
                      n3=n3, q3=q3)

    # --- 1. 判定「不可」の理由なし ---
    for r in lv:
        if ws.cell(r, 16).value == '不可':
            reason = ws.cell(r, 19).value
            if not reason:
                add(name, '重', '1 判定「不可」の理由なし', 'P%d' % r,
                    '%s「%s」の判定が「不可」だが、実施できなかった理由が備考欄(S列)にも'
                    'S1〜S16の未実施理由欄にも記載されていない。'
                    % (ws.cell(r, 1).value, str(ws.cell(r, 2).value).split('\n')[0]),
                    '「-」(テスト対象なし)と違い「不可」は対象があるのに実施できなかった意。'
                    '理由を S%d に記載するか、実施して判定を入れる必要がある。' % r, '個別')

    # --- 2. テスト件数・障害件数の整合 ---
    exp_n10 = 0 if not ng else '%d→%d' % (len(ng), len(ng) - len(fixed))
    if ws['E10'].value != len(judged):
        add(name, '重', '2 テスト件数の不一致', 'E10',
            'テスト件数 %s に対し、判定が記入されたテスト項目は %d 件。'
            % (ws['E10'].value, len(judged)),
            'テスト件数＝判定記入済み項目数（「-」「不可」を除く）。', '個別')
    if str(ws['N10'].value) not in (str(exp_n10), '0→0'):
        add(name, '重', '2 障害件数の不一致', 'N10',
            '障害件数 %s に対し、NG判定は %d 件（うち再テストOK %d 件）。'
            % (ws['N10'].value, len(ng), len(fixed)), '', '個別')

    # --- 3. 要／不要フラグと理由欄 ---
    for r in range(6, 10):
        label = ws.cell(r, 1).value
        req, notreq, rsn = ws.cell(r, 27).value, ws.cell(r, 28).value, ws.cell(r, 7).value
        if notreq is True and (rsn in (None, '', '-')):
            add(name, '重', '3 要／不要と理由欄の不整合', 'G%d' % r,
                '%s が「不要」だが、不要理由が記載されていない。' % label, '', '個別')
        if req is True and rsn not in (None, '', '-'):
            add(name, '中', '3 要／不要と理由欄の不整合', 'G%d' % r,
                '%s は「要」なのに不要理由欄に「%s」と記載がある。' % (label, rsn),
                '要／不要の選択と理由欄のどちらかが誤りの可能性。', '個別')
    for r in range(17, 33):
        if str(ws.cell(r, 7).value) == '不要' and not ws.cell(r, 9).value:
            add(name, '中', '3 要／不要と理由欄の不整合', 'I%d' % r,
                '%s が「不要」だが未実施理由が空欄。' % str(ws.cell(r, 1).value).strip(), '', '個別')

    # --- 4. テスト項目本文の誤記・表記揺れ ---
    for r in lv:
        no = str(ws.cell(r, 1).value)
        txt = str(ws.cell(r, 2).value)
        head = txt.split('\n')[0]
        if no == 'S1-3-1' and txt.count(')') > txt.count('('):
            add(name, '中', '4 テスト項目本文の誤記', 'B%d' % r,
                'S1-3-1 の文面で開き括弧が欠落している：「エラー発生時に後続処理が行われること'
                'エラーが発生しても…コミットされること)」。',
                '正しくは「…行われること(エラーが発生しても…コミットされること)」。'
                'テンプレート由来で複数ファイルに共通。', '共通')
        if no in major and head != major[no]:
            add(name, '軽', '4 テスト項目の表記揺れ', 'B%d' % r,
                '%s の文言が他ファイルと異なる（本書「%s」／多数派「%s」）。'
                % (no, head, major[no]),
                '同一No.の項目で文言が %d 種類に分かれている。テスト観点を揃えるか、'
                '差異が意図的なら残す。' % len(wording[no]), '個別')

    # --- 5. 文書種別・体裁 ---
    if str(ws['A1'].value) == '総合テスト仕様書／報告書':
        add(name, '中', '5 文書種別とファイル名の不一致', 'A1',
            '表題が「総合テスト仕様書／報告書」だが、ファイル名は「単体テスト仕様_」で始まる。',
            '同一フォルダの他ファイルも同じ状態であれば、総合テンプレートを流用した運用と思われる。'
            '意図的かどうかの確認のみ。', '共通')

    # --- 6. 日付・担当 ---
    for r in range(34, ws.max_row + 1):
        no = ws.cell(r, 1).value
        if not no:
            continue
        if str(no).count('-') == 1:
            dt = asdate(ws.cell(r, 16).value)
            if isinstance(dt, datetime.date) and n3 and q3 and not (n3 <= dt <= q3):
                add(name, '中', '6 実施日がヘッダの期間外', 'P%d' % r,
                    '%s の実施日 %s が、テスト実施日 %s〜%s の範囲外。'
                    % (no, dt, n3, q3), 'ヘッダの実施日か区分行の実施日のどちらかが誤り。', '個別')
        if str(no).count('-') == 2:
            if ws.cell(r, 16).value not in NOTRUN and ws.cell(r, 15).value in (None, '', '-'):
                add(name, '中', '6 担当の記入漏れ', 'O%d' % r,
                    '%s は判定「%s」が入っているのに担当が空欄（「-」）。'
                    % (no, ws.cell(r, 16).value), '', '個別')
            if ws.cell(r, 16).value == 'NG' and ws.cell(r, 18).value not in ('OK', 'NG'):
                add(name, '重', '6 NG項目の再テスト結果なし', 'R%d' % r,
                    '%s がNGだが、再テスト判定(R列)が記入されていない。' % no, '', '個別')

    # --- 7. テスト環境 ---
    env = str(ws['A14'].value)
    if 'Windows 11' in env and ('要' in (ws['G21'].value, ws['G22'].value)):
        add(name, '中', '7 性能・ﾘｿｰｽ測定の環境記載', 'A14',
            'S5 性能ﾃｽﾄ／S6 ﾘｿｰｽ測定を「要」としているが、テスト環境のOS記載が'
            'クライアントPC（Windows 11 Pro／メモリ8GB）になっている。',
            '他ファイルの多くは Windows Server 2022／96GB。測定をサーバ環境で行ったのであれば'
            '記載を修正、クライアントPCで行ったのであれば測定結果の妥当性の確認が必要。', '個別')
    if '_T1' in str(ws['A12'].value):
        add(name, '軽', '7 接続先DBサーバ名の表記揺れ', 'A12',
            'DBサーバ名が「SJX_T1／MPASC1_T1」表記（他ファイルは「SJX／MPASC1」表記）。',
            '同じIP・SID（10.25.94.212/SENSC101）を指しているため、別環境なのか'
            '表記揺れなのかの確認が必要。', '個別')

# ---------- 出力 ----------
wb = openpyxl.Workbook()
ws = wb.active
ws.title = '指摘一覧'

HDR = ['No.', 'ファイル名', 'テーブルID', '移行方式', '区分', '重要度',
       'チェック観点', 'セル', '指摘内容', '根拠・確認事項']
ws.append(HDR)

order = {'重': 0, '中': 1, '軽': 2}
findings.sort(key=lambda x: (0 if x['kind'] == '個別' else 1, order[x['level']], x['view'], x['file']))
for i, f in enumerate(findings, 1):
    md = meta[f['file']]
    ws.append([i, f['file'], md['tid'], '手動移行' if md['manual'] else 'SQL移行',
               f['kind'], f['level'], f['view'], f['cell'], f['body'], f['note']])

thin = Side(style='thin', color='BFBFBF')
border = Border(left=thin, right=thin, top=thin, bottom=thin)
hdr_fill = PatternFill('solid', fgColor='1F4E78')
for c in ws[1]:
    c.font = Font(bold=True, color='FFFFFF', size=10)
    c.fill = hdr_fill
    c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    c.border = border

lvl_fill = {'重': PatternFill('solid', fgColor='FFC7CE'),
            '中': PatternFill('solid', fgColor='FFEB9C'),
            '軽': PatternFill('solid', fgColor='E2EFDA')}
for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=len(HDR)):
    for c in row:
        c.border = border
        c.font = Font(size=10)
        c.alignment = Alignment(vertical='top', wrap_text=(c.column in (9, 10)))
    row[5].fill = lvl_fill[row[5].value]
    row[5].alignment = Alignment(horizontal='center', vertical='top')
    row[4].alignment = Alignment(horizontal='center', vertical='top')
    row[0].alignment = Alignment(horizontal='center', vertical='top')

widths = [5, 52, 14, 10, 7, 8, 28, 8, 62, 58]
for i, w in enumerate(widths, 1):
    ws.column_dimensions[get_column_letter(i)].width = w
ws.auto_filter.ref = 'A1:%s%d' % (get_column_letter(len(HDR)), ws.max_row)
ws.freeze_panes = 'C2'
ws.row_dimensions[1].height = 30

# ---- サマリ ----
s = wb.create_sheet('ファイル別サマリ')
s.append(['ファイル名', 'テーブルID', '移行方式', 'テスト項目数', '判定済', 'NG',
          'テスト件数(E10)', '障害件数(N10)', '実施日(自)', '実施日(至)',
          '重', '中', '軽', '指摘計'])
per = collections.Counter((f['file'], f['level']) for f in findings)
for name in sorted(meta):
    md = meta[name]
    a, b, c = per[(name, '重')], per[(name, '中')], per[(name, '軽')]
    s.append([name, md['tid'], '手動移行' if md['manual'] else 'SQL移行',
              md['leaves'], md['judged'], md['ng'], md['e10'], md['n10'],
              md['n3'], md['q3'], a, b, c, a + b + c])
for c in s[1]:
    c.font = Font(bold=True, color='FFFFFF', size=10)
    c.fill = hdr_fill
    c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    c.border = border
for row in s.iter_rows(min_row=2, max_row=s.max_row, max_col=14):
    for c in row:
        c.border = border
        c.font = Font(size=10)
for row in s.iter_rows(min_row=2, max_row=s.max_row, min_col=9, max_col=10):
    for c in row:
        c.number_format = 'yyyy/mm/dd'
        c.alignment = Alignment(horizontal='center')
s.column_dimensions['A'].width = 52
for col in 'BCDEFGHIJKLMN':
    s.column_dimensions[col].width = 12
s.auto_filter.ref = 'A1:N%d' % s.max_row
s.freeze_panes = 'B2'
s.row_dimensions[1].height = 30

wb.save(OUT)
print('WROTE', OUT)
print('findings:', len(findings), '(重%d 中%d 軽%d)'
      % (sum(1 for f in findings if f['level'] == '重'),
         sum(1 for f in findings if f['level'] == '中'),
         sum(1 for f in findings if f['level'] == '軽')))
print('個別:', sum(1 for f in findings if f['kind'] == '個別'),
      ' 共通:', sum(1 for f in findings if f['kind'] == '共通'))
for k, v in collections.Counter(f['view'] for f in findings).most_common():
    print('   ', k, v)
