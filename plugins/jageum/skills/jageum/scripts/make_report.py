#!/usr/bin/env python3
"""자금수지보고 만들기: 지난달 보고서(양식) + 위하고 분개장, 통장 잔액 → 이번 달 보고서.
사용: make_report.py <지난달 보고서.xlsx> <분개장.xlsx> <통장거래.xlsx> <연> <월> <저장할 파일.xlsx>
- 양식의 수식, 서식은 그대로 두고 값만 바꿈
- 계좌표: 전일 = 통장 지난달 말 잔액, 입금/출금 = 분개장, 통장시재 = 통장 이번 달 말 잔액
- 자금수지 줄: 분개장 '보통예금' 줄 (적요, 프로젝트 그대로)
- 결과 옆에 점검표(json)를 남김: 차이 나는 계좌, 프로젝트 없는 줄, 분개장에 없는 통장 거래
"""
import copy, json, re, sys
import openpyxl

dig = lambda s: re.sub(r'\D', '', str(s or ''))
ACC = '_-* #,##0_-;\\-* #,##0_-;_-* \\-_-;_-@_-'  # 엑셀 회계 서식 (0은 '-')


def acct_fmt(no, bank):
    """계좌번호 표기: 하나 3-6-5, 기업 3-6-2-3 (보고서에 쓰는 모양)"""
    if len(no) == 14 and '하나' in bank: return f'{no[:3]}-{no[3:9]}-{no[9:]}'
    if len(no) == 14: return f'{no[:3]}-{no[3:9]}-{no[9:11]}-{no[11:]}'
    return no


def load_bj(path):
    ws = openpyxl.load_workbook(path, read_only=True).worksheets[0]
    it = ws.iter_rows(values_only=True); h = next(it)
    return [dict(zip(h, r)) for r in it]


def load_bank(path):
    wb = openpyxl.load_workbook(path, read_only=True)
    def rows(name):
        it = wb[name].iter_rows(values_only=True); h = next(it)
        return [dict(zip(h, r)) for r in it]
    return rows('계좌잔액'), rows('통장거래_' + wb.sheetnames[0].split('_')[1])


def main(tpl, bj_path, bank_path, year, month, out, day=None):
    year, month = int(year), int(month); pm = 12 if month == 1 else month - 1
    ym = f'{year}{month:02d}'
    bj = [r for r in load_bj(bj_path) if str(r['일자']).startswith(ym)]
    accts, bank_tx = load_bank(bank_path)
    code = lambda s: re.match(r'\((\d+)\)', s).group(1)
    acct_no = {code(a['계좌']): dig(re.search(r'\[(\d+)\]', a['계좌']).group(1)) for a in accts}

    wb = openpyxl.load_workbook(tpl)
    top, sub = wb.worksheets[0], wb.worksheets[1]
    # 시트 이름과 수식 안 시트 참조를 이번 달로
    ren = {}
    for ws in wb.worksheets:
        new = re.sub(rf'(?<!\d){pm}월', f'{month}월', ws.title)
        if new != ws.title: ren[ws.title] = new
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith('='):
                    v = c.value
                    for o, n in ren.items(): v = v.replace(o, n)
                    c.value = v
    for o, n in ren.items(): wb[o].title = n
    # 기준월: '(9월1일~9월30일)', '(9월 1일~9월 30일)', '(9월01일~9월30일)' 처럼 지점마다 다른 모양을 그대로 두고 숫자만
    last = int(day) if day else (31 if month in (1, 3, 5, 7, 8, 10, 12) else 30 if month != 2 else (29 if year % 4 == 0 else 28))  # 기준일 (월 중간 '오늘자' 보고서면 그날)
    for row in top.iter_rows(max_row=3):
        for c in row:
            if isinstance(c.value, str) and re.match(r'\(\d+월\s*\d+일\s*~', c.value):
                c.value = re.sub(r'(~\s*)\d+(월\s*)\d+(일)', rf'\g<1>{month}\g<2>{last}\g<3>', re.sub(r'^\(\d+월', f'({month}월', c.value))
    top['K2'] = month

    check = {'계좌': [], '프로젝트없음': [], '통장만': [], '양식에없는계좌': [], '양식계좌번호다름': []}
    is_f = lambda v: isinstance(v, str) and v.startswith('=')
    # 계좌표. 지점마다 양식이 달라서 양식에 있는 수식, 메모는 그대로 두고 값 칸만 채움
    by_no = {dig(re.search(r'\[(\d+)\]', a['계좌']).group(1)): a for a in accts}
    used = set(); acct_rows = []; table = []; after_sub = False; other = set()
    fix_diff = any(isinstance(top.cell(r, 9).value, str) and '+G' in top.cell(r, 9).value for r in range(4, top.max_row + 1))
    for r in range(4, top.max_row + 1):
        label = str(top.cell(r, 1).value or '')
        if '합계' in label:
            if fix_diff: top.cell(r, 9).value = f'=F{r}-G{r}'
            break
        if '소계' in label: after_sub = True
        if '소계' in label:  # 숫자가 박혀 있거나 비어 있는 칸만 합계 수식으로, 범위가 어긋난 SUM은 바로잡음
            for c in range(3, 10):
                v = top.cell(r, c).value; L = openpyxl.utils.get_column_letter(c)
                if v is None or isinstance(v, (int, float)) or (is_f(v) and re.fullmatch(rf'=SUM\({L}\d+:{L}\d+\)', v)):
                    if v is not None or c != 9 or fix_diff: top.cell(r, c).value = f'=SUM({L}{acct_rows[0] if acct_rows else 4}:{L}{r - 1})'
            continue
        no = dig(top.cell(r, 2).value)
        if not no: continue
        hit = [k for k in by_no if k == no or (len(no) >= 5 and k.endswith(no))]
        for n in (8, 7, 6, 5):  # 양식 계좌번호가 틀린 경우 (다른 지점 MMF 번호를 복사해 둔 경우 등): 끝자리로 하나만 맞으면 그 계좌
            if hit: break
            hit = [k for k in by_no if k.endswith(no[-n:])]
            if len(hit) > 1: hit = []
            elif hit: check['양식계좌번호다름'].append({'행': r, '양식': top.cell(r, 2).value, '위하고': by_no[hit[0]]['계좌']})
        if not hit:
            check['계좌'].append({'행': r, '계좌번호': top.cell(r, 2).value, '문제': '위하고에 없는 계좌'}); continue
        a = by_no[hit[0]]; c = code(a['계좌']); used.add(c); acct_rows.append(r)
        if after_sub: other.add(c)  # MMF, 신탁: 보통예금 원장 대조에서 뺌
        vals = list(a.values()); start, bi, bo, end = vals[1], vals[2], vals[3], vals[4]  # 계좌, 지난달 말 잔액, 입금, 출금, 이번 달 말 잔액
        lines = [x for x in bj if str(x['거래처코드']) == c and x['계정과목'] == '보통예금']
        i, o = (sum(x['차변금액'] or 0 for x in lines), sum(x['대변금액'] or 0 for x in lines)) if lines else (bi, bo)  # MMF, 신탁(기타단기금융상품)은 보통예금 원장에 없어서 통장 입출금을 따로 적고 합산 (담당자 확인 기준)
        for col, v in ((3, start), (4, i), (5, o), (7, end)):
            old = top.cell(r, col).value
            if is_f(old): continue  # 지점 양식이 SUMIFS 수식이면 그대로
            if old is None and not v and not fix_diff: continue  # 지점 양식은 비어 있던 칸에 0을 안 씀 (차이 수식을 바로잡는 양식은 0을 적음)
            top.cell(r, col).value = v
        f6 = top.cell(r, 6).value  # 잔고는 전일 + 입금 - 출금. '=C7-E7' 처럼 빠진 수식은 바로잡고, 지점이 숫자로 적는 칸은 둠
        if f6 is None or is_f(f6) or (fix_diff and isinstance(f6, (int, float))): top.cell(r, 6).value = f'=C{r}+D{r}-E{r}'
        for k in range(3, 10):  # 일반 서식으로 남은 금액 칸은 회계 서식
            if top.cell(r, k).number_format == 'General': top.cell(r, k).number_format = ACC
        if fix_diff:  # 차이 수식이 잘못된 양식: 증감 =D-E, 차이 =F-G 로 통일
            if top.cell(r, 8).value is None: top.cell(r, 8).value = f'=D{r}-E{r}'
            top.cell(r, 9).value = f'=F{r}-G{r}'
        if len(dig(top.cell(r, 2).value)) < min(10, len(hit[0])):  # '04-032', 43904 처럼 줄여 쓴 계좌번호는 전체 번호로
            top.cell(r, 2).value = acct_fmt(hit[0], label)
        table.append({'은행': label.replace('\n', ' '), '계좌번호': str(top.cell(r, 2).value), '구분': '기타단기금융상품' if after_sub else '보통예금',
                      '전일': start, '입금': i, '출금': o, '시재': end})
        diff = start + i - o - end
        if diff: check['계좌'].append({'행': r, '계좌번호': top.cell(r, 2).value, '장부잔고': start + i - o, '통장시재': end, '차이': diff})
    for a in accts:
        c = code(a['계좌'])
        if c not in used and (list(a.values())[2] or list(a.values())[3]):
            check['양식에없는계좌'].append({'계좌': a['계좌'], '입금': list(a.values())[2], '출금': list(a.values())[3]})

    # 자금수지 줄
    hdr = {}
    for c in range(1, sub.max_column + 1):
        h = str(sub.cell(3, c).value or '').replace(' ', '').replace('\n', '')
        for k in ('적요', '입금', '출금', '계좌명', '프로젝트', '비고', '지출일자'):
            if k in h and k not in hdr and not (k == '입금' and '계좌' in h): hdr[k] = c
    tpl_row = {c: copy.copy(sub.cell(4, c)._style) for c in range(1, sub.max_column + 1)}
    tpl_f = {c: sub.cell(4, c).value for c in range(1, sub.max_column + 1)
             if isinstance(sub.cell(4, c).value, str) and sub.cell(4, c).value.startswith('=')}
    prefill = max([r for r in range(4, sub.max_row + 1) if sub.cell(r, 1).value is not None] or [3])
    pre_pat = {c: sub.cell(prefill, c).value for c in range(1, sub.max_column + 1) if sub.cell(prefill, c).value is not None}
    if sub.cell(prefill, hdr['지출일자']).value: prefill = 0  # 마지막 줄이 실제 거래면 미리 채운 빈 줄이 없는 양식
    # 지점 계좌명 ('하나은행(12345)○○', '기업은행017(○○)'): 지난달 줄에 쓴 이름을 계좌번호 끝자리로 위하고 계좌에 이어 붙임
    names = {}
    if '계좌명' in hdr:
        for nm in {sub.cell(r, hdr['계좌명']).value for r in range(4, sub.max_row + 1)} - {None, ''}:
            d = dig(nm); hit = [k for k, v in acct_no.items() if d and v.endswith(d)] or [k for k, v in acct_no.items() if d and d in v]
            if len(hit) == 1: names[hit[0]] = nm
    for r in range(4, sub.max_row + 1):
        for c in range(1, sub.max_column + 1): sub.cell(r, c).value = None
    rows = [x for x in bj if x['계정과목'] == '보통예금']
    # 담당자 순서: 날짜별로 입금 먼저, 그다음 출금, 각각 전표번호 순
    rows.sort(key=lambda x: (str(x['일자']), 0 if x['차변금액'] else 1, str(x['전표번호'])))
    for n, x in enumerate(rows):
        r = 4 + n; d = str(x['일자'])
        for c, st in tpl_row.items(): sub.cell(r, c)._style = copy.copy(st)
        for c, f in tpl_f.items(): sub.cell(r, c).value = re.sub(r'([A-Z])4\b', rf'\g<1>{r}', f)
        sub.cell(r, 1).value = year
        sub.cell(r, hdr['지출일자']).value = f'{d[4:6]}-{d[6:]}'
        sub.cell(r, hdr['적요']).value = x['적요']
        sub.cell(r, hdr['입금']).value = x['차변금액'] or ''
        sub.cell(r, hdr['출금']).value = x['대변금액'] or ''
        pj = re.sub(r'^\[\d+\]', '', x['프로젝트'] or '')
        sub.cell(r, hdr['프로젝트']).value = pj
        if '계좌명' in hdr: sub.cell(r, hdr['계좌명']).value = names.get(str(x['거래처코드']), x['거래처'])
        # 분류가 거래와 어긋나 보이는 줄 (예: 부가세 납부가 급여성지출, 환불 입금이 기타지출)
        why = ('입금인데 지출 분류' if x['차변금액'] and '지출' in pj else '출금인데 입금 분류' if x['대변금액'] and '입금' in pj else
               '세금 적요인데 제세금지출 아님' if re.search(r'부가가치세|원천세|법인세|지방세|소득세', x['적요'] or '') and pj and pj != '제세금지출' else '')
        if why: check.setdefault('분류확인', []).append({'행': r, '일자': d, '적요': x['적요'], '프로젝트': pj, '이유': why})
        if not pj: check['프로젝트없음'].append({'행': r, '일자': d, '적요': x['적요'], '금액': x['차변금액'] or -(x['대변금액'] or 0)})
    # 담당자처럼 아래 빈 줄도 지난달 파일 끝까지 연도와 월 수식을 미리 채워 둠
    for r in range(4 + len(rows), prefill + 1):
        for c, st in tpl_row.items(): sub.cell(r, c)._style = copy.copy(st)
        for c, v in pre_pat.items():
            sub.cell(r, c).value = year if c == 1 else re.sub(rf'([A-Z]){prefill}\b', rf'\g<1>{r}', v) if isinstance(v, str) else v
    # 분개장과 통장이 다른 날 (전표 처리 전 등). 한 거래가 원금, 이자로 나뉘기도 해서 날짜, 계좌별 합계로 비교
    net = {}
    for x in rows:
        k = (str(x['일자']), str(x['거래처코드'])); net[k] = net.get(k, 0) + (x['차변금액'] or 0) - (x['대변금액'] or 0)
    for t in bank_tx:
        c = code(t['계좌'])
        if c not in used or c in other or 'MMF' in t['계좌'] or '신탁' in t['계좌']: continue
        k = (str(t['일자']), c); net[k] = net.get(k, 0) - (t['입금'] or 0) + (t['출금'] or 0)
    for (d, c), v in sorted(net.items()):
        if v:
            tx = [t for t in bank_tx if str(t['일자']) == d and code(t['계좌']) == c]
            check['통장만'].append({'일자': d, '계좌코드': c, '분개장-통장': v,
                                   '통장거래': [f"{t['적요(통장)']} 입{t['입금']} 출{t['출금']}" for t in tx]})
    # 분류 줄의 비고만 비워 둠 (Claude가 채움). '비고' 머리글은 둠. 지난달 비고는 말투 참고용으로 남김
    prev_notes = {}; side = '입금'
    for r in range(1, top.max_row + 1):
        if '출금' in str(top.cell(r, 2).value or ''): side = '출금'
        f = top.cell(r, 4).value
        if isinstance(f, str) and 'SUMIFS' in f.upper() and isinstance(top.cell(r, 3).value, str):
            if top.cell(r, 5).value: prev_notes[f'{top.cell(r, 3).value}({side})'] = top.cell(r, 5).value
            top.cell(r, 5).value = None
    big = {}
    for x in rows:  # 분류별 큰 거래 (비고 재료)
        k = re.sub(r'^\[\d+\]', '', x['프로젝트'] or '') or '(없음)'
        g = big.setdefault(k, {}); g[x['적요']] = g.get(x['적요'], 0) + (x['차변금액'] or 0) + (x['대변금액'] or 0)
    big = {k: sorted(v.items(), key=lambda kv: -kv[1])[:6] for k, v in big.items()}
    # 지난달 파일마다 섞여 있는 표기를 한 모양으로: TRUE() 검사식, 회계 서식의 '-'
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if fix_diff and isinstance(c.value, str) and c.value.startswith('=IF(') and ',TRUE,FALSE)' in c.value:
                    c.value = c.value.replace(',TRUE,FALSE)', ',TRUE(),FALSE())')
                if '"-"' in c.number_format: c.number_format = c.number_format.replace('"-"', '\\-')
    wb.save(out)
    cats = {}
    for x in rows:
        k = re.sub(r'^\[\d+\]', '', x['프로젝트'] or '') or '(없음)'; v = cats.setdefault(k, [0, 0])
        v[0] += x['차변금액'] or 0; v[1] += x['대변금액'] or 0
    json.dump({'줄수': len(rows), '계좌표': table, '분류합계': cats, '지난달비고': prev_notes, '분류별큰거래': big, **check}, open(out.replace('.xlsx', '_점검.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print(json.dumps({'줄수': len(rows), **{k: len(v) for k, v in check.items()}}, ensure_ascii=False))
    return out.replace('.xlsx', '_점검.json')


if __name__ == '__main__':
    main(*sys.argv[1:7])
