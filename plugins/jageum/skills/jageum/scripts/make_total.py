#!/usr/bin/env python3
"""자금수지 통합: 지점별 보고서를 만들 때 남긴 점검 파일(계좌표, 분류합계)을 모아 전 지점 시재 합계 엑셀을 만듦.
사용: make_total.py <연> <월> <저장할 파일.xlsx> <지점명>=<점검.json> ...
시트: 통합시재(계좌별) / 지점별 요약 / 분류별 입출금. 합계는 엑셀 수식이라 열면 다시 계산됨.
"""
import json, sys
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

ACC = '_-* #,##0_-;\\-* #,##0_-;_-* \\-_-;_-@_-'
HEAD = PatternFill('solid', fgColor='DDEBF7'); SUBF = PatternFill('solid', fgColor='F2F2F2'); TOTF = PatternFill('solid', fgColor='FFF2CC')
thin = Side(style='thin', color='BFBFBF'); BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
B = Font(name='맑은 고딕', bold=True); N = Font(name='맑은 고딕')
CATS_IN = ['영업입금', '지원금입금', '이자입금', '본지점거래(입금)', '은행차입', '관계사차입', '투자입금', '기타입금', '대체거래']
CATS_OUT = ['급여성지출', '비용지출', '제세금지출', '장비/시설투자', '본지점거래(지출)', '배당지급', '은행상환', '관계사상환', '지출이자', '기타지출', '대체거래']


def style_row(ws, r, c1, c2, fill=None, bold=False, num_from=None):
    for c in range(c1, c2 + 1):
        x = ws.cell(r, c); x.border = BOX; x.font = B if bold else N
        if fill: x.fill = fill
        if num_from and c >= num_from: x.number_format = ACC


def main(year, month, out, *pairs, company=''):
    month = int(month); data = []
    for p in pairs:
        name, path = p.split('=', 1); data.append((name, json.load(open(path, encoding='utf-8'))))
    wb = openpyxl.Workbook()

    # 1. 계좌별
    ws = wb.active; ws.title = f'통합시재({month}월)'
    ws['A1'] = f'{company} 전 지점 자금 보유 현황 ({month}월 말 기준)'.strip(); ws['A1'].font = Font(name='맑은 고딕', bold=True, size=14)
    ws['A2'] = f'({month}월1일~{month}월 말일) · 지점별 자금수지보고 계좌표를 모은 것'; ws['A2'].font = N
    hd = ['지점', '구분', '은행', '계좌번호', '전일', '입금', '출금', '잔고', '통장시재', '차이']
    for c, h in enumerate(hd, 1): ws.cell(4, c, h)
    style_row(ws, 4, 1, 10, HEAD, True)
    r = 5; sub_rows = []
    for name, d in data:
        first = r
        for t in d['계좌표']:
            ws.cell(r, 1, name); ws.cell(r, 2, t['구분']); ws.cell(r, 3, t['은행']); ws.cell(r, 4, t['계좌번호'].strip())
            ws.cell(r, 5, t['전일']); ws.cell(r, 6, t['입금']); ws.cell(r, 7, t['출금'])
            ws.cell(r, 8, f'=E{r}+F{r}-G{r}'); ws.cell(r, 9, t['시재']); ws.cell(r, 10, f'=H{r}-I{r}')
            style_row(ws, r, 1, 10, num_from=5); r += 1
        ws.cell(r, 1, name); ws.cell(r, 3, '지점 계')
        for c in range(5, 11): ws.cell(r, c, f'=SUBTOTAL(9,{L(c)}{first}:{L(c)}{r - 1})')
        style_row(ws, r, 1, 10, SUBF, True, 5); sub_rows.append(r); r += 1
    r += 1; tot = {}
    for label, cond in (('보통예금 합계', '보통예금'), ('기타단기금융상품 합계 (MMF, 신탁)', '기타단기금융상품')):
        ws.cell(r, 1, label)
        for c in range(5, 11): ws.cell(r, c, f'=SUMIFS({L(c)}5:{L(c)}{sub_rows[-1]},$B$5:$B${sub_rows[-1]},"{cond}")')
        style_row(ws, r, 1, 10, TOTF, True, 5); tot[cond] = r; r += 1
    ws.cell(r, 1, '총계')
    for c in range(5, 11): ws.cell(r, c, f'={L(c)}{tot["보통예금"]}+{L(c)}{tot["기타단기금융상품"]}')
    style_row(ws, r, 1, 10, TOTF, True, 5); grand = r
    ws.cell(r + 2, 1, '원자료가 없는 지점은 빠져 있습니다. 본점과 지점 사이 이체는 양쪽에 다 잡혀 입금, 출금 합계에는 겹쳐 있습니다 (시재 합계에는 영향 없음).').font = Font(name='맑은 고딕', size=9, color='7F7F7F')
    for c, w in zip(range(1, 11), (12, 16, 30, 22, 16, 16, 16, 16, 16, 10)): ws.column_dimensions[L(c)].width = w
    ws.freeze_panes = 'E5'

    # 2. 지점별 요약
    s2 = wb.create_sheet('지점별 요약')
    s2['A1'] = f'지점별 {month}월 말 시재'; s2['A1'].font = Font(name='맑은 고딕', bold=True, size=14)
    hd = ['지점', '보통예금 시재', 'MMF, 신탁', '시재 합계', f'{month}월 입금', f'{month}월 출금', '전월 대비 증감', '장부와 통장 차이']
    for c, h in enumerate(hd, 1): s2.cell(3, c, h)
    style_row(s2, 3, 1, 8, HEAD, True)
    src = f"'{ws.title}'"
    for k, (name, d) in enumerate(data):
        rr = 4 + k; s2.cell(rr, 1, name)
        s2.cell(rr, 2, f'=SUMIFS({src}!$I:$I,{src}!$A:$A,$A{rr},{src}!$B:$B,"보통예금")')
        s2.cell(rr, 3, f'=SUMIFS({src}!$I:$I,{src}!$A:$A,$A{rr},{src}!$B:$B,"기타단기금융상품")')
        s2.cell(rr, 4, f'=B{rr}+C{rr}')
        s2.cell(rr, 5, f'=SUMIFS({src}!$F:$F,{src}!$A:$A,$A{rr},{src}!$B:$B,"<>")')
        s2.cell(rr, 6, f'=SUMIFS({src}!$G:$G,{src}!$A:$A,$A{rr},{src}!$B:$B,"<>")')
        s2.cell(rr, 7, f'=E{rr}-F{rr}')
        s2.cell(rr, 8, f'=SUMIFS({src}!$J:$J,{src}!$A:$A,$A{rr},{src}!$B:$B,"<>")')
        style_row(s2, rr, 1, 8, num_from=2)
    rr = 4 + len(data); s2.cell(rr, 1, '합계')
    for c in range(2, 9): s2.cell(rr, c, f'=SUM({L(c)}4:{L(c)}{rr - 1})')
    style_row(s2, rr, 1, 8, TOTF, True, 2)
    for c, w in zip(range(1, 9), (12, 18, 18, 18, 18, 18, 16, 16)): s2.column_dimensions[L(c)].width = w

    # 3. 분류별 입출금
    s3 = wb.create_sheet('분류별 입출금')
    s3['A1'] = f'{month}월 분류별 입출금 (프로젝트 검토 기준)'; s3['A1'].font = Font(name='맑은 고딕', bold=True, size=14)
    names = [n for n, _ in data]
    s3.cell(3, 1, '구분'); s3.cell(3, 2, '분류')
    for k, n in enumerate(names): s3.cell(3, 3 + k, n)
    s3.cell(3, 3 + len(names), '합계'); style_row(s3, 3, 1, 3 + len(names), HEAD, True)
    rr = 4
    for side, cats, idx in (('입금', CATS_IN, 0), ('출금', CATS_OUT, 1)):
        first = rr
        extra = sorted({c for _, d in data for c, v in d['분류합계'].items() if v[idx] and c not in cats})
        for cat in cats + extra:
            s3.cell(rr, 1, side); s3.cell(rr, 2, cat)
            for k, (_, d) in enumerate(data): s3.cell(rr, 3 + k, d['분류합계'].get(cat, [0, 0])[idx])
            s3.cell(rr, 3 + len(names), f'=SUM(C{rr}:{L(2 + len(names))}{rr})')
            style_row(s3, rr, 1, 3 + len(names), num_from=3); rr += 1
        s3.cell(rr, 1, side); s3.cell(rr, 2, f'{side} 합계')
        for c in range(3, 4 + len(names)): s3.cell(rr, c, f'=SUM({L(c)}{first}:{L(c)}{rr - 1})')
        style_row(s3, rr, 1, 3 + len(names), SUBF, True, 3); rr += 2
    s3.column_dimensions['A'].width = 8; s3.column_dimensions['B'].width = 18
    for k in range(len(names) + 1): s3.column_dimensions[L(3 + k)].width = 16
    s3.freeze_panes = 'C4'
    wb.save(out)
    # 화면용 숫자
    sums = {'보통예금': 0, '기타단기금융상품': 0}
    for name, d in data:
        for t in d['계좌표']: sums[t['구분']] += t['시재'] or 0
    print(json.dumps({'지점': len(data), **sums, '총계': sums['보통예금'] + sums['기타단기금융상품'],
                      '지점별': {n: sum(t['시재'] or 0 for t in d['계좌표']) for n, d in data}}, ensure_ascii=False))


if __name__ == '__main__':
    main(*sys.argv[1:])
