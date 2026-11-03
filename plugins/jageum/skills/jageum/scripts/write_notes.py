#!/usr/bin/env python3
"""자금수지보고 비고 넣기: Claude가 쓴 분류별 비고(json)를 계좌마감 시트 E열에 넣음.
사용: write_notes.py <보고서.xlsx> <비고.json>   비고.json = {"영업입금": "...", "비용지출": "...", ...}
분류 이름은 보고서 C열 그대로. 입금, 출금 양쪽에 있는 '대체거래'는 "대체거래(입금)", "대체거래(출금)"으로 구분.
"""
import json, sys
import openpyxl


def main(path, notes_path):
    notes = json.load(open(notes_path, encoding='utf-8'))
    wb = openpyxl.load_workbook(path); ws = wb.worksheets[0]
    side = '입금'
    for r in range(1, ws.max_row + 1):
        b = str(ws.cell(r, 2).value or '')
        if '출금' in b: side = '출금'
        f, k = ws.cell(r, 4).value, ws.cell(r, 3).value
        if not (isinstance(f, str) and 'SUMIFS' in f.upper() and k): continue
        v = notes.get(f'{k}({side})', notes.get(k))
        if v: ws.cell(r, 5).value = v
    wb.save(path)


if __name__ == '__main__':
    main(*sys.argv[1:3])
