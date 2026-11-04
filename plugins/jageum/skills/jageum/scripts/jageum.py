#!/usr/bin/env python3
"""자금수지보고 실행기. 담당자가 자금을 확정하고 "이제 해" 하면 Claude가 부름.
회사, 지점, 폴더 이름은 코드에 없고, 첫 설정 때 이 PC의 나스 폴더에서 찾아 설정 파일(~/.jageum/config.json)에 저장함.

첫 설정 (Claude가 담당자와 인터뷰하며 부름)
  jageum.py find-nas                          # 이 PC에서 자금수지보고 폴더가 있는 나스 폴더 후보 찾기
  jageum.py check --nas "<나스 폴더>"           # 구조 확인, 지점 목록, 쓰기 가능 여부
  jageum.py config --nas "<나스 폴더>" --branches "본점,○○지점" [--company "회사 이름"]
  jageum.py status                            # 저장된 설정 보기
매달
  jageum.py where  --branch <지점> --ym 2026-10
  jageum.py report --branch <지점> --ym 2026-10 [--day 3] [--raw <원자료 폴더>] [--force]   # --day: 월 중간 '오늘자' 기준일
  jageum.py notes  --report <보고서.xlsx> --notes <비고.json>
  jageum.py total  --ym 2026-10

폴더 규칙 (나스 폴더 기준, check 가 찾아 저장)
  보고서  <보고서 폴더>/<지점 폴더>/ 아래 지난달 보고서 옆      예: 01. 지점별 자금수지보고/01. ○○지점/2026년/
  원자료  <원자료 폴더>/<연>/<연-월>_마감분/자금수지/<지점 폴더>/   (분개장*.xlsx, 통장거래*.xlsx)
  통합    <보고서 폴더의 상위>/통합 자금수지보고/<연>/
"""
import argparse, contextlib, glob, io, json, os, re, shutil, string, sys, unicodedata

sys.stdout.reconfigure(encoding='utf-8')
HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, HERE)
import make_report, make_total, write_notes  # noqa: E402

CFG = os.path.join(os.path.expanduser('~'), '.jageum', 'config.json')
N = lambda s: unicodedata.normalize('NFC', s)
out_json = lambda d: print(json.dumps(d, ensure_ascii=False, indent=1))


def last_day(y, m): return 31 if m in (1, 3, 5, 7, 8, 10, 12) else 30 if m != 2 else (29 if y % 4 == 0 else 28)
def ym(s): y, m = map(int, re.match(r'(\d{4})-(\d{1,2})', s).groups()); return y, m
def short(folder): return re.sub(r'^\d+\.\s*', '', N(folder))  # '01. ○○지점' -> '○○지점'


def load_cfg():
    if not os.path.exists(CFG): sys.exit('설정이 없습니다. 첫 설정부터 하세요 (find-nas, check, config).')
    return json.load(open(CFG, encoding='utf-8'))


def has_report(d):
    """폴더 아래에 '자금수지보고' 엑셀이 있는지 (맥에서 받은 한글 이름은 자모가 풀려 있어 NFC로 맞춰 비교)"""
    for _, _, files in os.walk(d):
        if any('자금수지보고' in N(f) and '통합' not in N(f) and f.endswith('.xlsx') for f in files): return True
    return False


def scan(nas):
    """나스 폴더 안에서 보고서 폴더(지점 폴더들 아래 '자금수지보고' 엑셀이 있는 곳)와 원자료 폴더를 찾음"""
    rep = None; best = 0
    for d, subs, files in os.walk(nas):
        if d[len(nas):].count(os.sep) > 3: subs[:] = []; continue
        n = sum(1 for s in subs if has_report(os.path.join(d, s)))
        if n and n >= best: rep, best = d, n  # 같은 수면 더 깊은 폴더 (지점 폴더들이 바로 아래 있는 곳)
    raw = None
    for d, subs, _ in os.walk(nas):
        if d[len(nas):].count(os.sep) > 1: subs[:] = []; continue
        for s in subs:
            if '데이터 자동수집' in N(s) or '자동수집' in N(s): raw = os.path.join(d, s)
    branches = sorted(N(s) for s in os.listdir(rep) if os.path.isdir(os.path.join(rep, s))) if rep else []
    return rep, raw, branches


def branch_folder(cfg, name):
    n = N(name).replace(' ', '')
    for f in cfg['지점폴더']:
        s = short(f).replace(' ', '')
        if n in (s, f.replace(' ', '')) or n == s.replace('지점', '') or s == n + '지점': return f
    part = [f for f in cfg['지점폴더'] if n.replace('지점', '') in short(f).replace(' ', '')]
    if len(part) == 1: return part[0]  # '본점' -> '○○본점'처럼 하나로 정해지면
    sys.exit(f'지점을 모릅니다: {name} (가능: {", ".join(short(f) for f in cfg["지점폴더"])})')


def find(cfg, bf, y, m, raw=None, day=None):
    py, pm = (y - 1, 12) if m == 1 else (y, m - 1)
    rawd = raw or os.path.join(cfg['원자료폴더'] or '', str(y), f'{y}-{m:02d}_마감분', '자금수지', bf)
    fs = [f for f in glob.glob(os.path.join(rawd, '*.xlsx')) if not os.path.basename(f).startswith('~$')]
    bj = [f for f in fs if '분개장' in N(os.path.basename(f))]
    bk = [f for f in fs if '통장거래' in N(os.path.basename(f))]
    root = os.path.join(cfg['보고서폴더'], bf); cand = []
    for f in glob.glob(os.path.join(root, '**', '*.xlsx'), recursive=True):
        n = N(os.path.basename(f)); rel = N(os.path.relpath(f, root))
        if n.startswith('~$') or '양식' in n or '외감' in n or '(셈프로)' in n: continue
        if not re.search(rf'(?<!\d){pm}월[ _]*\d+일', n) or '기준' not in n: continue
        if os.sep in rel and str(py) not in rel.split(os.sep)[0]: continue  # 연도 폴더가 있으면 지난달 연도만
        cand.append(f)
    tpl = max(cand, key=os.path.getmtime) if cand else None; out = None
    if tpl:
        d = os.path.dirname(tpl)
        if py != y: d = os.path.join(os.path.dirname(d), N(os.path.basename(d)).replace(str(py), str(y)))  # 1월은 새 연도 폴더
        n = re.sub(rf'(?<!\d){pm}월([ _]*)\d+일', lambda mm: f'{m}월{mm.group(1)}{day or last_day(y, m)}일', N(os.path.basename(tpl)))
        out = os.path.join(d, n)
    return {'원자료폴더': rawd, '분개장': bj[0] if len(bj) == 1 else bj, '통장거래': bk[0] if len(bk) == 1 else bk,
            '지난달보고서': tpl, '만들보고서': out, '이미있음': bool(out and os.path.exists(out))}


def cmd_find_nas(_):
    homes = [os.path.expanduser('~')] + [f'{c}:\\' for c in string.ascii_uppercase if os.path.exists(f'{c}:\\')]
    found = []
    for h in homes:
        for d, subs, _ in os.walk(h):
            depth = d[len(h):].count(os.sep)
            subs[:] = [s for s in subs if not s.startswith('.') and s not in ('AppData', 'Library', 'node_modules', 'Windows', 'Program Files', 'Program Files (x86)')]
            if depth > 4: subs[:] = []; continue
            if any('자금수지보고' in N(s) or '자금마감' in N(s) for s in subs):
                found.append(os.path.dirname(d) if '자금' in N(os.path.basename(d)) else d); subs[:] = []
    out_json({'후보': sorted(set(found))[:10]})


def cmd_check(a):
    nas = a.nas; ok = os.path.isdir(nas)
    rep, raw, br = scan(nas) if ok else (None, None, [])
    writable = False
    if rep:
        t = os.path.join(rep, '.jageum_쓰기확인.tmp')
        try: open(t, 'w').write('x'); os.remove(t); writable = True
        except OSError: pass
    out_json({'나스폴더있음': ok, '보고서폴더': rep, '원자료폴더': raw, '지점폴더': br, '보고서폴더에쓰기가능': writable})


def cmd_config(a):
    rep, raw, br = scan(a.nas)
    if not rep: sys.exit('이 폴더 안에서 지점별 자금수지보고 폴더를 못 찾았습니다.')
    mine = [N(x.strip()) for x in (a.branches or '').split(',') if x.strip()]
    cfg = {'나스': a.nas, '보고서폴더': rep, '원자료폴더': raw, '지점폴더': br, '회사': a.company or '', '내지점': []}
    cfg['내지점'] = [branch_folder(cfg, x) for x in mine]
    os.makedirs(os.path.dirname(CFG), exist_ok=True); json.dump(cfg, open(CFG, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    out_json({'저장': CFG, **cfg})


def cmd_report(a):
    cfg = load_cfg(); bf = branch_folder(cfg, a.branch); y, m = ym(a.ym); day = int(a.day) if a.day else None; w = find(cfg, bf, y, m, a.raw, day)
    miss = [k for k in ('분개장', '통장거래', '지난달보고서') if not isinstance(w[k], str)]
    if miss: out_json({'오류': '파일을 못 찾음', '없는것': miss, **w}); sys.exit(2)
    out = w['만들보고서']
    if w['이미있음'] and not a.force: out = out.replace('.xlsx', ' (셈프로).xlsx')  # 담당자 파일을 덮지 않음
    os.makedirs(os.path.dirname(out), exist_ok=True)
    with contextlib.redirect_stdout(io.StringIO()): chk = make_report.main(w['지난달보고서'], w['분개장'], w['통장거래'], y, m, out, day, with_bank_only=bool(day))
    dst = os.path.join(w['원자료폴더'], f'{short(bf)}_점검_{y}-{m:02d}.json')  # 점검표는 원자료 폴더에
    try: shutil.move(chk, dst)
    except OSError: dst = chk
    d = json.load(open(dst, encoding='utf-8'))
    out_json({'보고서': out, '점검표': dst, '줄수': d['줄수'],
              '확인할것': {k: d[k] for k in ('계좌', '프로젝트없음', '통장만', '전표처리전', '양식에없는계좌', '양식계좌번호다름', '분류확인') if d.get(k)},
              '분류합계': d['분류합계'], '지난달비고': d['지난달비고'], '분류별큰거래': d['분류별큰거래']})


def cmd_total(a):
    cfg = load_cfg(); y, m = ym(a.ym); pairs, missing = [], []
    for bf in cfg['지점폴더']:
        p = os.path.join(cfg['원자료폴더'] or '', str(y), f'{y}-{m:02d}_마감분', '자금수지', bf, f'{short(bf)}_점검_{y}-{m:02d}.json')
        if os.path.exists(p): pairs.append(f'{short(bf)}={p}')
        else: missing.append(short(bf))
    out = os.path.join(os.path.dirname(cfg['보고서폴더']), '통합 자금수지보고', str(y),
                       f'{cfg.get("회사") or "전 지점"}_자금수지보고_통합_{m}월 {last_day(y, m)}일 기준.xlsx')
    os.makedirs(os.path.dirname(out), exist_ok=True)
    with contextlib.redirect_stdout(io.StringIO()) as buf: make_total.main(y, m, out, *pairs, company=cfg.get('회사') or '')
    out_json({'통합': out, '빠진지점': missing, '합계': json.loads(buf.getvalue().strip().splitlines()[-1])})


def main():
    ap = argparse.ArgumentParser(); sp = ap.add_subparsers(dest='c', required=True)
    sp.add_parser('find-nas'); sp.add_parser('status')
    p = sp.add_parser('check'); p.add_argument('--nas', required=True)
    p = sp.add_parser('config'); p.add_argument('--nas', required=True); p.add_argument('--branches'); p.add_argument('--company')
    for n in ('where', 'report'):
        p = sp.add_parser(n); p.add_argument('--branch', required=True); p.add_argument('--ym', required=True)
        p.add_argument('--raw'); p.add_argument('--force', action='store_true'); p.add_argument('--day', help='월 중간 기준일 (예: 3 → 11월 3일 기준)')
    p = sp.add_parser('notes'); p.add_argument('--report', required=True); p.add_argument('--notes', required=True)
    p = sp.add_parser('total'); p.add_argument('--ym', required=True)
    a = ap.parse_args()
    if a.c == 'find-nas': cmd_find_nas(a)
    elif a.c == 'check': cmd_check(a)
    elif a.c == 'config': cmd_config(a)
    elif a.c == 'status': out_json(load_cfg())
    elif a.c == 'where': cfg = load_cfg(); out_json(find(cfg, branch_folder(cfg, a.branch), *ym(a.ym), a.raw, int(a.day) if a.day else None))
    elif a.c == 'report': cmd_report(a)
    elif a.c == 'notes': write_notes.main(a.report, a.notes); print('비고 넣음', a.report)
    elif a.c == 'total': cmd_total(a)


if __name__ == '__main__':
    main()
