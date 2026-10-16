"""공고 분석 결과 → 분양개요에 넣을 값(초안). 화면에서 고친 뒤 finalize() 로 파생 값을 다시 계산한다."""
import math
import os
import re
from datetime import date

from .notice import SIDO
from .options import option_amounts, option_columns

SIDO_SHORT = set(SIDO.values())


def kdate(iso, with_year=True):
    if not iso:
        return ''
    y, m, d = iso[:10].split('-')
    return f'{y}년 {m}월 {d}일' if with_year else f'{m}월 {d}일'


def krange(ds):
    """['2026-08-10','2026-08-12'] → '2026년 08월 10일 ~ 08월 12일'"""
    if not ds:
        return ''
    if len(ds) == 1 or ds[0] == ds[-1]:
        return kdate(ds[0])
    return f'{kdate(ds[0])} ~ {kdate(ds[-1], False)}'


def parse_filename(fn):
    """'경기-오산-양산4지구(1BL)-GS-입주자모집공고.pdf' → ('경기','오산','양산4지구(1BL)','GS')"""
    base = os.path.splitext(os.path.basename(fn or ''))[0]
    parts, depth, cur = [], 0, ''
    for ch in base:                     # 괄호 안의 '-' (C-27L) 는 나누지 않는다
        if ch == '(':
            depth += 1
        elif ch == ')':
            depth = max(0, depth - 1)
        if ch == '-' and depth == 0:
            parts.append(cur.strip())
            cur = ''
        else:
            cur += ch
    parts.append(cur.strip())
    if len(parts) >= 4 and parts[0] in SIDO_SHORT:
        return parts[0], parts[1], parts[2], parts[3]
    return None


def output_name(fn, sido='', sigun='', district=''):
    base = os.path.splitext(os.path.basename(fn or ''))[0]
    if re.search(r'입주자\s*모집\s*공고', base):
        b = re.sub(r'입주자\s*모집\s*공고(문)?', '분양개요 및 청약결과', base)
        b = re.sub(r'\s*\(\d{2}\.\d{2}\.\d{2}\)$', '', b)
        return b + '.xlsx'
    parts = [x for x in (sido, sigun, district) if x]
    return '-'.join(parts + ['분양개요 및 청약결과']) + '.xlsx'


def region(location):
    toks = (location or '').split()
    sido = sigun = dong = ''
    for i, t in enumerate(toks):
        if not sido and (t in SIDO or t in SIDO_SHORT):
            sido = SIDO.get(t, t)
            continue
        if sido and not sigun and re.search(r'[시군]$', t):
            sigun = re.sub(r'[시군]$', '', t)
            continue
        if sido and not sigun and re.search(r'구$', t) and sido in ('서울', '인천', '부산', '대구', '광주', '대전', '울산'):
            sigun = t
            continue
        if not dong and re.search(r'(동|읍|면|리|가)$', t) and not re.search(r'\d', t):
            dong = t
    return sido, sigun, dong


def fmt_trunc2(x):
    return f'{math.floor(x * 100) / 100:.2f}'


def make_draft(parsed, groups, notice_file):
    info, sch = parsed['info'], parsed['schedule']
    sido, sigun, dong = region(info.get('location'))
    fnp = parse_filename(notice_file)
    district = ''
    if fnp:
        sido, sigun, district = fnp[0] or sido, fnp[1] or sigun, fnp[2]
    if not district:
        m = re.search(r'([A-Za-z]*-?\d+)\s*(?:블록|블럭|BL)', info.get('location') or '')
        district = (dong or '') + (f'({m.group(1)}BL)' if m else '')
    name = info.get('name') or ''
    title = f'[{sido} {sigun}] {district} {name}'.replace('  ', ' ').strip()
    loc = info.get('location') or ''
    if district and district not in loc and fnp:
        m = re.match(r'^(.*?)\((.*)\)$', district)
        if m and re.sub(r'\s', '', m.group(2)) in re.sub(r'\s', '', loc):
            pass                                               # '… 구름산지구 A6BL' + '구름산지구(A6BL)' → 그대로
        elif m and m.group(1) and m.group(1) in loc:
            loc = f'{loc}({m.group(2)})'                       # '내삼미동(내삼미2구역A2BL)' → (내삼미2구역A2BL)
        else:
            loc = f"{loc}({district.replace('(', ' ').replace(')', '').strip()})"
    dev = info.get('developers') or {}
    if dev.get('trustor') or dev.get('trustee'):
        devtxt = ' / '.join(x for x in (dev.get('trustor'), dev.get('trustee'), dev.get('builder')) if x)
    else:
        devtxt = ' / '.join(x for x in (dev.get('developer'), dev.get('builder')) if x)
    mh = info.get('model_house') or ''
    if info.get('tel'):
        mh = f"{mh} (☎ {info['tel']})".strip()
    # 청약일정 문구
    parts = []
    if sch.get('special'):
        parts.append(f"특별공급 : {krange(sch['special'])}")
    if sch.get('first_parts') and len(sch['first_parts']) > 1:
        for lab, ds in sch['first_parts']:
            lab = re.sub(r'지역$', '', re.sub(r'\s', '', lab or '')) or lab
            parts.append(f"1순위 {lab} : {krange(ds)}")
    elif sch.get('first'):
        parts.append(f"1순위 : {krange(sch['first'])}")
    if sch.get('second'):
        parts.append(f"2순위 : {krange(sch['second'])}")
    if not sch.get('first') and sch.get('general'):
        parts.append(f"일반공급 : {krange(sch['general'])}")
    types = []
    for t in parsed['types']:
        tt = dict(t)
        tt['price'] = t['tiers'][t['base']]['total'] if t.get('base') is not None and t['tiers'] else 0
        tt['balcony'] = t.get('balcony') or 0
        types.append(tt)
    households = sum(t['units'] or 0 for t in types)
    pay = parsed.get('payment') or {}
    d = {
        'notice_file': os.path.basename(notice_file or ''),
        'output_name': output_name(notice_file, sido, sigun, district),
        'title': title,
        'written': kdate(date.today().isoformat()),
        'homepage': info.get('homepage') or '',
        'location': loc,
        'scale': info.get('scale') or '',
        'features': '',
        'parking': None,
        'parking_sources': [],
        'households': households,
        'heating': info.get('heating') or '',
        'developers': devtxt,
        'model_house': mh,
        'mh_open': (sch.get('mh_open') or [None])[0],
        'notice_date': (sch.get('notice') or [None])[0],
        'subscription': ', '.join(parts),
        'winners': (sch.get('winners') or [None])[0],
        'contract': (', '.join(f'{lab} : {krange(ds)}' for lab, ds in sch['contract_parts'])
                     if sch.get('contract_parts') else krange(sch.get('contract'))),
        'move_in': info.get('move_in') or '',
        'label': make_label(dong, sigun, name),
        'types': types,
        'options': groups,
        'option_cols': option_columns(groups, [t['abbr'] for t in types]),
        'avg': parsed.get('avg_price_per_py'),
        'payment': pay,
        'payment_note': payment_note(pay),
        'option_note': '',                              # 비우면 양식 문구 그대로
        'reserve_pct': info.get('reserve_pct'),
        'delete_hidden': True,
        'remarks_auto': True,
        'parking_text_auto': True,
        'basis_note_auto': True,
        'warnings': list(parsed.get('warnings') or []),
    }
    return finalize(d)


def make_label(dong, sigun, name):
    """구분 칸: '양산동\n오산\n헤리티지자이\n1단지' — 긴 단지명은 앞의 시·군 이름에서 끊는다"""
    out = [dong] if dong else []
    for w in (name or '').split():
        if sigun and w.startswith(sigun) and len(w) > len(sigun) + 2:
            out += [sigun, w[len(sigun):]]
        elif len(w) > 7:
            out += [w[:len(w) // 2], w[len(w) // 2:]]
        else:
            out.append(w)
    return '\n'.join(out)


def payment_note(pay):
    s = '※ 납부조건 : '
    if pay.get('deposit_pct'):
        s += f"계약금 {pay['deposit_pct']}%"
        if pay.get('first_fixed'):
            s += f"(1차 {won_short(pay['first_fixed'])}, 2차 차액분)"
    if pay.get('mid_each') and pay.get('mid_count'):
        s += f", 중도금 : {pay['mid_each']}% x {pay['mid_count']}회"
    elif pay.get('mid_pct_total'):
        s += f", 중도금 : {pay['mid_pct_total']}%"
    if pay.get('interest') and '중도금' in s:
        s += f"({pay['interest']})"
    return s


def won_short(v):
    v = int(v)
    if v % 10_000_000 == 0:
        return f'{v // 10_000_000}천만원'
    if v % 1_000_000 == 0:
        return f'{v // 1_000_000}백만원'
    if v % 10_000 == 0:
        return f'{v // 10_000:,}만원'
    return f'{v:,}원'


def basis_note(types):
    labels = []
    for t in types:
        if t.get('base') is None or len(t.get('tiers') or []) < 2:     # 층 구간이 하나뿐인 타입(펜트 등)은 빼고 본다
            continue
        labels.append(re.sub(r'\s+', '', t['tiers'][t['base']]['floor']))
    if not labels:
        return ''
    common = max(set(labels), key=labels.count)
    if labels.count(common) == len(labels):
        return f' ※ 층별공급금액 세분화, {common} 분양가'
    return f' ※ 층별공급금액 세분화, 타입별 최다세대 층({common} 등) 분양가'


def remarks_text(groups):
    inc = [g for g in groups if g.get('include')]
    exc = [g for g in groups if not g.get('include')]
    def names(gs):
        out = []
        for g in gs:
            n = g.get('label') or g['name']
            if n not in out:
                out.append(n)
        return out
    s = ' [유상옵션 금액 포함 품목]\n' + ''.join(f"  - {n}\n" for n in names(inc))
    if exc:
        s += '\n [유상옵션 금액 미포함 품목]\n' + ''.join(f"  - {n}\n" for n in names(exc))
    return s


def finalize(d):
    """화면에서 바꾼 값(기준층·옵션 선택·주차대수 …)으로 파생 값을 다시 계산"""
    for t in d['types']:
        if t.get('tiers') and t.get('base') is not None and 0 <= int(t['base']) < len(t['tiers']):
            t['base'] = int(t['base'])
            if not t.get('price_manual'):
                t['price'] = t['tiers'][t['base']]['total']
        t['options'] = option_amounts(d.get('options') or [], t['abbr'])
        t['option_total'] = sum(t['options'].values())        # 화면 참고용 (엑셀 L열은 비움 — 사용자가 직접 입력)
    if 'option_cols' not in d:
        d['option_cols'] = option_columns(d.get('options') or [], [t['abbr'] for t in d['types']])
    tot_price = sum(x['total'] * x['units'] for t in d['types'] for x in (t.get('tiers') or []))
    tot_py = sum((t.get('supply') or 0) * 0.3025 * sum(x['units'] for x in (t.get('tiers') or [])) for t in d['types'])
    d['avg'] = tot_price / tot_py if tot_py else None
    d['avg_note'] = f"※ 평균평당가 : {int(d['avg'] // 1000):,}천원/평" if d['avg'] else '※ 평균평당가 : '
    d['households'] = sum(int(t.get('units') or 0) for t in d['types'])
    if d.get('basis_note_auto', True):
        d['basis_note'] = basis_note(d['types'])
    if d.get('remarks_auto', True):
        d['remarks'] = remarks_text(d.get('options') or [])
    if d.get('parking_text_auto', True):
        pk = d.get('parking')
        heat = d.get('heating') or ''
        if pk and d['households']:
            d['parking_text'] = f"총 {int(pk):,}대(세대당 {fmt_trunc2(int(pk) / d['households'])}대) / {heat}".rstrip(' /')
        else:
            d['parking_text'] = f"총 대(세대당 대) / {heat}"
    return d
