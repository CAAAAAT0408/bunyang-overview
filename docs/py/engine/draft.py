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


def output_name(sido='', sigun='', district='', builder=''):
    """'경기-화성-동탄1(M1-1-2BL)-포스코-분양개요.xlsx' (2026-10-18 사용자 지정 — 시도-시군-지구(블럭)-시공사)"""
    parts = [x for x in (sido, sigun, district, builder) if x]
    return '-'.join(parts + ['분양개요']) + '.xlsx'


# 블록 번호: 'M1-1-2블록', 'A-48블록', 'Abc-14블록', '공동1BL', 'M1, M2BL', '1,2,3BL'
BLOCK_RX = re.compile(r'(?<![A-Za-z0-9])((?:[A-Za-z]{1,3}\s?-?\s?)?\d+(?:-\d+)*(?:\s*,\s*[A-Za-z]{0,3}\d+)*)\s*(?:블록|블럭|BL)(?![A-Za-z])')
DISTRICT_FILLER = re.compile(r'^(공동주택용지|공동주택|공동|도시개발사업|택지개발사업지구|택지개발지구|택지개발사업|공공주택지구|'
                             r'일반산업단지|산업단지|지구단위계획|내|일원|일대|종전부동산|부지)$')


def district_of(location, sigun='', dong=''):
    """공급위치 → 지구(블럭): '반송동 95번지[동탄(1) M1-1-2블록]' → '동탄1(M1-1-2BL)',
    '화성동탄2 택지개발사업지구 A78BL' → '동탄2(A78BL)', 블록이 없으면 동 이름"""
    loc = re.sub(r'\((\d)\)', r'\1', location or '')          # 동탄(1) → 동탄1
    m = BLOCK_RX.search(loc)
    if not m:
        return dong
    code = re.sub(r'\s', '', m.group(1))
    if re.fullmatch(r'[A-Za-z]-\d+', code):
        code = code.replace('-', '')                          # A-48 → A48 (Abc-14·M1-1-2 는 그대로)
    name, skipped = '', []
    for t in reversed(re.split(r'[\s\(\)\[\],]+', loc[:m.start()])):
        if not t or DISTRICT_FILLER.match(t) or re.search(r'\d+(-\d+)?번지$|^\d', t):
            skipped.append(t)
            continue
        name = t
        break
    if sigun and name.startswith(sigun) and len(name) > len(sigun) + 1:
        name = name[len(sigun):]                              # 화성동탄2 → 동탄2, 평택브레인시티 → 브레인시티
    name = re.sub(r'(국제화계획지구|택지개발사업지구|택지개발지구|택지개발사업|도시개발사업)$', '', name)
    name = re.sub(r'공공주택지구$', '지구', name)               # 고천공공주택지구 → 고천지구
    if name and '공공주택지구' in skipped and not name.endswith('지구'):
        name += '지구'                                        # 군포대야미 공공주택지구 → 대야미지구
    name = name or dong
    return f'{name}({code}BL)' if name else f'{code}BL'


COMPANY_SHORT = [('지에스', 'GS'), ('GS', 'GS'), ('디엘', 'DL'), ('DL', 'DL'), ('HDC', '현산'), ('에이치디씨', '현산'),
                 ('현대산업개발', '현산'), ('케이씨씨', 'KCC'), ('KCC', 'KCC'), ('비에스한양', '한양'), ('에이치엘디앤아이한라', '한라'),
                 ('HL디앤아이한라', '한라'), ('이이치엘디앤아이한라', '한라'), ('에스케이에코플랜트', 'SK'), ('SK에코플랜트', 'SK'),
                 ('한국토지주택공사', 'LH'), ('LH', 'LH'), ('경기주택도시공사', 'GH'), ('삼성물산', '삼성'), ('한화', '한화'),
                 ('현대엔지니어링', '현대엔지니어링'), ('포스코', '포스코')]


def company_short(name):
    """'㈜포스코이앤씨' → '포스코', '지에스건설㈜' → 'GS', '계룡건설산업㈜' → '계룡'"""
    n = re.sub(r'㈜|\(주\)|\(유\)|주식회사|\s', '', name or '')
    for k, v in COMPANY_SHORT:
        if n.startswith(k):
            return v
    prev = None
    while prev != n:
        prev = n
        n = re.sub(r'(건설산업|종합건설|산업개발|건설|이앤씨|이엔씨|E&C|토건|중공업|글로벌|건영|공영|주택|산업|개발|기업)$', '', n) or prev
    return n


def builder_short(dev):
    """파일 이름·보고자료의 시공사: 시행이 LH 면 'LH', 아니면 시공사 짧은 이름 (여럿이면 쉼표, 3곳까지)"""
    dev = dev or {}
    if re.search(r'한국토지주택공사|^LH', dev.get('developer') or ''):
        return 'LH'
    names = []
    for x in re.split(r'[,/·]', dev.get('builder') or dev.get('developer') or ''):
        s = company_short(x)
        if s and s not in names:
            names.append(s)
    return ','.join(names[:3])


def region(location):
    toks = (location or '').split()
    sido = sigun = dong = ''
    for i, t in enumerate(toks):
        if not sido and (t in SIDO or t in SIDO_SHORT):
            sido = SIDO.get(t, t)
            continue
        if sido and not sigun and re.search(r'[시군]$', t):
            sigun = re.sub(r'(특례)?[시군]$', '', t)               # 화성특례시 → 화성
            continue
        if sido and not sigun and re.search(r'구$', t) and sido in ('서울', '인천', '부산', '대구', '광주', '대전', '울산'):
            sigun = t
            continue
        if not dong and re.search(r'(동|읍|면|리|가)$', t) and not re.search(r'\d', t):
            dong = t
        elif not dong and re.fullmatch(r'[가-힣]{2,}\d+동', t):
            dong = re.sub(r'\d+동$', '동', t)                    # 광명5동 → 광명동
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
        district = district_of(info.get('location'), sigun, dong)
    builder = fnp[3] if fnp else builder_short(info.get('developers'))
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
        # 위탁 / 수탁 / 시공 은 빈 칸이 있어도 3칸 그대로 (2026-10-18 사용자 지정)
        devtxt = ' / '.join(x or '' for x in (dev.get('trustor'), dev.get('trustee'), dev.get('builder')))
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
        'output_name': output_name(sido, sigun, district, builder),
        'sido': sido, 'sigun': sigun, 'district': district, 'builder_short': builder,
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
        'developers_label_auto': True,
        'report_auto': True,
        'option_note': '',                              # 비우면 양식 문구 그대로
        'reserve_pct': info.get('reserve_pct'),
        'delete_hidden': True,
        'remarks_auto': True,
        'parking_text_auto': True,
        'basis_note_auto': True,
        'warnings': list(parsed.get('warnings') or []),
    }
    if pay.get('first_fixed_mixed'):
        d['warnings'].append('1차 계약금이 타입마다 다릅니다(' + '·'.join(f'{v // 1000:,}' for v in pay['first_fixed_mixed']) +
                             '천원). 납부조건에는 첫 타입 금액을 적었습니다.')
    miss = payment_missing(pay)
    if miss:
        d['warnings'].append(f"납부조건의 {', '.join(miss)}을(를) 공고에서 찾지 못해 비워 두었습니다. 화면 3번 '납부조건'에서 채워 주세요.")
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


def _mid_split(pay):
    """중도금 회당 % 와 회수. 합계와 회수만 있으면 나눠서 구한다"""
    each, cnt = pay.get('mid_each'), pay.get('mid_count')
    if not each and cnt and pay.get('mid_pct_total'):
        each = pay['mid_pct_total'] / cnt
    return each, cnt


def _num(v):
    if v is None or v == '':
        return ''
    return f'{v:g}' if isinstance(v, float) else str(v)


def payment_note(pay):
    """항상 이 모양 (2026-10-18 사용자 지정) — 공고에서 못 찾은 숫자는 비워 둔다:
    ※ 납부조건 : 계약금 10%(1차 : 10,000천원, 2차 : 차액분), 중도금 : 10% x 6회(이자후불제)"""
    each, cnt = _mid_split(pay)
    ff = f"{round(pay['first_fixed'] / 1000):,}" if pay.get('first_fixed') else ''
    return (f"※ 납부조건 : 계약금 {_num(pay.get('deposit_pct'))}%(1차 : {ff}천원, 2차 : 차액분), "
            f"중도금 : {_num(each)}% x {_num(cnt)}회({pay.get('interest') or ''})")


def payment_missing(pay):
    each, cnt = _mid_split(pay)
    names = [('계약금 %', pay.get('deposit_pct')), ('1차 계약금', pay.get('first_fixed')),
             ('중도금 회당 %', each), ('중도금 회수', cnt), ('이자 조건', pay.get('interest'))]
    return [n for n, v in names if not v]


def developers_label(text):
    """시행/시공 칸 이름: 업체가 3곳(위탁·수탁·시공)이면 '위탁/수탁/시공', 아니면 '시행/시공' (2026-10-18 사용자 지정)"""
    return '위탁/수탁/시공' if len((text or '').split(' / ')) >= 3 else '시행/시공'


def pyeong_groups(types):
    """오픈보고의 평형 구성: 전용면적이 2㎡ 안으로 비슷한 타입끼리 묶어 첫 타입의 공급면적(평)을 내림
    (84C·84F → 34, 116A·116B → 47, 128A·129B → 52). 펜트하우스(PH…)는 뺀다"""
    groups, out = [], []
    for t in types:
        if re.match(r'^PH', t.get('abbr') or '', re.I) or not t.get('supply') or not t.get('excl'):
            continue
        if any(abs(g - t['excl']) <= 2 for g in groups):
            continue
        groups.append(t['excl'])
        py = int(t['supply'] * 0.3025)
        if py not in out:
            out.append(py)
    return out


def expansion_tag(types):
    return '확장무상' if types and all(not (t.get('balcony') or 0) for t in types) else '확장별도'


def report_texts(d):
    """보고자료 시트 문구 (2026-10-18 사용자 지정 양식)
    · 오픈동향 : 경기 화성 동탄1(M1-1-2BL) 포스코(총 812세대 , 46,458천원/평, 이자후불제, 확장별도)"""
    if not d.get('district') and d.get('location'):
        sido, sigun, dong = region(d['location'])
        d.setdefault('sido', sido)
        d.setdefault('sigun', sigun)
        d['district'] = district_of(d['location'], sigun, dong)
    if not d.get('builder_short') and d.get('developers'):
        d['builder_short'] = ','.join(dict.fromkeys(company_short(x) for x in d['developers'].split(' / ')[-1].split(',') if x.strip()))
    place = ' '.join(x for x in (d.get('sido'), d.get('sigun'), d.get('district'), d.get('builder_short')) if x)
    hh = int(d.get('households') or 0)
    tags = [f"{int(d['avg'] // 1000):,}천원/평" if d.get('avg') else '천원/평']
    tags += [x for x in ((d.get('payment') or {}).get('interest'), expansion_tag(d.get('types') or [])) if x]
    trend = f"· 오픈동향 : {place}(총 {hh:,}세대 , {', '.join(tags)})"
    py = ', '.join(f'{p}평형' for p in pyeong_groups(d.get('types') or []))
    opening = '\n'.join(['오픈보고입니다.', f'{place} 총 {hh:,}세대로 공급하였습니다.',
                         f"{py} 구성으로 평균평당가 {', '.join(tags)}로 공급하였습니다."])
    return trend, opening


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
    if d.get('developers_label_auto', True):
        d['developers_label'] = developers_label(d.get('developers'))
    if d.get('report_auto', True):
        d['report_trend'], d['report_open'] = report_texts(d)
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


def unit_pair(u):
    """단위세대 두 칸(왼쪽 D:J · 오른쪽 K:N)에 넣을 그림 (2026-10-18 사용자 지정)
      기본형(비확장)이 있으면      → 기본형 · 확장형   (셋 다 있어도 이것)
      기본형 없고 확장형·아이소    → 확장형 · 아이소
      그 밖                         → X 표   · 확장형(없으면 아이소)
    → (왼쪽, 오른쪽, 확장형·아이소 배치인지)"""
    b, e, i = u.get('basic'), u.get('ext'), u.get('iso')
    if b:
        return b, e, False
    if e and i:
        return e, i, True
    return None, e or i, False


def unit_headers(units, abbrs):
    """단위세대 머리글: 그림이 있는 타입이 모두 확장형·아이소 배치면 '확장형 | 아이소', 아니면 '기본형 | 확장형'"""
    modes = [unit_pair(units.get(a) or {})[2] for a in abbrs if any((units.get(a) or {}).values())]
    return ('확장형', '아이소') if modes and all(modes) else ('기본형', '확장형')


REPORT_HEAD = ['단지명', '시', '군구', '시공사', '세대수', '입주일자', '평형(공급면적)', '평형', '매매가', '평균평당가']
REPORT_WIDTHS = {'A': 2, 'B': 34, 'C': 8, 'D': 9, 'E': 14, 'F': 8, 'G': 9, 'H': 14, 'I': 8, 'J': 12, 'K': 12}


def report_layout(d):
    """보고자료 시트 줄 위치: 제목 줄들, 주변시세 표 (머리글 줄, 마지막 줄)"""
    lines = (d.get('report_open') or '').split('\n')
    t3 = max(13, 8 + len(lines) + 2)
    hr = t3 + 3
    return {'titles': [3, 7, t3], 'lines': lines, 't3': t3, 'table': (hr, hr + max(1, len(d.get('nearby') or [])))}


def report_cells(d):
    """보고자료 시트 칸 [(줄, 열, 값, 숫자서식)] — 사용자 파일(경기-화성-동탄1(M1-1-2BL)-포스코-분양개요) 양식
    주변시세: 매매가(천원) · 평균평당가 = 매매가 ÷ 평형(공급면적×0.3025)"""
    lay = report_layout(d)
    out = [(3, 2, '1. 업무일지', None), (4, 2, d.get('report_trend') or '', None), (7, 2, '2. 오픈보고', None)]
    out += [(8 + i, 2, ln, None) for i, ln in enumerate(lay['lines'])]
    t3 = lay['t3']
    out.append((t3, 2, '3. 주변시세', None))
    if d.get('nearby_note'):
        out.append((t3 + 1, 2, d['nearby_note'], None))
    hr = lay['table'][0]
    out += [(hr, 2 + j, h, None) for j, h in enumerate(REPORT_HEAD)]
    rows = d.get('nearby') or []
    if not rows:
        out.append((hr + 1, 2, '(주변 단지를 찾지 못했습니다 — 네이버 부동산에서 직접 넣어 주세요)', None))
    for i, x in enumerate(rows):
        r = hr + 1 + i
        out += [(r, 2, x.get('name') or '', None), (r, 3, x.get('si') or '', None), (r, 4, x.get('gu') or '', None),
                (r, 5, x.get('builder') or '', None), (r, 6, x.get('households') or '', '#,##0'),
                (r, 7, x.get('movein') or '', '@'), (r, 8, x.get('supply') or '', '0.00'),
                (r, 9, f'=H{r}*0.3025', '0.0'), (r, 10, x.get('price') or '', '#,##0'),
                (r, 11, f'=IFERROR(J{r}/I{r},"")', '#,##0')]
    return out


def chosen_option_cols(d):
    """엑셀 Q열부터 적을 유상옵션 칸 (2026-10-18 사용자 파일 기준): '비고 포함' 품목만, 타입마다 고른 안(일반형 기준)의
    금액만 — 같은 품목의 다른 안(식기세척기 삼성·LG)은 함께 더하지 않는다. 값이 하나도 없는 칸은 뺀다"""
    groups = {g['id']: g for g in d.get('options') or [] if g.get('include')}
    out = []
    for c in d.get('option_cols') or []:
        g = groups.get(c.get('group'))
        if not g:
            continue
        vals = {}
        for a, v in (c.get('values') or {}).items():
            if not v:
                continue
            vid = (g.get('choice') or {}).get(a)
            if (c.get('vids') or {}).get(a) is not None:
                ok = c['vids'][a] == vid
            else:                                       # 예전 초안 (vids 없음): 고른 안의 금액과 같으면
                chosen = next((x for x in g.get('variants') or [] if x['id'] == vid), None)
                ok = chosen is not None and chosen['prices'].get(a) == v
            if ok:
                vals[a] = v
        if vals:
            out.append({**c, 'values': vals})
    return out
