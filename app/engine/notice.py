"""입주자모집공고(PDF) 분석 → 분양개요 작성에 필요한 값.

결과는 dict 하나(JSON 으로 그대로 웹 화면에 보낸다). 못 찾은 값은 None 이고 warnings 에 이유를 적는다.
"""
import re
from datetime import date

from .pdfdoc import PdfDoc

# ------------------------------------------------------------------ 공통 도구

CODE_RE = re.compile(r'^0?\d{2,3}\.\d{2,4}\s?[A-Za-z]{0,3}\d?$')     # 주택형 084.9796A
DATE_RE = re.compile(r'(20\d{2})\s*[.\-/년]\s*(\d{1,2})\s*[.\-/월]\s*(\d{1,2})')
SIDO = {'서울특별시': '서울', '서울시': '서울', '인천광역시': '인천', '경기도': '경기', '부산광역시': '부산',
        '대구광역시': '대구', '광주광역시': '광주', '대전광역시': '대전', '울산광역시': '울산',
        '세종특별자치시': '세종', '강원도': '강원', '강원특별자치도': '강원', '충청북도': '충북',
        '충청남도': '충남', '전라북도': '전북', '전북특별자치도': '전북', '전라남도': '전남',
        '경상북도': '경북', '경상남도': '경남', '제주특별자치도': '제주'}
SKIP_URL = ('applyhome', 'iros.go.kr', 'wetax', 'e-revenuestamp', 'eum.go.kr', 'molit.go.kr', 'law.go.kr',
            'gov.kr', 'hometax', 'nts.go.kr', 'khug.or.kr', 'hug.', 'reb.or.kr', 'kab.co.kr', 'lh.or.kr/', 'kakao',
            'naver.me', 'nhuf.molit', 'myhome.go.kr', 'apps.apple', 'play.google', 'mois.go.kr', 'scourt',
            'kcomwel', 'nhis.or.kr', 'nps.or.kr', 'bokjiro', 'hf.go.kr', 'kinfa', 'fss.or.kr', 'koreanair')


def num(s):
    """'1,234,000' → 1234000, '23.6528' → 23.6528, '-'/'' → None"""
    if s is None:
        return None
    if isinstance(s, (int, float)):
        return s
    t = str(s).strip().replace(',', '').replace(' ', '')
    t = re.sub(r'(원|㎡|세대|호|대)$', '', t)
    if not re.fullmatch(r'-?\d+(\.\d+)?', t):
        return None
    v = float(t)
    return int(v) if re.fullmatch(r'-?\d+', t) else v


def is_num(s):
    return num(s) is not None


SHORT_DATE_RE = re.compile(r"[‘'’`]\s?(\d{2})\s*[.\-/]\s*(\d{1,2})\s*[.\-/]\s*(\d{1,2})")


def _iter_dates(s):
    found = []
    for m in DATE_RE.finditer(s or ''):
        found.append((m.start(), int(m.group(1)), int(m.group(2)), int(m.group(3))))
    for m in SHORT_DATE_RE.finditer(s or ''):
        found.append((m.start(), 2000 + int(m.group(1)), int(m.group(2)), int(m.group(3))))
    found.sort()
    for _, y, mo, d in found:
        try:
            yield date(y, mo, d)
        except ValueError:
            pass


def to_date(s):
    return next(_iter_dates(s), None)


def all_dates(s):
    out = list(_iter_dates(s))
    # '2026.07.31.(금) ~ 08.04.(화)' 처럼 뒤쪽 날짜에 연도가 없는 경우
    m = re.search(r'[~〜∼～]\s*(\d{1,2})\s*[.월]\s*(\d{1,2})', s or '')
    if len(out) == 1 and m:
        try:
            out.append(date(out[0].year, int(m.group(1)), int(m.group(2))))
        except ValueError:
            pass
    return out


def floor_rank(label, idx=0):
    """층 구분 글자 → 높이 순위 ('20층 이상' 20, '17~19층' 19, '최상층' 1000). 숫자가 없으면 표 순서"""
    t = str(label or '')
    if re.search(r'최상|최고|펜트|PH', t):
        return 1000 + idx
    ns = [int(x) for x in re.findall(r'\d+', t)]
    if ns:
        return max(ns) + idx / 1000
    if '기준' in t:
        return 500 + idx
    return idx / 1000


def norm_type(s):
    t = str(s or '').upper()
    t = re.sub(r'\s+', '', t)
    t = t.replace('㎡', '').replace('M²', '').replace('M2', '').replace('TYPE', '').replace('타입', '')
    t = re.sub(r'형$', '', t)
    t = t.replace('형', '') if re.fullmatch(r'\d+형[A-Z]*', t) else t
    t = re.sub(r'^0+(\d)', r'\1', t)
    return t


def code_area(code):
    m = re.match(r'0?(\d{2,3}\.\d+)', str(code).strip())
    return float(m.group(1)) if m else None


def code_suffix(code):
    m = re.match(r'0?\d{2,3}\.\d+\s?([A-Za-z]{0,3}\d?)$', str(code).strip())
    return m.group(1).upper() if m else ''


def abbr_from_code(code):
    a = code_area(code)
    return f'{int(a)}{code_suffix(code)}' if a else str(code)


class TypeMatcher:
    """공고 곳곳의 타입 표기(84A, 084.9796A, 84㎡A, 84B,D, 전타입 …)를 공급표의 타입으로 맞춘다."""

    def __init__(self, types):
        self.types = types                    # [{'abbr','code',...}]
        self.by = {}
        for t in types:
            for k in (t['abbr'], t['code'], abbr_from_code(t['code'])):
                self.by[norm_type(k)] = t['abbr']
            self.by[norm_type(str(code_area(t['code'])) + code_suffix(t['code']))] = t['abbr']

    def one(self, s):
        k = norm_type(s)
        if not k:
            return None
        if k in self.by:
            return self.by[k]
        m = re.fullmatch(r'0?(\d{2,3})(\.\d+)?([A-Z]{0,3}\d?)', k)
        if m:
            cand = norm_type(m.group(1) + m.group(3))
            if cand in self.by:
                return self.by[cand]
        return None

    def many(self, s):
        s = str(s or '').strip()
        if not s:
            return []
        if re.search(r'전\s*(타입|주택형|세대|평형)|공통|전체', s):
            return [t['abbr'] for t in self.types]
        one = self.one(s)
        if one:
            return [one]
        out, prefix = [], ''
        for tok in re.split(r'[,·/&+\n]|및|\s', s):
            tok = tok.strip()
            if not tok:
                continue
            if re.fullmatch(r'[A-Za-z]{1,2}', tok) and prefix:
                tok = prefix + tok
            hit = self.one(tok)
            if hit:
                out.append(hit)
                m = re.match(r'(\d{2,3})', norm_type(tok))
                prefix = m.group(1) if m else prefix
        seen = []
        for o in out:
            if o not in seen:
                seen.append(o)
        return seen


def header_columns(rows, nhead):
    """머리글 여러 줄을 칸별 문자열로 합친다. 병합된 묶음 머리글은 오른쪽 빈칸에도 물려준다.
    표 위에 붙은 '※ …' 안내 문장 줄은 머리글에서 뺀다."""
    ncols = max(len(r) for r in rows)
    grid = []
    for r in rows[:nhead]:
        cells = [c for c in r if c]
        if len(cells) == 1 and (cells[0].lstrip().startswith(('※', '*', '■')) or len(cells[0]) > 30):
            continue
        grid.append(list(r) + [''] * (ncols - len(r)))
    if not grid:
        return [''] * ncols
    for r in range(len(grid) - 1):
        last = ''
        for c in range(ncols):
            if grid[r][c]:
                last = grid[r][c]
            elif last and any(grid[rr][c] for rr in range(r + 1, len(grid))):
                grid[r][c] = last
            else:
                last = '' if not any(grid[rr][c] for rr in range(r + 1, len(grid))) else last
    cols = []
    for c in range(ncols):
        parts = []
        for r in range(len(grid)):
            v = grid[r][c].replace('\n', '')
            if v and (not parts or parts[-1] != v):
                parts.append(v)
        cols.append(' '.join(parts))
    return cols


class NoticeError(Exception):
    pass


# ------------------------------------------------------------------ 분석기

class NoticeParser:
    def __init__(self, data, filename=''):
        self.pdf = PdfDoc(data)
        self.filename = filename
        self.warn = []
        self.lines = self.pdf.all_lines()
        self.text = '\n'.join(l.text for l in self.lines)

    # -------------------------------------------------------------- 표 찾기
    def find_tables(self, pred, pages=None):
        out = []
        for p in (pages if pages is not None else range(self.pdf.npages)):
            for t in self.pdf.tables(p):
                try:
                    if pred(t):
                        out.append(t)
                except Exception:
                    pass
        return out

    def pages_with(self, *words):
        out = []
        for p in range(self.pdf.npages):
            tx = self.pdf.page_text(p).replace(' ', '')
            if any(w.replace(' ', '') in tx for w in words):
                out.append(p)
        return out

    # -------------------------------------------------------------- 기본 정보
    def complex_name(self):
        for l in self.pdf.lines(0)[:12]:
            t = l.text
            if '입주자모집공고' in t and '입주자모집공고일' not in t and not t.startswith(('■', '본 ', '※')):
                n = re.sub(r'\s*입주자\s*모집\s*공고(문)?.*$', '', t).strip()
                n = re.sub(r'^\(?\s*\d{10}\s*\)?\s*', '', n).strip()       # 주택관리번호
                if n:
                    return n
        m = re.search(r'["“]([^"”]{3,40})["”]\s*견본주택', self.text)
        return m.group(1).strip() if m else ''

    def bullet(self, *keys):
        """'■ 공급위치 : …' 같은 줄의 값"""
        for l in self.lines:
            t = l.text.replace(' ', ' ')
            for k in keys:
                pat = r'^[^\w가-힣]*' + r'\s*'.join(map(re.escape, k)) + r'\s*[:：]\s*(.+)$'
                m = re.match(pat, t)
                if m:
                    return m.group(1).strip(), l
        return None, None

    def basic(self):
        info = {}
        loc, _ = self.bullet('공급위치', '사업위치', '건설위치', '대지위치')
        info['location'] = re.sub(r'\s+', ' ', loc or '').strip()
        scale, ln = self.bullet('공급규모', '건설규모', '사업규모')
        if scale:
            scale = re.split(r'\[|\(특별공급|※', scale)[0].strip().rstrip(',')
        if not scale:
            # LH: '■ 화성동탄2 C-27블록 공공분양주택 1층 ~ 20층, 7개동, 전용면적 85㎡이하 473세대 (…)'
            for l in self.lines[:2000]:
                if re.search(r'\d+\s*개\s*동', l.text) and re.search(r'[\d,]+\s*세대', l.text) and len(l.text) < 160:
                    scale = re.sub(r'^[^\w가-힣]*', '', l.text)
                    scale = re.split(r'\s*\((?:일반|특별)', scale)[0].strip()
                    break
        info['scale'] = scale or ''
        mv, _ = self.bullet('입주시기', '입주예정', '입주예정일', '입주예정시기')
        if not mv:
            m = re.search(r'입주\s*예정\s*시기는\s*[‘\'’]?\s*(\d{2,4})\s*년\s*(\d{1,2})\s*월', self.text)
            if m:
                y = int(m.group(1))
                mv = f'{y + 2000 if y < 100 else y}년 {int(m.group(2))}월'
        info['move_in_raw'] = mv or ''
        m = re.search(r'(20\d{2})\s*년\s*(\d{1,2})\s*월', mv or '') or re.search(r'(20\d{2})\s*[.]\s*(\d{1,2})', mv or '')
        info['move_in'] = f'{m.group(1)}년 {int(m.group(2)):02d}월' if m else ''
        # 견본주택
        mh = ''
        best = -1
        ADDR = r'(?:[가-힣]+(?:시|군|구))\s+[가-힣0-9]+(?:동|읍|면|로|길|가)\b'
        for m in re.finditer(r'(?:견본주택|모델하우스|주택전시관|위치안내)\s*(?:위치|주소|장소)?\s*[:：]\s*([^\n]+)', self.text):
            v = m.group(1).strip()
            sc = 0
            if re.search(ADDR, v):
                sc += 3
            if re.search(r'\d+(-\d+)?\s*(번지)?', v):
                sc += 1
            if re.search(r'\d{1,2}:\d{2}|~|청약|접수|방문', v):
                sc -= 3
            if sc > best:
                best, mh = sc, v
        if best < 1:
            mh = ''
            # '견본주택 위치  경기도 …' 처럼 쌍점 없이 적힌 경우: 같은 줄(또는 다음 줄)의 주소 모양 글
            ADDR2 = r'((?:서울|경기|인천|부산|대구|광주|대전|울산|세종|강원|충청|충북|충남|전라|전북|전남|경상|경북|경남|제주)\S*\s+' \
                    r'[가-힣]+(?:시|군|구)\s+[^\n]{2,60})'
            for i, l in enumerate(self.lines):
                if not re.search(r'견본주택|모델하우스|주택전시관', l.text):
                    continue
                for cand in (l.text, self.lines[i + 1].text if i + 1 < len(self.lines) else ''):
                    m = re.search(ADDR2, cand)
                    if m and re.search(r'\d', m.group(1)) and not re.search(r'\d{1,2}:\d{2}', m.group(1)):
                        mh = re.split(r'\s*[(（☎]|\s{3,}|\s+[•※▶■❚]', m.group(1))[0].strip(' ,')
                        m2 = re.match(r'^(.*?\d+(?:-\d+)?(?:번지)?)(?=[\s,)]|일원|$)', mh)
                        if m2:
                            mh = m2.group(1)            # 지번(번지)까지만
                        break
                if mh:
                    break
        mh = re.sub(r'\s*[\(（]?\s*(☎|tel|TEL|분양문의).*$', '', mh or '').strip()
        mh = re.sub(r',\s*LH\s.*$|,\s*[^,]*주택전시관\s*$', '', mh).strip()
        mh = re.split(r'\s+[•※▶■❚]\s*|\s{3,}', mh)[0].strip()      # 같은 높이의 옆 칸 글이 붙은 경우
        info['model_house'] = mh
        # 분양문의
        tel = ''
        m = re.search(r'분양\s*문의\s*\)?\s*[:：]?\s*\(?\s*([0-9]{2,4}-[0-9]{3,4}(?:-[0-9]{4})?)', self.text)
        if m:
            tel = m.group(1)
        else:
            m = re.search(r'(?:주택전시관|견본주택|분양사무소|홍보관)\s*[:：]?\s*\(?\s*(0\d{1,2}-\d{3,4}-\d{4}|1\d{3}-\d{4})', self.text) or \
                re.search(r'☎\s*([0-9]{2,4}-[0-9]{3,4}(?:-[0-9]{4})?)', self.text)
            tel = m.group(1) if m else ''
        info['tel'] = tel
        info['homepage'] = self.homepage()
        info['heating'] = self.heating()
        info['name'] = self.complex_name()
        info['developers'] = self.developers()
        info['reserve_pct'] = self.reserve_pct()
        return info

    def homepage(self):
        cands = {}
        for l in self.lines:
            for m in re.finditer(r'(https?://[^\s)\]>"\'”,]+|www\.[^\s)\]>"\'”,]+)', l.text):
                u = m.group(1).rstrip('.')
                low = u.lower()
                if any(s in low for s in SKIP_URL):
                    continue
                score = cands.get(u, 0) + 1
                if '홈페이지' in l.text or '홈 페 이 지' in l.text:
                    score += 5
                cands[u] = score
        if not cands:
            return ''
        u = max(cands.items(), key=lambda kv: kv[1])[0]
        return u if u.startswith('http') else 'https://' + u

    def heating(self):
        cnt = {k: len(re.findall(k, self.text.replace(' ', ''))) for k in ('지역난방', '개별난방', '중앙난방')}
        best = max(cnt.items(), key=lambda kv: kv[1])
        if best[1]:
            return best[0]
        if re.search(r'도시가스\s*(?:보일러|개별)', self.text):
            return '개별난방'
        return ''

    def reserve_pct(self):
        m = re.search(r'공급세대수의\s*(\d{2,3})\s*%', self.text)
        return int(m.group(1)) if m else None

    def developers(self):
        """사업주체·시행위탁·수탁·시공사 상호"""
        roles = {}
        for t in self.find_tables(lambda t: any(re.sub(r'\s', '', r[0] if r else '') in ('상호', '회사명', '업체명', '상호명')
                                                for r in t.rows)):
            head = None
            for i, r in enumerate(t.rows):
                k = re.sub(r'\s', '', r[0] if r else '')
                if k in ('상호', '회사명', '업체명', '상호명'):
                    head = t.rows[i - 1] if i > 0 else None
                    if head is None:
                        continue
                    lasth = ''
                    for c in range(1, min(len(head), len(r))):
                        h = re.sub(r'\s', '', head[c] or '') or lasth
                        lasth = h
                        name = (r[c] or '').replace('\n', ' ').strip()
                        if not name:
                            continue
                        if '위탁' in h:
                            roles.setdefault('trustor', name)
                        elif '수탁' in h or '신탁' in h:
                            roles.setdefault('trustee', name)
                        elif '시공' in h:
                            roles.setdefault('builder', [])
                            if name not in roles['builder']:
                                roles['builder'].append(name)
                        elif '사업주체' in h or '시행' in h:
                            roles.setdefault('developer', name)
                    break
            if roles:
                break
        if not roles:
            for t in self.find_tables(lambda t: '사업주체' in t.flat(2) and '시공' in t.flat(2)):
                head = None
                for i, r in enumerate(t.rows[:3]):
                    if any('사업주체' in (c or '') for c in r) and any('시공' in (c or '') for c in r):
                        head = i
                        break
                if head is None or head + 1 >= len(t.rows):
                    continue
                data = t.rows[head + 1]
                for c, h in enumerate(t.rows[head]):
                    if c >= len(data) or not data[c]:
                        continue
                    v = re.sub(r'\([^)]*\d{3}-\d{2}-\d{5}[^)]*\)', '', data[c])
                    v = ', '.join(x.strip(' ,') for x in v.split('\n') if x.strip(' ,'))
                    v = re.sub(r'\s+', ' ', v).strip(' ,')
                    if '사업주체' in h:
                        roles.setdefault('developer', v)
                    elif '시공' in h:
                        roles.setdefault('builder', [v])
                if roles:
                    break
        if not roles:
            m = re.search(r'시행위탁자[^,]*?(?:인|은)?\s*([㈜(주)]*\s*주식회사\s*\S+|\S+\s*㈜|\S+\(주\)|㈜\S+)', self.text)
            if m:
                roles['trustor'] = m.group(1).strip()
            m = re.search(r'시행수탁자\s*([㈜(주)]*\s*주식회사\s*\S+|\S+㈜|\S+\(주\)|㈜\S+)', self.text)
            if m:
                roles['trustee'] = m.group(1).strip()
            m = re.search(r'시공사\s*[:：]?\s*([^\s,]+(?:㈜|\(주\))|㈜\s*\S+|주식회사\s*\S+)', self.text)
            if m:
                roles['builder'] = [m.group(1).strip()]
        b = roles.get('builder') or []
        roles['builder'] = ', '.join(b) if isinstance(b, list) else b
        return roles

    # -------------------------------------------------------------- 일정
    SCHED_KEYS = [('notice', ('입주자모집공고일', '모집공고일', '공고일', '공고')),
                  ('mh_open', ('견본주택',)),
                  ('docs', ('서류',)),
                  ('special', ('특별공급',)),
                  ('first', ('1순위',)),
                  ('second', ('2순위',)),
                  ('general', ('일반공급',)),
                  ('winners', ('당첨자',)),
                  ('contract', ('계약',))]

    def _sched_key(self, h):
        h = re.sub(r'\s', '', h or '')
        if not h:
            return None
        for key, words in self.SCHED_KEYS:
            if any(w in h for w in words):
                if key == 'general' and ('1순위' in h or '2순위' in h):
                    continue
                return key
        return None

    def schedule(self):
        out, parts = {}, {}
        for t in self.find_tables(lambda t: '당첨' in t.flat() and ('특별공급' in t.flat() or '순위' in t.flat()
                                                                   or '접수' in t.flat()),
                                  pages=range(min(12, self.pdf.npages))):
            rows = t.rows
            di = next((i for i, r in enumerate(rows) if sum(1 for c in r if to_date(c)) >= 2), None)
            if not di:
                continue
            heads = header_columns(rows, di)
            hits = [(c, self._sched_key(h), h) for c, h in enumerate(heads)]
            hits = [x for x in hits if x[1]]
            if len(hits) < 3:
                continue
            r2 = rows[di]
            got = {}
            for c, k, h in hits:
                v = r2[c] if c < len(r2) else ''
                if not v or not to_date(v):
                    continue
                got.setdefault(k, []).append((h, v))
            for k, lst in got.items():
                if k not in out:
                    out[k] = lst[0][1]
                    if len(lst) > 1:
                        parts[k] = lst
            if len(out) >= 4:
                break
        res = {}
        for k, v in out.items():
            ds = all_dates(v)
            res[k] = [d.isoformat() for d in ds]
        for k, lst in parts.items():                    # 예) 계약 전자 / 현장
            res[k + '_parts'] = [[re.sub(r'^.*?(\S+)$', r'\1', re.sub(r'\s', ' ', h).strip()),
                                  [d.isoformat() for d in all_dates(v)]] for h, v in lst]
        if 'notice' not in res:
            m = re.search(r'(?:최초\s*)?입주자\s*모집\s*공고일\s*(?:은|:|：)?\s*(20\d{2}\s*[.]\s*\d{1,2}\s*[.]\s*\d{1,2})', self.text)
            if m and to_date(m.group(1)):
                res['notice'] = [to_date(m.group(1)).isoformat()]
            else:
                m = re.search(r'\((20\d{2}\.\s*\d{1,2}\.\s*\d{1,2}\.?)\)\s*로?\s*입주자\s*모집\s*공고\s*승인', self.text)
                if m and to_date(m.group(1)):
                    res['notice'] = [to_date(m.group(1)).isoformat()]
        if 'mh_open' not in res:
            m = re.search(r'견본주택\s*(?:오픈|개관|오픈일|개관일)\s*[:：]?\s*(20\d{2}[.\s]+\d{1,2}[.\s]+\d{1,2})', self.text)
            if m and to_date(m.group(1)):
                res['mh_open'] = [to_date(m.group(1)).isoformat()]
        return res

    # -------------------------------------------------------------- 공급대상
    def supply(self):
        types, seen = [], set()
        tabs = self.find_tables(lambda t: ('약식' in t.flat(4) or '주택형' in t.flat(4)) and '세대' in t.flat(4)
                                and re.search(r'공용|계약\s*면적', t.flat(4)) and '대지비' not in t.flat(4)
                                and any(CODE_RE.match(c.replace('\n', '')) for r in t.rows for c in r if c))
        for t in tabs:
            rows = t.rows
            first = next(i for i, r in enumerate(rows) if any(CODE_RE.match(c.replace('\n', '')) for c in r if c))
            heads = header_columns(rows, first) if first else [''] * t.ncols
            # 칸 역할
            def col(pred, start=0):
                for c in range(start, len(heads)):
                    if pred(heads[c].replace(' ', '')):
                        return c
                return None
            for r in rows[first:]:
                cells = [c.replace('\n', ' ').strip() for c in r]
                cc = next((c for c, v in enumerate(cells) if CODE_RE.match(v.replace(' ', ''))), None)
                if cc is None:
                    continue
                code = cells[cc].replace(' ', '')
                area = code_area(code)
                nums = [(c, num(v)) for c, v in enumerate(cells) if c > cc and num(v) is not None]
                # 전용 = 주택형 면적과 같은 첫 숫자
                ex = next(((c, v) for c, v in nums if isinstance(v, float) and abs(v - area) < 0.01), None)
                if ex is None:
                    ex = (None, area)
                floats = [(c, v) for c, v in nums if isinstance(v, float) or (v and v < 1000 and c != (ex[0] or -1))]
                pub = sup = None
                after = [(c, v) for c, v in nums if ex[0] is None or c > ex[0]]
                for i, (c1, v1) in enumerate(after):
                    for c2, v2 in after[i + 1:i + 3]:
                        if v1 is not None and v2 is not None and abs(ex[1] + v1 - v2) < 0.0025:
                            pub, sup = (c1, v1), (c2, v2)
                            break
                    if sup:
                        break
                if sup is None:
                    pub_c = col(lambda h: '공용' in h and '주거' in h)
                    pv = num(cells[pub_c]) if pub_c is not None and pub_c < len(cells) else None
                    if isinstance(pv, (int, float)) and 0 < pv < ex[1]:
                        pub = (pub_c, pv)
                        sup = (None, round(ex[1] + pv, 4))       # 표에 공급면적 합계 칸이 없음 → 전용+공용
                etc = con = None
                if sup and sup[0] is not None:
                    rest = [(c, v) for c, v in nums if c > sup[0]]
                    for i, (c1, v1) in enumerate(rest):
                        for c2, v2 in rest[i + 1:i + 3]:
                            if abs(sup[1] + v1 - v2) < 0.0025:
                                etc, con = (c1, v1), (c2, v2)
                                break
                        if con:
                            break
                # 대지지분, 세대수: 머리글로
                land_c = col(lambda h: '대지' in h)
                units_c = col(lambda h: ('총공급' in h) or (('공급세대' in h or h.endswith('세대수') or '세대수' in h)
                                                         and '특별' not in h and '일반' not in h and '최하층' not in h
                                                         and '우선' not in h))
                gen_c = col(lambda h: '일반공급' in h)
                low_c = col(lambda h: '최하층' in h or '우선배정' in h)
                abbr_c = col(lambda h: '약식' in h or h in ('타입', '(타입)') or (h.endswith('타입') and '주택형' not in h))
                spec = {}
                for c, h in enumerate(heads):
                    hh = h.replace(' ', '')
                    if '특별' in hh and c > (units_c or 0):
                        for key, words in (('agency', ('기관',)), ('multi', ('다자녀',)), ('newlywed', ('신혼',)),
                                           ('parents', ('노부모',)), ('first', ('생애',)), ('newborn', ('신생아',)),
                                           ('youth', ('청년',)), ('total', ('계', '소계', '합계'))):
                            if any(hh.endswith(w) or (w in hh and w != '계') for w in words):
                                spec.setdefault(key, c)
                                break
                units = num(cells[units_c]) if units_c is not None and units_c < len(cells) else None
                if units is None:
                    # 머리글을 못 읽으면 계약면적(없으면 공급면적) 뒤의 첫 정수
                    after_c = (con or sup or (None,))[0]
                    ints = [v for c, v in nums if isinstance(v, int) and after_c is not None and c > after_c]
                    units = ints[0] if ints else None
                abbr = cells[abbr_c] if abbr_c is not None and abbr_c < len(cells) and cells[abbr_c] else abbr_from_code(code)
                if re.fullmatch(r'0?\d{2,3}\.\d+[A-Za-z]*', abbr.replace(' ', '')):
                    abbr = abbr_from_code(abbr)
                abbr = abbr.replace(' ', '')
                key = (code, abbr)
                if key in seen:
                    continue
                seen.add(key)
                d = {
                    'code': code, 'abbr': abbr, 'area': area, 'suffix': code_suffix(code),
                    'excl': ex[1], 'pub': pub[1] if pub else None, 'supply': sup[1] if sup else None,
                    'etc': etc[1] if etc else None, 'contract': con[1] if con else None,
                    'land': num(cells[land_c]) if land_c is not None and land_c < len(cells) else None,
                    'units': units,
                    'special': {k: (num(cells[c]) or 0) for k, c in spec.items() if c < len(cells)},
                    'general': num(cells[gen_c]) if gen_c is not None and gen_c < len(cells) else None,
                    'lowest': num(cells[low_c]) if low_c is not None and low_c < len(cells) else None,
                    'page': t.page + 1,
                }
                if d['supply'] is None:
                    self.warn.append(f'{abbr}: 공급면적(전용+공용)을 표에서 확인하지 못했습니다.')
                types.append(d)
        if not types:
            self.warn.append('공급대상 표(주택형·약식표기·세대수)를 찾지 못했습니다.')
        return types

    # -------------------------------------------------------------- 공급금액
    @staticmethod
    def _split_multiline(row):
        """칸 안 줄바꿈으로 여러 층이 한 줄에 들어 있는 표 → 층별 줄로 나눈다."""
        parts = [c.split('\n') if c else [''] for c in row]
        n = max(len(p) for p in parts)
        if n <= 1:
            return [row]
        numeric_multi = [p for p in parts if len(p) == n and sum(1 for x in p if is_num(x)) >= n - 1]
        if len(numeric_multi) < 2:
            return [[c.replace('\n', ' ') for c in row]]
        out = []
        for i in range(n):
            out.append([(p[i] if len(p) == n else (p[0] if i == 0 else '')).strip() for p in parts])
        return out

    FLOOR_RE = re.compile(r'(\d+\s*층|층\s*이상|최상층|최고층|필로티|\d+\s*~\s*\d+|\d+\s*-\s*\d+\s*층|\d+\s*F\b|이상|이하|기준층|저층)')

    def _price_map(self, rows, first):
        """공급금액 표 머리글 → 칸 역할. 못 찾으면 None"""
        heads = header_columns(rows, first) if first else []
        if not heads:
            return None
        hs = [h.replace(' ', '') for h in heads]
        floor_c, best = None, 0
        for c in range(len(hs)):
            hdr = ('층' in hs[c] and ('구분' in hs[c] or '별' in hs[c])) or hs[c] in ('층', '층별', '해당층', '층수')
            k = 0
            for r in rows[first:]:
                v = r[c] if c < len(r) else ''
                if not v:
                    continue
                if self.FLOOR_RE.search(v) and not is_num(v):
                    k += 1
                elif hdr and re.fullmatch(r'\d{1,2}(\s*[~\-]\s*\d{1,2})?', v.strip()):
                    k += 1
            if hdr:
                k += 3
            if k > best:
                best, floor_c = k, c
        if floor_c is None:
            return None
        units_c = None
        for c in range(floor_c + 1, min(floor_c + 4, len(hs))):
            if '세대' in hs[c]:
                units_c = c
                break
        if units_c is None:
            vals = [r[floor_c + 1] for r in rows[first:] if floor_c + 1 < len(r) and r[floor_c + 1]]
            ints = [v for v in vals if isinstance(num(v), int) and 0 < num(v) <= 3000]
            units_c = floor_c + 1 if vals and len(ints) >= len(vals) * 0.7 else None   # LH: 층별 세대수 없음
        land_c = next((c for c, h in enumerate(hs) if '대지' in h), None)
        build_c = next((c for c, h in enumerate(hs) if '건축' in h), None)
        vat_c = next((c for c, h in enumerate(hs) if '부가' in h), None)
        def val(r, c):
            return num(r[c]) if c is not None and c < len(r) else None
        total_c = None
        lo = max([x for x in (land_c, build_c, vat_c, units_c, floor_c) if x is not None])
        for c, h in enumerate(hs):
            if c <= lo:
                continue
            if re.search(r'(계|공급금액|분양금액|분양가격|합계|공급가격|분양가|총액|주택가격)$', h):
                total_c = c
                break
        if land_c is not None and build_c is not None:
            for c in range(max(land_c, build_c) + 1, min(len(hs), max(land_c, build_c) + 4)):
                good = sum(1 for r in rows[first:]
                           if val(r, land_c) and val(r, build_c) and val(r, c)
                           and abs(val(r, land_c) + val(r, build_c) + (val(r, vat_c) or 0) - val(r, c)) <= 10)
                if good:
                    total_c = c
                    break
        if total_c is None:
            return None
        dep_c = None
        for c, h in enumerate(hs):
            if '계약금' in h and ('1차' in h or '계약시' in h or '체결시' in h):
                dep_c = c
                break
        return {'ncols': len(hs), 'floor': floor_c, 'units': units_c, 'land': land_c, 'build': build_c,
                'vat': vat_c, 'total': total_c, 'dep': dep_c, 'heads': heads}

    def _continue_map(self, prev, p, rows, first):
        """머리글 없이 다음 쪽으로 이어진 표: 앞 표의 칸 역할을 빠진 앞칸 수만큼 밀어 맞춘다 (금액 검산으로 확인)"""
        mp0, p0 = prev
        if p - p0 > 3:
            return None
        ncols = max(len(r) for r in rows)
        def val(r, c):
            return num(r[c]) if c is not None and 0 <= c < len(r) else None
        best, best_s = 0, None
        for sh in range(0, max(1, mp0['ncols'] - ncols + 1) + 1):
            m = {k: (v - sh if isinstance(v, int) else v) for k, v in mp0.items() if k not in ('heads', 'ncols')}
            if m['floor'] is None or m['floor'] < 0 or m['total'] is None or m['total'] >= ncols:
                continue
            good = 0
            for r in rows[first:]:
                tt, u = val(r, m['total']), val(r, m['units'])
                if not tt or not isinstance(u, int):
                    continue
                if m['land'] is not None and m['build'] is not None:
                    a, b = val(r, m['land']), val(r, m['build'])
                    if a and b and abs(a + b + (val(r, m['vat']) or 0) - tt) <= 10:
                        good += 1
                elif tt >= 1_000_000:
                    good += 1
            if good > best:
                best, best_s = good, m
        if not best_s:
            return None
        best_s['ncols'] = ncols
        best_s['heads'] = []
        best_s['scale'] = mp0.get('scale', 1)
        return best_s

    def prices(self, types):
        tm = TypeMatcher(types)
        units_of = {t['abbr']: t['units'] for t in types}
        order = [t['abbr'] for t in types]
        tiers = {a: [] for a in order}
        payment_heads, fixed_first = [], []
        hits = self.pages_with('대지비', '분양금액', '분양가격', '공급가격')
        cand_pages = []
        for p in hits:                  # 금액표 쪽 + 머리글 없이 이어지는 다음 쪽들
            for q in range(p, min(p + 4, self.pdf.npages)):
                if q not in cand_pages:
                    cand_pages.append(q)
        cand_pages.sort()
        prev = None                     # (칸 역할, 쪽) — 다음 쪽 표가 머리글 없이 이어질 때 씀
        cur = None
        def val(r, c):
            return num(r[c]) if c is not None and c < len(r) else None
        for p in cand_pages:
            # 금액 단위: '(단위 : 천원)' 이면 1,000 배
            ptxt = self.pdf.page_text(p).replace(' ', '')
            scale = 1000 if re.search(r'단위[:：]?[^)\n]{0,12}천원', ptxt) else 1
            if scale == 1 and prev and prev[0].get('scale') == 1000 and p - prev[1] <= 1:
                scale = 1000
            big = 1_000_000 // scale
            for t in self.pdf.tables(p):
                rows = []
                for r in t.rows:
                    rows.extend(self._split_multiline(r))
                first = next((i for i, r in enumerate(rows)
                              if sum(1 for c in r if (num(c) or 0) >= big) >= 2), None)
                if first is None:
                    continue
                head_txt = ' '.join(' '.join(r) for r in rows[:first])
                mp = self._price_map(rows, first) if first > 0 and re.search(r'층|구분|대지|공급금액|분양', head_txt) else None
                if mp is None:
                    # 머리글 없이 이어진 표만 앞 표의 칸 배치로 읽는다 (다른 표의 머리글이 있으면 건너뜀)
                    own_head = first > 0 and any(c and not is_num(c) for r in rows[:first] for c in r)
                    mp = self._continue_map(prev, p, rows, first) if prev and not own_head else None
                    if mp is None:
                        continue
                else:
                    cur = None
                    payment_heads.extend(h for h in mp['heads'] if re.search(r'계약금|중도금|잔금', h))
                mp['scale'] = scale
                prev = (mp, p)
                type_cs = list(range(0, mp['floor']))
                for r in rows[first:]:
                    if any('마이너스' in (c or '') for c in r[:mp['total']]):
                        continue                          # 마이너스옵션 금액 줄은 빼고 기본형만
                    lab = None
                    for c in type_cs:
                        if c < len(r) and r[c]:
                            c0 = r[c].split('\n')[0]
                            hit = tm.one(c0) or tm.one(re.sub(r'\(.*$', '', c0))
                            if hit:
                                lab = hit
                                break
                    if lab is None:
                        for c in type_cs:
                            v = val(r, c)
                            if isinstance(v, int) and v > 0:
                                cand = [a for a in order if units_of.get(a) == v and not tiers[a]]
                                if cand:
                                    lab = cand[0]
                                    break
                    fl = (r[mp['floor']] if mp['floor'] < len(r) else '').replace('\n', ' ').strip()
                    u, tt = val(r, mp['units']), val(r, mp['total'])
                    if mp['units'] is None:
                        u = None                          # 층별 세대수가 공고에 없음 (LH)
                        if tt is None:
                            continue
                    elif tt is None or u is None or not isinstance(u, int) or u <= 0 or u > 3000:
                        continue
                    tt = tt * scale
                    if tt < 1_000_000:
                        continue
                    if lab is None and cur is None and not any(tiers[a] for a in order):
                        cur = order[0]                    # 타입 이름이 병합 칸 가운데에 있어 첫 줄에 없는 표
                    if lab is not None:
                        cur = lab
                    elif cur is not None and units_of.get(cur) and u is not None and \
                            sum(x['units'] or 0 for x in tiers[cur]) >= units_of[cur]:
                        # 타입 이름 칸이 비어 있고 앞 타입은 세대수가 다 찼다 → 다음 타입
                        nxt = [a for a in order[order.index(cur) + 1:] if not tiers[a]]
                        cur = nxt[0] if nxt else cur
                    if cur is None:
                        continue
                    if not fl:
                        fl = '-'
                    if re.fullmatch(r'\d{1,2}(\s*[~\-]\s*\d{1,2})?', fl):
                        fl += '층'
                    def sv(c):
                        v = val(r, c)
                        return v * scale if v is not None else None
                    tiers[cur].append({'floor': fl, 'units': u, 'land': sv(mp['land']), 'build': sv(mp['build']),
                                       'vat': sv(mp['vat']), 'total': tt, 'page': p + 1})
                    if mp['dep'] is not None:
                        fixed_first.append(sv(mp['dep']))
        # 표로 다 못 읽은 타입은 줄 글자로 다시 읽어 세대수가 맞으면 그것을 쓴다
        bad = [a for a in order if units_of.get(a) and (not tiers[a] or all(x['units'] is not None for x in tiers[a]))
               and sum(x['units'] or 0 for x in tiers[a]) != units_of[a]]
        if bad:
            by_lines = self.prices_by_lines(types, cand_pages)
            for a in bad:
                lst = by_lines.get(a) or []
                if lst and sum(x['units'] for x in lst) == units_of[a]:
                    tiers[a] = lst
            if not payment_heads:
                payment_heads.extend(getattr(self, '_line_heads', []))
            if not fixed_first:
                fixed_first.extend(getattr(self, '_line_first', []))
        for a, lst in tiers.items():
            s = sum(x['units'] or 0 for x in lst)
            if not lst:
                self.warn.append(f'{a}: 공급금액 표에서 층별 금액을 찾지 못했습니다.')
            elif any(x['units'] is None for x in lst):
                self.no_tier_units = True
            elif units_of.get(a) and s != units_of[a]:
                self.warn.append(f'{a}: 층별 세대수 합 {s} ≠ 공급세대수 {units_of[a]} (공급금액 표를 확인하세요)')
        self._payment_heads = payment_heads
        self._fixed_first = [x for x in fixed_first if x]
        return tiers

    LINE_PRICE_RE = re.compile(
        r'(?P<floor>(?:\d+\s*(?:층|F)?\s*[~\-]\s*\d+\s*(?:층|F)(?:\s*이상)?|\d+(?:\s*,\s*\d+)+\s*(?:층|F)|'
        r'\d+\s*(?:층|F)(?:\s*이상|\s*이하)?|기준층|최상층|최고층|필로티))\s+(?P<units>\d{1,4})\s+'
        r'(?P<nums>(?:(?:[\d,]{5,}|-)\s+){2,}[\d,]{5,})')

    def prices_by_lines(self, types, pages):
        """칸 선이 없어 표로 안 잡히는 공급금액 표: '11~20층 60 265,861,460 346,138,540 612,000,000 …' 줄을 읽는다"""
        tm = TypeMatcher(types)
        units_of = {t['abbr']: t['units'] for t in types}
        order = [t['abbr'] for t in types]
        tiers = {a: [] for a in order}
        cur = None
        heads, firsts = [], []
        for p in pages:
            ptxt = self.pdf.page_text(p).replace(' ', '')
            scale = 1000 if re.search(r'단위[:：]?[^)\n]{0,12}천원', ptxt) else 1
            for l in self.pdf.lines(p):
                tx = l.text
                if re.search(r'계약금|중도금|잔금', tx) and '%' in tx:
                    heads.append(tx)
                toks = tx.split()
                # 줄 앞머리의 타입 이름 (숫자만인 이름은 그 타입 세대수가 같은 줄에 있을 때만)
                for tok in toks[:2]:
                    hit = tm.one(tok)
                    if hit and (not re.fullmatch(r'\d+', tok) or str(units_of.get(hit)) in toks
                                or f'{units_of.get(hit) or 0:,}' in toks or len(toks) == 1):
                        cur = hit
                        break
                m = self.LINE_PRICE_RE.search(tx)
                if not m or '마이너스' in tx:
                    continue
                nums = [0 if x == '-' else num(x) for x in m.group('nums').split()]
                nums = [x for x in nums if x is not None]
                total = None
                for k in (2, 3):
                    if len(nums) > k and nums[0] and nums[1] and abs(sum(nums[:k]) - nums[k]) <= 10:
                        total = nums[k]
                        land, build, vat = nums[0], nums[1], (nums[2] if k == 3 else None)
                        dep = nums[k + 1] if len(nums) > k + 1 else None
                        break
                if not total:
                    continue
                u = int(m.group('units'))
                if cur is None or (units_of.get(cur) and sum(x['units'] for x in tiers[cur]) >= units_of[cur]):
                    nxt = [a for a in (order[order.index(cur) + 1:] if cur else order) if not tiers[a]]
                    if nxt:
                        cur = nxt[0]
                if cur is None:
                    continue
                fl = re.sub(r'\s+', ' ', m.group('floor')).strip()
                tiers[cur].append({'floor': fl, 'units': u, 'land': land * scale, 'build': build * scale,
                                   'vat': vat * scale if vat else vat, 'total': total * scale, 'page': p + 1})
                if dep:
                    firsts.append(dep * scale)
        self._line_heads, self._line_first = heads, firsts
        return tiers

    def payment_terms(self):
        heads = ' '.join(getattr(self, '_payment_heads', []))
        dep = re.search(r'계약금\s*\(?\s*(\d+)\s*%', heads)
        mid = re.findall(r'(\d+)\s*(?:회차|차|회)\s*\(?\s*(\d+)\s*%', heads) + \
            re.findall(r'(\d+)\s*차\s*중도금\s*\(?\s*(\d+)\s*%', heads)
        midtot = re.search(r'중도금\s*\(?\s*(\d+)\s*%', heads)
        bal = re.search(r'잔금\s*\(?\s*(\d+)\s*%', heads)
        out = {'deposit_pct': int(dep.group(1)) if dep else None,
               'mid_pct_total': int(midtot.group(1)) if midtot else None,
               'balance_pct': int(bal.group(1)) if bal else None}
        # 중도금 회차: 'n차(10%)' 중 계약금 1·2차는 % 가 없거나 같을 수 있으니 중도금 합으로 맞춘다
        pcts = [int(p) for _, p in mid]
        if out['mid_pct_total'] and pcts:
            each = max(set(pcts), key=pcts.count)
            cnt = round(out['mid_pct_total'] / each) if each else None
            out['mid_each'], out['mid_count'] = each, cnt
        elif out['mid_pct_total']:
            out['mid_each'], out['mid_count'] = None, None
        lh = sorted(set(re.findall(r'(\d+)\s*차\s*중도금\s*\(?\s*(\d+)\s*%', heads)))
        if lh:                                          # LH: '1차 중도금 15%' … 로 회차가 따로 적힘
            pc = [int(p) for _, p in lh]
            out['mid_each'] = max(set(pc), key=pc.count)
            out['mid_count'] = len({n for n, _ in lh})
            out['mid_pct_total'] = out['mid_each'] * out['mid_count']
        ff = [x for x in getattr(self, '_fixed_first', []) if x]
        out['first_fixed'] = max(set(ff), key=ff.count) if ff and ff.count(max(set(ff), key=ff.count)) >= len(ff) * 0.8 else None
        if ff and out['first_fixed'] is None:
            # 타입마다 다르면(84 타입 1천만원·큰 타입 3천만원) 첫 타입 금액 (2026-10-18 사용자 파일 기준)
            out['first_fixed'] = ff[0]
            out['first_fixed_mixed'] = sorted(set(ff))
        mid_txt = ' '.join(l.text for l in self.lines if '중도금' in l.text)
        if '이자후불' in mid_txt.replace(' ', ''):
            out['interest'] = '이자후불제'
        elif '무이자' in mid_txt.replace(' ', ''):
            out['interest'] = '무이자'
        else:
            out['interest'] = ''
        return out

    # -------------------------------------------------------------- 발코니 확장
    def balcony(self, types):
        tm = TypeMatcher(types)
        res = {}
        pages = self.pages_with('발코니 확장', '발코니확장')
        def take(abbrs, v):
            for a in abbrs:
                if a not in res and v:
                    res[a] = v
        def balc_table(t):
            head = t.title + ' ' + t.flat(2)
            return ('확장' in head and '대지비' not in head and '납부계좌' not in t.flat(2)
                    and '계좌번호' not in t.flat(2) and not re.search(r'에어컨|공기청정|가전|중문', head))
        for t in self.find_tables(balc_table, pages=pages):
            rows = t.rows
            # (1) 가로형: 타입이 머리글에
            for i, r in enumerate(rows):
                hit = [(c, tm.many(v)) for c, v in enumerate(r) if v and tm.many(v)]
                if len(hit) >= 2 or (len(hit) == 1 and len(types) == 1):
                    for r2 in rows[i + 1:]:
                        lab = ' '.join(r2[:max(1, hit[0][0])]).replace(' ', '')
                        if ('확장' in lab or '금액' in lab or '공급' in lab or '합계' in lab) and '계약금' not in lab                                 and '중도금' not in lab and '잔금' not in lab:
                            for c, abbrs in hit:
                                if c < len(r2) and num(r2[c]):
                                    take(abbrs, num(r2[c]))
                            break
                    break
            # (2) 세로형: 타입이 첫 칸에, 금액 칸은 머리글 '확장 금액/공사비' 
            amt_c = None
            for c, h in enumerate(rows[0] if rows else []):
                hh = h.replace(' ', '').replace('\n', '')
                if ('확장' in hh and re.search(r'금액|공사비|대금|공급가|비용', hh)) or hh in ('공급금액', '금액', '합계', '총액'):
                    amt_c = c
                    break
            if amt_c is None:
                continue
            unit = 1000 if re.search(r'천\s*원', t.title + ' ' + self.pdf.page_text(t.page)[:3000]) else 1
            cur_abbrs = []
            for r in rows[1:]:
                abbrs = tm.many(r[0]) if r else []
                if not abbrs and len(r) > 1 and amt_c > 1:
                    abbrs = tm.many(r[1])
                if abbrs:
                    cur_abbrs = abbrs
                elif r and not r[0] and cur_abbrs:
                    abbrs = []                            # 같은 타입의 다음 줄 (마이너스옵션 등)
                if not abbrs or any('마이너스' in (c or '') for c in r):
                    continue
                v = num(r[amt_c]) if amt_c < len(r) else None
                if v and unit == 1000 and v < 100_000:
                    v *= 1000
                if v and 100_000 <= v <= 200_000_000:
                    take(abbrs, v)
        # (3) 줄 글자: 발코니 확장 구역 안의 '84A 4,500,000 450,000 …' 줄 (표에서 타입 칸이 비어 읽힌 경우, 쪽을 넘어 이어짐)
        if len(res) < len(types):
            HEAD = re.compile(r'^(■|▣|□|◆|●|▶|\d{1,2}\s*[.)]|\d{1,2}-\d|[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]|\[)')
            in_sec, last_p = False, None
            for p in pages:
                if in_sec and last_p is not None and p != last_p + 1:
                    in_sec = False
                for l in self.pdf.lines(p):
                    tx = l.text.strip()
                    if re.search(r'발코니\s*확장', tx) and (re.search(r'금액|공사비|공급가|대금|비용|계약', tx) or HEAD.match(tx))                             and not re.search(r'납부계좌|계좌', tx):
                        in_sec = True
                        continue
                    if in_sec and HEAD.match(tx) and '발코니' not in tx:
                        in_sec = False
                        continue
                    if not in_sec:
                        continue
                    m = re.match(r'^((?:\S+?)(?:\s*,\s*[A-Z0-9]+)*)\s+([\d,]{6,})\b', tx)
                    if m:
                        v = num(m.group(2))
                        if v and 100_000 <= v <= 200_000_000:
                            take(tm.many(m.group(1)), v)
                last_p = p
        # 전 세대 무상 확장
        if len(res) < len(types) and re.search(r'발코니\s*확장[^.\n]{0,30}무상|무상\s*(으로)?\s*(발코니\s*)?확장', self.text):
            for t in types:
                res.setdefault(t['abbr'], 0)
            self.warn.append('공고에 발코니 확장이 무상(전 세대 확장형)으로 되어 있어 확장비를 0 으로 둡니다.')
        for t in types:
            if t['abbr'] not in res:
                self.warn.append(f"{t['abbr']}: 발코니 확장 금액을 찾지 못했습니다 (확장비 0 으로 둡니다).")
        return res

    # -------------------------------------------------------------- 전체
    def run(self):
        if len(self.text.replace(' ', '')) < 2000:
            raise NoticeError('이 PDF 에는 글자가 없습니다 (스캔한 그림 PDF). 청약홈(applyhome.co.kr)에서 받은 '
                              '글자 있는 입주자모집공고 PDF 를 넣어 주세요.')
        info = self.basic()
        sched = self.schedule()
        types = self.supply()
        tiers = self.prices(types) if types else {}
        pay = self.payment_terms()
        balc = self.balcony(types) if types else {}
        for t in types:
            lst = tiers.get(t['abbr']) or []
            t['tiers'] = lst
            if lst and any(x['units'] is None for x in lst):
                t['base'] = len(lst) - 1                  # 층별 세대수가 없으면 맨 아래(가장 높은 층) 구간
            elif lst:
                # 세대수가 가장 많은 층, 같으면 가장 높은 층 (2026-10-15 사용자 지정)
                best = max(range(len(lst)), key=lambda i: (lst[i]['units'], floor_rank(lst[i]['floor'], i)))
                t['base'] = best
            else:
                t['base'] = None
            t['balcony'] = balc.get(t['abbr'])
        # 평균평당가: 전체 분양가 합 ÷ 전체 공급면적(평)
        if getattr(self, 'no_tier_units', False):
            self.warn.append('공고 공급금액 표에 층별 세대수가 없어, 공급금액은 가장 높은 층 구간 금액으로 두고 평균평당가는 '
                             '그 금액 × 타입 세대수로 계산했습니다 (확인 필요).')
            for t in types:
                for x in t['tiers']:
                    x['units'] = x['units'] if x['units'] is not None else 0
                if t['tiers'] and t['base'] is not None and not any(x['units'] for x in t['tiers']):
                    t['tiers'][t['base']]['units'] = t['units'] or 0
        tot_price = sum(x['total'] * x['units'] for t in types for x in t['tiers'])
        tot_py = sum((t['supply'] or 0) * 0.3025 * sum(x['units'] for x in t['tiers']) for t in types)
        avg = tot_price / tot_py if tot_py else None
        return {
            'info': info, 'schedule': sched, 'types': types, 'payment': pay,
            'avg_price_per_py': avg, 'total_price': tot_price, 'total_py': tot_py,
            'pages': self.pdf.npages, 'warnings': self.warn,
        }


def parse_notice(data, filename=''):
    return NoticeParser(data, filename).run()
