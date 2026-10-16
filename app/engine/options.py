"""모집공고의 유상옵션(추가 선택품목·시스템에어컨 등) 표 → 품목 묶음.

묶음(group) = 분양개요 '유상옵션(C)' 의 한 칸. 묶음마다 대안(variant) 이 있고 타입별로 대안 하나를 고른다.
  - 가로형 표 (타입이 머리글): 줄마다 품목. '○○ 선택 시 / 미선택 시', '선택1/2' 같은 줄은 한 묶음의 대안.
  - 세로형 표 (타입이 첫 칸): 같은 타입의 여러 줄(2대/4대 …)은 한 묶음의 대안.
기본 선택 규칙 (화면에서 바꿀 수 있음)
  - 공기청정기·환기·주방가전(냉장고·오븐 등)·마이너스옵션·무상품목은 기본 제외.
  - 대수가 있는 품목(에어컨 등)은 그 타입에 해당하는 가장 많은 대수. '○○ 선택 시 미해당' 은 ○○ 를 넣으면 뺀다.
  - '○○ 선택 시 / 미선택 시' 는 ○○ 를 넣었으면 '선택 시' 금액(그 타입 금액이 없으면 '미선택 시').
"""
import re

from .notice import CODE_RE, TypeMatcher, num

EXCLUDE_RE = re.compile(r'공기\s*청정|환기|냉장고(?!장)|김치|오븐|광파|세탁기|건조기|스타일러|의류\s*관리|와인|정수기|주방\s*가전|'
                        r'가전\s*옵션|마이너스|무상|TV|티비|음식물|비데|청소기|커튼|블라인드')
ALT_RE = re.compile(r'(미?선택\s*시|선택\s*\d+|선택안|미적용|적용\s*시|TYPE\s*[A-Z]|타입\s*[A-Z]|일반형|고급형|기본형|확장형)', re.I)
SKIP_LABEL_RE = re.compile(r'계약금|중도금|잔금|납부|계좌|은행|예금주|합\s*계$|^계$')
TEXT_ZERO_RE = re.compile(r'기본|무상|미해당|해당\s*없|선택\s*불가|불가|미적용|없음|^-$|^‐$|^—$')
FLOOR_ROW_RE = re.compile(r'^\s*\d+\s*~\s*\d+\s*층|\d+\s*층\s+\d+\s+[\d,]{5,}|\d+\s*호\s*라인|^\s*\d+\s*층\s+\d+\s')
TITLE_HINT = re.compile(r'공급대금|공급금액|공급가격|설치\s*내역|\(단위|단위\s*:|품목|옵션|에어컨|공기청정|가전|Style|스타일')


def clean_title(t):
    t = re.sub(r'\(\s*단위.*$', '', t)
    t = re.sub(r'단위\s*[:：].*$', '', t)
    t = re.sub(r'^[\s■□▣◆●○▶◎※·\-\d.)]+', '', t)
    t = re.sub(r'설치\s*내역\s*및\s*공급\s*(대금|금액)|및\s*공급\s*(대금|금액)|공급\s*(대금|금액)', '', t)
    t = re.sub(r'\[[^\]]*\]', '', t)
    t = re.sub(r'\(\s*※[^)]*\)', '', t)
    t = re.sub(r'\s+', ' ', t).strip(' -:：')
    return t


def display_name(name):
    """비고 칸 목록용 짧은 이름: '주방 [주방특화] 엔지니어드스톤 벽/상판 +…' → '주방 주방특화'"""
    if '에어컨' in (name or ''):
        return '시스템에어컨'
    s = re.sub(r'\([^)]*\)', '', name or '')
    m = re.search(r'\[([^\]]+)\]\s*([^\s+/,]*)', s)
    if m:
        head = s[:m.start()].strip()
        tail = m.group(2) if m.group(2) and len(m.group(2)) <= 8 else ''
        s = ' '.join(x for x in (head, m.group(1), tail) if x)
    s = re.split(r'\s*[+/]', s)[0] if len(s) > 16 else re.split(r'\s*\+', s)[0]
    s = re.sub(r'\s+', ' ', s).strip()
    return s[:22]


def short_name(name):
    s = re.sub(r'\([^)]*\)', '', name)
    s = re.sub(r'\s+', ' ', s).strip()
    rep = [('천장형 시스템에어컨', '에어컨'), ('시스템에어컨', '에어컨'), ('시스템 에어컨', '에어컨'),
           ('현관 중문', '중문'), ('현관중문', '중문'), ('욕실 복합 환풍기', '복합환풍기'), ('Style Up', '스타일업'),
           ('Style up', '스타일업'), ('STYLE UP', '스타일업')]
    for a, b in rep:
        s = s.replace(a, b)
    return s[:14]


def price_of(v):
    """표 칸 값 → 금액(원). 기본제공·-·무상 → None. 옵션 금액은 1천원~1억원 사이로 본다"""
    if v is None:
        return None
    s = str(v).replace('\n', ' ').strip()
    if not s or TEXT_ZERO_RE.search(s):
        return None
    n = num(s.split(' ')[0])
    return n if isinstance(n, (int, float)) and 1000 <= n <= 100_000_000 else None


def units_count(text):
    m = re.search(r'(\d+)\s*대', text or '')
    return int(m.group(1)) if m else None


class OptionReader:
    def __init__(self, parser, types):
        self.p = parser
        self.types = types
        self.order = [t['abbr'] for t in types]
        self.tm = TypeMatcher(types)
        self.groups = []

    # 표 위쪽에서 제목 줄 찾기
    def title_for(self, t):
        lines = [l for l in self.p.pdf.lines(t.page) if l.y < t.bbox[1] - 1]
        if t.page > 0 and len(lines) < 2:
            lines = self.p.pdf.lines(t.page - 1)[-6:] + lines
        for l in reversed(lines[-12:]):
            tx = l.text.strip()
            if tx.startswith(('⦁', '•', '·', '-', '※', '*')) or tx.startswith('설치'):
                continue
            if len(self.tm.many(tx)) >= 2:          # 다른 표의 타입 머리글 줄
                continue
            if TITLE_HINT.search(tx) or tx.startswith(('■', '▣', '□', '◆', '●')):
                ct = clean_title(tx)
                if ct and len(ct) >= 2:
                    return ct
        return clean_title(t.title) if t.title else ''

    def option_pages(self):
        """옵션 표가 있는 쪽: '추가 선택품목·유상옵션·에어컨·공기청정기·가전' 이 나오는 쪽들의 처음~끝"""
        hit = []
        for p in range(self.p.pdf.npages):
            tx = self.p.pdf.page_text(p).replace(' ', '')
            if re.search(r'추가선택|유상옵션|플러스옵션|선택품목|에어컨|공기청정|가전옵션|옵션품목|빌트인|유상\)', tx):
                tables = self.p.pdf.tables(p)
                if any(price_of(c) for t in tables for r in t.rows for c in r):
                    hit.append(p)
                elif tables and re.search(r'추가선택|유상옵션|플러스옵션|선택품목|옵션품목|유상\)', tx) \
                        and len(re.findall(r'\d{1,3}(?:,\d{3}){1,3}(?!\d)', self.p.pdf.page_text(p))) >= 4:
                    hit.append(p)               # 표 칸에 금액이 안 잡히는 공고(칸 안 점선 등) — 금액 글자가 여럿인 옵션 쪽
        if not hit:
            return []
        return list(range(hit[0], min(self.p.pdf.npages, hit[-1] + 3)))   # 표가 다음 쪽으로 이어질 수 있음

    # -------------------------------------------------------------- 가로형
    def read_wide(self, t, hrow, tcols, title):
        rows = t.rows
        first_type_c = min(tcols)
        group_label = ''
        items = []
        for r in rows[hrow + 1:]:
            labels = [(r[c] or '').replace('\n', ' ').strip() for c in range(min(first_type_c, len(r)))]
            if labels and labels[0]:
                group_label = labels[0]
            elif labels and not labels[0] and len(labels) > 1 and labels[1]:
                pass                                     # 병합된 묶음 이름을 이어 씀
            else:
                group_label = group_label if any(labels) else group_label
            sub = ' '.join(x for x in labels[1:] if x)
            full = (group_label + ' ' + sub).strip()
            if not full or SKIP_LABEL_RE.search(group_label.replace(' ', '')) and not sub:
                continue
            if SKIP_LABEL_RE.search(full):
                continue
            prices = {}
            raw = {}
            for c, abbrs in tcols.items():
                if c >= len(r):
                    continue
                for a in abbrs:
                    raw[a] = (r[c] or '').replace('\n', ' ').strip()
                    prices[a] = price_of(r[c])
            if not any(v for v in prices.values()) and not any(raw.values()):
                continue
            is_value_row = re.sub(r'\s', '', group_label) in ('공급금액', '공급대금', '금액', '공급가격', '판매가', '공급가', '분양가')
            name = title if is_value_row or not group_label else group_label
            items.append({'group': name, 'sub': sub if not is_value_row or sub else '', 'prices': prices, 'raw': raw,
                          'value_row': is_value_row})
        # 묶기: 같은 이름 줄이 여럿이면 대안(…선택 시) 이거나 따로 고르는 품목(붙박이장 침실2/침실3)
        count = {}
        for it in items:
            count[it['group']] = count.get(it['group'], 0) + 1
        by = {}
        for it in items:
            alt = bool(ALT_RE.search(it['sub'])) or (it['value_row'] and bool(it['sub']))
            several = count[it['group']] > 1 and not alt
            key = it['group'] + ('|' + it['sub'] if several else '')
            g = by.get(key)
            if g is None:
                base = re.sub(r'\s*\((개별|별도)?\s*선택[^)]*\)', '', it['group']).strip()
                sub_b = _brief(it['sub'])
                for w in base.split():
                    sub_b = sub_b.replace(w, '').strip()
                nm = f"{base} {sub_b}".strip() if several else base
                g = {'name': nm, 'detail': '' if alt else it['sub'], 'variants': [], 'page': t.page + 1, 'kind': 'wide',
                     'title': title}
                by[key] = g
                self.groups.append(g)
            g['variants'].append({'label': it['sub'] if alt else (it['sub'] or '공급금액'), 'prices': it['prices'],
                                  'raw': it['raw']})

    # -------------------------------------------------------------- 세로형
    def backfill_amounts(self, t, amt_c, first, key_c=None):
        """표 인식에서 금액 칸이 비어 나오는 공고(칸 안 점선 등): 그 줄의 금액 칸 자리에 찍힌 숫자를 위치로 읽어 채운다"""
        rows = [list(r) for r in t.rows]
        if not t.cells or all((r[amt_c] if amt_c < len(r) else '') for r in rows[first:] if any(r)):
            return rows
        words = self.p.pdf.words(t.page)
        col_box = next((cb[amt_c] for cb in t.cells if amt_c < len(cb) and cb[amt_c]), None)
        for i in range(first, len(rows)):
            r = rows[i]
            if amt_c >= len(r) or r[amt_c]:
                continue
            cb = t.cells[i]
            box = cb[amt_c] if amt_c < len(cb) and cb[amt_c] else None
            kb = cb[key_c] if key_c is not None and key_c < len(cb) and cb[key_c] else None
            if kb:
                y0, y1 = kb[1], kb[3]                 # 그 줄의 타입(또는 품목) 칸 높이 — 병합된 '구분' 칸은 여러 줄에 걸쳐 있음
            else:
                hs = sorted(((c[3] - c[1]), c) for c in cb if c)
                if not hs:
                    continue
                c = hs[0][1]
                y0, y1 = c[1], c[3]
            if not (box or col_box):
                continue
            x0, x1 = (box or col_box)[0], (box or col_box)[2]
            found = [w for w in words if x0 - 2 <= (w[0] + w[2]) / 2 <= x1 + 2 and y0 - 1 <= (w[1] + w[3]) / 2 <= y1 + 1
                     and re.fullmatch(r'[\d,]{4,}', w[4]) and (num(w[4]) or 0) >= 1000]
            found.sort(key=lambda w: (round(w[1]), w[0]))
            if found:
                r[amt_c] = '\n'.join(w[4] for w in found)
        return rows

    @staticmethod
    def expand_rows(rows, key_cols, amt_c=None):
        """'59B⏎84A⏎84B' | '1,500,000⏎1,530,000⏎…' 처럼 한 칸에 여러 줄 → 줄마다 한 행.
        금액 칸(amt_c)의 줄 수가 기준: 같은 줄 수인 칸은 나누고, 아닌 칸은 그대로 복사('1안)…2안)…' 은 안 단위로 나눔)"""
        out = []
        for r in rows:
            lines = lambda c: [x.strip() for x in (r[c] or '').split('\n') if x.strip()] if c is not None and c < len(r) else []
            n = len(lines(amt_c)) if amt_c is not None else max((len(lines(c)) for c in key_cols), default=1)
            if n <= 1:
                out.append(r)
                continue
            split = {}
            for c in range(len(r)):
                ls = lines(c)
                if len(ls) == n:
                    split[c] = ls
                else:
                    parts = [x.strip() for x in re.split(r'(?=\d\s*안\s*\))', (r[c] or '').replace('\n', ' ')) if x.strip()]
                    if len(parts) == n:
                        split[c] = parts
            if amt_c is None and not all(c in split for c in key_cols if c is not None and c < len(r)):
                out.append(r)
                continue
            for i in range(n):
                nr = list(r)
                for c, v in split.items():
                    nr[c] = v[i]
                out.append(nr)
        return out

    def read_long(self, t, hrow, type_c, amt_c, item_c, title, carry=None):
        rows = self.expand_rows(self.backfill_amounts(t, amt_c, (hrow or 0) + 1 if hrow is not None else 0, type_c),
                                (type_c,), amt_c)
        heads = rows[hrow] if hrow is not None else []
        desc_cs = [c for c in range(len(rows[0])) if c not in (type_c, amt_c, item_c)]
        cur_types, cur_cond, cur_item = (carry or (None, '', ''))
        start = (hrow + 1) if hrow is not None else 0
        # 머리글 두 줄인 경우 (대수/설치장소)
        if hrow is not None and start < len(rows) and not any(price_of(x) for x in rows[start]):
            start += 1
        for r in rows[start:]:
            if type_c < len(r) and r[type_c]:
                cell = r[type_c].replace('\n', ' ')
                m = re.search(r'\(([^)]*)\)', cell)
                cond = m.group(1).strip() if m else ''
                ts = self.tm.many(re.sub(r'\([^)]*\)', ' ', cell))
                if ts:
                    cur_types, cur_cond = ts, cond
            if item_c is not None and item_c < len(r) and r[item_c]:
                cur_item = r[item_c].replace('\n', ' ').strip()
            if not cur_types:
                continue
            amt = price_of(r[amt_c]) if amt_c < len(r) else None
            if amt is None:
                continue
            desc = ' '.join((r[c] or '').replace('\n', ' ').strip() for c in desc_cs if c < len(r) and r[c]
                            and not re.search(r'비고|공급업체|제조사', heads[c] if c < len(heads) else ''))
            remark = ''
            for c in desc_cs:
                if c < len(heads) and '비고' in (heads[c] or '') and c < len(r):
                    remark = (r[c] or '').replace('\n', ' ').strip()
            gname = cur_item or title or '유상옵션'
            g = next((g for g in self.groups if g['name'] == gname and g['kind'] == 'long' and g.get('table') == (t.page, title)), None)
            if g is None:
                g = next((g for g in self.groups if g['name'] == gname and g['kind'] == 'long'), None)
            if g is None:
                g = {'name': gname, 'detail': '', 'variants': [], 'page': t.page + 1, 'kind': 'long', 'table': (t.page, title),
                     'title': title}
                self.groups.append(g)
            label = (f'[{cur_cond}] ' if cur_cond else '') + desc
            g['variants'].append({'label': label.strip(), 'types': list(cur_types), 'cond': cur_cond, 'remark': remark,
                                  'prices': {a: amt for a in cur_types}, 'units': units_count(desc)})
        return (cur_types, cur_cond, cur_item)



    def read_alltypes(self, t, hrow, amt_c, item_c, title):
        rows = self.expand_rows(self.backfill_amounts(t, amt_c, hrow + 1, item_c), (), amt_c)
        heads = rows[hrow]
        desc_cs = [c for c in range(len(heads)) if c not in (amt_c, item_c)
                   and not re.search(r'제조사|모델|비고|공급업체', heads[c] or '')]
        cur_item = ''
        alltypes = list(self.order)
        for r in rows[hrow + 1:]:
            if item_c is not None and item_c < len(r) and r[item_c]:
                cur_item = re.sub(r'\s+', ' ', r[item_c].replace('\n', ' ')).strip()
            amt = price_of(r[amt_c]) if amt_c < len(r) else None
            if amt is None or not cur_item:
                continue
            desc = ' '.join((r[c] or '').replace('\n', ' ').strip() for c in desc_cs if c < len(r) and r[c])
            g = next((g for g in self.groups if g['name'] == cur_item and g.get('table') == (t.page, title)), None)
            if g is None:
                g = {'name': cur_item, 'detail': '', 'variants': [], 'page': t.page + 1, 'kind': 'long',
                     'table': (t.page, title), 'title': title, 'all_types': True}
                self.groups.append(g)
            g['variants'].append({'label': desc, 'types': alltypes, 'cond': '', 'remark': '',
                                  'prices': {a: amt for a in alltypes}, 'units': units_count(desc)})

    # -------------------------------------------------------------- LH 식 선택품목 표
    #  주택형 | 유형 | 선택품목(-1, -2 …) | 기본형 | 합계(천원) | 계약금 …
    #  기본형 칸이 비어 있는 줄 = 바로 위 품목의 대안(택1), 채워진 줄 = 새 품목
    def lh_cols(self, rows):
        """칸 내용으로 (타입, 유형, 품목, 기본형, 금액) 칸을 찾는다 — 쪽마다 앞쪽 빈 칸 수가 달라도 맞게"""
        ncols = max(len(r) for r in rows)
        def colvals(c):
            return [(r[c] if c < len(r) else '') for r in rows]
        numc = [sum(1 for v in colvals(c) if num(v) is not None) for c in range(ncols)]
        itemc = [sum(1 for v in colvals(c) if v and re.search(r'[가-힣]', v) and num(v) is None
                     and (re.match(r'^\s*-\s*\d', v) or '[' in v or len(v) > 6)) for c in range(ncols)]
        item = max(range(ncols), key=lambda c: (itemc[c], -c))
        if itemc[item] < 2:
            return None
        amt = next((c for c in range(item + 1, ncols) if numc[c] >= max(2, len(rows) * 0.3)), None)
        if amt is None:
            return None
        def best(rng, score):
            sc = [(score(c), -c, c) for c in rng]
            sc = [x for x in sc if x[0] > 0]
            return max(sc)[2] if sc else None
        typec = best(range(0, item), lambda c: sum(1 for v in colvals(c) if v and CODE_RE.match(v.replace(' ', '').replace('\n', ''))))
        if typec is None:
            typec = best(range(0, item), lambda c: sum(1 for v in colvals(c) if v and len(v) < 12 and self.tm.many(v)))
        kind = best(range((typec if typec is not None else -1) + 1, item),
                    lambda c: sum(1 for v in colvals(c) if v and re.search(r'[가-힣]', v) and len(v) < 12))
        base = best(range(item + 1, amt), lambda c: sum(1 for v in colvals(c) if v and num(v) is None))
        return {'type': typec, 'kind': kind, 'item': item, 'base': base, 'amt': amt}

    def read_lh(self, t, cols, scale, state):
        rows = t.rows
        for r in rows:
            def cell(k):
                c = cols.get(k)
                return (r[c] if c is not None and c < len(r) else '').replace('\n', ' ').strip()
            tv = cell('type')
            if tv:
                ts = self.tm.many(tv)
                if ts:
                    state['types'] = ts
            if cell('kind'):
                state['kind'] = re.sub(r'\s+', '', cell('kind'))
            item = re.sub(r'^[\ue000-\uf8ff\s·•]+', '', cell('item'))      # 앞에 붙은 기호(아이콘 글꼴) 빼기
            if not item or not state.get('types'):
                continue
            raw = cell('amt')
            v = num(raw.split(' ')[0]) if raw else None
            if v is None and not re.search(r'무상', raw):
                continue
            amount = (v * scale) if isinstance(v, (int, float)) else None
            name = re.sub(r'^\s*-\s*\d+\s*', '', item).strip()
            new_group = bool(cell('base')) or not state.get('group') or not re.match(r'^\s*-\s*\d', item)
            key = (state.get('kind'), tuple(state['types']))
            if new_group or state.get('key_types') != key:
                gname = f"{state.get('kind', '')} {name}".strip()
                g = next((g for g in self.groups if g.get('lh_key') == (state.get('kind'), name) and g['kind'] == 'long'), None)
                if g is None:
                    g = {'name': gname, 'detail': cell('base'), 'variants': [], 'page': t.page + 1, 'kind': 'long',
                         'title': '추가선택품목', 'lh_key': (state.get('kind'), name), 'item_name': name}
                    self.groups.append(g)
                state['group'] = g
                state['key_types'] = key
            g = state['group']
            if amount is not None and 1000 <= amount <= 100_000_000:
                g['variants'].append({'label': name, 'types': list(state['types']), 'cond': '', 'remark': '',
                                      'prices': {a: amount for a in state['types']}, 'units': units_count(name)})

    # -------------------------------------------------------------- 전체
    def read(self):
        last_long = None          # 다음 쪽으로 이어지는 세로형 표
        lh = None                 # LH 식 선택품목 표 상태
        for p in self.option_pages():
            for t in self.p.pdf.tables(p):
                rows = t.rows
                if not rows or len(rows) < 2 and not last_long:
                    continue
                head = ' '.join(' '.join(r) for r in rows[:3]).replace(' ', '')
                acct = any(c and len(c) < 30 and re.search(r'계좌|예금주', c) for r in rows[:3] for c in r)
                if ('선택품목' in head and '기본형' in head) or (lh and lh['page'] >= p - 1 and not acct):
                    cols = self.lh_cols(rows)
                    if cols:
                        if not lh or ('선택품목' in head and '기본형' in head):
                            unit = 1000 if re.search(r'천\s*원', head + ' ' + (t.title or '')) else 1
                            lh = {'scale': unit, 'state': {}, 'page': p}
                        lh['page'] = p
                        self.read_lh(t, cols, lh['scale'], lh['state'])
                        continue
                flat = t.flat()
                if re.search(r'발코니', flat[:200]) or re.search(r'계좌번호|예금주|금융기관', flat[:120]):
                    last_long = None
                    continue
                head4 = t.flat(4)
                if re.search(r'대지비|건축비|층\s*구분|해당\s*세대|주거\s*전용|공용면적', head4):
                    last_long = None                      # 공급금액·공급대상 표
                    continue
                title = self.title_for(t)
                if re.search(r'발코니\s*확장', title):
                    last_long = None
                    continue
                # 가로형?
                done = False
                for i, r in enumerate(rows[:4]):
                    tcols = {c: self.tm.many(v) for c, v in enumerate(r) if v and c > 0}
                    tcols = {c: v for c, v in tcols.items() if v}
                    if len(tcols) >= 2 or (len(tcols) == 1 and len(self.types) == 1):
                        if any(price_of(x) for rr in rows[i + 1:] for x in rr):
                            self.read_wide(t, i, tcols, title)
                            done = True
                        break
                if done:
                    last_long = None
                    continue
                # 세로형?
                hrow = type_c = amt_c = item_c = None
                for i, r in enumerate(rows[:3]):
                    hs = [re.sub(r'\s', '', x or '') for x in r]
                    tc = next((c for c, h in enumerate(hs) if re.search(r'약식|주택형|타입|TYPE', h, re.I)), None)
                    ac = next((c for c, h in enumerate(hs) if re.search(r'공급금액|공급가|금액|판매가|공급대금', h)
                               and not re.search(r'계약금|중도금|잔금', h)), None)
                    if tc is not None and ac is not None:
                        hrow, type_c, amt_c = i, tc, ac
                        item_c = next((c for c, h in enumerate(hs[:tc]) if re.search(r'품목|구분|옵션', h)), None)
                        break
                if hrow is None:
                    # 타입 칸 없이 '구분 … 공급금액' 만 있는 표: 모든 타입 공통 (빌트인 가전 등)
                    for i, r in enumerate(rows[:2]):
                        hs = [re.sub(r'\s', '', x or '') for x in r]
                        ac = next((c for c, h in enumerate(hs) if re.search(r'공급금액|공급가|판매가|금액\(원\)', h)
                                   and not re.search(r'계약금|중도금|잔금|납부', h)), None)
                        ic = next((c for c, h in enumerate(hs) if re.search(r'^(구분|품목|품명|항목)', h)), None)
                        if ac is not None and ic is not None and ic < ac \
                                and not re.search(r'계좌|인지세|감리|발코니|납부|면적|세대', ' '.join(hs) + title):
                            self.read_alltypes(t, i, ac, ic, title)
                            hrow = -1
                            break
                    if hrow == -1:
                        last_long = None
                        continue
                if hrow is not None:
                    last_long = (t, type_c, amt_c, item_c, title, len(rows[hrow]))
                    carry = self.read_long(t, hrow, type_c, amt_c, item_c, title)
                    last_long = last_long + (carry,)
                    continue
                if last_long and last_long[0].page == t.page - 1 and max(len(r) for r in rows) == last_long[5]:
                    lt, tc, ac, ic, ttl, nc, carry = last_long
                    carry = self.read_long(t, None, tc, ac, ic, ttl, carry)
                    last_long = (t, tc, ac, ic, ttl, nc, carry)
                    continue
                last_long = None
        self.finish()
        return self.groups

    def finish(self):
        groups = []
        for i, g in enumerate(self.groups):
            if not g['variants']:
                continue
            if not any(v['prices'].get(a) for v in g['variants'] for a in self.order):
                continue                        # 어느 타입에도 금액이 없는 줄 (특별공급 표 등이 섞여 읽힌 것)
            if FLOOR_ROW_RE.search(g['name']):
                continue                        # 공급금액 표의 층·호라인 줄 ('6~10층 30 214,750', '101동2호라인')
            g['id'] = f'g{len(groups) + 1}'
            for j, v in enumerate(g['variants']):
                v['id'] = f'{g["id"]}v{j + 1}'
            g.pop('table', None)
            g.pop('lh_key', None)
            judge = g.get('item_name') or g['name']
            g['include'] = not EXCLUDE_RE.search(judge + ' ' + (g.get('title') or ''))
            g['short'] = short_name(g['name'])
            g['label'] = display_name(g['name'])
            groups.append(g)
        self.groups = groups
        included = [g['name'] for g in groups if g['include']]
        for g in groups:
            g['choice'] = {a: self.default_choice(g, a, included) for a in self.order}

    def default_choice(self, g, abbr, included):
        cands = [v for v in g['variants'] if v['prices'].get(abbr)]
        if not cands:
            return None
        if g['kind'] == 'long':
            def ok(v):
                rm = (v.get('remark') or '') + ' ' + (v.get('cond') or '')
                m = re.search(r'(\S+)\s*에\s*해당', rm)
                if m and abbr not in self.tm.many(m.group(1)):
                    return False
                m = re.search(r'(.+?)\s*선택\s*시\s*(미해당|선택\s*불가|불가)', rm)
                if m and any(_similar(m.group(1), n) for n in included):
                    return False
                if v.get('cond') and re.search(r'선택\s*시', v['cond']):
                    return False                  # 다른 옵션을 고를 때만 쓰는 금액은 기본에서 뺌
                return True
            okv = [v for v in cands if ok(v)]
            plain = [v for v in okv if not v.get('cond')]       # '(주방 냉방 강화)' 같은 조건 없는 기본 항목 먼저
            pool = plain or okv or cands
            if any(v.get('units') for v in pool):
                best = max(pool, key=lambda v: (v.get('units') or 0, -pool.index(v)))
            else:
                best = max(pool, key=lambda v: (v['prices'][abbr], -pool.index(v)))
            return best['id']
        # 가로형 대안
        if len(cands) == 1:
            return cands[0]['id']
        sel = [v for v in cands if re.search(r'(?<!미)선택\s*시', v['label'])]
        unsel = [v for v in cands if re.search(r'미\s*선택\s*시', v['label'])]
        if sel or unsel:
            cond = re.sub(r'\s*(미)?선택\s*시.*$', '', (sel or unsel)[0]['label']).strip()
            if any(_similar(cond, n) for n in included) and sel:
                return sel[0]['id']
            return (unsel or sel)[0]['id']
        return cands[0]['id']


def _brief(sub, base=''):
    s = re.sub(r'\([^)]*\)', '', sub or '').strip()
    return s[:16]


def _similar(a, b):
    na = re.sub(r'[\s()]', '', a).lower()
    nb = re.sub(r'[\s()]', '', b).lower()
    return bool(na) and bool(nb) and (na in nb or nb in na)


def read_options(parser, types):
    try:
        return OptionReader(parser, types).read()
    except Exception as e:
        parser.warn.append(f'유상옵션 표를 읽다가 오류가 났습니다: {e}')
        return []


def option_amounts(groups, abbr):
    """타입 하나의 묶음별 금액(원). 제외 묶음은 빠짐"""
    out = {}
    for g in groups:
        if not g.get('include'):
            continue
        vid = (g.get('choice') or {}).get(abbr)
        v = next((v for v in g['variants'] if v['id'] == vid), None)
        out[g['id']] = (v['prices'].get(abbr) if v else None) or 0
    return out


def variant_key(label):
    """대안 이름 → 칸 구분 낱말: '1안', '5대', '선택1', '선택시/미선택시'"""
    t = label or ''
    m = re.search(r'(\d+)\s*안\b|(\d+)\s*안\)', t)
    if m:
        return f"{m.group(1) or m.group(2)}안"
    m = re.search(r'선택\s*(\d+)', t)
    if m:
        return f'선택{m.group(1)}'
    if re.search(r'미\s*선택\s*시', t):
        return '미선택시'
    if re.search(r'선택\s*시', t):
        return '선택시'
    u = units_count(t)
    if u:
        return f'{u}대'
    return None


def column_name(g):
    n = re.sub(r'\s+', ' ', re.sub(r'\([^)]*\)', '', g.get('name') or '')).strip()
    return n if len(n) <= 16 else (g.get('label') or display_name(g.get('name') or ''))


def option_columns(groups, order):
    """모든 유상옵션을 엑셀 칸으로 펼친다. 타입마다 고를 것이 하나인 품목은 한 칸,
    여러 안(1안/2안, 2대/4대 …)이 있는 품목은 안마다 한 칸. 값은 원 단위 {타입: 금액}"""
    cols = []
    for g in groups:
        per = {a: [v for v in g['variants'] if v['prices'].get(a)] for a in order}
        if not any(len(vs) > 1 for vs in per.values()):
            cols.append({'id': g['id'], 'group': g['id'], 'header': column_name(g), 'name': g['name'],
                         'page': g.get('page'), 'values': {a: (vs[0]['prices'][a] if vs else None) for a, vs in per.items()}})
            continue
        def key_of(v):
            return variant_key(v['label']) or re.sub(r'\s+', ' ', v['label'])[:14]
        keys = []
        for v in g['variants']:
            k = key_of(v)
            if k not in keys:
                keys.append(k)
        if all(re.fullmatch(r'\d+(안|대)|선택\d+', k) for k in keys):
            keys.sort(key=lambda k: (k[-1], int(re.sub(r'\D', '', k))))
        for j, k in enumerate(keys):
            vals = {}
            for a in order:
                vs = [v for v in per[a] if key_of(v) == k]
                vals[a] = vs[0]['prices'][a] if vs else None
            cols.append({'id': f"{g['id']}k{j + 1}", 'group': g['id'], 'header': f"{column_name(g)} {k}",
                         'name': f"{g['name']} {k}", 'page': g.get('page'), 'values': vals})
    return cols
