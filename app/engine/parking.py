"""주차대수 찾기: 분양 홈페이지 글 → Daum 검색 요약문 → DuckDuckGo 요약문.

후보마다 출처를 남기고, 세대수×세대당 대수와 맞고 '1단지/1BL' 같은 문맥이 맞는 총대수를 고른다.
총대수가 없고 세대당 대수만 있으면 추정값(세대수×세대당)을 낸다 — 화면에서 확인.
"""
import html as htmlmod
import math
import re
import time
import urllib.parse
import urllib.request

UA = ('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) '
      'Chrome/130.0.0.0 Safari/537.36')

TOTAL_RATIO = re.compile(r'([\d,]{3,6})\s*대\s*\(?\s*(?:세대|가구)\s*당\s*(?:약\s*)?([0-3]\.\d{1,2})\s*대')
TOTAL_NEAR = [re.compile(r'주\s*차[^\d\n]{0,25}?(?:총\s*)?([\d,]{3,6})\s*대'),
              re.compile(r'(?:총\s*)([\d,]{3,6})\s*대[^.\n]{0,12}주\s*차')]
RATIO = re.compile(r'(?:세대|가구)\s*당\s*(?:약\s*)?([0-3]\.\d{1,2})\s*대')


def _get(url, timeout=15):
    req = urllib.request.Request(url, headers={'User-Agent': UA, 'Accept-Language': 'ko-KR,ko;q=0.9'})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read().decode('utf-8', 'replace')


def _plain(h):
    h = re.sub(r'(?is)<(script|style)[^>]*>.*?</\1>', ' ', h)
    h = re.sub(r'(?i)</?(b|strong|em|mark|span)[^>]*>', '', h)
    h = re.sub(r'(?s)<[^>]+>', '\n', h)
    h = htmlmod.unescape(h)
    h = re.sub(r'[ \t ]+', ' ', h)
    return re.sub(r'\n\s*\n+', '\n', h)


def daum(query):
    h = _get('https://search.daum.net/search?w=tot&q=' + urllib.parse.quote(query))
    out = []
    for line in _plain(h).split('\n'):
        line = line.strip()
        if '주차' in line and len(line) > 15 and '광고' not in line:
            out.append(line)
    return out


def ddg(query):
    h = _get('https://html.duckduckgo.com/html/?q=' + urllib.parse.quote(query))
    out = []
    for m in re.finditer(r'class="result__a" href="([^"]+)"[^>]*>(.*?)</a>.*?class="result__snippet"[^>]*>(.*?)</a>', h, re.S):
        sn = _plain(m.group(3)).replace('\n', ' ').strip()
        if '주차' in sn:
            u = m.group(1)
            mm = re.search(r'uddg=([^&]+)', u)
            out.append((urllib.parse.unquote(mm.group(1)) if mm else u, sn))
    return out


def extract(text):
    """글 한 덩어리 → [{'total','ratio','pre'(숫자 앞 글), 'snippet'}]"""
    res = []
    used = set()
    for m in TOTAL_RATIO.finditer(text):
        tot = int(m.group(1).replace(',', ''))
        used.add(m.start(1))
        res.append({'total': tot, 'ratio': float(m.group(2)), 'pre': text[max(0, m.start() - 30):m.start()],
                    'snippet': text[max(0, m.start() - 50):m.end() + 20]})
    for rx in TOTAL_NEAR:
        for m in rx.finditer(text):
            if m.start(1) in used:
                continue
            used.add(m.start(1))
            tot = int(m.group(1).replace(',', ''))
            res.append({'total': tot, 'ratio': None, 'pre': text[max(0, m.start(1) - 30):m.start(1)],
                        'snippet': text[max(0, m.start() - 40):m.end() + 30]})
    for m in RATIO.finditer(text):
        if '주차' in text[max(0, m.start() - 40):m.end() + 40]:
            res.append({'total': None, 'ratio': float(m.group(1)), 'pre': text[max(0, m.start() - 30):m.start()],
                        'snippet': text[max(0, m.start() - 50):m.end() + 20]})
    return res


def block_tokens(name, block):
    toks = set()
    for m in re.finditer(r'(\d)\s*단지', name or ''):
        toks.add(f'{m.group(1)}단지')
    m = re.search(r'([A-Za-z]*-?\d+)\s*(?:BL|블록|블럭)', (block or '') + ' ' + (name or ''), re.I)
    if m:
        toks.add(m.group(1).upper() + 'BL')
        n = re.sub(r'\D', '', m.group(1))
        if n and len(n) == 1:
            toks.add(f'{n}단지')
    return toks


def search_parking(name, households, homepage_text='', block='', log=None, notice_text=''):
    log = log or (lambda *a: None)
    cands = []
    for line in (notice_text or '').split('\n'):            # 공고 본문 (유의사항에 '아파트 주차대수는 총 583대' 등)
        if '주차' in line:
            for c in extract(line):
                cands.append(dict(c, source='모집공고', snippet=line[:200], weight=4))
    for c in extract(homepage_text or ''):
        cands.append(dict(c, source='분양 홈페이지', weight=3))
    base = re.sub(r'\s+', ' ', re.sub(r'\(([^)]*)\)', r' \1', name or '')).strip()
    queries = [f'{base} 주차대수', f'{base} 세대당 주차']
    if block and block not in base:
        queries.insert(1, f'{base} {block} 주차대수')
    for q in queries:
        try:
            for line in daum(q):
                for c in extract(line):
                    cands.append(dict(c, source=f'Daum 검색 "{q}"', snippet=line[:220], weight=1))
        except Exception as e:
            log(f'Daum 검색 실패: {e}')
        time.sleep(0.5)
    if not any(c['total'] for c in cands):
        try:
            for url, sn in ddg(f'{base} 주차대수'):
                for c in extract(sn):
                    cands.append(dict(c, source=url, snippet=sn[:220], weight=1))
        except Exception as e:
            log(f'DuckDuckGo 검색 실패: {e}')
    return choose(cands, households, block_tokens(name, block))


def choose(cands, households, mine=frozenset()):
    hh = households or 0
    ratios = [c['ratio'] for c in cands if c['ratio']]
    ratio_mode = max(set(ratios), key=ratios.count) if ratios else None
    scored = {}
    for c in cands:
        t = c['total']
        if not t or (hh and not (0.5 * hh <= t <= 3.5 * hh)):
            continue
        s = scored.setdefault(t, {'total': t, 'score': 0.0, 'sources': [], 'snips': set()})
        key = c['snippet'][:80]
        if key in s['snips']:
            continue
        s['snips'].add(key)
        s['score'] += c['weight']
        if hh and c['ratio'] and abs(t / hh - c['ratio']) < 0.02:
            s['score'] += 2
        pre = c.get('pre', '')
        tail = pre[-18:]
        if re.search(r'(아파트|공동주택)[^,\d]{0,10}$', pre):
            s['score'] += 4                          # '아파트 지하 1,604대'
        if re.search(r'근린|오피스텔|상가|판매시설', tail) or re.search(r'총\s*주차\s*대수[^\d]{0,6}$', tail) and '아파트' in c['snippet']:
            s['score'] -= 4                          # '…근린생활시설 지상 5대, 총 1,609대' / '총 주차대수 688대'
        others = re.findall(r'\d\s*단지|[A-Za-z]*-?\d+\s*BL', pre, re.I)
        if others:
            last = re.sub(r'\s', '', others[-1]).upper()
            if mine:
                s['score'] += 3 if last in {m.upper() for m in mine} else -3
        if len(s['sources']) < 4:
            s['sources'].append({'source': c['source'], 'snippet': c['snippet']})
    for s in scored.values():
        if hh and ratio_mode:
            d = abs(s['total'] / hh - ratio_mode)
            s['score'] += 5 if d < 0.02 else (-4 if d > 0.1 else 0)
        s.pop('snips', None)
    ranked = sorted(scored.values(), key=lambda s: -s['score'])
    best = ranked[0] if ranked and ranked[0]['score'] > 0 else None
    out = {'best': best['total'] if best else None, 'estimate': False, 'ratio': ratio_mode,
           'candidates': ranked[:6],
           'ratio_sources': [{'source': c['source'], 'snippet': c['snippet']} for c in cands if c['ratio'] and not c['total']][:3]}
    if not best and ratio_mode and hh:
        out['best'] = math.ceil(ratio_mode * hh)          # 올림: 세대당 대수가 원래 값보다 작아지지 않게
        out['estimate'] = True
    return out
