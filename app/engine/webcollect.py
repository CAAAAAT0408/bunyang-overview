"""분양 홈페이지에서 그림 후보 모으기 + 그림 자르기.

- 홈페이지 첫 화면과 같은 사이트 안의 링크(메뉴)를 따라가며 그림을 모은다 (frame/iframe 포함).
- 일반 요청으로 그림이 거의 없으면(스크립트로 그리는 사이트) headless Edge 로 다시 읽는다.
- 그림마다 메뉴 이름·설명(alt)·주소 낱말로 위치도/조감도/배치도/커뮤니티/평면도 점수를 매긴다.
- 평면도는 '기본형'·'확장형(확장 기본형)' 표기가 있는 것만 그 칸에 넣는다 (설명·탭 글자, 없으면 그림 속 글자를 OCR).
  표기가 없는 평면도 한 장은 확장형 칸에만 넣고 기본형 칸은 X 표로 둔다 (2026-10-15 사용자 지정).
- 사진+표가 붙은 사업개요 그림은 흰 여백 기준으로 사진만 잘라 후보를 늘린다.
- 2026-10-18 사용자 지정 (imgcut 으로 그림 속 덩어리를 나눔):
  · 평면 그림에서 2D 평면·3D 아이소를 나누고 키맵(작은 단지배치도)·면적표는 잘라 낸다.
    기본형·확장형이 있으면 그 둘, 기본형이 없고 아이소가 있으면 확장형·아이소, 둘 다 없으면 기본형 칸 X.
  · 단지배치도는 위에서 본 배치도만, 커뮤니티는 도면만 자른다. 동호수배치도는 배치도로 쓰지 않는다.
  · 견본주택 위치도(오시는 길 약도)를 찾아 양식 O6 에 넣는다.
  · 한 홈페이지에 단지가 여럿(95·99)이면 우리 단지 번호가 주소에 있는 그림을 먼저 쓴다.
"""
import hashlib
import html as htmlmod
import io
import json
import os
import re
import shutil
import socket
import subprocess
import tempfile
import time
import urllib.parse
import urllib.request

from . import imgcut

UA = ('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) '
      'Chrome/130.0.0.0 Safari/537.36')
SKIP_IMG = re.compile(r'logo|icon|ico_|btn|button|arrow|sns|favicon|kakao|naver|blog|insta|youtube|facebook|'
                      r'loading|spinner|blank|spacer|qr|bi\.|/bi_|_bi|top_|quick|gnb|lnb|footer|header|'
                      r'keymap|key_map|키맵|_m\.|_mo\.|_m_\d|_mobile|/mobile/|/mo/', re.I)
STRONG = {'location': r'위치도|입지\s*환경|광역\s*위치|입지도', 'birdseye': r'조감도|투시도',
          'siteplan': r'단지\s*배치도', 'community': r'커뮤니티', 'units': r'평면도',
          'mh_map': r'견본\s*주택|모델\s*하우스|주택\s*전시관|오시는\s*길'}
DONGHO = re.compile(r'동\s*[·.]?\s*호수|동호수|동호\s*배치|dongho|(?<![a-z])dh\d*[._]', re.I)
MH_WORDS = re.compile(r'견본\s*주택|모델\s*하우스|주택\s*전시관|홍보관', re.I)
DECOR = re.compile(r'bg|leaf|deco|visual|banner|photo|thumb|_img_?\d', re.I)
PLAN_SKIP = re.compile(r'keymap|key_map|키맵|마감|texture|finish|interior|인테리어|동\s*호|dongho|조망|view_', re.I)
SKIP_LINK = re.compile(r'\.(pdf|zip|hwp|hwpx|docx?|xlsx?|pptx?|mp4|avi|jpg|jpeg|png|gif)(\?|$)|^(mailto|tel|javascript|#)', re.I)
SLOTS = {
    'location': (r'입지|위치|교통|입지환경|location|map|지도|생활권|premium|프리미엄', 3),
    'birdseye': (r'조감|투시|사업개요|단지개요|개요|overview|outline|bird|perspective|main_?visual|visual', 3),
    'siteplan': (r'단지\s*배치|배치도|단지배치|site\s*plan|siteplan|layout|배치|danzi|danji', 3),
    'community': (r'커뮤니티|community|club|클럽|주민공동|부대시설|amenity', 3),
    'units': (r'평면|평형|타입|type|unit|floor|세대안내|유니트|세대\s*정보|plan', 3),
    'mh_map': (r'오시는\s*길|찾아\s*오시는|견본\s*주택|모델\s*하우스|주택\s*전시관|홍보관|contact|directions|way_?to|'
               r'model_?house|(?<![a-z])mh(?![a-z])', 3),
}
IMAGE_SLOTS = ('location', 'birdseye', 'siteplan', 'community', 'mh_map')


def _host(u):
    try:
        return urllib.parse.urlparse(u).netloc.lower().split(':')[0]
    except Exception:
        return ''


def _same_site(a, b):
    ha, hb = _host(a), _host(b)
    if not ha or not hb:
        return False
    strip = lambda h: h[4:] if h.startswith('www.') else h
    if strip(ha) != strip(hb):
        return False
    # 홈페이지 주소에 경로가 있으면(예: /osxi) 그 경로 안만
    pb = urllib.parse.urlparse(b).path.rstrip('/')
    if pb and not re.search(r'\.(html?|php|aspx?|jsp)$', pb, re.I):
        pa = urllib.parse.urlparse(a).path
        return pa == pb or pa.startswith(pb + '/') or pa.startswith(pb + '?')
    return True


JS_EXPR = re.compile(r'^[A-Za-z_$][\w$]*(\.[A-Za-z_$][\w$]*)+$'
                     r'(?<!\.html)(?<!\.htm)(?<!\.php)(?<!\.jsp)(?<!\.asp)(?<!\.aspx)'
                     r'(?<!\.jpg)(?<!\.png)(?<!\.gif)(?<!\.jpeg)(?<!\.webp)')     # 'item.url' 같은 틀 글자
AD_HOST = re.compile(r'googletagmanager|doubleclick|google-analytics|facebook\.com|youtube|kakao|naver\.com|'
                     r'criteo|adservice|nexsmartmedia', re.I)


def iri(url):
    """한글 도메인(광명시티프라디움.com)·한글 경로 → 요청할 수 있는 ASCII 주소"""
    try:
        p = urllib.parse.urlsplit((url or '').strip())
        host = p.hostname or ''
        try:
            host = host.encode('idna').decode('ascii')
        except Exception:
            pass
        netloc = host + (f':{p.port}' if p.port else '')
        path = urllib.parse.quote(p.path, safe="/%:@!$&'()*+,;=-._~")
        query = urllib.parse.quote(p.query, safe="=&%:@!$'()*+,;/?-._~")
        return urllib.parse.urlunsplit((p.scheme, netloc, path, query, p.fragment))
    except Exception:
        return url


def urllib_get(url, referer=None, timeout=20):
    """주소 내용 받기 → (최종 주소, Content-Type, bytes). 웹 버전(브라우저)은 bridge 가 중계 서버로 받는 함수로 바꿔 끼운다"""
    req = urllib.request.Request(url, headers={'User-Agent': UA, 'Accept-Language': 'ko-KR,ko;q=0.9',
                                               **({'Referer': referer} if referer else {})})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.geturl(), r.headers.get('Content-Type', ''), r.read()


http_get = urllib_get


def fetch(url, referer=None, timeout=20, binary=False):
    url = iri(url)
    final, ctype, data = http_get(url, referer, timeout)
    if binary:
        return final, ctype, data
    enc = 'utf-8'
    m = re.search(r'charset=([\w-]+)', ctype, re.I) or re.search(rb'charset=["\']?([\w-]+)', data[:3000], re.I)
    if m:
        enc = m.group(1).decode() if isinstance(m.group(1), bytes) else m.group(1)
    try:
        text = data.decode(enc, 'replace')
    except LookupError:
        text = data.decode('utf-8', 'replace')
    return final, ctype, text


def _text_of(h):
    h = re.sub(r'(?is)<(script|style|noscript)[^>]*>.*?</\1>', ' ', h)
    h = re.sub(r'(?s)<[^>]+>', ' ', h)
    return re.sub(r'\s+', ' ', htmlmod.unescape(h)).strip()


def parse_page(url, h):
    """링크, 그림, 제목, 본문 글자"""
    links, imgs = [], []
    for m in re.finditer(r'(?is)<a\b[^>]*href=["\']([^"\']+)["\'][^>]*>(.*?)</a>', h):
        href, inner = m.group(1).strip(), m.group(2)
        if SKIP_LINK.search(href) or JS_EXPR.match(href):
            continue
        text = _text_of(inner)
        alt = ' '.join(re.findall(r'alt=["\']([^"\']*)["\']', inner))
        links.append((urllib.parse.urljoin(url, htmlmod.unescape(href)), (text or alt)[:40]))
    for m in re.finditer(r'(?is)<(?:i?frame)\b[^>]*src=["\']([^"\']+)["\']', h):
        if not AD_HOST.search(m.group(1)):
            links.append((urllib.parse.urljoin(url, m.group(1)), 'frame'))
    for m in re.finditer(r'(?is)(?<![:\w-])data-(?:href|url|link)\s*=\s*["\']([^"\']+)["\']|'
                         r'location\.href\s*=\s*["\']([^"\']+)["\']', h):
        v = (m.group(1) or m.group(2) or '').strip()
        if re.match(r'^(https?:)?/|^\./|^[\w\-/]+\.(html?|php|aspx?|jsp)', v) and not SKIP_LINK.search(v):
            links.append((urllib.parse.urljoin(url, htmlmod.unescape(v)), ''))
    for m in re.finditer(r'(?is)<img\b([^>]*)>', h):
        attrs = m.group(1)
        src = None
        for k in ('data-src', 'data-original', 'data-lazy', 'data-lazy-src', 'src'):
            mm = re.search(k + r'\s*=\s*["\']([^"\']+)["\']', attrs, re.I)
            if mm and not mm.group(1).startswith('data:'):
                src = mm.group(1)
                break
        if not src:
            mm = re.search(r'srcset\s*=\s*["\']([^"\']+)["\']', attrs, re.I)
            if mm:
                src = mm.group(1).split(',')[-1].strip().split(' ')[0]
        if not src or JS_EXPR.match(src):
            continue
        alt = (re.search(r'alt\s*=\s*["\']([^"\']*)["\']', attrs, re.I) or [None, ''])[1]
        cls = (re.search(r'class\s*=\s*["\']([^"\']*)["\']', attrs, re.I) or [None, ''])[1]
        if re.search(r'\bmobile\b|\bmo\b|_m\b', cls):
            continue                                           # 모바일용 중복 그림
        imgs.append({'src': urllib.parse.urljoin(url, htmlmod.unescape(src.replace('~/', '/'))), 'alt': alt})
    for m in re.finditer(r'background(?:-image)?\s*:\s*url\(\s*["\']?([^"\')]+)["\']?\s*\)', h, re.I):
        imgs.append({'src': urllib.parse.urljoin(url, m.group(1)), 'alt': ''})
    t = re.search(r'(?is)<title[^>]*>(.*?)</title>', h)
    return {'url': url, 'title': _text_of(t.group(1)) if t else '', 'links': links, 'imgs': imgs,
            'text': _text_of(h)[:20000]}


class Collector:
    def __init__(self, home, out_dir, types=None, log=None, max_pages=45, prefer=(), prefer_tokens=()):
        home = (home or '').strip()
        self.home = iri(home if re.match(r'^https?://', home) else ('https://' + home if home else ''))
        self.out = out_dir
        os.makedirs(self.out, exist_ok=True)
        self.types = types or []
        self.log = log or (lambda *a: None)
        self.max_pages = max_pages
        self.pages = []
        self.cands = []                 # 그림 후보
        self.seen_img = {}
        self.errors = []
        self.prefer = set(prefer)          # 우리 단지 번호 ('1') — '1단지'·'1BL' 그림 우선
        # 단지명 속 번호 ('아크메르 동탄 99' → '99') — 주소에 '/99/'·'complex99' 가 있는 그림 우선
        self.prefer_rx = [re.compile(rf'(?<!\d){re.escape(t)}(?!\d)') for t in prefer_tokens]

    # -------------------------------------------------------------- 둘러보기
    def crawl_static(self):
        start = self.home
        queue = [(start, '첫 화면')]
        done = set()
        while queue and len(self.pages) < self.max_pages:
            url, menu = queue.pop(0)
            key = url.split('#')[0]
            if key in done:
                continue
            done.add(key)
            try:
                final, ctype, h = fetch(url, referer=start)
            except Exception as e:
                self.errors.append(f'{url}: {e}')
                continue
            if 'html' not in ctype and not h.lstrip().startswith('<'):
                continue
            pg = parse_page(final, h)
            pg['menu'] = menu
            self.pages.append(pg)
            # 링크: 메뉴 낱말이 있는 것부터
            nxt = []
            for u, txt in pg['links']:
                if not _same_site(u, start) and txt != 'frame':
                    continue
                if u.split('#')[0] in done:
                    continue
                pri = 0 if any(re.search(p, txt + ' ' + u, re.I) for p, _ in SLOTS.values()) else 1
                nxt.append((pri, u, txt or menu))
            nxt.sort(key=lambda x: x[0])
            for pri, u, txt in nxt:
                if all(u != q[0] for q in queue):
                    queue.append((u, txt))
        self.log(f'홈페이지 {len(self.pages)}쪽 읽음')

    def crawl_edge(self):
        """스크립트로 그리는 사이트: headless Edge(CDP) 로 쪽마다 그림 주소를 모은다"""
        exe = None
        for p in (r'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe',
                  r'C:\Program Files\Microsoft\Edge\Application\msedge.exe'):
            if os.path.exists(p):
                exe = p
                break
        if not exe:
            self.log('Edge 를 찾지 못해 스크립트 사이트는 건너뜁니다')
            return
        import websocket
        s = socket.socket()
        s.bind(('127.0.0.1', 0))
        port = s.getsockname()[1]
        s.close()
        prof = tempfile.mkdtemp(prefix='bunyang_edge_')
        # headless 표식이 있는 기본 UA 는 막는 사이트가 있다 (광명시티프라디움 403) → 일반 Edge UA 로
        proc = subprocess.Popen([exe, '--headless=new', f'--remote-debugging-port={port}', f'--user-data-dir={prof}',
                                 '--no-first-run', '--disable-gpu', '--window-size=1400,1000',
                                 f'--user-agent={UA} Edg/130.0.0.0', 'about:blank'],
                                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        try:
            ws_url = None
            for _ in range(60):
                try:
                    with urllib.request.urlopen(f'http://127.0.0.1:{port}/json', timeout=2) as r:
                        tabs = json.loads(r.read())
                    pg = next((t for t in tabs if t.get('type') == 'page'), None)
                    if pg:
                        ws_url = pg['webSocketDebuggerUrl']
                        break
                except Exception:
                    pass
                time.sleep(0.25)
            if not ws_url:
                self.log('Edge 를 띄우지 못했습니다')
                return
            ws = websocket.create_connection(ws_url, timeout=30, suppress_origin=True)
            mid = [0]

            def call(method, **params):
                mid[0] += 1
                ws.send(json.dumps({'id': mid[0], 'method': method, 'params': params}))
                while True:
                    msg = json.loads(ws.recv())
                    if msg.get('id') == mid[0]:
                        return msg.get('result', {})
            call('Page.enable')
            js = r'''(() => {
              const out = {title: document.title, text: (document.body ? document.body.innerText : '').slice(0, 20000),
                           imgs: [], links: []};
              for (const i of document.images) out.imgs.push({src: i.currentSrc || i.src, alt: i.alt || '',
                                                              w: i.naturalWidth, h: i.naturalHeight});
              for (const e of document.querySelectorAll('*')) {
                const b = getComputedStyle(e).backgroundImage;
                if (b && b.startsWith('url(')) {
                  const r = e.getBoundingClientRect();
                  if (r.width > 300 && r.height > 200) out.imgs.push({src: b.slice(4, -1).replace(/["']/g, ''), alt: e.getAttribute('aria-label') || '', w: 0, h: 0});
                }
              }
              const txt = e => (e.innerText || e.textContent || e.title || '').trim().replace(/\s+/g, ' ').slice(0, 40);
              for (const a of document.querySelectorAll('a[href]')) out.links.push([a.href, txt(a)]);
              // 스크립트 메뉴: <a href="javascript:;" data-href="/pages/location"> · onclick="location.href='…'"
              for (const e of document.querySelectorAll('[data-href],[data-url],[data-link],[onclick]')) {
                let v = e.getAttribute('data-href') || e.getAttribute('data-url') || e.getAttribute('data-link') || '';
                if (!v) { const m = (e.getAttribute('onclick') || '').match(/location\.href\s*=\s*['"]([^'"]+)['"]/); if (m) v = m[1]; }
                if (!v || /^(javascript|#|mailto|tel)/i.test(v)) continue;
                try { out.links.push([new URL(v, location.href).href, txt(e)]); } catch (err) {}
              }
              for (const f of document.querySelectorAll('iframe[src],frame[src]')) out.links.push([f.src, 'frame']);
              return JSON.stringify(out);
            })()'''
            # 평면도 탭(주소가 바뀌지 않는 탭)을 눌러 가며 그림 모으기
            click_js = r'''((re, exact) => {
              const rx = new RegExp(re, 'i');
              const els = [...document.querySelectorAll('a,button,li,span,div,p,strong,label,em')].filter(e => {
                const t = (e.innerText || '').trim();
                if (!t || t.length > 16 || !rx.test(t) || e.children.length > 3) return false;
                if (exact && t.replace(/\s/g, '') !== exact) return false;
                const r = e.getBoundingClientRect(); return r.width > 0 && r.height > 0;
              });
              if (!els.length) return 'none';
              const el = els[els.length - 1];
              const a = el.closest('a'); const href = a ? (a.getAttribute('href') || '') : '';
              if (href && !/^(#|javascript)/i.test(href)) return 'link';
              el.click(); return 'ok';
            })'''
            vis_js = r'''JSON.stringify([...document.images].filter(i => i.offsetParent !== null && i.naturalWidth >= 500)
                         .map(i => ({src: i.currentSrc || i.src, alt: i.alt || '', w: i.naturalWidth, h: i.naturalHeight})))'''

            def wait_ready(limit=8.0):
                t0 = time.time()
                while time.time() - t0 < limit:
                    r = call('Runtime.evaluate', expression='document.readyState', returnByValue=True)
                    if r.get('result', {}).get('value') == 'complete':
                        break
                    time.sleep(0.3)
                time.sleep(1.2)

            def click_tabs(pg):
                """타입 탭(59A·84A…)과 그 안의 기본형/확장형 탭을 눌러 보이는 그림에 이름표를 붙인다"""
                extra = []
                for t in self.types:
                    r = call('Runtime.evaluate', expression=f'{click_js}({json.dumps(type_pattern(t["abbr"]))}, "")',
                             returnByValue=True)
                    if r.get('result', {}).get('value') != 'ok':
                        continue
                    time.sleep(1.2)
                    subs = []
                    for sub in ('기본형', '확장형', '확장기본형', '기본확장형'):
                        rr = call('Runtime.evaluate', expression=f'{click_js}({json.dumps(sub)}, {json.dumps(sub)})',
                                  returnByValue=True)
                        if rr.get('result', {}).get('value') == 'ok':
                            time.sleep(1.0)
                            subs.append(sub)
                            v = call('Runtime.evaluate', expression=vis_js, returnByValue=True)
                            for im in json.loads(v.get('result', {}).get('value') or '[]'):
                                im['alt'] = f"{t['abbr']} {sub} {im['alt']}".strip()
                                extra.append(im)
                    if not subs:
                        v = call('Runtime.evaluate', expression=vis_js, returnByValue=True)
                        for im in json.loads(v.get('result', {}).get('value') or '[]'):
                            im['alt'] = f"{t['abbr']} {im['alt']}".strip()
                            extra.append(im)
                # 탭마다 그대로 보이는 그림(배경·로고)은 이름표 없이
                cnt = {}
                for im in extra:
                    cnt[im['src']] = cnt.get(im['src'], 0) + 1
                pg['imgs'] = [im for im in extra if cnt[im['src']] == 1 and im['src'].startswith('http')] + pg['imgs']
            queue = [(self.home, '첫 화면', 0)]
            done = set()
            visited = 0
            clicked = 0
            while queue and visited < min(self.max_pages, 30):
                url, menu, pri = queue.pop(0)
                if url.split('#')[0] in done:
                    continue
                if pri and visited >= 10:              # 메뉴 낱말이 없는 쪽은 앞의 몇 쪽만
                    continue
                done.add(url.split('#')[0])
                visited += 1
                call('Page.navigate', url=url)
                wait_ready()
                for y in range(0, 5):
                    call('Runtime.evaluate', expression=f'window.scrollTo(0, {y * 900})')
                    time.sleep(0.25)
                r = call('Runtime.evaluate', expression=js, returnByValue=True)
                try:
                    data = json.loads(r['result']['value'])
                except Exception:
                    continue
                pg = {'url': url, 'title': data.get('title', ''), 'menu': menu, 'text': data.get('text', ''),
                      'links': [tuple(x) for x in data.get('links', [])],
                      'imgs': [i for i in data.get('imgs', []) if i.get('src', '').startswith('http')]}
                # 평면 쪽이면 타입 탭을 눌러 본다 (탭이 주소 링크면 그 링크를 따라가므로 건너뜀)
                if clicked < 2 and self.types and re.search(SLOTS['units'][0], f"{menu} {url}", re.I) \
                        and sum(1 for t in self.types if re.search(type_pattern(t['abbr']), pg['text'])) >= 2 \
                        and not any(re.search(type_pattern(t['abbr']), txt or '') and _same_site(u, self.home)
                                    for u, txt in pg['links'] for t in self.types[:1]):
                    clicked += 1
                    try:
                        click_tabs(pg)
                    except Exception as e:
                        self.errors.append(f'탭 누르기 실패: {e}')
                self.pages.append(pg)
                for u, txt in pg['links']:
                    if (_same_site(u, self.home) or txt == 'frame') and not SKIP_LINK.search(u) and not AD_HOST.search(u):
                        if u.split('#')[0] not in done and all(u != q[0] for q in queue):
                            p0 = 0 if any(re.search(p, txt + ' ' + u, re.I) for p, _ in SLOTS.values()) else 1
                            (queue.insert(0, (u, txt, 0)) if p0 == 0 else queue.append((u, txt, 1)))
            ws.close()
            self.log(f'Edge 로 {visited}쪽 읽음')
        finally:
            try:
                proc.kill()
            except Exception:
                pass
            shutil.rmtree(prof, ignore_errors=True)

    # -------------------------------------------------------------- 그림 받기
    def image_list(self, limit=240):
        """읽은 쪽들의 그림 주소 (아이콘·팝업·이미 받은 것 빼고) → [{src, page, menu, title, alt}]"""
        out = []
        for pg in self.pages:
            for im in pg['imgs']:
                src = im['src']
                if SKIP_IMG.search(src) or src in self.seen_img or re.search(r'popup|/pop_|pop\d', src, re.I):
                    continue
                self.seen_img[src] = None
                out.append({'src': src, 'page': pg['url'], 'menu': pg.get('menu', ''), 'title': pg.get('title', ''),
                            'alt': im.get('alt', '')})
                if len(out) >= limit:
                    return out
        return out

    def add_image(self, data, it):
        """받은 그림 하나 → 후보 (작은 그림·같은 그림 빼고). 넣었으면 후보, 아니면 None"""
        from PIL import Image
        if not data or len(data) < 15_000:
            return None
        try:
            img = Image.open(io.BytesIO(data))
            w, h = img.size
            if w < 500 or h < 300 or h > 7000 or w * h > 16_000_000:
                return None                     # 작은 그림, 마감재 목록 같은 아주 긴 그림 (풀면 메모리를 많이 씀)
            img.load()
        except Exception:
            return None
        cid = 'w' + hashlib.md5(data).hexdigest()[:10]
        if any(c['id'] == cid for c in self.cands):
            return None
        path = self.save(img, cid)
        self.seen_img[it['src']] = cid
        c = {'id': cid, 'file': path, 'w': w, 'h': h, 'src': it['src'], 'page': it['page'], 'menu': it.get('menu', ''),
             'title': it.get('title', ''), 'alt': it.get('alt', ''), 'kind': 'web'}
        self.cands.append(c)
        return c

    def download(self, max_imgs=160):
        n = 0
        for it in self.image_list():
            if n >= max_imgs:
                break
            try:
                _, ctype, data = fetch(it['src'], referer=it['page'], binary=True, timeout=25)
            except Exception as e:
                self.errors.append(f"{it['src']}: {e}")
                continue
            if self.add_image(data, it):
                n += 1
        self.log(f'그림 {len(self.cands)}장 받음')

    def save(self, img, cid):
        from PIL import Image
        if img.mode not in ('RGB', 'L'):
            bg = Image.new('RGB', img.size, (255, 255, 255))
            try:
                bg.paste(img, mask=img.convert('RGBA').split()[-1])
            except Exception:
                bg.paste(img.convert('RGB'))
            img = bg
        elif img.mode == 'L':
            img = img.convert('RGB')
        if img.width > 2000:
            img = img.resize((2000, round(img.height * 2000 / img.width)))
        path = os.path.join(self.out, cid + '.jpg')
        img.save(path, 'JPEG', quality=90)
        thumb = img.copy()
        thumb.thumbnail((360, 360))
        thumb.save(os.path.join(self.out, cid + '_t.jpg'), 'JPEG', quality=80)
        return path

    # -------------------------------------------------------------- 분류
    def preferred(self, c):
        """주소(폴더·파일 이름)나 쪽 주소에 우리 단지 번호가 있는 그림"""
        where = urllib.parse.unquote(urllib.parse.urlparse(c.get('src') or '').path) + ' ' + \
            urllib.parse.urlparse(c.get('page') or '').path + ' ' + (c.get('menu') or '') + ' ' + (c.get('alt') or '')
        return any(rx.search(where) for rx in self.prefer_rx)

    def page_text(self, url):
        return next((p.get('text') or '' for p in self.pages if p.get('url') == url), '')

    def score(self, c, slot):
        pat, w = SLOTS[slot]
        s = 0
        for field, fw in (('menu', 4), ('alt', 3), ('title', 1), ('src', 2), ('page', 2)):
            if re.search(pat, c.get(field) or '', re.I):
                s += fw
        base = os.path.basename(urllib.parse.urlparse(c.get('src') or '').path)
        if slot == 'mh_map':
            if s == 0 and re.search(r'(?<![a-z])map', base, re.I) and not re.search(r'location|입지|위치', base, re.I) \
                    and MH_WORDS.search(self.page_text(c.get('page'))):
                s = 3                               # 첫 화면의 'map.jpg' + 쪽 글에 '주택전시관' (LH)
            if s and re.search(r'map|약도|지도', base + ' ' + (c.get('alt') or ''), re.I):
                s += 2
            if s and DECOR.search(base):
                s -= 4                              # 배경·장식 그림 (contact_bg, contact_leaf)
            if s and re.search(r'location|입지', base, re.I):
                s -= 4
        if s <= 0:
            return 0
        if slot == 'siteplan' and (DONGHO.search(f"{c.get('alt') or ''} {base}") or
                                   (DONGHO.search(c.get('menu') or '') and '단지' not in (c.get('menu') or ''))):
            s = max(1, s - 6)                       # 동·호수 배치도는 다른 배치도가 없을 때만 (2026-10-18 사용자 지정)
        if STRONG.get(slot) and re.search(STRONG[slot], c.get('menu') or ''):
            s += 1
        if self.preferred(c):
            s += 3
        if slot == 'birdseye' and c.get('menu') == '첫 화면' and c['w'] >= 1200:
            s += 1
        if re.search(STRONG[slot], c.get('alt') or ''):
            s += 2                                  # 그림 설명에 '단지배치도'·'조감도' 같은 이름이 그대로 있으면
        label = (c.get('alt') or '') + ' ' + (c.get('menu') or '')
        found = re.findall(r'(\d)\s*단지|(\d)\s*BL', label, re.I)
        if found and self.prefer:
            nums = {a or b for a, b in found}
            s += 3 if nums & self.prefer else -1    # 다른 블록 표기는 약하게만 (설명 오타가 있는 사이트도 있음)
        return s

    def unit_hits(self, abbr):
        """타입 이름이 설명·파일 이름·폴더 이름(/59A/unit.jpg)에 있는 평면도 후보 (키맵·마감재 제외)"""
        pat = type_pattern(abbr)
        def where(c):
            parts = urllib.parse.urlparse(c.get('src') or '').path.split('/')
            return (c.get('alt') or '') + ' ' + ' '.join(urllib.parse.unquote(x) for x in parts[-2:])
        pool = [c for c in self.cands if c.get('kind') == 'web' and not PLAN_SKIP.search(where(c) + ' ' + (c.get('menu') or ''))]
        hits = [c for c in pool if re.search(pat, where(c) + ' ' + (c.get('ocr_text') or ''), re.I)]
        if not hits:
            hits = [c for c in pool if re.search(pat, c.get('menu') or '', re.I) and self.score(c, 'units') > 0]
        return sorted(hits, key=lambda c: (not self.preferred(c), -c['w'] * c['h']))

    def ocr_type_names(self, limit=16):
        """타입 이름이 그림 이름·설명에 없는 평면 쪽 그림(unit01.jpg …)은 그림 속 글자('74㎡')에서 타입 이름을 찾는다"""
        if not self.types or all(self.unit_hits(t['abbr']) for t in self.types):
            return
        base = lambda c: os.path.basename(urllib.parse.urlparse(c.get('src') or '').path)
        desktop = {base(c) for c in self.cands if c.get('kind') == 'web' and '/m/' not in (c.get('src') or '')}
        pats = [type_pattern(t['abbr']) for t in self.types]
        pool = [c for c in self.cands if c.get('kind') == 'web' and 'ocr_text' not in c and self.score(c, 'units') >= 4
                and not ('/m/' in (c.get('src') or '') and base(c) in desktop)            # 모바일용 중복 그림
                and not any(re.search(p, f"{c.get('alt', '')} {base(c)}", re.I) for p in pats)]
        pool = sorted(pool, key=lambda c: -c['w'] * c['h'])[:limit]
        if not pool:
            return
        res = ocr_lines([c['file'] for c in pool], self.log)
        for c in pool:
            c['ocr_text'] = ' '.join(ln[2] for ln in res.get(c['file']) or [])

    def classify(self):
        auto = {}
        for slot in IMAGE_SLOTS:
            ranked = sorted(self.cands, key=lambda c: (self.score(c, slot), c['w'] * c['h']), reverse=True)
            ranked = [c for c in ranked if self.score(c, slot) > 0]
            auto[slot] = [ranked[0]['id']] if ranked else []
        auto['units'] = {t['abbr']: {'hits': [c['id'] for c in self.unit_hits(t['abbr'])]}
                         for t in self.types if self.unit_hits(t['abbr'])}
        return auto

    def missing(self):
        """정적 읽기로 빠진 것이 많은지 (→ Edge 로 다시)"""
        auto = self.classify()
        empty = [k for k in ('location', 'birdseye', 'siteplan', 'community') if not auto.get(k)]
        units = len(auto.get('units') or {})
        return len(empty) >= 2 or (self.types and units < (len(self.types) + 1) // 2)

    def assign_plans(self, auto):
        """타입별 평면 → {'basic': id, 'ext': id, 'iso': id} (2026-10-18 사용자 지정)
        그림 속 덩어리를 나눠 2D 평면·3D 아이소만 자르고 키맵(작은 단지배치도)·면적표·표기 단추는 뺀다.
        2D 평면의 기본형(비확장)/확장형은 설명·탭·파일 이름, 없으면 그림 속 글자(OCR) 표기로 가리고,
        표기 없는 평면은 확장형으로 본다. 어느 칸에 넣을지는 엑셀 작성 때 unit_pair() 가 정한다."""
        from PIL import Image
        byid = {c['id']: c for c in self.cands}
        per = {abbr: [byid[i] for i in u.get('hits', []) if i in byid][:4] for abbr, u in (auto.get('units') or {}).items()}
        parts = {}
        for c in {c['id']: c for hs in per.values() for c in hs}.values():
            try:
                with Image.open(c['file']) as im:
                    parts[c['id']] = imgcut.plan_parts(im.convert('RGB'))
            except Exception as e:
                self.errors.append(f"평면 나누기 실패 {c['id']}: {e}")
        def dominance(c):
            p = parts.get(c['id']) or {}
            return max([b['area'] for b in (p.get('plans') or []) + (p.get('isos') or [])] or [0])
        for abbr in per:                                # 도면이 그림의 큰 부분인 것부터 (옵션 위치도·표 그림은 뒤로)
            per[abbr].sort(key=lambda c: (not self.preferred(c), -dominance(c)))
        def label_of(c):
            return text_label(f"{c.get('alt', '')} {c.get('menu', '')} {os.path.basename(c.get('src', ''))}")
        need = sorted({c['file'] for hs in per.values() for c in hs
                       if label_of(c) is None and (parts.get(c['id']) or {}).get('plans')})
        ocr = ocr_lines(need, self.log) if need else {}
        ocr_ok = bool(ocr) or not need                  # 웹 버전(브라우저)은 Windows OCR 이 없다
        names = {'basic': '기본형', 'ext': '확장형', 'iso': '아이소'}
        out = {}
        nsheet = 0
        for abbr, hits in per.items():
            got = {}
            unl = None
            for c in hits:
                p = parts.get(c['id'])
                if not p:
                    continue
                if p['sheet']:
                    nsheet += 1                         # 유상옵션 안내처럼 작은 도면이 여럿인 그림
                    continue
                lab = label_of(c)
                if lab == 'option':
                    continue
                plans = p['plans']
                if plans:
                    if lab in ('basic', 'ext'):
                        got.setdefault(lab, (c, plans[0]['box']))
                    else:
                        done = set()
                        for k, i in label_plans(plans, ocr.get(c['file'])):
                            if k not in got:
                                got[k] = (c, plans[i]['box'])
                            done.add(i)
                        if not ocr_ok and not done and len(plans) == 2 and plans[1]['area'] >= 0.7 * plans[0]['area']:
                            # OCR 없이 표기 없는 비슷한 2D 평면 두 장 → 읽는 순서대로 기본형·확장형 (자이 등 관례)
                            a, b = sorted(plans, key=lambda p: (p['box'][1] // 60, p['box'][0]))
                            got.setdefault('basic', (c, a['box']))
                            got.setdefault('ext', (c, b['box']))
                            done = {0, 1}
                        rest = [b for i, b in enumerate(plans) if i not in done]
                        if rest and unl is None:
                            unl = (c, rest[0]['box'])
                if p['isos'] and 'iso' not in got:
                    got['iso'] = (c, p['isos'][0]['box'])
            if 'ext' not in got and unl:
                got['ext'] = unl                        # 표기 없는 평면 한 장은 확장형 (기본형 칸은 X)
            res = {}
            for k, (c, box) in got.items():
                cc = crop_save(c, imgcut.pad(box, (c['w'], c['h']), 10), self.out, self.cands, 'plan', names[k])
                res[k] = cc['id']
            if res:
                out[abbr] = res
        if nsheet:
            self.log(f'평면 후보 중 옵션 안내 그림 {nsheet}장은 뺐습니다')
        return out

    def crop_slots(self, auto):
        """배치도는 위에서 본 배치도만, 커뮤니티는 도면만, 견본주택 위치도는 지도만 자른다 (2026-10-18 사용자 지정).
        잘라도 원래 그림은 후보에 남겨 화면에서 바꿀 수 있다"""
        from PIL import Image
        byid = {c['id']: c for c in self.cands}
        for slot, fn, label in (('siteplan', imgcut.siteplan_box, '배치도'), ('community', imgcut.community_boxes, '도면'),
                                ('mh_map', imgcut.main_box, '약도')):
            if not auto.get(slot) or auto[slot][0] not in byid:
                continue
            src = byid[auto[slot][0]]
            try:
                with Image.open(src['file']) as im:
                    found = fn(im.convert('RGB'))
            except Exception as e:
                self.errors.append(f'{slot} 자르기 실패: {e}')
                continue
            boxes = found if isinstance(found, list) else ([found] if found else [])
            if boxes:
                auto[slot] = [crop_save(src, imgcut.pad(b, (src['w'], src['h']), 8), self.out, self.cands, 'crop', label)['id']
                              for b in boxes]
        return auto

    def all_text(self):
        return '\n'.join(p.get('text', '') for p in self.pages)

    def run(self):
        if not self.home:
            return {'cands': [], 'auto': {}, 'pages': [], 'errors': ['홈페이지 주소가 없습니다']}
        try:
            self.crawl_static()
        except Exception as e:
            self.errors.append(f'홈페이지 읽기 실패: {e}')
        try:
            self.download()
        except Exception as e:
            self.errors.append(f'그림 받기 실패: {e}')
        if len(self.cands) < 4 or self.missing():
            self.log('스크립트로 그리는 쪽이 있어 Edge 로 다시 읽습니다 (1~3분)')
            try:
                self.pages_static = len(self.pages)
                self.crawl_edge()
                self.download()
            except Exception as e:
                self.errors.append(f'Edge 읽기 실패: {e}')
                self.log(f'Edge 읽기 실패: {e}')
        return self.finish()

    def finish(self):
        """모은 그림 → 자동 배정 (분류·평면 나누기·자르기)"""
        try:
            self.ocr_type_names()
        except Exception as e:
            self.errors.append(f'평면 타입 글자 읽기 실패: {e}')
        auto = self.classify()
        # 평면도: 2D 기본형·확장형 + 3D 아이소 (키맵 빼고)
        try:
            auto['units'] = self.assign_plans(auto)
        except Exception as e:
            self.errors.append(f'평면도 나누기 실패: {e}')
            self.log(f'평면도 나누기 실패: {e}')
            auto['units'] = {}
        if auto['units']:
            n = {k: sum(1 for u in auto['units'].values() if u.get(k)) for k in ('basic', 'ext', 'iso')}
            self.log(f"평면도: 기본형 {n['basic']}타입, 확장형 {n['ext']}타입, 아이소 {n['iso']}타입")
        # 조감도가 사업개요처럼 사진+표인 그림이면 사진 부분만
        if auto.get('birdseye'):
            src = next(c for c in self.cands if c['id'] == auto['birdseye'][0])
            photo = split_photo(src, self.out, self.cands)
            if photo:
                auto['birdseye'] = [photo['id']]
        try:
            auto = self.crop_slots(auto)
        except Exception as e:
            self.errors.append(f'그림 자르기 실패: {e}')
        return {'cands': self.cands, 'auto': auto,
                'pages': [{'url': p['url'], 'menu': p.get('menu', ''), 'title': p.get('title', '')} for p in self.pages],
                'errors': self.errors[:20], 'text': self.all_text()[:200000]}


def prefer_nums(draft):
    """우리 단지 번호 ('1단지'·'1BL' 의 1) — 그림 설명에 이 번호가 있으면 우선"""
    nums = set(re.findall(r'(\d)\s*단지', draft.get('title', '')))
    m = re.search(r'\(([A-Za-z]*-?\d+)\s*BL\)', draft.get('title', '')) or \
        re.search(r'([A-Za-z]*-?\d+)\s*(?:BL|블록|블럭)', draft.get('location', ''))
    b = (m.group(1) + 'BL') if m else ''
    if b and re.fullmatch(r'\d+BL', b) and len(b) == 3:
        nums.add(b[0])
    return nums


def prefer_tokens(draft):
    """단지명 속 번호 ('아크메르 동탄 99' → ['99']) — 한 홈페이지에 단지가 여럿일 때 우리 단지 그림 고르기"""
    name = draft.get('complex_name') or draft.get('title') or ''
    return sorted(set(re.findall(r'(?<![\d.])(\d{2,3})(?![\d])', name)))


def text_label(text):
    """설명·탭·파일 이름의 평면 종류: 'basic' | 'ext' | 'option'(쓰지 않음) | None(표기 없음)"""
    t = re.sub(r'\s', '', text or '')
    if re.search(r'옵션|선택형|option|\+', t, re.I):
        return 'option'
    if re.search(r'비확장|basic', t, re.I):
        return 'basic'
    if re.search(r'확장|expand|(?<![a-z])ext(?![a-z])|_ex[._\d]', t, re.I):
        return 'ext'
    if re.search(r'기본형', t):
        return 'basic'
    return None


def ocr_label_pos(lines):
    """OCR 표제 줄의 평면 종류와 y 위치 → [('basic'|'ext', y), …] (같은 종류는 처음 것만)"""
    out = []
    for ln in lines or []:
        if not isinstance(ln, (list, tuple)) or len(ln) < 3:
            continue
        labs = ocr_labels([ln])
        if len(labs) == 1:
            k = labs.pop()
            if all(k != o[0] for o in out):
                out.append((k, ln[1]))
    return out


def label_plans(plans, lines):
    """OCR 표기 줄 → 2D 평면 덩어리 짝 [(kind, 평면 번호)]. 표기 수와 평면 수가 같으면 읽는 순서대로 짝짓고,
    아니면 표기 바로 아래(제목) 또는 바로 위(단추)의 가장 가까운 평면"""
    labs = []
    for ln in lines or []:
        if isinstance(ln, (list, tuple)) and len(ln) >= 3:
            k = ocr_labels([ln])
            if len(k) == 1:
                labs.append((k.pop(), ln[0], ln[1]))
    if not labs:
        return []
    out, used = [], set()
    for k, x, y in sorted(labs, key=lambda t: (t[2], t[1])):
        best = None
        for i, b in enumerate(plans):
            x0, y0, x1, y1 = b['box']
            pw, ph = x1 - x0, y1 - y0
            if not (x0 - 0.25 * pw <= x <= x1 - 0.05 * pw):
                continue                                # 평면 폭 안의 표기만 (옆 표의 '기본형' 열 머리글 등 제외)
            if y <= y0:
                d = y0 - y if y0 - y <= 0.6 * ph else None          # 위의 제목
            elif y >= y1:
                d = (y - y1) * 1.2 if y - y1 <= 0.4 * ph else None  # 아래의 단추
            else:
                d = (y - y0) if y - y0 <= 0.15 * ph else None       # 도면 안 위쪽 표기
            if d is not None and (best is None or d < best[0]):
                best = (d, i)
        if best and best[1] not in used and all(k != o[0] for o in out):
            out.append((k, best[1]))
            used.add(best[1])
    return out


def crop_below(src, y, out_dir, cands, label=''):
    """표제 글자(y) 아래의 첫 도면 덩어리를 잘라 후보로. 그림 전체와 거의 같으면 None"""
    from PIL import Image
    with Image.open(src['file']) as img:
        img = img.convert('RGB')
        segs = _segments(img, min_gap=8, min_h=80)
        pick = None
        for y0, y1 in segs:
            if y1 <= y + 12:
                continue
            if (y1 - y0) < img.height * 0.12:
                if y0 >= y - 4:
                    continue                    # 표제 줄 자체 등 작은 덩어리
            pick = (y0, y1)
            break
        if not pick:
            return None
        y0, y1 = pick
        if y0 < y < y1 and (y - y0) > 30:
            y0 = int(y) + 25                    # 표제가 덩어리 안에 있으면 그 아래부터
        x0, x1 = _trim_x(img, y0, y1)
        if (y1 - y0) * (x1 - x0) > img.width * img.height * 0.85:
            return None
    return crop_save(src, (x0, max(0, y0 - 4), x1, y1 + 4), out_dir, cands, 'plan', label)


def ocr_labels(lines):
    """OCR 줄들에서 평면 표기 모으기 → {'basic', 'ext'} 의 부분집합 (옵션 표기 줄은 무시)"""
    labs = set()
    for ln in lines or []:
        t = re.sub(r'\s', '', ln[2] if isinstance(ln, (list, tuple)) else str(ln))
        if not re.search(r'기본|확장', t):
            continue
        if len(t) > 16 or re.search(r'옵션|선택', t):
            continue                                # 긴 줄은 품목표 글자 ('기본형 배기휀 …') — 표제만 본다
        if '비확장' in t:
            labs.add('basic')
        elif re.search(r'확장형|발코니확장|확장기본', t):
            labs.add('ext')
        elif '기본형' in t:
            labs.add('basic')
    return labs


OCR_PS1 = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ocr.ps1')


def ocr_lines(files, log=None):
    """Windows 내장 OCR(한국어)로 그림 글자 읽기 → {파일: [(x, y, 글자), …]}. 안 되면 빈 dict"""
    if not files or not os.path.exists(OCR_PS1):
        return {}
    tmp = tempfile.mkdtemp(prefix='bunyang_ocr_')
    try:
        lst = os.path.join(tmp, 'list.txt')
        out = os.path.join(tmp, 'out.jsonl')
        with open(lst, 'w', encoding='utf-8') as f:
            f.write('\n'.join(os.path.abspath(x) for x in files))
        t = time.time()
        subprocess.run(['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', OCR_PS1, '-List', lst, '-Out', out],
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=60 + 5 * len(files),
                       creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
        res = {}
        if os.path.exists(out):
            for line in open(out, encoding='utf-8'):
                try:
                    o = json.loads(line)
                except Exception:
                    continue
                if o.get('file'):
                    res[os.path.abspath(o['file'])] = [tuple(x) for x in (o.get('lines') or []) if len(x) == 3]
        if log:
            log(f'평면도 글자 읽기(OCR) {len(files)}장 ({time.time() - t:.0f}초)')
        return {f: res.get(os.path.abspath(f), []) for f in files}
    except Exception as e:
        if log:
            log(f'OCR 실패: {e}')
        return {}
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def type_pattern(abbr):
    m = re.match(r'^(\d{2,3})([A-Za-z]*\d?)$', abbr)
    if not m:
        return re.escape(abbr)
    n, suf = m.group(1), m.group(2)
    if suf:
        return rf'(?<!\d){n}\s*(?:㎡|m²|m2|형|타입|TYPE|\.\d+)?\s*-?\s*{suf}(?![A-Za-z0-9])'
    return rf'(?<!\d){n}\s*(?:㎡|m²|m2|형|타입|TYPE|\.\d+)?(?![\dA-Za-z])'


# ------------------------------------------------------------------ 자르기

def _segments(img, white=243, min_gap=10, min_h=120):
    """흰(밝은) 가로줄로 나뉜 덩어리들의 (y0, y1)"""
    g = img.convert('L')
    w, h = g.size
    step = max(1, w // 300)
    px = g.load()
    rows = []
    for y in range(h):
        bright = 0
        tot = 0
        for x in range(0, w, step):
            tot += 1
            if px[x, y] >= white:
                bright += 1
        rows.append(bright / tot >= 0.985)
    segs, y = [], 0
    while y < h:
        if rows[y]:
            y += 1
            continue
        y0 = y
        gap = 0
        while y < h:
            if rows[y]:
                gap += 1
                if gap >= min_gap:
                    break
            else:
                gap = 0
            y += 1
        y1 = y - gap
        if y1 - y0 >= min_h:
            segs.append((y0, y1))
    return segs


def _trim_x(img, y0, y1, white=243):
    g = img.convert('L').crop((0, y0, img.width, y1))
    w, h = g.size
    px = g.load()
    step = max(1, h // 200)
    def col_white(x):
        return all(px[x, y] >= white for y in range(0, h, step))
    x0 = 0
    while x0 < w - 1 and col_white(x0):
        x0 += 1
    x1 = w - 1
    while x1 > x0 and col_white(x1):
        x1 -= 1
    return max(0, x0 - 6), min(w, x1 + 7)


def _stats(img):
    from PIL import ImageStat
    st = ImageStat.Stat(img.convert('RGB'))
    lum = sum(st.mean) / 3
    sat = max(st.stddev)
    hsv = ImageStat.Stat(img.convert('HSV'))
    return lum, sat, hsv.mean[1]


def crop_save(src, box, out_dir, cands, kind='crop', label=''):
    from PIL import Image
    with Image.open(src['file']) as img:
        im = img.crop(box)
        cid = 'c' + hashlib.md5(f"{src['id']}{box}".encode()).hexdigest()[:10]
        if any(c['id'] == cid for c in cands):
            return next(c for c in cands if c['id'] == cid)
        path = os.path.join(out_dir, cid + '.jpg')
        im.convert('RGB').save(path, 'JPEG', quality=92)
        t = im.convert('RGB')
        t.thumbnail((360, 360))
        t.save(os.path.join(out_dir, cid + '_t.jpg'), 'JPEG', quality=80)
        c = {'id': cid, 'file': path, 'w': im.width, 'h': im.height, 'src': src.get('src', ''), 'page': src.get('page', ''),
             'menu': src.get('menu', ''), 'title': src.get('title', ''), 'alt': (src.get('alt', '') + ' ' + label).strip(),
             'kind': kind, 'parent': src['id'], 'box': list(box)}
        cands.append(c)
        return c


def split_plan(src, out_dir, cands):
    """평면도 그림에서 도면 덩어리(밝은 배경 위 큰 덩어리)를 위→아래 순서로 잘라 낸다"""
    from PIL import Image
    with Image.open(src['file']) as img:
        img = img.convert('RGB')
        segs = _segments(img)
        parts = []
        for y0, y1 in segs:
            sub = img.crop((0, y0, img.width, y1))
            lum, sd, sat = _stats(sub)
            if lum < 110:                          # 어두운 머리 띠
                continue
            if (y1 - y0) < img.height * 0.12:
                continue
            x0, x1 = _trim_x(img, y0, y1)
            if (x1 - x0) < img.width * 0.3:
                continue
            parts.append((y0, y1, x0, x1, (y1 - y0) * (x1 - x0)))
    narrow = [p for p in parts if (p[3] - p[2]) < src['w'] * 0.96]
    if narrow:
        parts = narrow                              # 도면은 좌우에 흰 여백이 있다
    if len(parts) > 2:
        parts = sorted(sorted(parts, key=lambda p: -p[4])[:2], key=lambda p: p[0])
    out = []
    for i, (y0, y1, x0, x1, _) in enumerate(parts):
        lab = ('기본형' if i == 0 else '확장형') if len(parts) >= 2 else ''
        out.append(crop_save(src, (x0, max(0, y0 - 4), x1, y1 + 4), out_dir, cands, 'plan', lab))
    return out


def split_photo(src, out_dir, cands):
    """사진 아래에 표가 붙은 그림이면 사진 덩어리만"""
    from PIL import Image
    with Image.open(src['file']) as img:
        img = img.convert('RGB')
        segs = _segments(img, min_gap=6, min_h=150)
        if len(segs) < 2:
            return None
        best = None
        for y0, y1 in segs:
            sub = img.crop((0, y0, img.width, y1))
            lum, sd, sat = _stats(sub)
            score = sat * (y1 - y0)
            if best is None or score > best[0]:
                best = (score, y0, y1)
    if not best:
        return None
    _, y0, y1 = best
    if (y1 - y0) > src['h'] * 0.92:
        return None
    return crop_save(src, (0, y0, src['w'], y1), out_dir, cands, 'photo', '사진')
