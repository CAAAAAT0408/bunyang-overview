"""웹 버전 홈페이지 그림 모으기 — 공고 분석과 다른 작업 공간(collect.js)에서 돈다.
(공고 PDF 를 읽은 Python 은 메모리가 커서 그 뒤에 그림 부품(Pillow)이 멈추는 일이 있어 따로 둔다)

collect.js 가 부른다:
  set_logger(fn)
  collect_pages(home, proxy, draft_json) → JSON {ok, items:[{i, src, ref}]}  홈페이지 쪽 읽기 (중계 서버로)
  collect_add(i, data)                   → 1/0  받은 그림 하나를 바로 후보로 (쌓아 두지 않음)
  collect_finish()                       → JSON {ok, cands, auto}  분류·평면 나누기·자르기 (서버 버전과 같은 규칙)
  collect_image(id)                      → bytes
브라우저는 다른 사이트를 직접 못 읽어서 Cloudflare 중계(proxy/worker.js)를 거친다.
"""
import json
import shutil
import traceback
import urllib.parse

from engine import webcollect

WEB_DIR = '/home/pyodide/web'
SLOTS = ('location', 'birdseye', 'siteplan', 'community', 'mh_map')
_S = {'log': None}


def set_logger(fn):
    _S['log'] = fn


def log(msg):
    fn = _S.get('log')
    if fn is not None:
        try:
            fn(str(msg))
        except Exception:
            pass


def _ok(**kw):
    return json.dumps({'ok': True, **kw}, ensure_ascii=False, default=str)


def _fail(msg):
    return json.dumps({'ok': False, 'error': str(msg)}, ensure_ascii=False)


def _proxy_get(proxy):
    """webcollect.http_get 대신: 중계 서버로 동기 요청 (웹 워커 안이라 동기 XHR 가능)"""
    from js import Uint8Array, XMLHttpRequest

    def get(url, referer=None, timeout=20):
        q = proxy + '?url=' + urllib.parse.quote(url, safe='')
        if referer:
            q += '&ref=' + urllib.parse.quote(referer, safe='')
        x = XMLHttpRequest.new()
        x.open('GET', q, False)
        x.responseType = 'arraybuffer'
        x.send()
        if x.status != 200:
            raise IOError(f'HTTP {x.status}')
        return (x.getResponseHeader('X-Final-Url') or url, x.getResponseHeader('Content-Type') or '',
                Uint8Array.new(x.response).to_bytes())
    return get


def collect_pages(home, proxy, draft_json):
    try:
        d = json.loads(draft_json)
        webcollect.http_get = _proxy_get(proxy)
        shutil.rmtree(WEB_DIR, ignore_errors=True)
        col = webcollect.Collector(home, WEB_DIR + '/img', d['types'], log=log, max_pages=30,
                                   prefer=webcollect.prefer_nums(d), prefer_tokens=webcollect.prefer_tokens(d))
        if not col.home:
            return _fail('홈페이지 주소를 넣어 주세요.')
        log(f'분양 홈페이지 읽는 중: {home}')
        col.crawl_static()
        if not col.pages:
            return _fail('홈페이지를 읽지 못했습니다. 주소를 확인해 주세요.')
        items = col.image_list()
        _S['col'], _S['items'] = col, items
        log(f'그림 {len(items)}장 받는 중')
        return _ok(items=[{'i': i, 'src': webcollect.iri(it['src']), 'ref': it['page']} for i, it in enumerate(items)])
    except Exception as e:
        log(traceback.format_exc().splitlines()[-1])
        return _fail(f'홈페이지 읽기 실패: {e}')


def collect_add(i, data, max_imgs=160):
    col = _S.get('col')
    if col is None or len(col.cands) >= max_imgs:
        return 0
    try:
        raw = bytes(data.to_py()) if hasattr(data, 'to_py') else bytes(data)
        return 1 if col.add_image(raw, _S['items'][int(i)]) else 0
    except Exception as e:
        col.errors.append(f'그림 {i}: {e}')
        return 0


def collect_finish():
    try:
        col = _S['col']
        log(f'그림 {len(col.cands)}장 받음')
        r = col.finish()
        auto = r['auto']
        imgs = {k: auto.get(k) or [] for k in SLOTS}
        imgs['units'] = {a: {k: v for k, v in u.items() if k in ('basic', 'ext', 'iso')}
                         for a, u in (auto.get('units') or {}).items()}
        _S['cands'] = {c['id']: c for c in col.cands}
        log(f"그림 후보 {len(col.cands)}장 (자동 배정: 위치도 {len(imgs['location'])}, 조감도 {len(imgs['birdseye'])}, "
            f"배치도 {len(imgs['siteplan'])}, 커뮤니티 {len(imgs['community'])}, 견본주택 위치도 {len(imgs['mh_map'])}, "
            f"평면도 {len(imgs['units'])}타입)")
        if len(col.cands) < 4:
            log('그림을 거의 못 받았습니다. 스크립트로 그리는 홈페이지는 웹 버전에서 못 읽을 수 있습니다 — 그림을 저장해 직접 넣어 주세요.')
        keep = ('id', 'w', 'h', 'menu', 'alt', 'kind', 'parent')
        return _ok(cands=[{k: c.get(k) for k in keep} for c in col.cands], auto=imgs)
    except Exception as e:
        log(traceback.format_exc().splitlines()[-1])
        return _fail(f'그림 정리 실패: {e}')


def collect_image(cid):
    c = (_S.get('cands') or {}).get(cid)
    return open(c['file'], 'rb').read() if c else b''
