"""브라우저(Pyodide) 화면과 엔진을 잇는다. 공고·양식·그림은 이 브라우저 밖으로 나가지 않는다.

worker.js 가 부른다:
  set_logger(fn)                              진행 기록을 화면으로 보낼 함수
  analyze(notice, nname, template, tname)     → JSON {ok, draft} (공고 분석 + 양식 확인)
  pdf_images()                                → JSON [{id, page, w, h, ext}] 공고 PDF 속 큰 그림 목록
  pdf_image(id)                               → bytes
  page_count() / render_page(pno, zoom)       → 공고 쪽 그림(PNG) — 그림을 잘라 쓸 때
  build(draft_json, images_json)              → JSON {ok, name, sec}; 결과는 result() 로
"""
import json
import os
import re
import time
import traceback

import build_web
import xlsxio
from engine.draft import finalize, make_draft
from engine.notice import NoticeError, NoticeParser
from engine.options import read_options

IMG_DIR = '/home/pyodide/img'
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


def _bytes(x):
    if isinstance(x, (bytes, bytearray)):
        return bytes(x)
    if hasattr(x, 'to_py'):
        x = x.to_py()
    return bytes(x)


def _ok(**kw):
    return json.dumps({'ok': True, **kw}, ensure_ascii=False, default=str)


def _fail(msg):
    return json.dumps({'ok': False, 'error': str(msg)}, ensure_ascii=False)


class UserError(Exception):
    pass


def check_template(data):
    """양식에 분양개요 시트·공급개요 표가 있는지 가볍게 확인"""
    try:
        book = xlsxio.Book(data)
    except Exception as e:
        raise UserError(f'양식 엑셀을 열 수 없습니다: {e}')
    names = [s['name'] for s in book.sheet_list()]
    name = next((n for n in names if re.sub(r'\s', '', n) == '분양개요'), names[0])
    ws = book.sheet(name)
    found = {'공급개요': False, '소계': False, '단위세대': False}
    for ref, c in ws.cells.items():
        r, col = xlsxio.cell_key(ref)
        if r > 300 or col > 6:
            continue
        v = ws.value(ref)
        t = re.sub(r'\s', '', str(v)) if v is not None else ''
        if t.startswith('3.공급개요'):
            found['공급개요'] = True
        if t == '소계':
            found['소계'] = True
        if re.match(r'^\d\.단위세대', t):
            found['단위세대'] = True
    if not found['공급개요'] or not found['소계']:
        raise UserError('양식에서 "3. 공 급 개 요" 표(구분·소계)를 찾지 못했습니다. 분양개요 및 청약결과 양식(.xlsx)을 넣어 주세요.')
    return {'sheets': names, 'has_result': '청약결과' in names, **found}


def analyze(notice, nname, template, tname):
    try:
        t0 = time.time()
        ndata, tdata = _bytes(notice), _bytes(template)
        if not tname.lower().endswith(('.xlsx', '.xlsm')):
            raise UserError('양식은 엑셀(.xlsx) 파일이어야 합니다.')
        tinfo = check_template(tdata)
        tinfo['name'] = tname
        log('모집공고 읽는 중')
        try:
            np_ = NoticeParser(ndata, nname)
            parsed = np_.run()
        except NoticeError as e:
            raise UserError(str(e))
        if not parsed['types']:
            raise UserError('공고에서 공급대상 표(주택형·세대수)를 찾지 못했습니다.')
        groups = read_options(np_, parsed['types'])
        draft = make_draft(parsed, groups, nname)
        draft['complex_name'] = parsed['info'].get('name') or ''
        draft.setdefault('warnings', parsed.get('warnings') or [])
        log(f"공고 분석 끝 ({time.time() - t0:.0f}초): 타입 {len(draft['types'])}개, 총 {draft['households']:,}세대, "
            f"유상옵션 품목 {len(groups)}개")
        for w in parsed.get('warnings') or []:
            log('⚠ ' + w)
        draft['homepage'] = ''
        draft['template'] = tinfo
        draft['images'] = {'location': [], 'birdseye': [], 'siteplan': [], 'community': [], 'units': {}}
        draft = finalize(draft)
        _S.update(notice=ndata, nname=nname, template=tdata, tname=tname, doc=np_.pdf.doc, pdfimgs={})
        log('분석 끝 — 아래에서 값과 그림을 확인한 뒤 "엑셀 만들기"를 누르세요.')
        return _ok(draft=draft)
    except UserError as e:
        log('✖ ' + str(e))
        return _fail(e)
    except Exception as e:
        log('✖ 분석 중 오류: ' + str(e))
        log(traceback.format_exc().splitlines()[-1])
        return _fail(f'분석 중 오류: {e}')


# ---------------------------------------------------------------- 공고 PDF 속 그림 · 쪽
def pdf_images(min_w=360, min_h=240, limit=60):
    """공고 PDF 에 들어 있는 큰 그림(배치도·평면도 등) 목록"""
    import pymupdf
    doc = _S.get('doc')
    if doc is None:
        return json.dumps([])
    out, seen = [], set()
    for pno in range(doc.page_count):
        try:
            infos = doc[pno].get_images(full=True)
        except Exception:
            continue
        for info in infos:
            xref = info[0]
            if xref in seen:
                continue
            seen.add(xref)
            w, h = info[2], info[3]
            if w < min_w or h < min_h:
                continue
            try:
                pix = pymupdf.Pixmap(doc, xref)
                if pix.alpha:
                    pix = pymupdf.Pixmap(pix, 0)
                if pix.n - pix.alpha >= 4:
                    pix = pymupdf.Pixmap(pymupdf.csRGB, pix)
                try:
                    data, ext = pix.tobytes('jpeg', jpg_quality=88), 'jpeg'
                except Exception:
                    data, ext = pix.tobytes('png'), 'png'
            except Exception:
                continue
            iid = f'p{pno + 1}x{xref}'
            _S['pdfimgs'][iid] = data
            out.append({'id': iid, 'page': pno + 1, 'w': pix.width, 'h': pix.height, 'ext': ext})
            if len(out) >= limit:
                return json.dumps(out)
    return json.dumps(out)


def pdf_image(iid):
    return _S.get('pdfimgs', {}).get(iid, b'')


def page_count():
    doc = _S.get('doc')
    return doc.page_count if doc is not None else 0


def render_page(pno, zoom=2.0):
    import pymupdf
    doc = _S['doc']
    pno = max(0, min(int(pno), doc.page_count - 1))
    return doc[pno].get_pixmap(matrix=pymupdf.Matrix(zoom, zoom)).tobytes('png')


# ---------------------------------------------------------------- 엑셀 만들기
def _img(iid):
    if not iid:
        return None
    p = os.path.join(IMG_DIR, iid)
    if not os.path.exists(p):
        return None
    data = open(p, 'rb').read()
    try:
        w, h, _ = build_web.img_size(data)
    except Exception:
        return None
    return {'data': data, 'w': w, 'h': h}


def safe_filename(name):
    name = re.sub(r'[\\/:*?"<>|]', '_', name or '').strip() or '분양개요 및 청약결과.xlsx'
    if not name.lower().endswith('.xlsx'):
        name += '.xlsx'
    return name


def build(draft_json):
    try:
        if not _S.get('template'):
            raise UserError('먼저 공고와 양식을 분석해 주세요.')
        d = finalize(json.loads(draft_json))
        im = d.get('images') or {}
        images = {k: [x for x in (_img(i) for i in im.get(k) or []) if x] for k in ('location', 'birdseye', 'siteplan', 'community')}
        images['units'] = {}
        for a, u in (im.get('units') or {}).items():
            images['units'][a] = {k: _img(v) for k, v in u.items() if k in ('basic', 'ext') and _img(v)}
        log('엑셀 작성 중')
        data, sec = build_web.build_workbook(_S['template'], d, images, log=log)
        _S['result'] = data
        name = safe_filename(d.get('output_name') or '분양개요 및 청약결과.xlsx')
        log(f'엑셀 완성 ({sec:.1f}초): {name}')
        return _ok(name=name, sec=round(sec, 1), draft=d)
    except (UserError, build_web.BuildError) as e:
        log('✖ ' + str(e))
        return _fail(e)
    except Exception as e:
        log('✖ 엑셀 작성 중 오류: ' + str(e))
        log(traceback.format_exc().splitlines()[-1])
        return _fail(f'엑셀 작성 중 오류: {e}')


def result():
    return _S.get('result') or b''
