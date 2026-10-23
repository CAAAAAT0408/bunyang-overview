"""분양개요 양식(.xlsx) 에 값을 채워 새 파일을 만든다 — 브라우저(Pyodide)용, Excel 없이 xlsx 내부 XML 을 고친다.

app/engine/build_xlsx.py (Excel COM 판) 와 같은 규칙으로 채운다.
  분양개요 시트: 1. 일반개요 / 2. 사업일정 / 3. 공급개요(타입당 2줄 + 소계 + ※ 3줄) /
                4~7 그림 구역(테두리 사각형 안) / 8. 단위세대(타입당 6줄, 기본형 D:J · 확장형 K:N)
  청약결과 시트: 타입당 1줄 + 소계. 수식만 넣고 입력값은 비운다.
  공고자료 시트(숨김, 새로 만듦): 공급표·층별 공급금액·평균평당가 계산 근거·확장비·유상옵션.
계산식 값은 파일을 Excel 로 열 때 다시 계산된다(전체 재계산 표시).
"""
import copy
import math
import posixpath
import re
import time
from datetime import date

from lxml import etree

import xlsxio as X
from xlsxio import q, NS, RNS, PNS, CNS, XDR, col2num, num2col, cell_key
from engine.draft import REPORT_WIDTHS, chosen_option_cols, report_cells, report_layout, unit_headers, unit_pair

A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R_IMG = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
R_DRAW = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing'
DRAW_CT = 'application/vnd.openxmlformats-officedocument.drawing+xml'
WS_CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'
EMU_PT = 12700
EMU_PX = 9525

# CT_Worksheet 자식 순서 (새 요소를 알맞은 자리에 넣을 때)
WS_ORDER = ['sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData', 'sheetCalcPr', 'sheetProtection',
            'protectedRanges', 'scenarios', 'autoFilter', 'sortState', 'dataConsolidate', 'customSheetViews', 'mergeCells',
            'phoneticPr', 'conditionalFormatting', 'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup',
            'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags',
            'drawing', 'legacyDrawing', 'legacyDrawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems',
            'tableParts', 'extLst']


class BuildError(Exception):
    pass


def ns(v):
    """비교용: 공백 제거한 문자열"""
    return re.sub(r'\s', '', str(v)) if v is not None else ''


def cref(r, c):
    return f'{num2col(c)}{r}'


def excel_date(iso):
    if not iso:
        return None
    try:
        y, m, d = map(int, str(iso)[:10].split('-'))
        return (date(y, m, d) - date(1899, 12, 30)).days
    except Exception:
        return None


def localname(el):
    return etree.QName(el).localname if isinstance(el.tag, str) else ''


def insert_ordered(root, el):
    """worksheet 자식 el 을 정해진 순서 자리에 넣는다."""
    idx = WS_ORDER.index(localname(el))
    after = None
    for child in root:
        ln = localname(child)
        if ln in WS_ORDER and WS_ORDER.index(ln) < idx:
            after = child
    if after is None:
        root.insert(0, el)
    else:
        after.addnext(el)


def img_size(data):
    """PNG/JPEG 머리에서 (폭, 높이, 확장자)"""
    if data[:8] == b'\x89PNG\r\n\x1a\n':
        return int.from_bytes(data[16:20], 'big'), int.from_bytes(data[20:24], 'big'), 'png'
    if data[:2] == b'\xff\xd8':
        i = 2
        while i < len(data) - 9:
            if data[i] != 0xFF:
                i += 1
                continue
            m = data[i + 1]
            if m in (0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF):
                return int.from_bytes(data[i + 7:i + 9], 'big'), int.from_bytes(data[i + 5:i + 7], 'big'), 'jpeg'
            if m in (0xD8, 0x01) or 0xD0 <= m <= 0xD7:
                i += 2
                continue
            i += 2 + int.from_bytes(data[i + 2:i + 4], 'big')
        return 0, 0, 'jpeg'
    raise BuildError('그림은 PNG 또는 JPG 여야 합니다.')


def _next_rid(rels_root):
    ids = [int(re.sub(r'\D', '', r.get('Id') or '') or 0) for r in rels_root]
    return f'rId{max(ids + [0]) + 1}'


# ====================================================================== 칸 쓰기
def put(sh, r, c, v):
    """값 쓰기: '=…' 은 수식, 숫자는 숫자, '' / None 은 내용 지우기(서식은 둠)"""
    ref = cref(r, c)
    if v is None or (isinstance(v, str) and v == ''):
        sh.clear(ref)
        return
    if isinstance(v, bool):
        v = int(v)
    if isinstance(v, (int, float)):
        if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
            sh.clear(ref)
            return
        sh.set_number(ref, v)
        return
    s = str(v)
    if s.startswith('=') and len(s) > 1:
        set_formula(sh, ref, s[1:])
        return
    sh.set_text(ref, s)


def set_formula(sh, ref, f):
    c = sh._cell(ref)
    sh._clear(c)
    e = etree.SubElement(c, q('f'))
    e.text = f
    sh.book.touched.add(sh.path)


def clear_range(sh, r1, c1, r2, c2):
    for ref, c in list(sh.cells.items()):
        r, cc = cell_key(ref)
        if r1 <= r <= r2 and c1 <= cc <= c2:
            sh._clear(c)
    sh.book.touched.add(sh.path)


# ====================================================================== 병합
def merge_area(sh, r, c):
    for (a, b) in sh.merged():
        if a[0] <= r <= b[0] and a[1] <= c <= b[1]:
            return a[0], a[1], b[0], b[1]
    return r, c, r, c


def unmerge_range(sh, r1, c1, r2, c2):
    mc = sh.root.find(q('mergeCells'))
    if mc is None:
        return
    for m in list(mc):
        a, _, b = m.get('ref').partition(':')
        (ra, ca), (rb, cb) = cell_key(a), cell_key(b or a)
        if ra <= r2 and rb >= r1 and ca <= c2 and cb >= c1:
            mc.remove(m)
    mc.set('count', str(len(mc)))
    sh.book.touched.add(sh.path)


def merge_range(sh, r1, c1, r2, c2):
    unmerge_range(sh, r1, c1, r2, c2)
    if (r1, c1) == (r2, c2):
        return
    sh.merge(f'{cref(r1, c1)}:{cref(r2, c2)}')


# ====================================================================== 줄 묶음 늘리기·줄이기
def resize_blocks(sh, first, nold, nnew, size):
    """first 줄부터 size 줄씩 nold 개 있는 묶음을 nnew 개로. 마지막 묶음을 복사해 늘린다(서식·높이·병합·수식)."""
    if nnew > nold:
        sh.unshare()
        last = first + (nold - 1) * size
        for _ in range(nnew - nold):
            snaps = [sh.row_copy(r) for r in range(last, last + size)]
            ms = [(a, b) for a, b in sh.merged() if a[0] >= last and b[0] <= last + size - 1]
            sh.insert_rows(last + size, size)
            for i, sn in enumerate(snaps):
                sh.put_row(sn, last + size + i)
            for a, b in ms:
                sh.merge(f'{cref(a[0] + size, a[1])}:{cref(b[0] + size, b[1])}')
            last += size
    elif nnew < nold:
        sh.delete_rows(first + nnew * size, (nold - nnew) * size)
    return (nnew - nold) * size


# ====================================================================== 시트 크기 (열 너비·행 높이 → EMU)
class Geo:
    def __init__(self, sh, mdw=None):
        root = sh.root
        fp = root.find(q('sheetFormatPr'))
        self.def_h = 16.5
        self.dcw = None
        self.bcw = 8.0
        if fp is not None:
            if fp.get('defaultRowHeight'):
                self.def_h = float(fp.get('defaultRowHeight'))
            if fp.get('defaultColWidth'):
                self.dcw = float(fp.get('defaultColWidth'))
            if fp.get('baseColWidth'):
                self.bcw = float(fp.get('baseColWidth'))
        self.cols = {}
        cols = root.find(q('cols'))
        if cols is not None:
            for col in cols.findall(q('col')):
                w = float(col.get('width')) if col.get('width') else None
                hid = col.get('hidden') in ('1', 'true')
                for c in range(int(col.get('min')), min(int(col.get('max')), 2000) + 1):
                    self.cols[c] = (w, hid)
        self.rows = {}
        for r, row in sh.rows.items():
            ht = float(row.get('ht')) if row.get('ht') else None
            self.rows[r] = (ht, row.get('hidden') in ('1', 'true'))
        self.mdw = mdw or 7
        self._cx = {}
        self._ry = {}

    def col_px(self, c, mdw=None):
        mdw = mdw or self.mdw
        w, hid = self.cols.get(c, (None, False))
        if hid:
            return 0
        if w is None:
            w = self.dcw if self.dcw else math.trunc((self.bcw * mdw + 5) / mdw * 256) / 256
        return math.trunc(((256 * w + math.trunc(128 / mdw)) / 256) * mdw)

    def col_emu(self, c):
        return self.col_px(c) * EMU_PX

    def row_emu(self, r):
        ht, hid = self.rows.get(r, (None, False))
        if hid:
            return 0
        return round((ht if ht is not None else self.def_h) * EMU_PT)

    def x_of(self, c, off=0, mdw=None):
        """c 열(1부터) 왼쪽 끝 + off 의 절대 위치(EMU)"""
        if mdw:
            return sum(self.col_px(k, mdw) * EMU_PX for k in range(1, c)) + off
        if c not in self._cx:
            self._cx[c] = sum(self.col_emu(k) for k in range(1, c))
        return self._cx[c] + off

    def y_of(self, r, off=0):
        if r not in self._ry:
            self._ry[r] = sum(self.row_emu(k) for k in range(1, r))
        return self._ry[r] + off

    def at_x(self, x):
        c, acc = 1, 0
        while c < 16384:
            w = self.col_emu(c)
            if acc + w > x:
                return c, int(x - acc)
            acc += w
            c += 1
        return c, 0

    def at_y(self, y):
        r, acc = 1, 0
        while r < 1048576:
            h = self.row_emu(r)
            if acc + h > y:
                return r, int(y - acc)
            acc += h
            r += 1
        return r, 0


def calibrate_mdw(sh, drawing):
    """그림의 절대 위치(xfrm)와 칸 위치를 맞춰 보며 글꼴 최대 숫자폭(px)을 고른다."""
    if drawing is None:
        return 7
    g = Geo(sh, 7)
    samples = []
    for an in drawing.anchors():
        if an.kind == 'abs' or an.xfrm is None or an.fr is None or an.otype == 'grp':
            continue
        samples.append((an.fr[0] + 1, an.fr[1], an.xfrm[0]))
        if len(samples) >= 16:
            break
    if not samples:
        return 7
    best, err = 7, None
    for mdw in range(5, 14):
        e = sum(abs(g.x_of(c, off, mdw) - x) for c, off, x in samples)
        if err is None or e < err:
            best, err = mdw, e
    return best


# ====================================================================== 그림 (drawing 부품)
class Anchor:
    def __init__(self, el):
        self.el = el
        ln = localname(el)
        self.kind = {'twoCellAnchor': 'two', 'oneCellAnchor': 'one', 'absoluteAnchor': 'abs'}.get(ln, 'other')
        self.fr = self._pos('from')
        self.to = self._pos('to')
        ext = el.find('{%s}ext' % XDR)
        self.ext = (int(ext.get('cx')), int(ext.get('cy'))) if ext is not None else None
        pos = el.find('{%s}pos' % XDR)
        self.pos = (int(pos.get('x')), int(pos.get('y'))) if pos is not None else None
        self.obj, self.otype = None, 'other'
        kids = list(el)
        for child in el:
            if localname(child) == 'AlternateContent':      # mc:AlternateContent 로 감싼 그림·도형
                for alt in child:
                    kids.extend(list(alt))
        for child in kids:
            t = localname(child)
            if t in ('pic', 'sp', 'grpSp', 'graphicFrame', 'cxnSp', 'contentPart'):
                self.obj = child
                self.otype = {'pic': 'pic', 'sp': 'sp', 'grpSp': 'grp', 'graphicFrame': 'frame', 'cxnSp': 'cxn'}.get(t, 'other')
                break
        self.xfrm = None
        if self.obj is not None:
            xf = next(self.obj.iter('{%s}xfrm' % A_NS, '{%s}xfrm' % XDR), None)
            if xf is not None:
                off = xf.find('{%s}off' % A_NS)
                ex = xf.find('{%s}ext' % A_NS)
                if off is not None and ex is not None:
                    self.xfrm = (int(off.get('x')), int(off.get('y')), int(ex.get('cx')), int(ex.get('cy')))

    def _pos(self, tag):
        e = self.el.find('{%s}%s' % (XDR, tag))
        if e is None:
            return None
        vals = []
        for k in ('col', 'colOff', 'row', 'rowOff'):
            x = e.find('{%s}%s' % (XDR, k))
            vals.append(int(float(x.text)) if x is not None and x.text else 0)
        return tuple(vals)

    @property
    def is_textbox(self):
        if self.otype != 'sp':
            return False
        nv = self.obj.find('{%s}nvSpPr' % XDR)
        cnv = nv.find('{%s}cNvSpPr' % XDR) if nv is not None else None
        return cnv is not None and cnv.get('txBox') == '1'

    def box(self, geo):
        """(x, y, cx, cy) EMU"""
        if self.kind == 'abs':
            return self.pos[0], self.pos[1], self.ext[0], self.ext[1]
        x = geo.x_of(self.fr[0] + 1, self.fr[1])
        y = geo.y_of(self.fr[2] + 1, self.fr[3])
        if self.kind == 'two' and self.to is not None:
            x2 = geo.x_of(self.to[0] + 1, self.to[1])
            y2 = geo.y_of(self.to[2] + 1, self.to[3])
            return x, y, max(0, x2 - x), max(0, y2 - y)
        if self.ext:
            return x, y, self.ext[0], self.ext[1]
        return x, y, 0, 0

    def top_left(self, geo):
        if self.kind == 'abs':
            c, _ = geo.at_x(self.pos[0])
            r, _ = geo.at_y(self.pos[1])
            return r, c
        return self.fr[2] + 1, self.fr[0] + 1


class Drawing:
    """시트에 딸린 drawing 부품. 바꾼 뒤 save() 로 book.files 에 다시 쓴다."""

    def __init__(self, book, sh, path):
        self.book, self.sh, self.path = book, sh, path
        self.root = etree.fromstring(book.files[path])

    @classmethod
    def of(cls, book, sh, create=False):
        for kind, part in sh.rel_parts():
            if kind == 'drawing' and part in book.files:
                return cls(book, sh, part)
        if not create:
            return None
        return cls._create(book, sh)

    @classmethod
    def _create(cls, book, sh):
        nums = [int(m.group(1)) for n in list(book.files) + book.order for m in [re.match(r'xl/drawings/drawing(\d+)\.xml$', n)] if m]
        path = f'xl/drawings/drawing{max(nums + [0]) + 1}.xml'
        xml = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               f'<xdr:wsDr xmlns:xdr="{XDR}" xmlns:a="{A_NS}"/>')
        book.files[path] = xml.encode('utf-8')
        book.order.append(path)
        ov = etree.SubElement(book.ct, q('Override', CNS))
        ov.set('PartName', '/' + path)
        ov.set('ContentType', DRAW_CT)
        rp = X.rels_path(sh.path)
        if rp in book.files:
            rr = etree.fromstring(book.files[rp])
        else:
            rr = etree.Element(q('Relationships', PNS), nsmap={None: PNS})
            book.order.append(rp)
        rid = _next_rid(rr)
        rel = etree.SubElement(rr, q('Relationship', PNS))
        rel.set('Id', rid)
        rel.set('Type', R_DRAW)
        rel.set('Target', posixpath.relpath(path, posixpath.dirname(sh.path)))
        book.files[rp] = etree.tostring(rr, xml_declaration=True, encoding='UTF-8', standalone=True)
        de = etree.Element(q('drawing'))
        de.set(q('id', RNS), rid)
        insert_ordered(sh.root, de)
        book.touched.add(sh.path)
        return cls(book, sh, path)

    def anchors(self):
        out = []
        for el in self.root:
            ln = localname(el)
            if ln in ('twoCellAnchor', 'oneCellAnchor', 'absoluteAnchor'):
                out.append(Anchor(el))
            elif ln == 'AlternateContent':                  # mc:AlternateContent > Choice/Fallback > 앵커
                inner = next((x for alt in el for x in alt
                              if localname(x) in ('twoCellAnchor', 'oneCellAnchor', 'absoluteAnchor')), None)
                if inner is not None:
                    an = Anchor(inner)
                    an.wrapper = el
                    out.append(an)
        return out

    def remove(self, an):
        w = getattr(an, 'wrapper', None)
        if w is not None:
            self.root.remove(w)
        else:
            self.root.remove(an.el)

    def save(self):
        self.book.files[self.path] = etree.tostring(self.root, xml_declaration=True, encoding='UTF-8', standalone=True)

    def _rels(self):
        rp = X.rels_path(self.path)
        if rp in self.book.files:
            return rp, etree.fromstring(self.book.files[rp])
        if rp not in self.book.order:
            self.book.order.append(rp)
        return rp, etree.Element(q('Relationships', PNS), nsmap={None: PNS})

    def add_picture(self, data, x, y, cx, cy, geo, name='그림'):
        w, h, ext = img_size(data)
        book = self.book
        nums = [int(m.group(1)) for n in list(book.files) + book.order for m in [re.match(r'xl/media/image(\d+)\.\w+$', n)] if m]
        media = f'xl/media/image{max(nums + [0]) + 1}.{ext}'
        book.files[media] = data
        book.order.append(media)
        if not any(localname(d) == 'Default' and d.get('Extension', '').lower() == ext for d in book.ct):
            de = etree.Element(q('Default', CNS))
            de.set('Extension', ext)
            de.set('ContentType', 'image/png' if ext == 'png' else 'image/jpeg')
            book.ct.insert(0, de)
        rp, rr = self._rels()
        rid = _next_rid(rr)
        rel = etree.SubElement(rr, q('Relationship', PNS))
        rel.set('Id', rid)
        rel.set('Type', R_IMG)
        rel.set('Target', posixpath.relpath(media, posixpath.dirname(self.path)))
        book.files[rp] = etree.tostring(rr, xml_declaration=True, encoding='UTF-8', standalone=True)
        ids = [int(e.get('id')) for e in self.root.iter('{%s}cNvPr' % XDR) if (e.get('id') or '').isdigit()]
        nid = max(ids + [1]) + 1
        cx, cy = int(max(1, cx)), int(max(1, cy))
        c, coff = geo.at_x(x)
        r, roff = geo.at_y(y)
        c2, coff2 = geo.at_x(x + cx)
        r2, roff2 = geo.at_y(y + cy)
        # 시작·끝을 칸에 묶는다(editAs=oneCell: 셀 따라 이동, 크기 고정) — 인쇄 때 칸과 같은 비율로 커지고 작아진다
        xml = (f'<xdr:twoCellAnchor editAs="oneCell" xmlns:xdr="{XDR}" xmlns:a="{A_NS}" xmlns:r="{RNS}">'
               f'<xdr:from><xdr:col>{c - 1}</xdr:col><xdr:colOff>{coff}</xdr:colOff>'
               f'<xdr:row>{r - 1}</xdr:row><xdr:rowOff>{roff}</xdr:rowOff></xdr:from>'
               f'<xdr:to><xdr:col>{c2 - 1}</xdr:col><xdr:colOff>{coff2}</xdr:colOff>'
               f'<xdr:row>{r2 - 1}</xdr:row><xdr:rowOff>{roff2}</xdr:rowOff></xdr:to>'
               f'<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="{nid}" name="{name} {nid}"/>'
               f'<xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>'
               f'<xdr:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>'
               f'<xdr:spPr><a:xfrm><a:off x="{int(x)}" y="{int(y)}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
               f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>'
               f'<xdr:clientData/></xdr:twoCellAnchor>')
        self.root.append(etree.fromstring(xml))
        return w, h


# ====================================================================== 서식(styles.xml)
STYLE_ORDER = ['numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs', 'cellStyles', 'dxfs', 'tableStyles',
               'colors', 'extLst']
BORDER_ORDER = ['start', 'left', 'end', 'right', 'top', 'bottom', 'diagonal', 'vertical', 'horizontal']


def _insert_in(parent, el, order):
    idx = order.index(localname(el))
    after = None
    for ch in parent:
        ln = localname(ch)
        if ln in order and order.index(ln) < idx:
            after = ch
    if after is None:
        parent.insert(0, el)
    else:
        after.addnext(el)


class Styles:
    def __init__(self, book):
        if book._styles is None:
            book._styles = etree.fromstring(book.files['xl/styles.xml'])
        self.book = book
        self.root = book._styles
        self.cache = {}

    def _list(self, tag):
        e = self.root.find(q(tag))
        if e is None:
            e = etree.Element(q(tag))
            _insert_in(self.root, e, STYLE_ORDER)
        return e

    def xf(self, s):
        return self._list('cellXfs')[s]

    def font_of(self, s):
        return self._list('fonts')[int(self.xf(s).get('fontId') or 0)]

    def _append(self, tag, el):
        lst = self._list(tag)
        lst.append(el)
        lst.set('count', str(len(lst)))
        return len(lst) - 1

    def numfmt_id(self, code):
        builtin = {'General': 0, '0': 1, '0.00': 2, '#,##0': 3, '#,##0.00': 4}
        if code in builtin:
            return builtin[code]
        nf = self._list('numFmts')
        for e in nf:
            if e.get('formatCode') == code:
                return int(e.get('numFmtId'))
        nid = max([int(e.get('numFmtId')) for e in nf] + [163]) + 1
        e = etree.SubElement(nf, q('numFmt'))
        e.set('numFmtId', str(nid))
        e.set('formatCode', code)
        nf.set('count', str(len(nf)))
        return nid

    def derive(self, s, bold=None, size=None, fill=None, no_fill=False, border=None, numfmt=None, align=None):
        """s 서식을 바탕으로 일부만 바꾼 새 서식 번호.
        border: 'thin' (네 변 얇은 선) / 'diag' (대각선 둘) / 'nodiag' (대각선 지움)"""
        key = (s, bold, size, fill, no_fill, border, numfmt, tuple(sorted((align or {}).items())))
        if key in self.cache:
            return self.cache[key]
        nx = copy.deepcopy(self.xf(s))
        if bold is not None or size is not None:
            f = copy.deepcopy(self.font_of(s))
            if bold is not None:
                b = f.find(q('b'))
                if bold and b is None:
                    f.insert(0, etree.Element(q('b')))
                elif not bold and b is not None:
                    f.remove(b)
            if size is not None:
                sz = f.find(q('sz'))
                if sz is None:
                    sz = etree.Element(q('sz'))
                    f.insert(0, sz)
                sz.set('val', str(size))
            nx.set('fontId', str(self._append('fonts', f)))
            nx.set('applyFont', '1')
        if fill or no_fill:
            if no_fill:
                nx.set('fillId', '0')
            else:
                fe = etree.Element(q('fill'))
                pf = etree.SubElement(fe, q('patternFill'))
                pf.set('patternType', 'solid')
                fg = etree.SubElement(pf, q('fgColor'))
                fg.set('rgb', 'FF' + fill)
                bg = etree.SubElement(pf, q('bgColor'))
                bg.set('indexed', '64')
                nx.set('fillId', str(self._append('fills', fe)))
            nx.set('applyFill', '1')
        if border:
            borders = self._list('borders')
            be = copy.deepcopy(borders[int(nx.get('borderId') or 0)])
            if border == 'thin':
                for side in ('left', 'right', 'top', 'bottom'):
                    e = be.find(q(side))
                    if e is None:
                        e = etree.Element(q(side))
                        _insert_in(be, e, BORDER_ORDER)
                    e.set('style', 'thin')
                    for ch in list(e):
                        e.remove(ch)
                    col = etree.SubElement(e, q('color'))
                    col.set('indexed', '64')
            else:
                d = be.find(q('diagonal'))
                if d is None:
                    d = etree.Element(q('diagonal'))
                    _insert_in(be, d, BORDER_ORDER)
                for ch in list(d):
                    d.remove(ch)
                if border == 'diag':
                    be.set('diagonalUp', '1')
                    be.set('diagonalDown', '1')
                    d.set('style', 'thin')
                    col = etree.SubElement(d, q('color'))
                    col.set('indexed', '64')
                else:
                    be.attrib.pop('diagonalUp', None)
                    be.attrib.pop('diagonalDown', None)
                    d.attrib.pop('style', None)
            nx.set('borderId', str(self._append('borders', be)))
            nx.set('applyBorder', '1')
        if numfmt is not None:
            nx.set('numFmtId', str(self.numfmt_id(numfmt)))
            nx.set('applyNumberFormat', '1')
        if align:
            al = nx.find(q('alignment'))
            if al is None:
                al = etree.Element(q('alignment'))
                nx.insert(0, al)
            for k, v in align.items():
                if v is None:
                    al.attrib.pop(k, None)
                else:
                    al.set(k, str(v))
            nx.set('applyAlignment', '1')
        idx = self._append('cellXfs', nx)
        self.cache[key] = idx
        return idx

    def has_diag(self, s):
        be = self._list('borders')[int(self.xf(s).get('borderId') or 0)]
        return be.get('diagonalUp') in ('1', 'true') or be.get('diagonalDown') in ('1', 'true')


def set_rich(sh, styles, ref, text, red_from):
    """칸 글자를 바꾸고 red_from(0부터) 이후만 빨간색으로 (앞은 검정). 글꼴은 칸 서식 그대로."""
    c = sh._cell(ref)
    font = styles.font_of(int(c.get('s') or 0))

    def rpr(rgb):
        r = etree.Element(q('rPr'))
        for ch in font:
            ln = localname(ch)
            if ln == 'color':
                continue
            if ln == 'name':
                e = etree.SubElement(r, q('rFont'))
                e.set('val', ch.get('val'))
            elif ln in ('b', 'i', 'strike', 'u', 'vertAlign', 'sz', 'family', 'charset', 'scheme', 'outline',
                        'shadow', 'condense', 'extend'):
                r.append(copy.deepcopy(ch))
        col = etree.SubElement(r, q('color'))
        col.set('rgb', rgb)
        return r
    si = etree.Element(q('si'))
    parts = [(text[:red_from], 'FF000000'), (text[red_from:], 'FFFF0000')] if red_from is not None else [(text, 'FF000000')]
    for t, rgb in parts:
        if not t:
            continue
        run = etree.SubElement(si, q('r'))
        run.append(rpr(rgb))
        te = etree.SubElement(run, q('t'))
        te.text = t
        te.set(X.XML_SPACE, 'preserve')
    if not len(si):
        etree.SubElement(si, q('t')).text = ''
    book = sh.book
    book.sst_root.append(si)
    book.sst.append(si)
    sh._clear(c)
    c.set('t', 's')
    v = etree.SubElement(c, q('v'))
    v.text = str(len(book.sst) - 1)
    book.touched.add(sh.path)


def set_col_width(sh, c1, c2, width):
    cols = sh.root.find(q('cols'))
    if cols is None:
        cols = etree.Element(q('cols'))
        insert_ordered(sh.root, cols)
    keep = []
    for col in list(cols):
        mn, mx = int(col.get('min')), int(col.get('max'))
        cols.remove(col)
        if mx < c1 or mn > c2:
            keep.append(col)
            continue
        if mn < c1:
            a = copy.deepcopy(col)
            a.set('max', str(c1 - 1))
            keep.append(a)
        if mx > c2:
            b = copy.deepcopy(col)
            b.set('min', str(c2 + 1))
            keep.append(b)
    new = etree.Element(q('col'))
    new.set('min', str(c1))
    new.set('max', str(c2))
    new.set('width', str(width))
    new.set('customWidth', '1')
    keep.append(new)
    keep.sort(key=lambda e: int(e.get('min')))
    for e in keep:
        cols.append(e)
    sh.book.touched.add(sh.path)


# ====================================================================== 인쇄영역 · 쪽 나누기
def sheet_index(book, name):
    return [s['name'] for s in book.sheet_list()].index(name)


def get_print_area(book, name):
    dn = book.wb.find(q('definedNames'))
    if dn is None:
        return None
    idx = str(sheet_index(book, name))
    for d in dn:
        if d.get('name') == '_xlnm.Print_Area' and d.get('localSheetId') == idx:
            return d.text
    return None


def set_print_area(book, name, rng):
    dn = book.wb.find(q('definedNames'))
    if dn is None:
        dn = etree.Element(q('definedNames'))
        anchor = book.wb.find(q('sheets'))
        for tag in ('functionGroups', 'externalReferences'):
            e = book.wb.find(q(tag))
            if e is not None:
                anchor = e
        anchor.addnext(dn)
    idx = str(sheet_index(book, name))
    text = f'{X._qname(name)}!{rng}'
    for d in dn:
        if d.get('name') == '_xlnm.Print_Area' and d.get('localSheetId') == idx:
            d.text = text
            break
    else:
        d = etree.SubElement(dn, q('definedName'))
        d.set('name', '_xlnm.Print_Area')
        d.set('localSheetId', idx)
        d.text = text
    book.touched.add('xl/workbook.xml')


def set_row_breaks(sh, rows):
    old = sh.root.find(q('rowBreaks'))
    if old is not None:
        sh.root.remove(old)
    rows = sorted({r for r in rows if r > 1})
    if not rows:
        return
    rb = etree.Element(q('rowBreaks'))
    rb.set('count', str(len(rows)))
    rb.set('manualBreakCount', str(len(rows)))
    for r in rows:
        b = etree.SubElement(rb, q('brk'))
        b.set('id', str(r - 1))
        b.set('max', '16383')
        b.set('man', '1')
    insert_ordered(sh.root, rb)
    sh.book.touched.add(sh.path)


# ====================================================================== 새 시트
def add_sheet(book, name, rows, hidden=True, widths=None, styles=None):
    """rows(줄 목록)로 새 시트를 맨 끝에 만든다. 값이 '=…' 이면 수식. styles = {(r, c): 서식번호}"""
    nums = [int(m.group(1)) for n in list(book.files) + book.order for m in [re.match(r'xl/worksheets/sheet(\d+)\.xml$', n)] if m]
    path = f'xl/worksheets/sheet{max(nums + [0]) + 1}.xml'
    root = etree.Element(q('worksheet'), nsmap={None: NS, 'r': RNS})
    sv = etree.SubElement(etree.SubElement(root, q('sheetViews')), q('sheetView'))
    sv.set('workbookViewId', '0')
    fp = etree.SubElement(root, q('sheetFormatPr'))
    fp.set('defaultRowHeight', '16.5')
    if widths:
        cols = etree.SubElement(root, q('cols'))
        for (c1, c2), w in sorted(widths.items()):
            col = etree.SubElement(cols, q('col'))
            col.set('min', str(c1))
            col.set('max', str(c2))
            col.set('width', str(w))
            col.set('customWidth', '1')
    data = etree.SubElement(root, q('sheetData'))
    styles = styles or {}
    for i, row in enumerate(rows, 1):
        re_ = etree.SubElement(data, q('row'))
        re_.set('r', str(i))
        for j, v in enumerate(row, 1):
            empty = v is None or v == ''
            if empty and (i, j) not in styles:
                continue
            c = etree.SubElement(re_, q('c'))
            c.set('r', cref(i, j))
            if (i, j) in styles:
                c.set('s', str(styles[(i, j)]))
            if empty:
                continue
            if isinstance(v, bool):
                v = int(v)
            if isinstance(v, (int, float)):
                e = etree.SubElement(c, q('v'))
                e.text = str(int(v)) if float(v).is_integer() else repr(float(v))
            elif isinstance(v, str) and v.startswith('=') and len(v) > 1:
                e = etree.SubElement(c, q('f'))
                e.text = v[1:]
            else:
                c.set('t', 's')
                e = etree.SubElement(c, q('v'))
                e.text = str(book.add_string(str(v)))
    pm = etree.SubElement(root, q('pageMargins'))
    for k, v in (('left', '0.7'), ('right', '0.7'), ('top', '0.75'), ('bottom', '0.75'), ('header', '0.3'), ('footer', '0.3')):
        pm.set(k, v)
    book.files[path] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    book.order.append(path)
    rid = _next_rid(book.rels)
    rel = etree.SubElement(book.rels, q('Relationship', PNS))
    rel.set('Id', rid)
    rel.set('Type', X.WS_TYPE)
    rel.set('Target', path[len('xl/'):])
    ov = etree.SubElement(book.ct, q('Override', CNS))
    ov.set('PartName', '/' + path)
    ov.set('ContentType', WS_CT)
    sheets = book.wb.find(q('sheets'))
    sid = max(int(s.get('sheetId')) for s in sheets) + 1
    el = etree.SubElement(sheets, q('sheet'))
    el.set('name', name)
    el.set('sheetId', str(sid))
    if hidden:
        el.set('state', 'hidden')
    el.set(q('id', RNS), rid)
    book.touched.add('xl/workbook.xml')
    return book.sheet(name)


# ====================================================================== 채우기
class Builder:
    SECTIONS = [('location', r'^4\.위치도'), ('birdseye', r'^5\.조감도'), ('siteplan', r'^6\.단지배치도'),
                ('community', r'^7\.커뮤니티')]

    NEW_SHEETS = ('공고자료', '보고자료')               # build 가 새로 만드는 시트

    def __init__(self, data, log=None):
        self.book = X.Book(data)
        self.styles = Styles(self.book)
        self.log = log or (lambda *a: None)
        self.mdw = 7
        self.dc = 0                                     # 분양개요 시트가 오른쪽으로 밀린 열 수 (find_offset)

    # -------------------------------------------------------------- 찾기
    def sheet(self, *names):
        for s in self.book.sheet_list():
            if ns(s['name']) in [ns(n) for n in names]:
                return self.book.sheet(s['name'])
        return None

    @staticmethod
    def val(ws, r, c):
        return ws.value(cref(r, c))

    def find_row(self, ws, pattern, col=1, start=1, end=400, cols=1):
        rx = re.compile(pattern)
        for r in range(start, end + 1):
            for c in range(col, col + cols):
                v = ws.value(cref(r, c))
                if v is not None and rx.search(ns(v)):
                    return r
        return None

    def find_cell(self, ws, pattern, r1, c1, r2, c2):
        rx = re.compile(pattern)
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                v = ws.value(cref(r, c))
                if v is not None and rx.search(ns(v)):
                    return r, c
        return None

    def find_offset(self, ws):
        """양식 맨 앞에 열이 더 끼어 있으면(폭 0 으로 숨긴 빈 A열 등) 그만큼 오른쪽으로 밀어 쓴다.
        보통 양식은 '1. 일반개요'·'3. 공급개요' 제목이 A열 (2026-10-23 M1-2-2BL 양식은 B열)"""
        rc = self.find_cell(ws, r'^[13]\.(일반개요|공급개요)', 1, 1, 60, 6)
        return rc[1] - 1 if rc else 0

    # -------------------------------------------------------------- 그림
    def geo(self, ws):
        return Geo(ws, self.mdw)

    def drawing(self, ws, create=False):
        return Drawing.of(self.book, ws, create)

    def pictures_in(self, ws, r1, r2, max_col=None, min_col=None):
        dr = self.drawing(ws)
        if dr is None:
            return 0
        g = self.geo(ws)
        n = 0
        for an in dr.anchors():
            if an.otype not in ('pic', 'grp'):
                continue
            r, c = an.top_left(g)
            if not (r1 <= r <= r2):
                continue
            if max_col and c > max_col:
                continue
            if min_col and c < min_col:
                continue
            dr.remove(an)
            n += 1
        if n:
            dr.save()
        return n

    def frames_in(self, ws, r1, r2):
        dr = self.drawing(ws)
        if dr is None:
            return []
        g = self.geo(ws)
        out = []
        for an in dr.anchors():
            if an.otype != 'sp' or an.is_textbox:
                continue
            r, _ = an.top_left(g)
            x, y, w, h = an.box(g)
            if r1 <= r <= r2 and w / EMU_PT > 200 and h / EMU_PT > 100:
                out.append((x / EMU_PT, y / EMU_PT, w / EMU_PT, h / EMU_PT))
        return out

    def place_images(self, ws, images, left, top, width, height, margin=6, gap=6, align='center'):
        """images = [{'data': bytes, 'w', 'h'}] — 틀(점 단위) 안에 비율 유지. 여러 장이면 가장 크게 보이는 바둑판 배치"""
        images = [im for im in images if im and im.get('data')]
        if not images:
            return 0
        sizes = []
        for im in images:
            w, h, _ = img_size(im['data'])
            sizes.append((im.get('w') or w or 1, im.get('h') or h or 1))
        L, T = left + margin, top + margin
        W, H = width - 2 * margin, height - 2 * margin
        n = len(images)
        best = None
        for cols in range(1, n + 1):
            rows = -(-n // cols)
            cw, ch = (W - gap * (cols - 1)) / cols, (H - gap * (rows - 1)) / rows
            if cw <= 0 or ch <= 0:
                continue
            sc = [min(cw / w, ch / h) for (w, h) in sizes]
            area = sum(s * s * w * h for s, (w, h) in zip(sc, sizes))
            if best is None or area > best[0]:
                best = (area, cols, cw, ch, sc)
        _, cols, cw, ch, sc = best
        dr = self.drawing(ws, create=True)
        g = self.geo(ws)
        for k, (im, (w, h), s) in enumerate(zip(images, sizes, sc)):
            ww, hh = w * s, h * s
            cx, cy = L + (k % cols) * (cw + gap), T + (k // cols) * (ch + gap)
            x, y = (cx, cy) if align == 'topleft' else (cx + (cw - ww) / 2, cy + (ch - hh) / 2)
            dr.add_picture(im['data'], x * EMU_PT, y * EMU_PT, ww * EMU_PT, hh * EMU_PT, g)
        dr.save()
        return len(images)

    def area_box(self, ws, r1, c1, r2, c2):
        g = self.geo(ws)
        x, y = g.x_of(c1), g.y_of(r1)
        return x / EMU_PT, y / EMU_PT, (g.x_of(c2 + 1) - x) / EMU_PT, (g.y_of(r2 + 1) - y) / EMU_PT

    # -------------------------------------------------------------- 메인
    def build(self, d, images):
        ws = self.sheet('분양개요') or self.book.sheet(self.book.sheet_list()[0]['name'])
        self.ws = ws
        for s in self.book.all_sheets():
            s.unshare()
        self.mdw = calibrate_mdw(ws, self.drawing(ws))
        self.dc = self.find_offset(ws)
        if self.dc:
            self.log(f'양식 맨 앞에 열이 {self.dc}개 더 있어(표가 {num2col(1 + self.dc)}열부터 시작) 그만큼 밀어서 씁니다')

        def step(name, fn, *a):
            t = time.time()
            r = fn(*a)
            self.log(f'{name} ({time.time() - t:.1f}초)')
            return r
        old_hidden = [s['name'] for s in self.book.sheet_list() if s['state'] != 'visible']
        step('일반개요·사업일정', self.fill_top, ws, d)
        r0, sum_row = step('공급개요', self.fill_supply, ws, d)
        step('위치도·조감도·배치도·커뮤니티 그림', self.fill_images, ws, d, images, sum_row)
        step('견본주택 위치도(O6)', self.fill_mh_map, ws, d, images)
        step('단위세대', self.fill_units, ws, d, images)
        step('인쇄영역·쪽 나누기', self.set_pages, ws)
        step('공고자료 시트', self.data_sheet, d)
        step('보고자료 시트', self.report_sheet, d)
        step('청약결과 시트', self.fill_result, d, r0)
        if d.get('delete_hidden', True):
            step('이전 현장 숨김 시트 정리', self.delete_sheets, old_hidden)
        wv = next(self.book.wb.iter(q('workbookView')), None)
        if wv is not None:
            wv.set('activeTab', str(sheet_index(self.book, ws.name)))
        self.book._structure_changed = True
        for sv in ws.root.iter(q('sheetView')):
            sv.attrib.pop('topLeftCell', None)
            for sel in sv.findall(q('selection')):
                sel.set('activeCell', 'A1')
                sel.set('sqref', 'A1')

    # -------------------------------------------------------------- 1·2. 일반개요 · 사업일정
    def fill_top(self, ws, d):
        dc = self.dc
        put(ws, 1, 1 + dc, d['title'])
        r = self.find_row(ws, r'^\(?작성일', 1 + dc, 1, 6)
        if r:
            put(ws, r, 1 + dc, f"(작성일 : {d['written']})")
        rc = self.find_cell(ws, r'^https?://|^www\.', 1, 15 + dc, 8, 30 + dc)
        if rc:
            put(ws, rc[0], rc[1], d.get('homepage') or '')
        elif d.get('homepage'):
            put(ws, 5, 16 + dc, d['homepage'])                                # P5
        labels = [('공급위치', 'location'), ('단지규모', 'scale'), ('단지특징', 'features'),
                  ('주차대수', 'parking_text'), ('위탁', 'developers'), ('시행', 'developers'), ('견본주택', 'model_house'),
                  ('오픈일', 'mh_open'), ('모집공고', 'notice_date'), ('청약일정', 'subscription'),
                  ('당첨자', 'winners'), ('계약일정', 'contract'), ('입주예정', 'move_in')]
        gcol = 7 + dc
        at = {}
        for r in range(5, 21):
            lab = ns(self.val(ws, r, 2 + dc))
            if not lab:
                continue
            for key, field in labels:
                if key in lab:
                    at[field] = r
                    v = d.get(field)
                    if field in ('mh_open', 'notice_date', 'winners'):
                        ser = excel_date(v)
                        put(ws, r, gcol, ser if ser else '')
                    else:
                        put(ws, r, gcol, v or '')
                    if field == 'developers' and d.get('developers_label'):
                        put(ws, r, 2 + dc, d['developers_label'])           # 위탁/수탁/시공 · 시행/시공
                    break
        if not d.get('mh_open') and at.get('mh_open') and at.get('notice_date'):
            # M/H 오픈일이 공고에 없으면 모집공고일과 같게 (2026-10-18 사용자 파일: =G15)
            put(ws, at['mh_open'], gcol, '=' + num2col(gcol) + str(at['notice_date']))
        rc = self.find_cell(ws, r'^세대수$', 12, 15 + dc, 20, 30 + dc)
        if rc:
            put(ws, rc[0], rc[1] + 1, d.get('households') or '')
            put(ws, rc[0] + 1, rc[1] + 1, d.get('parking') or '')
            a, b = num2col(rc[1] + 1) + str(rc[0] + 1), num2col(rc[1] + 1) + str(rc[0])
            put(ws, rc[0] + 2, rc[1] + 1, f'=IFERROR(TRUNC({a}/{b},2),"")')

    # -------------------------------------------------------------- 3. 공급개요
    def big_merges(self, ws, r1, r2, c1, c2):
        """r1~r2 안에서 2줄 넘게 걸친 병합(구분·비고 칸)을 풀고 위치를 기억"""
        found = {}
        for c in range(c1, c2 + 1):
            for r in (r1, r1 + 1):
                a = merge_area(ws, r, c)
                if a[2] - a[0] + 1 > 2 and a[0] <= r1 + 1:
                    key = (a[0], a[1])
                    if key not in found:
                        found[key] = {'row': a[0], 'col': a[1], 'ncols': a[3] - a[1] + 1, 'last': a[2]}
        for m in found.values():
            unmerge_range(ws, m['row'], m['col'], m['last'], m['col'] + m['ncols'] - 1)
        return list(found.values())

    def fill_supply(self, ws, d):
        dc = self.dc
        c1 = 1 + dc                                   # 구분 칸 (보통 A열)
        sec = self.find_row(ws, r'^3\.공급개요', c1, 15, 60)
        hdr = self.find_row(ws, r'^구분$', c1, (sec or 18) + 1, (sec or 18) + 5)
        if not hdr:
            raise BuildError(f'양식에서 "3. 공급개요" 표의 "구  분" 머리글을 찾지 못했습니다 '
                             f'(공급개요 제목 바로 아래 {num2col(c1)}열에 있어야 합니다).')
        r0 = hdr + 2
        sum_row = self.find_row(ws, r'^소계$', 3 + dc, r0, r0 + 120)
        if not sum_row:
            raise BuildError('양식의 공급개요 표에서 "소 계" 줄을 찾지 못했습니다.')

        def hcol(pat, default):                       # 못 찾으면 보통 양식의 열 + dc
            rx = re.compile(pat)
            for rr in (hdr, hdr + 1):
                for c in range(c1, 15 + dc):
                    v = self.val(ws, rr, c)
                    if v is not None and rx.search(ns(v)):
                        return c
            return default + dc
        cC = hcol(r'전용면적', 3)
        cE = hcol(r'공급면적', 5)
        cG = hcol(r'^\(평\)$', 7)
        cH = hcol(r'세대', 8)
        cI = hcol(r'공급금액', 9)
        cJ = hcol(r'확장', 10)
        cK = hcol(r'소계', 11)
        cL = hcol(r'유상옵션|옵션', 12)
        cM = hcol(r'합계', 13)
        cN = hcol(r'비고', 14)
        L = num2col
        types = d['types']
        n = len(types)
        nold = (sum_row - r0) // 2
        self.pictures_in(ws, sec or hdr, sum_row + 12, min_col=cN + 1)   # 위쪽 O6 견본주택 위치도는 fill_mh_map 이 바꾼다
        merges = self.big_merges(ws, r0, sum_row, c1, cN)
        self.log(f'공급개요 {nold}개 타입 → {n}개')
        delta = resize_blocks(ws, r0, nold, n, 2)
        sum_row += delta
        for m in merges:
            merge_range(ws, m['row'], m['col'], m['last'] + delta, m['col'] + m['ncols'] - 1)
        # 유상옵션 내역 칸: Q 열부터 '비고 포함' 품목의 고른 안만, 유상옵션(C) = 그 합계 (2026-10-18 사용자 파일 기준)
        ocols = chosen_option_cols(d)
        oc0 = cN + 3
        far = max(60, oc0 + len(ocols) + 2)
        clear_range(ws, hdr, cN + 1, sum_row + 1, far)
        if ocols:
            unmerge_range(ws, hdr, oc0, sum_row + 1, far)
            for j, c in enumerate(ocols):
                put(ws, hdr + 1, oc0 + j, c.get('header') or c.get('name') or '')
        for i, t in enumerate(types):
            r = r0 + 2 * i
            r1 = r + 1
            row_a = {cC: t['excl'], cE: t['supply'], cG: f'={L(cE)}{r}*0.3025', cH: t['units'],
                     cI: round((t.get('price') or 0) / 1000, 1), cJ: round((t.get('balcony') or 0) / 1000, 1),
                     cK: f'=+{L(cI)}{r}+{L(cJ)}{r}',
                     cL: f'=SUM({L(oc0)}{r}:{L(oc0 + len(ocols) - 1)}{r})' if ocols else '', cM: f'=+{L(cL)}{r}+{L(cK)}{r}'}
            if cC + 1 < cE:
                row_a[cC + 1] = t.get('suffix') or ''
            row_b = {cI: f'=+{L(cI)}{r}/{L(cG)}{r}', cJ: f'=+{L(cJ)}{r}/{L(cG)}{r}', cK: f'=+{L(cI)}{r1}+{L(cJ)}{r1}',
                     cL: f'=+{L(cL)}{r}/{L(cG)}{r}', cM: f'=+{L(cK)}{r1}+{L(cL)}{r1}'}
            for c in range(cC, cM + 1):
                put(ws, r, c, row_a.get(c, ''))
                put(ws, r1, c, row_b.get(c, ''))
        if ocols and types:
            for i, t in enumerate(types):
                for j, c in enumerate(ocols):
                    v = (c.get('values') or {}).get(t['abbr'])
                    put(ws, r0 + 2 * i, oc0 + j, round(v / 1000, 1) if v else '')
            self.option_block_format(ws, hdr, oc0, r0, r0 + 2 * len(types) - 1, len(ocols))
        put(ws, r0, c1, d.get('label') or '')
        put(ws, r0, cN, d.get('remarks') or '')
        put(ws, sum_row, cH, f'=SUM({L(cH)}{r0}:{L(cH)}{sum_row - 1})')
        put(ws, sum_row, cI, d.get('basis_note') or '')
        for rr in range(sum_row + 1, sum_row + 6):
            v = ns(self.val(ws, rr, c1))
            if '평균평당가' in v:
                self.rich(ws, rr, c1, d.get('avg_note') or '', r'\d')
            elif '납부조건' in v:
                self.rich(ws, rr, c1, d.get('payment_note') or '', r'중도금')
            elif (v.startswith('※유상옵션') or '유상옵션:' in v) and d.get('option_note'):
                put(ws, rr, c1, d['option_note'])
        self.cols = dict(C=cC, E=cE, G=cG, H=cH, I=cI, J=cJ, K=cK, L=cL, M=cM, N=cN)
        self.r0, self.sum_row = r0, sum_row
        return r0, sum_row

    def option_block_format(self, ws, hdr, c0, r1, r2, n):
        """T열 옵션표 모양: 머리글 줄바꿈·회색·굵게, 금액 천단위, 얇은 테두리 (인쇄영역 밖)"""
        try:
            st = self.styles
            c1 = c0 + n - 1
            ref = cref(hdr, c0)
            ws.set_style(ref, st.derive(ws.style(ref) or 0, bold=True, align={'wrapText': None}))
            for c in range(c0, c1 + 1):
                ref = cref(hdr + 1, c)
                ws.set_style(ref, st.derive(ws.style(ref) or 0, bold=True, fill='EFEFEF', border='thin',
                                            align={'wrapText': '1', 'horizontal': 'center', 'vertical': 'center'}))
                for r in range(r1, r2 + 1):
                    ref = cref(r, c)
                    ws.set_style(ref, st.derive(ws.style(ref) or 0, no_fill=True, border='thin', numfmt='#,##0;-#,##0;;@',
                                                align={'horizontal': 'right'}))
            set_col_width(ws, c0, c1, 9.5)
        except Exception as e:
            self.log(f'유상옵션 칸 서식 실패: {e}')

    def rich(self, ws, r, c, text, red_from_pat):
        m = re.search(red_from_pat, text)
        try:
            set_rich(ws, self.styles, cref(r, c), text, m.start() if m else None)
        except Exception as e:
            self.log(f'글자색 지정 실패: {e}')
            put(ws, r, c, text)

    # -------------------------------------------------------------- 4~7. 그림 구역
    def section_rows(self, ws):
        rows = {}
        c1 = 1 + self.dc
        for key, pat in self.SECTIONS:
            r = self.find_row(ws, pat, c1, 1, 600)
            if r is None:
                r = self.find_row(ws, re.sub(r'^\^\d\\\.', '', pat), c1, 1, 600)
            rows[key] = r
        rows['units'] = self.find_row(ws, r'^\d\.단위세대', c1, 1, 800)
        return rows

    def fill_images(self, ws, d, images, sum_row):
        rows = self.section_rows(ws)
        order = [k for k in ('location', 'birdseye', 'siteplan', 'community', 'units') if rows.get(k)]
        order.sort(key=lambda k: rows[k])
        cN = self.cols['N']
        for i, key in enumerate(order):
            if key == 'units':
                continue
            r1 = rows[key]
            r2 = (rows[order[i + 1]] - 1) if i + 1 < len(order) else r1 + 60
            files = images.get(key) or []
            if not files:
                self.log(f'{key}: 넣을 그림이 없어 양식 그림을 그대로 둡니다')
                continue
            frames = self.frames_in(ws, r1, r2)
            self.pictures_in(ws, r1, r2, max_col=cN)
            if frames:
                left, top, w, h = max(frames, key=lambda f: f[2] * f[3])
            else:
                g = self.geo(ws)
                top = g.y_of(r1 + 1) / EMU_PT
                left = g.x_of(2 + self.dc) / EMU_PT
                w = g.x_of(cN + 1) / EMU_PT - left
                h = g.y_of(r2 + 1) / EMU_PT - top
            self.place_images(ws, files, left, top, w, h)

    # -------------------------------------------------------------- 8. 단위세대
    def fill_units(self, ws, d, images):
        rows = self.section_rows(ws)
        h = rows.get('units')
        if not h:
            self.log('단위세대 구역을 찾지 못했습니다')
            return
        first = h + 2
        c1, cD = 1 + self.dc, 4 + self.dc               # 타입 칸(A) · 기본형 그림 칸(D)
        a = merge_area(ws, first, cD)
        size = a[2] - a[0] + 1 if a[2] > a[0] else 6
        nold = 0
        r = first
        while nold < 60:
            m = merge_area(ws, r, cD)
            if m[2] - m[0] + 1 != size or m[0] != r:
                break
            nold += 1
            r += size
        if nold == 0:
            self.log('단위세대 칸 모양을 알 수 없어 건너뜁니다')
            return
        cN = self.cols['N']
        self.pictures_in(ws, h, first + nold * size + 2, max_col=cN)
        types = d['types']
        resize_blocks(ws, first, nold, len(types), size)
        unit_imgs = images.get('units') or {}
        kb = merge_area(ws, h + 1, cD)
        ke_col = kb[1] + (kb[3] - kb[1] + 1)
        heads = unit_headers(unit_imgs, [t['abbr'] for t in types])      # 기본형|확장형 또는 확장형|아이소
        put(ws, h + 1, cD, heads[0])
        put(ws, h + 1, ke_col, heads[1])
        for i, t in enumerate(types):
            b = first + i * size
            lab = t['abbr']
            put(ws, b + 2, c1, int(lab) if re.fullmatch(r'\d+', lab) else lab)
            sr = getattr(self, 'r0', None)
            if sr:
                put(ws, b + 3, c1, f"={num2col(self.cols['C'])}{sr + 2 * i}*0.3025")
            else:
                put(ws, b + 3, c1, round(t['excl'] * 0.3025, 2))
            left, right, _ = unit_pair(unit_imgs.get(t['abbr']) or {})
            for f, col, is_left in ((left, cD, True), (right, ke_col, False)):
                area = merge_area(ws, b, col)
                if is_left:
                    self.cross(ws, area, not f)          # 왼쪽 칸 그림이 없으면 X 표 (기본형 없음)
                if not f:
                    continue
                self.place_images(ws, [f], *self.area_box(ws, *area), margin=4)
        self.units_first, self.units_size, self.units_n = first, size, len(types)

    # -------------------------------------------------------------- O6 견본주택 위치도
    def fill_mh_map(self, ws, d, images):
        """견본주택 위치도 → O6, 인쇄영역 밖 (2026-10-18 사용자 지정). 양식에 있던 약도 크기에 맞춘다"""
        files = images.get('mh_map') or []
        cN = self.cols['N']
        dr = self.drawing(ws)
        g = self.geo(ws)
        old = []
        if dr is not None:
            for an in dr.anchors():
                if an.otype in ('pic', 'grp'):
                    r, c = an.top_left(g)
                    if r <= 14 and cN + 1 <= c <= cN + 10:
                        old.append(an)
        if not files:
            self.log('견본주택 위치도: 넣을 그림이 없어 ' + ('양식 그림을 그대로 둡니다' if old else 'O6 를 비워 둡니다'))
            return
        w, h = 300.0, 240.0
        if old:
            _, _, bw, bh = max((an.box(g) for an in old), key=lambda b: b[2] * b[3])
            w, h = max(bw / EMU_PT, 200.0), max(bh / EMU_PT, 150.0)
            for an in old:
                dr.remove(an)
            dr.save()
        self.place_images(ws, files[:1], g.x_of(cN + 1) / EMU_PT, g.y_of(6) / EMU_PT, w, h, margin=0, align='topleft')

    # -------------------------------------------------------------- 보고자료 시트
    def report_sheet(self, d):
        """보고자료 시트 (2026-10-18 사용자 지정 양식): 1. 업무일지(오픈동향) / 2. 오픈보고 / 3. 주변시세"""
        old = next((s['name'] for s in self.book.sheet_list() if ns(s['name']) == '보고자료'), None)
        if old:
            self.book.delete_sheet(old)
        cells = report_cells(d)
        nr = max(r for r, _, _, _ in cells)
        rows = [[''] * 11 for _ in range(nr)]
        st = self.styles
        styles = {}
        for r, c, v, n in cells:
            rows[r - 1][c - 1] = v
            if n:
                styles[(r, c)] = st.derive(0, numfmt=n, border='thin')
        lay = report_layout(d)
        for r in lay['titles']:
            styles[(r, 2)] = st.derive(0, bold=True)
        hr, last = lay['table']
        head = st.derive(0, bold=True, fill='EFEFEF', border='thin', align={'horizontal': 'center'})
        cell = st.derive(0, border='thin')
        for c in range(2, 12):
            styles[(hr, c)] = head
            for r in range(hr + 1, last + 1):
                styles.setdefault((r, c), cell)
        widths = {(X.col2num(k), X.col2num(k)): w for k, w in REPORT_WIDTHS.items()}
        add_sheet(self.book, '보고자료', rows, hidden=False, widths=widths, styles=styles)

    def cross(self, ws, area, on):
        """기본형 평면이 없으면 칸에 X 표 (양식 관례)"""
        try:
            r1, c1, r2, c2 = area
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    ref = cref(r, c)
                    if ref not in ws.cells and (r, c) != (r1, c1):
                        continue
                    s = ws.style(ref) or 0
                    if on or self.styles.has_diag(s):
                        ws.set_style(ref, self.styles.derive(s, border='diag' if on else 'nodiag'))
            ws.book.touched.add(ws.path)
        except Exception as e:
            self.log(f'X 표 지정 실패: {e}')

    def set_pages(self, ws):
        rows = self.section_rows(ws)
        if getattr(self, 'units_n', None):
            last = self.units_first + self.units_n * self.units_size - 1
        else:
            last = ws.max_row()
        try:
            set_print_area(self.book, ws.name, f'$A$1:${num2col(self.cols["N"])}${last}')
        except Exception as e:
            self.log(f'인쇄영역 설정 실패: {e}')
        try:
            brk = [rows[k] for k in ('location', 'birdseye', 'siteplan', 'community', 'units') if rows.get(k)]
            if getattr(self, 'units_n', None):
                for k in range(5, self.units_n, 5):
                    brk.append(self.units_first + k * self.units_size)
            set_row_breaks(ws, brk)
        except Exception as e:
            self.log(f'쪽 나누기 설정 실패: {e}')

    # -------------------------------------------------------------- 공고자료 시트
    def data_sheet(self, d):
        if self.sheet('공고자료') is not None:
            self.book.delete_sheet(next(s['name'] for s in self.book.sheet_list() if ns(s['name']) == '공고자료'))
        rows = []
        rows.append(['분양개요 자동작성 근거 (입주자모집공고 원문 값)'])
        rows.append([f"공고 파일 : {d.get('notice_file', '')}", '', '', '', '', '', f"작성 : {d['written']}"])
        rows.append([])
        rows.append(['1. 공급대상 (단위 : ㎡, 세대)'])
        rows.append(['타입', '주택형', '주거전용', '주거공용', '공급면적', '기타공용', '계약면적', '대지지분',
                     '총공급세대', '특별공급 계', '일반공급', '공고 쪽'])
        t0 = len(rows) + 1
        for t in d['types']:
            sp = t.get('special') or {}
            rows.append([t['abbr'], t.get('code', ''), t['excl'], t.get('pub'), t['supply'], t.get('etc'),
                         t.get('contract'), t.get('land'), t['units'], sp.get('total', ''), t.get('general'),
                         t.get('page', '')])
        t1 = len(rows)
        rows.append(['합계', '', '', '', '', '', '', '', f'=SUM(I{t0}:I{t1})', f'=SUM(J{t0}:J{t1})', f'=SUM(K{t0}:K{t1})'])
        self.data_special = {t['abbr']: f'공고자료!$J${t0 + i}' for i, t in enumerate(d['types'])}
        rows.append([])
        rows.append(['2. 층별 공급금액 (단위 : 원) — ◀ 표시 = 분양개요에 쓴 기준층(세대수가 가장 많은 층)'])
        rows.append(['타입', '층 구분', '세대수', '대지비', '건축비', '부가세', '공급금액 계', '세대수×금액', '공급면적(㎡)',
                     '세대수×면적', '기준층', '공고 쪽'])
        p0 = len(rows) + 1
        for t in d['types']:
            sup = t['supply']
            for k, x in enumerate(t.get('tiers') or []):
                r = len(rows) + 1
                rows.append([t['abbr'], x['floor'], x['units'], x.get('land'), x.get('build'), x.get('vat'), x['total'],
                             f'=C{r}*G{r}', sup, f'=C{r}*I{r}', '◀ 기준층' if k == t.get('base') else '', x.get('page', '')])
        p1 = len(rows)
        rows.append([])
        rows.append(['3. 평균평당가 = 전체 분양가 합 ÷ 전체 공급면적(평)'])
        a = len(rows) + 1
        rows.append(['전체 분양가 합 (원)', f'=SUM(H{p0}:H{p1})'])
        rows.append(['전체 공급면적 (㎡)', f'=SUM(J{p0}:J{p1})'])
        rows.append(['전체 공급면적 (평)', f'=B{a + 1}*0.3025'])
        rows.append(['평균평당가 (원/평)', f'=B{a}/B{a + 2}'])
        rows.append(['평균평당가 (천원/평, 버림)', f'=TRUNC(B{a + 3}/1000)'])
        rows.append([])
        rows.append(['4. 발코니 확장 금액 (원)'])
        for t in d['types']:
            rows.append([t['abbr'], t.get('balcony') or 0])
        rows.append([])
        rows.append(['5. 유상옵션 (원) — 비고 포함 품목은 분양개요 Q열부터 같은 내용(천원), 유상옵션(C) = 그 합계'])
        groups = {g['id']: g for g in (d.get('options') or [])}
        rows.append(['품목', '비고 포함'] + [t['abbr'] for t in d['types']] + ['공고 쪽'])
        o0 = len(rows) + 1
        for c in d.get('option_cols') or []:
            g = groups.get(c.get('group')) or {}
            line = [c.get('name') or c.get('header'), '포함' if g.get('include') else '제외']
            line += [(c.get('values') or {}).get(t['abbr']) or '' for t in d['types']]
            line.append(c.get('page', ''))
            rows.append(line)
        o1 = len(rows)
        rows.append([])
        rows.append(['  ↳ 품목별 원문 (대안·적용 타입)'])
        for g in d.get('options') or []:
            for v in g.get('variants') or []:
                rows.append([g['name'], v.get('label', ''), ', '.join(v.get('types') or [])])
        rows.append([])
        rows.append(['6. 주차대수 (직접 입력)'])
        rows.append(['주차대수', d.get('parking') or '', '세대수', d.get('households') or ''])
        st = self.styles
        styles = {(1, 1): st.derive(0, bold=True, size=13)}
        num = st.derive(0, numfmt='#,##0')
        for r in range(p0, p1 + 1):
            for c in range(4, 9):
                styles[(r, c)] = num
        for r in range(a, a + 5):
            styles[(r, 2)] = num
        if o1 >= o0:
            for r in range(o0, o1 + 1):
                for c in range(3, 3 + len(d['types'])):
                    styles[(r, c)] = num
        add_sheet(self.book, '공고자료', rows, hidden=True, widths={(1, 1): 26, (2, 26): 13}, styles=styles)
        self.data_avg_cell = f'공고자료!$B${a + 4}'

    # -------------------------------------------------------------- 청약결과
    def fill_result(self, d, r0):
        ws = self.sheet('청약결과')
        if ws is None:
            self.log('청약결과 시트가 없어 건너뜁니다')
            return
        hdr = self.find_row(ws, r'^구분$', 2, 1, 15)
        if not hdr:
            self.log('청약결과 머리글을 찾지 못해 건너뜁니다')
            return
        first = None
        for rr in range(hdr + 1, hdr + 6):
            ref = cref(rr, 3)
            if ws.formula(ref) is not None or isinstance(ws.value(ref), (int, float)):
                first = rr
                break
        if first is None:
            first = hdr + 3
        sum_row = self.find_row(ws, r'^소계$', 3, first, first + 80)
        if not sum_row:
            self.log('청약결과 소계 줄을 찾지 못해 건너뜁니다')
            return
        n = len(d['types'])
        nold = sum_row - first
        unmerge_range(ws, *merge_area(ws, first, 2))
        unmerge_range(ws, first, 1, sum_row - 1, 27)
        delta = resize_blocks(ws, first, nold, n, 1)
        sum_row += delta
        clear_range(ws, first, 2, sum_row - 1, 27)
        merge_range(ws, first, 2, sum_row, 2)
        src = self.ws.name
        C = self.cols
        L = num2col
        for i, t in enumerate(d['types']):
            R = first + i
            r = r0 + 2 * i
            f = {
                'C': f'=TRUNC({src}!{L(C["C"])}{r},1)',
                'D': f'={src}!{L(C["C"] + 1)}{r}' if t.get('suffix') else '',
                'E': f'={src}!{L(C["E"])}{r}',
                'F': f'=E{R}/3.3058',
                'G': f'={src}!{L(C["I"])}{r}',
                'H': f'=G{R}/F{R}',
                'I': f'={src}!{L(C["H"])}{r}',
                'J': f'={self.data_special.get(t["abbr"], 0)}' if getattr(self, 'data_special', None) else '',
                'L': f'=IFERROR(K{R}/J{R},0)',
                'M': f'=+IF(K{R}>J{R},I{R}-J{R},I{R}-K{R})',
                'O': f'=IFERROR(N{R}/M{R},0)',
                'R': f'=N{R}+P{R}+Q{R}',
                'S': f'=IFERROR(R{R}/M{R},0)',
                'T': f'=+IF(R{R}>M{R},0,M{R}-R{R})',
                'X': f'=SUM(U{R}:W{R})',
                'Y': f'=IFERROR((R{R}+X{R})/M{R},0)',
            }
            for c in range(3, 26):
                put(ws, R, c, f.get(L(c), ''))
        put(ws, first, 2, f'={src}!{num2col(1 + self.dc)}{r0}')
        S = sum_row
        a, b = first, sum_row - 1
        fs = {'I': f'=SUM(I{a}:I{b})', 'J': f'=SUM(J{a}:J{b})', 'K': f'=SUM(K{a}:K{b})', 'L': f'=IFERROR(K{S}/J{S},0)',
              'M': f'=SUM(M{a}:M{b})', 'N': f'=SUM(N{a}:N{b})', 'O': f'=IFERROR(N{S}/M{S},0)', 'P': f'=SUM(P{a}:P{b})',
              'Q': f'=SUM(Q{a}:Q{b})', 'R': f'=N{S}+P{S}+Q{S}', 'S': f'=IFERROR(R{S}/M{S},0)', 'T': f'=SUM(T{a}:T{b})',
              'U': f'=SUM(U{a}:U{b})', 'V': f'=SUM(V{a}:V{b})', 'W': f'=SUM(W{a}:W{b})', 'X': f'=SUM(U{S}:W{S})',
              'Y': f'=IFERROR((R{S}+X{S})/M{S},0)'}
        for col, fx in fs.items():
            put(ws, S, col2num(col), fx)
        r2 = self.find_row(ws, r'^\(?작성일', 2, 1, 4)
        if r2:
            put(ws, r2, 2, '( 작성일 :         년     월     일 )')
        r3 = self.find_row(ws, r'예비', 2, S + 1, S + 4)
        if r3 and d.get('reserve_pct'):
            put(ws, r3, 2, f"※ 예비당첨자(예비입주자) {d['reserve_pct']}%")
        try:
            pa = get_print_area(self.book, ws.name) or ''
            m = re.search(r'\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$', pa)
            if m:
                set_print_area(self.book, ws.name, f'${m.group(1)}${m.group(2)}:${m.group(3)}${S + 1}')
        except Exception:
            pass

    # -------------------------------------------------------------- 이전 현장 숨김 시트
    def delete_sheets(self, names):
        # 공고자료·보고자료는 이번에 새로 만든 시트 — 양식에서 숨겨져 있던 같은 이름이라도 지우지 않는다
        # (2026-10-23 M1-2-2BL 양식: 숨긴 보고자료 때문에 새 보고자료가 지워졌음)
        names = [n for n in names if ns(n) not in self.NEW_SHEETS]
        if not names:
            return
        pats = []
        for n in names:
            pats.append(n + '!')
            pats.append("'" + n.replace("'", "''") + "'!")
        for s in self.book.all_sheets():
            if s.name in names or ns(s.name) in self.NEW_SHEETS:
                continue
            for ref, c in list(s.cells.items()):
                f = c.find(q('f'))
                if f is None or not f.text:
                    continue
                if any(p in f.text for p in pats):
                    c.remove(f)          # 값(<v>)만 남긴다 (Excel 의 c.Value = c.Value)
                    s.book.touched.add(s.path)
        for n in names:
            try:
                self.book.delete_sheet(n)
                self.log(f'이전 현장 숨김 시트 "{n}" 삭제')
            except Exception as e:
                self.log(f'숨김 시트 "{n}" 삭제 실패: {e}')


def build_workbook(template_bytes, d, images, log=None):
    """images = {'location': [img…], 'birdseye': […], 'siteplan': […], 'community': […],
                 'units': {타입: {'basic': img, 'ext': img}}},  img = {'data': bytes, 'w': 폭, 'h': 높이}
    결과 xlsx bytes 와 걸린 초를 돌려준다."""
    t = time.time()
    b = Builder(template_bytes, log)
    b.build(d, images)
    out = b.book.save()
    return out, time.time() - t
