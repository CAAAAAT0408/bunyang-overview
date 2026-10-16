"""모집공고 PDF 읽기 도구.

- 줄 텍스트: 한글 PDF 중에는 숫자·괄호가 다른 글꼴 조각으로 엉뚱한 순서에 들어 있는 것이 있어
  (예: '지하 층지상 최고 2 ~ 25층'), 글자 하나하나의 위치로 줄을 다시 세운다.
- 표: PyMuPDF find_tables. 칸 안 줄바꿈은 '\n' 으로 남는다.
"""
import re

import pymupdf

try:
    pymupdf.TOOLS.mupdf_display_errors(False)
    pymupdf.TOOLS.mupdf_display_warnings(False)
except Exception:
    pass


class Line:
    __slots__ = ('page', 'y', 'x0', 'x1', 'text', 'size')

    def __init__(self, page, y, x0, x1, text, size):
        self.page, self.y, self.x0, self.x1, self.text, self.size = page, y, x0, x1, text, size

    def __repr__(self):
        return f'<p{self.page + 1} y{self.y:.0f} {self.text[:40]!r}>'


class Table:
    __slots__ = ('page', 'bbox', 'rows', 'title', 'index', 'cells')

    def __init__(self, page, bbox, rows, title, index, cells=None):
        self.page, self.bbox, self.rows, self.title, self.index = page, bbox, rows, title, index
        self.cells = cells or []           # 줄마다 칸 위치 [(x0,y0,x1,y1) | None, …]

    @property
    def ncols(self):
        return max((len(r) for r in self.rows), default=0)

    def flat(self, nrows=None):
        rows = self.rows if nrows is None else self.rows[:nrows]
        return ' '.join(c for r in rows for c in r if c)

    def __repr__(self):
        return f'<Table p{self.page + 1} {len(self.rows)}x{self.ncols} {self.title[:30]!r}>'


def _page_lines(page, pno, ytol=2.6):
    chars = []
    d = page.get_text('rawdict', flags=pymupdf.TEXT_PRESERVE_WHITESPACE | pymupdf.TEXT_MEDIABOX_CLIP)
    for b in d['blocks']:
        for ln in b.get('lines', []):
            if abs(ln['dir'][0]) < 0.9:          # 세로 글자 제외
                continue
            for s in ln['spans']:
                sy = (s['bbox'][1] + s['bbox'][3]) / 2
                for ch in s['chars']:
                    x0, _, x1, _ = ch['bbox']
                    chars.append((sy, x0, x1, ch['c'], s['size']))
    chars.sort(key=lambda c: c[0])
    groups = []
    for c in chars:
        if groups and abs(groups[-1][0] - c[0]) <= ytol:
            groups[-1][1].append(c)
        else:
            groups.append([c[0], [c]])
    out = []
    for y, cs in groups:
        cs.sort(key=lambda c: c[1])
        dedup = []
        for c in cs:                    # 그림자·굵게 효과로 같은 글자를 거의 같은 자리에 여러 번 찍은 PDF
            if dedup and dedup[-1][3] == c[3] and abs(dedup[-1][1] - c[1]) < c[4] * 0.3:
                continue
            dedup.append(c)
        cs = dedup
        txt, lastx, size = '', None, 0
        for c in cs:
            size = max(size, c[4])
            if c[3] in (' ', '　', '\xa0'):
                if txt and not txt.endswith(' '):
                    txt += ' '
                lastx = c[2]
                continue
            if lastx is not None and c[1] - lastx > c[4] * 0.32 and txt and not txt.endswith(' '):
                txt += ' '
            txt += c[3]
            lastx = c[2]
        txt = txt.strip()
        if txt:
            out.append(Line(pno, y, cs[0][1], cs[-1][2], txt, size))
    return out


def clean_cell(v):
    if v is None:
        return ''
    v = str(v).replace('\r', '').replace('　', ' ').replace('\xa0', ' ')
    v = '\n'.join(re.sub(r'[ \t]+', ' ', s).strip() for s in v.split('\n'))
    return v.strip('\n ').strip()


class PdfDoc:
    def __init__(self, data):
        self.doc = pymupdf.open(stream=data, filetype='pdf')
        self.npages = self.doc.page_count
        self._lines = {}
        self._tables = {}

    def lines(self, pno):
        if pno not in self._lines:
            self._lines[pno] = _page_lines(self.doc[pno], pno)
        return self._lines[pno]

    def all_lines(self):
        out = []
        for p in range(self.npages):
            out.extend(self.lines(p))
        return out

    def words(self, pno):
        """[(x0, y0, x1, y1, 글자)] — 칸 위치로 글자를 찾을 때"""
        if not hasattr(self, '_words'):
            self._words = {}
        if pno not in self._words:
            self._words[pno] = [w[:5] for w in self.doc[pno].get_text('words')]
        return self._words[pno]

    def page_text(self, pno):
        return '\n'.join(l.text for l in self.lines(pno))

    def tables(self, pno):
        if pno not in self._tables:
            page = self.doc[pno]
            res = []
            try:
                found = page.find_tables()
                tabs = found.tables
            except Exception:
                tabs = []
            lines = self.lines(pno)
            for i, t in enumerate(tabs):
                try:
                    rows = [[clean_cell(c) for c in r] for r in t.extract()]
                except Exception:
                    continue
                if not rows:
                    continue
                x0, y0, x1, y1 = t.bbox
                # 표 바로 위 (최대 40pt) 의 줄을 제목으로
                above = [l for l in lines if y0 - 40 <= l.y < y0 - 1 and l.x1 > x0 and l.x0 < x1]
                title = ' '.join(l.text for l in above[-2:])
                cells = []
                try:
                    for tr in t.rows:
                        cells.append([tuple(c) if c else None for c in tr.cells])
                except Exception:
                    cells = []
                res.append(Table(pno, tuple(t.bbox), rows, title, i, cells if len(cells) == len(rows) else []))
            self._tables[pno] = res
        return self._tables[pno]

    def render_clip(self, pno, rect, zoom=2.0):
        page = self.doc[pno]
        pix = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), clip=pymupdf.Rect(rect))
        return pix.tobytes('png')

    def render_page(self, pno, zoom=1.5):
        return self.doc[pno].get_pixmap(matrix=pymupdf.Matrix(zoom, zoom)).tobytes('png')
