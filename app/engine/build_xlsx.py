"""분양개요 양식(.xlsx) 에 값을 채워 새 파일을 만든다 — Excel COM.

양식은 이전 현장의 '분양개요 및 청약결과' 파일이면 된다. 줄 위치는 글자로 찾는다.
  분양개요 시트: 1. 일반개요 / 2. 사업일정 / 3. 공급개요(타입당 2줄 + 소계 + ※ 3줄) /
                4~7 그림 구역(테두리 사각형 안) / 8. 단위세대(타입당 6줄, 기본형 D:J · 확장형 K:N)
  청약결과 시트: 타입당 1줄 + 소계. 수식만 넣고 입력값은 비운다.
  공고자료 시트(숨김, 새로 만듦): 공급표·층별 공급금액·평균평당가 계산 근거·확장비·유상옵션 선택 내역.
"""
import os
import re
import time
from datetime import date

import pythoncom
import win32com.client as win32
from win32com.client import constants as _c  # noqa: F401  (상수는 숫자로 직접 씀)

XL_DOWN = -4121
XL_UP = -4162
XL_CENTER = -4108
XL_LEFT = -4131
XL_PAGEBREAK_MANUAL = -4135
MSO_PICTURE = 13
MSO_LINKED_PICTURE = 11
MSO_AUTOSHAPE = 1
RED = 255


class BuildError(Exception):
    pass


def col_letter(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def ns(v):
    """비교용: 공백 제거한 문자열"""
    return re.sub(r'\s', '', str(v)) if v is not None else ''


def excel_date(iso):
    if not iso:
        return None
    y, m, d = map(int, iso[:10].split('-'))
    return (date(y, m, d) - date(1899, 12, 30)).days


class Builder:
    def __init__(self, path, log=None):
        self.path = os.path.abspath(path)
        self.log = log or (lambda *a: None)
        self.xl = None
        self.wb = None

    # -------------------------------------------------------------- Excel 열고 닫기
    def open(self):
        pythoncom.CoInitialize()
        self.xl = win32.DispatchEx('Excel.Application')
        self.xl.Visible = False
        self.xl.DisplayAlerts = False
        self.xl.AskToUpdateLinks = False
        self.xl.ScreenUpdating = False
        self.xl.EnableEvents = False
        try:
            self.xl.AutomationSecurity = 3          # 매크로 끔
        except Exception:
            pass
        self.wb = self.xl.Workbooks.Open(self.path, 0, False)
        try:
            self.xl.Calculation = -4135             # 수동
        except Exception:
            pass

    def close(self, save=True):
        try:
            if self.wb is not None:
                try:
                    self.xl.Calculation = -4105     # 자동
                except Exception:
                    pass
                if save:
                    self.wb.Save()
                self.wb.Close(False)
        finally:
            self.wb = None
            if self.xl is not None:
                try:
                    self.xl.Quit()
                except Exception:
                    pass
            self.xl = None
            pythoncom.CoUninitialize()

    # -------------------------------------------------------------- 찾기
    def sheet(self, *names):
        for ws in self.wb.Worksheets:
            if ns(ws.Name) in [ns(n) for n in names]:
                return ws
        return None

    @staticmethod
    def grid(ws, r1, c1, r2, c2):
        v = ws.Range(ws.Cells(r1, c1), ws.Cells(r2, c2)).Value
        if not isinstance(v, tuple):
            v = ((v,),)
        return v

    def find_row(self, ws, pattern, col=1, start=1, end=400, cols=1):
        rx = re.compile(pattern)
        g = self.grid(ws, start, col, end, col + cols - 1)
        for i, row in enumerate(g):
            for v in row:
                if v is not None and rx.search(ns(v)):
                    return start + i
        return None

    def find_cell(self, ws, pattern, r1, c1, r2, c2):
        rx = re.compile(pattern)
        g = self.grid(ws, r1, c1, r2, c2)
        for i, row in enumerate(g):
            for j, v in enumerate(row):
                if v is not None and rx.search(ns(v)):
                    return r1 + i, c1 + j
        return None

    # -------------------------------------------------------------- 그림
    def pictures_in(self, ws, r1, r2, max_col=None, min_col=None):
        out = []
        for sh in ws.Shapes:
            try:
                if sh.Type not in (MSO_PICTURE, MSO_LINKED_PICTURE) and not (sh.Type == 14 or sh.Type == 6):
                    continue
                tl = sh.TopLeftCell
                if not (r1 <= tl.Row <= r2):
                    continue
                if max_col and tl.Column > max_col:
                    continue
                if min_col and tl.Column < min_col:
                    continue
                if sh.Type == 6:                      # 그룹: 안에 그림이 있으면 지움
                    pass
                out.append(sh)
            except Exception:
                pass
        return out

    def frames_in(self, ws, r1, r2):
        out = []
        for sh in ws.Shapes:
            try:
                if sh.Type == MSO_AUTOSHAPE and r1 <= sh.TopLeftCell.Row <= r2 and sh.Width > 200 and sh.Height > 100:
                    out.append(sh)
            except Exception:
                pass
        return out

    def place_images(self, ws, files, left, top, width, height, margin=6, gap=6):
        files = [f for f in files if f and os.path.exists(f)]
        if not files:
            return []
        from PIL import Image
        sizes = []
        for f in files:
            with Image.open(f) as im:
                sizes.append(im.size)
        L, T = left + margin, top + margin
        W, H = width - 2 * margin, height - 2 * margin - gap * (len(files) - 1)
        # 너비를 W 로 맞췄을 때의 높이 합이 H 를 넘으면 함께 줄인다
        hs = [W * h / w for (w, h) in sizes]
        scale = min(1.0, H / sum(hs)) if sum(hs) else 1.0
        y = T + max(0, (H - sum(hs) * scale) / 2)
        shapes = []
        for f, (w, h), hh in zip(files, sizes, hs):
            ww, hh = W * scale, hh * scale
            x = L + (W - ww) / 2
            sh = ws.Shapes.AddPicture(f, False, True, x, y, ww, hh)
            try:
                sh.Placement = 2                     # 셀에 따라 이동(크기 고정)
            except Exception:
                pass
            shapes.append(sh)
            y += hh + gap
        return shapes

    # -------------------------------------------------------------- 줄 넣고 빼기
    def big_merges(self, ws, r1, r2, c1, c2):
        """r1~r2 안에서 2줄 넘게 걸친 병합(구분·비고 칸)을 풀고 위치를 기억"""
        found = {}
        for c in range(c1, c2 + 1):
            for r in (r1, r1 + 1):
                ma = ws.Cells(r, c).MergeArea
                if ma.Rows.Count > 2 and ma.Row <= r1 + 1:
                    key = (ma.Row, ma.Column)
                    if key not in found:
                        found[key] = {'row': ma.Row, 'col': ma.Column, 'ncols': ma.Columns.Count,
                                      'last': ma.Row + ma.Rows.Count - 1,
                                      'value': ws.Cells(ma.Row, ma.Column).Value}
        for m in found.values():
            ws.Range(ws.Cells(m['row'], m['col']), ws.Cells(m['last'], m['col'] + m['ncols'] - 1)).UnMerge()
        return list(found.values())

    def resize_blocks(self, ws, first, nold, nnew, size):
        """first 줄부터 size 줄씩 nold 개 있는 묶음을 nnew 개로. 마지막 묶음을 복사해 늘린다."""
        if nnew > nold:
            last = first + (nold - 1) * size
            for _ in range(nnew - nold):
                ws.Rows(f'{last}:{last + size - 1}').Copy()
                ws.Rows(f'{last + size}').Insert(XL_DOWN)
                last += size
            try:
                self.xl.CutCopyMode = 0
            except Exception:
                pass
        elif nnew < nold:
            a = first + nnew * size
            b = first + nold * size - 1
            ws.Rows(f'{a}:{b}').Delete(XL_UP)
        return (nnew - nold) * size

    # -------------------------------------------------------------- 메인
    def build(self, d, images):
        ws = self.sheet('분양개요') or self.wb.Worksheets(1)
        self.ws = ws
        t0 = time.time()
        def step(name, fn, *a):
            t = time.time()
            r = fn(*a)
            self.log(f'{name} ({time.time() - t:.1f}초)')
            return r
        old_hidden = [s.Name for s in self.wb.Worksheets if s.Visible != -1]
        step('일반개요·사업일정', self.fill_top, ws, d)
        r0, sum_row = step('공급개요', self.fill_supply, ws, d)
        step('위치도·조감도·배치도·커뮤니티 그림', self.fill_images, ws, d, images, sum_row)
        step('단위세대', self.fill_units, ws, d, images)
        step('인쇄영역·쪽 나누기', self.set_pages, ws)
        step('공고자료 시트', self.data_sheet, d)
        step('청약결과 시트', self.fill_result, d, r0)
        if d.get('delete_hidden', True):
            step('이전 현장 숨김 시트 정리', self.delete_sheets, old_hidden)
        try:
            self.xl.Calculation = -4105
            self.wb.Application.CalculateFull()
        except Exception:
            pass
        for s in self.wb.Worksheets:
            if s.Visible == -1:
                try:
                    s.Activate()
                    s.Range('A1').Select()
                    self.xl.ActiveWindow.ScrollRow = 1
                    self.xl.ActiveWindow.ScrollColumn = 1
                except Exception:
                    pass
        ws.Activate()

    # -------------------------------------------------------------- 1·2. 일반개요 · 사업일정
    def fill_top(self, ws, d):
        ws.Range('A1').Value = d['title']
        r = self.find_row(ws, r'^\(?작성일', 1, 1, 6)
        if r:
            ws.Cells(r, 1).Value = f"(작성일 : {d['written']})"
        # 홈페이지 주소 (인쇄영역 밖)
        rc = self.find_cell(ws, r'^https?://|^www\.', 1, 15, 8, 30)
        if rc:
            ws.Cells(*rc).Value = d.get('homepage') or ''
        elif d.get('homepage'):
            ws.Range('P5').Value = d['homepage']
        labels = [('공급위치', 'location'), ('단지규모', 'scale'), ('단지특징', 'features'),
                  ('주차대수', 'parking_text'), ('위탁', 'developers'), ('시행', 'developers'), ('견본주택', 'model_house'),
                  ('오픈일', 'mh_open'), ('모집공고', 'notice_date'), ('청약일정', 'subscription'),
                  ('당첨자', 'winners'), ('계약일정', 'contract'), ('입주예정', 'move_in')]
        g = self.grid(ws, 5, 2, 20, 2)
        gcol = 7
        for i, row in enumerate(g):
            lab = ns(row[0])
            if not lab:
                continue
            for key, field in labels:
                if key in lab:
                    v = d.get(field)
                    cell = ws.Cells(5 + i, gcol)
                    if field in ('mh_open', 'notice_date', 'winners'):
                        ser = excel_date(v)
                        cell.Value = ser if ser else ''
                    else:
                        cell.Value = v or ''
                    break
        # 세대당 주차 보조칸
        rc = self.find_cell(ws, r'^세대수$', 12, 15, 20, 30)
        if rc:
            ws.Cells(rc[0], rc[1] + 1).Value = d.get('households') or ''
            ws.Cells(rc[0] + 1, rc[1] + 1).Value = d.get('parking') or ''
            a, b = col_letter(rc[1] + 1) + str(rc[0] + 1), col_letter(rc[1] + 1) + str(rc[0])
            ws.Cells(rc[0] + 2, rc[1] + 1).Formula = f'=IFERROR(TRUNC({a}/{b},2),"")'

    # -------------------------------------------------------------- 3. 공급개요
    def fill_supply(self, ws, d):
        sec = self.find_row(ws, r'^3\.공급개요', 1, 15, 60)
        hdr = self.find_row(ws, r'^구분$', 1, (sec or 18) + 1, (sec or 18) + 5)
        if not hdr:
            raise BuildError('양식에서 "3. 공급개요" 표의 "구  분" 머리글을 찾지 못했습니다.')
        r0 = hdr + 2
        sum_row = self.find_row(ws, r'^소계$', 3, r0, r0 + 120)
        if not sum_row:
            raise BuildError('양식의 공급개요 표에서 "소 계" 줄을 찾지 못했습니다.')
        # 머리글로 칸 찾기
        hg = self.grid(ws, hdr, 1, hdr + 1, 14)
        def hcol(pat, default):
            rx = re.compile(pat)
            for rr in hg:
                for j, v in enumerate(rr):
                    if v is not None and rx.search(ns(v)):
                        return j + 1
            return default
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
        L = col_letter
        types = d['types']
        n = len(types)
        nold = (sum_row - r0) // 2
        # 인쇄영역 밖 보조칸·그림 정리 (이전 현장 옵션표 캡처 등)
        for sh in self.pictures_in(ws, 1, sum_row + 12, min_col=cN + 1):
            sh.Delete()
        merges = self.big_merges(ws, r0, sum_row, 1, cN)
        self.log(f'공급개요 {nold}개 타입 → {n}개')
        delta = self.resize_blocks(ws, r0, nold, n, 2)
        sum_row += delta
        for m in merges:
            last = m['last'] + delta
            rng = ws.Range(ws.Cells(m['row'], m['col']), ws.Cells(last, m['col'] + m['ncols'] - 1))
            rng.Merge()
        # 유상옵션 내역 칸 (T 열부터, 2026-10-15 사용자 지정: 공고의 유상옵션 전체를 타입별 한 줄로, 합계(L)는 사용자가 직접)
        ocols = d.get('option_cols') or []
        oc0 = cN + 6                                  # T 열
        # 보조칸 비우기 (O 열부터)
        far = max(60, oc0 + len(ocols) + 2)
        ws.Range(ws.Cells(hdr, cN + 1), ws.Cells(sum_row + 1, far)).ClearContents()
        if ocols:
            ws.Range(ws.Cells(hdr, oc0), ws.Cells(sum_row + 1, far)).UnMerge()
            ws.Cells(hdr, oc0).Value = '유상옵션 (공고 원문, 단위: 천원) — 합계는 유상옵션(C) 칸에 직접 입력'
            ws.Range(ws.Cells(hdr + 1, oc0), ws.Cells(hdr + 1, oc0 + len(ocols) - 1)).Value = \
                [[c.get('header') or c.get('name') or '' for c in ocols]]
        # 이 양식들은 셀 한 칸 쓸 때마다 Excel 이 0.1초씩 걸려서, 타입 줄 전체(C~M)를 한 번에 쓴다.
        # 병합(C24:C25, E24:F25 …)은 범위 안에 통째로 들어가므로 괜찮다.
        width = cM - cC + 1
        block = []
        for i, t in enumerate(types):
            r = r0 + 2 * i
            r1 = r + 1
            a = [''] * width
            b = [''] * width
            def put(row, col, v):
                row[col - cC] = v
            put(a, cC, t['excl'])
            if cC + 1 < cE:
                put(a, cC + 1, t.get('suffix') or '')
            put(a, cE, t['supply'])
            put(a, cG, f'={L(cE)}{r}*0.3025')
            put(a, cH, t['units'])
            put(a, cI, round((t.get('price') or 0) / 1000, 1))
            put(a, cJ, round((t.get('balcony') or 0) / 1000, 1))
            put(a, cK, f'=+{L(cI)}{r}+{L(cJ)}{r}')
            put(a, cL, '')                                  # 유상옵션(C) 합계 — 사용자가 직접
            put(a, cM, f'=+{L(cL)}{r}+{L(cK)}{r}')
            put(b, cI, f'=+{L(cI)}{r}/{L(cG)}{r}')
            put(b, cJ, f'=+{L(cJ)}{r}/{L(cG)}{r}')
            put(b, cK, f'=+{L(cI)}{r1}+{L(cJ)}{r1}')
            put(b, cL, f'=+{L(cL)}{r}/{L(cG)}{r}')
            put(b, cM, f'=+{L(cK)}{r1}+{L(cL)}{r1}')
            block += [a, b]
        if block:
            ws.Range(ws.Cells(r0, cC), ws.Cells(r0 + len(block) - 1, cM)).Formula = block
        if ocols and types:
            ob = []
            for t in types:
                line = []
                for c in ocols:
                    v = (c.get('values') or {}).get(t['abbr'])
                    line.append(round(v / 1000, 1) if v else '')
                ob.append(line)
                ob.append([''] * len(ocols))
            rng = ws.Range(ws.Cells(r0, oc0), ws.Cells(r0 + len(ob) - 1, oc0 + len(ocols) - 1))
            rng.Value = ob
            self.option_block_format(ws, hdr, oc0, r0, r0 + len(ob) - 1, len(ocols))
        ws.Cells(r0, 1).Value = d.get('label') or ''
        ws.Cells(r0, cN).Value = d.get('remarks') or ''
        ws.Cells(sum_row, cH).Formula = f'=SUM({L(cH)}{r0}:{L(cH)}{sum_row - 1})'
        ws.Cells(sum_row, cI).Value = d.get('basis_note') or ''
        # ※ 3줄
        for rr in range(sum_row + 1, sum_row + 6):
            v = ns(ws.Cells(rr, 1).Value)
            if '평균평당가' in v:
                self.rich(ws.Cells(rr, 1), d['avg_note'], r'\d')
            elif '납부조건' in v:
                self.rich(ws.Cells(rr, 1), d['payment_note'], r'중도금')
            elif (v.startswith('※유상옵션') or '유상옵션:' in v) and d.get('option_note'):
                ws.Cells(rr, 1).Value = d['option_note']
        self.cols = dict(C=cC, E=cE, G=cG, H=cH, I=cI, J=cJ, K=cK, L=cL, M=cM, N=cN)
        self.r0, self.sum_row = r0, sum_row
        return r0, sum_row

    def option_block_format(self, ws, hdr, c0, r1, r2, n):
        """T열 옵션표 모양: 머리글 줄바꿈·회색, 금액 천단위, 얇은 테두리 (인쇄영역 밖)"""
        try:
            c1 = c0 + n - 1
            head = ws.Range(ws.Cells(hdr + 1, c0), ws.Cells(hdr + 1, c1))
            head.WrapText = True
            head.HorizontalAlignment = -4108
            head.VerticalAlignment = -4108
            head.Interior.Color = 0xEFEFEF
            head.Font.Bold = True
            ws.Cells(hdr, c0).Font.Bold = True
            ws.Cells(hdr, c0).WrapText = False
            body = ws.Range(ws.Cells(r1, c0), ws.Cells(r2, c1))
            body.NumberFormat = '#,##0;-#,##0;;@'
            body.HorizontalAlignment = -4152
            whole = ws.Range(ws.Cells(hdr + 1, c0), ws.Cells(r2, c1))
            whole.Interior.Pattern = -4142
            head.Interior.Color = 0xEFEFEF
            for b in (7, 8, 9, 10, 11, 12):
                whole.Borders(b).LineStyle = 1
                whole.Borders(b).Weight = 1
            ws.Range(ws.Cells(1, c0), ws.Cells(1, c1)).EntireColumn.ColumnWidth = 9.5
        except Exception as e:
            self.log(f'유상옵션 칸 서식 실패: {e}')

    def rich(self, cell, text, red_from_pat):
        cell.Value = text
        m = re.search(red_from_pat, text)
        if not m:
            return
        start = m.start() + 1
        try:
            cell.GetCharacters(1, len(text)).Font.Color = 0
            cell.GetCharacters(start, len(text) - start + 1).Font.Color = RED
        except Exception as e:
            self.log(f'글자색 지정 실패: {e}')

    # -------------------------------------------------------------- 4~7. 그림 구역
    SECTIONS = [('location', r'^4\.위치도'), ('birdseye', r'^5\.조감도'), ('siteplan', r'^6\.단지배치도'),
                ('community', r'^7\.커뮤니티')]

    def section_rows(self, ws):
        rows = {}
        for key, pat in self.SECTIONS:
            r = self.find_row(ws, pat.replace(r'^\d\.', r'^\d\.'), 1, 1, 600)
            if r is None:
                r = self.find_row(ws, pat.replace(r'^4\.', '').replace(r'^5\.', '').replace(r'^6\.', '').replace(r'^7\.', ''), 1, 1, 600)
            rows[key] = r
        rows['units'] = self.find_row(ws, r'^\d\.단위세대', 1, 1, 800)
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
            frames = self.frames_in(ws, r1, r2)
            if not files:
                self.log(f'{key}: 넣을 그림이 없어 양식 그림을 그대로 둡니다')
                continue
            for sh in self.pictures_in(ws, r1, r2, max_col=cN):
                sh.Delete()
            if frames:
                f = max(frames, key=lambda s: s.Width * s.Height)
                left, top, w, h = f.Left, f.Top, f.Width, f.Height
            else:
                top = ws.Cells(r1 + 1, 1).Top
                left = ws.Cells(r1 + 1, 2).Left
                w = ws.Cells(r1 + 1, cN + 1).Left - left
                h = ws.Cells(r2 + 1, 1).Top - top
            self.place_images(ws, files, left, top, w, h)

    # -------------------------------------------------------------- 8. 단위세대
    def fill_units(self, ws, d, images):
        rows = self.section_rows(ws)
        h = rows.get('units')
        if not h:
            self.log('단위세대 구역을 찾지 못했습니다')
            return
        first = h + 2
        ma = ws.Cells(first, 4).MergeArea
        size = ma.Rows.Count if ma.Rows.Count > 1 else 6
        nold = 0
        r = first
        while nold < 60:
            m = ws.Cells(r, 4).MergeArea
            if m.Rows.Count != size or m.Row != r:
                break
            nold += 1
            r += size
        if nold == 0:
            self.log('단위세대 칸 모양을 알 수 없어 건너뜁니다')
            return
        cN = self.cols['N']
        for sh in self.pictures_in(ws, h, first + nold * size + 2, max_col=cN):
            sh.Delete()
        types = d['types']
        self.resize_blocks(ws, first, nold, len(types), size)
        unit_imgs = images.get('units') or {}
        # 기본형/확장형 칸 열 (머리글 줄 h+1 의 병합)
        kb = ws.Cells(h + 1, 4).MergeArea
        ke_col = kb.Column + kb.Columns.Count
        for i, t in enumerate(types):
            b = first + i * size
            lab = t['abbr']
            ws.Cells(b + 2, 1).Value = int(lab) if re.fullmatch(r'\d+', lab) else lab
            sr = getattr(self, 'r0', None)
            if sr:
                ws.Cells(b + 3, 1).Formula = f"={col_letter(self.cols['C'])}{sr + 2 * i}*0.3025"
            else:
                ws.Cells(b + 3, 1).Value = round(t['excl'] * 0.3025, 2)
            ui = unit_imgs.get(t['abbr']) or {}
            for kind, col in (('basic', 4), ('ext', ke_col)):
                f = ui.get(kind)
                area = ws.Cells(b, col).MergeArea
                if kind == 'basic':
                    self.cross(area, not f)                # 기본형 평면이 없으면 X 표 (양식 관례)
                if not f:
                    continue
                self.place_images(ws, [f], area.Left, area.Top, area.Width, area.Height, margin=4)
        self.units_first, self.units_size, self.units_n = first, size, len(types)

    def cross(self, area, on):
        try:
            for b in (5, 6):                           # xlDiagonalDown, xlDiagonalUp
                if on:
                    area.Borders(b).LineStyle = 1
                    area.Borders(b).Weight = 2
                else:
                    area.Borders(b).LineStyle = -4142
        except Exception as e:
            self.log(f'X 표 지정 실패: {e}')

    def set_pages(self, ws):
        rows = self.section_rows(ws)
        last = None
        if getattr(self, 'units_n', None):
            last = self.units_first + self.units_n * self.units_size - 1
        else:
            last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
        try:
            ws.PageSetup.PrintArea = f'$A$1:${col_letter(self.cols["N"])}${last}'
        except Exception:
            pass
        try:
            ws.ResetAllPageBreaks()
            for key in ('location', 'birdseye', 'siteplan', 'community', 'units'):
                if rows.get(key):
                    ws.HPageBreaks.Add(ws.Cells(rows[key], 1))
            if getattr(self, 'units_n', None):
                per = 5
                for k in range(per, self.units_n, per):
                    ws.HPageBreaks.Add(ws.Cells(self.units_first + k * self.units_size, 1))
        except Exception as e:
            self.log(f'쪽 나누기 설정 실패: {e}')

    # -------------------------------------------------------------- 공고자료 시트
    def data_sheet(self, d):
        old = self.sheet('공고자료')
        if old is not None:
            old.Visible = -1
            old.Delete()
        last = self.wb.Worksheets(self.wb.Worksheets.Count)
        ws = self.wb.Worksheets.Add(None, last)
        ws.Name = '공고자료'
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
        rows.append(['5. 유상옵션 (원) — 분양개요 T열부터 같은 내용(천원), 유상옵션(C) 합계는 직접 입력'])
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
        width = max(len(r) for r in rows)
        data = [list(r) + [''] * (width - len(r)) for r in rows]
        data = [[('' if v is None else v) for v in r] for r in data]
        rng = ws.Range(ws.Cells(1, 1), ws.Cells(len(data), width))
        rng.Formula = data
        ws.Range('A1').Font.Bold = True
        ws.Range('A1').Font.Size = 13
        ws.Columns('A').ColumnWidth = 26
        ws.Columns('B:Z').ColumnWidth = 13
        ws.Range(ws.Cells(p0, 4), ws.Cells(p1, 8)).NumberFormat = '#,##0'
        ws.Range(ws.Cells(a, 2), ws.Cells(a + 4, 2)).NumberFormat = '#,##0'
        if o1 >= o0:
            ws.Range(ws.Cells(o0, 3), ws.Cells(o1, 2 + len(d['types']))).NumberFormat = '#,##0'
        ws.Visible = 0                                  # 숨김 (양식의 Sheet1 처럼)
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
            f = ws.Cells(rr, 3).Formula
            if isinstance(f, str) and f.startswith('=') or isinstance(ws.Cells(rr, 3).Value, (int, float)):
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
        # 데이터 줄의 병합을 모두 풀고(구분 칸은 다시 병합), 비고 값 지우기
        ws.Cells(first, 2).MergeArea.UnMerge()
        ws.Range(ws.Cells(first, 1), ws.Cells(sum_row - 1, 27)).UnMerge()
        delta = self.resize_blocks(ws, first, nold, n, 1)
        sum_row += delta
        ws.Range(ws.Cells(first, 2), ws.Cells(sum_row - 1, 27)).ClearContents()
        ws.Range(ws.Cells(first, 2), ws.Cells(sum_row, 2)).Merge()
        src = "'" + self.ws.Name + "'" if re.search(r'\W', self.ws.Name) else self.ws.Name
        src = self.ws.Name
        C = self.cols
        L = col_letter
        cols = [col_letter(c) for c in range(3, 26)]          # C ~ Y
        rows = []
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
            rows.append([f.get(c, '') for c in cols])         # 입력칸(K·N·P·Q·U·V·W)은 비워 둠
        if rows:
            ws.Range(ws.Cells(first, 3), ws.Cells(first + len(rows) - 1, 25)).Formula = rows
        ws.Cells(first, 2).Formula = f'={src}!A{r0}'
        S = sum_row
        a, b = first, sum_row - 1
        fs = {'I': f'=SUM(I{a}:I{b})', 'J': f'=SUM(J{a}:J{b})', 'K': f'=SUM(K{a}:K{b})', 'L': f'=IFERROR(K{S}/J{S},0)',
              'M': f'=SUM(M{a}:M{b})', 'N': f'=SUM(N{a}:N{b})', 'O': f'=IFERROR(N{S}/M{S},0)', 'P': f'=SUM(P{a}:P{b})',
              'Q': f'=SUM(Q{a}:Q{b})', 'R': f'=N{S}+P{S}+Q{S}', 'S': f'=IFERROR(R{S}/M{S},0)', 'T': f'=SUM(T{a}:T{b})',
              'U': f'=SUM(U{a}:U{b})', 'V': f'=SUM(V{a}:V{b})', 'W': f'=SUM(W{a}:W{b})', 'X': f'=SUM(U{S}:W{S})',
              'Y': f'=IFERROR((R{S}+X{S})/M{S},0)'}
        ws.Range(f'I{S}:Y{S}').Formula = [[fs.get(col_letter(c), ws.Cells(S, c).Formula) for c in range(9, 26)]]
        r2 = self.find_row(ws, r'^\(?작성일', 2, 1, 4)
        if r2:
            ws.Cells(r2, 2).Value = '( 작성일 :         년     월     일 )'
        r3 = self.find_row(ws, r'예비', 2, S + 1, S + 4)
        if r3 and d.get('reserve_pct'):
            ws.Cells(r3, 2).Value = f"※ 예비당첨자(예비입주자) {d['reserve_pct']}%"
        try:
            pa = ws.PageSetup.PrintArea
            m = re.match(r'^\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$', pa or '')
            if m:
                ws.PageSetup.PrintArea = f'${m.group(1)}${m.group(2)}:${m.group(3)}${S + 1}'
        except Exception:
            pass

    # -------------------------------------------------------------- 이전 현장 숨김 시트
    def delete_sheets(self, names):
        names = [n for n in names if n != '공고자료']
        if not names:
            return
        # 남은 시트의 수식이 지울 시트를 가리키면 값으로 바꾼다
        for s in self.wb.Worksheets:
            if s.Name in names or s.Name == '공고자료':
                continue
            try:
                ur = s.UsedRange
                cells = ur.SpecialCells(-4123)          # xlCellTypeFormulas
            except Exception:
                continue
            for c in cells:
                f = c.Formula
                if any((n + '!') in f or (f"'{n}'!") in f for n in names):
                    c.Value = c.Value
        for n in names:
            try:
                sh = self.wb.Worksheets(n)
                sh.Visible = -1
                sh.Delete()
                self.log(f'이전 현장 숨김 시트 "{n}" 삭제')
            except Exception as e:
                self.log(f'숨김 시트 "{n}" 삭제 실패: {e}')

    # -------------------------------------------------------------- 미리보기
    def export_pdf(self, pdf_path):
        ws = self.sheet('분양개요') or self.wb.Worksheets(1)
        ws.ExportAsFixedFormat(0, pdf_path)


def build_workbook(template_path, out_path, d, images, log=None, preview_pdf=None):
    import shutil
    shutil.copyfile(template_path, out_path)
    b = Builder(out_path, log)
    t = time.time()
    b.open()
    ok = False
    try:
        b.build(d, images)
        if preview_pdf:
            try:
                b.export_pdf(preview_pdf)
            except Exception as e:
                (log or print)(f'미리보기 PDF 실패: {e}')
        ok = True
    finally:
        b.close(save=ok)
    return time.time() - t
