"""그림 속 덩어리(도면·사진·글자·키맵) 찾기 — 홈페이지 그림에서 필요한 부분만 자르기 (2026-10-18)

홈페이지 그림은 한 장에 여러 가지가 붙어 있는 일이 많다.
  평면 쪽    : 2D 평면 + 3D 아이소 + 키맵(작은 단지배치도) + 면적표·표기 단추
  배치도 쪽  : 홍보 문구 + 위에서 본 배치도 + 범례 + 사진
  커뮤니티 쪽: 층별 도면 + 사진 + 키맵 + 설명 글
흰(옅은 회색 포함) 배경 위의 덩어리를 찾고, 모양·색으로 종류를 가린다.
  plan  위에서 본 2D 도면 (벽이 가로·세로라 경계가 줄마다 같은 자리에 이어짐)
  iso   비스듬히 본 3D 도면 (경계가 사선)
  photo 사진 (네모 꽉 참, 색이 다양함)
  keymap 옅은 선 그림 (키맵·약도)
  label / text  검은 상자·글자 줄
"""
from collections import deque

from PIL import Image, ImageChops, ImageFilter

BG = 236            # 이보다 밝으면 배경 (흰색·옅은 회색 띠)
WORK_W = 520        # 덩어리 찾기는 이 폭으로 줄여서 (속도)


def _mask(sm):
    r, g, b = sm.split()
    mn = ImageChops.darker(ImageChops.darker(r, g), b)
    return mn.point(lambda v: 255 if v < BG else 0)


def _components(mask):
    """255 픽셀 덩어리(상하좌우 이웃)들의 상자 [(x0, y0, x1, y1)]"""
    w, h = mask.size
    data = mask.tobytes()
    seen = bytearray(w * h)
    out = []
    for start in range(w * h):
        if not data[start] or seen[start]:
            continue
        seen[start] = 1
        q = deque([start])
        x0 = x1 = start % w
        y0 = y1 = start // w
        while q:
            i = q.popleft()
            x, y = i % w, i // w
            if x < x0:
                x0 = x
            elif x > x1:
                x1 = x
            if y < y0:
                y0 = y
            elif y > y1:
                y1 = y
            if x > 0 and data[i - 1] and not seen[i - 1]:
                seen[i - 1] = 1
                q.append(i - 1)
            if x < w - 1 and data[i + 1] and not seen[i + 1]:
                seen[i + 1] = 1
                q.append(i + 1)
            if y > 0 and data[i - w] and not seen[i - w]:
                seen[i - w] = 1
                q.append(i - w)
            if y < h - 1 and data[i + w] and not seen[i + w]:
                seen[i + w] = 1
                q.append(i + w)
        out.append((x0, y0, x1 + 1, y1 + 1))
    return out


def _features(ob, lb, rgb, w, x0, y0, x1, y1):
    bw, bh = x1 - x0, y1 - y0
    cnt = dark = mid = 0
    blank_rows = 0
    cols = {}
    same = tot = 0
    prev = None
    edge = 0
    for y in range(y0, y1):
        row = ob[y * w + x0:y * w + x1]
        n = bw - row.count(0)
        cnt += n
        if n == 0:
            blank_rows += 1
            prev = None
            continue
        lrow = lb[y * w + x0:y * w + x1]
        crow = rgb[(y * w + x0) * 3:(y * w + x1) * 3]
        tr = set()
        last = 0
        for k in range(bw):
            v = row[k]
            if v:
                if lrow[k] < 70:
                    dark += 1
                elif lrow[k] < 170:
                    mid += 1
                if not (k & 1) and not (y & 1):
                    key = (crow[3 * k] // 24, crow[3 * k + 1] // 24, crow[3 * k + 2] // 24)
                    cols[key] = cols.get(key, 0) + 1
                v = 1
            if v != last:
                tr.add(k)
            last = v
        if prev is not None:
            for t in tr:
                tot += 1
                if t in prev or t - 1 in prev or t + 1 in prev:
                    same += 1
        prev = tr
    # 상자 네 변(잇기 전 경계라 3px 안쪽)이 채워진 정도 — 사진은 네 변이 꽉 참
    ix0, iy0, ix1, iy1 = x0 + 3, y0 + 3, max(x0 + 4, x1 - 4), max(y0 + 4, y1 - 4)
    for yy in (iy0, iy1 - 1):
        row = ob[yy * w + ix0:yy * w + ix1]
        edge += len(row) - row.count(0)
    for yy in range(iy0, iy1):
        edge += (1 if ob[yy * w + ix0] else 0) + (1 if ob[yy * w + ix1 - 1] else 0)
    csum = sum(cols.values()) or 1
    return {'fill': cnt / (bw * bh), 'dark': dark / (cnt or 1), 'mid': mid / (cnt or 1),
            'top': sum(sorted(cols.values(), reverse=True)[:12]) / csum,
            'axis': same / tot if tot else 0.0, 'gaps': blank_rows / bh,
            'rect': edge / max(1, 2 * (ix1 - ix0) + 2 * (iy1 - iy0))}


def kind_of(f):
    if f['area'] < 0.012:
        return 'text'
    if f.get('wide') and f['fill'] >= 0.8:
        return 'photo'                      # 그림 폭 전체를 덮는 띠 (머리 사진·안내 띠)
    if f['dark'] > 0.6 and f['fill'] > 0.6:
        return 'label'                      # 검은 바탕 상자 (타입 이름·표기 단추)
    if f['gaps'] > 0.15:
        return 'text'                       # 글자 줄 사이가 비어 있는 묶음 (설명 글·면적표)
    if f['dark'] < 0.09 and f['fill'] < 0.6 and f['top'] >= 0.85:
        return 'keymap'                     # 옅은 선 그림 (키맵·약도) — 색이 몇 가지뿐
    if f['fill'] >= 0.95 or (f['fill'] >= 0.86 and f['top'] < 0.9) or \
            (f['axis'] >= 0.88 and f['top'] < 0.8 and f['fill'] >= 0.68):
        return 'photo'                      # 네모 꽉 참(색이 다양함), 또는 가로세로 경계인데 색이 다양함(사진 묶음)
    if f['axis'] < 0.78:
        return 'iso'                        # 경계가 사선
    if f['top'] < 0.84 and f['mid'] > 0.25:
        return 'iso'                        # 벽면 음영(중간 밝기)이 많고 색이 다양함 — 거의 정면에서 본 3D
    return 'plan'                           # 위에서 본 2D: 색이 몇 가지뿐 (벽 검정·바닥 옅은 색)


def blobs(img):
    """덩어리 목록 [{box(원본 좌표), area(비율), kind, fill, dark, top, axis, gaps, rect}] — 위→아래, 왼→오른"""
    im = img.convert('RGB')
    s = min(1.0, WORK_W / im.width)
    sm = im.resize((max(1, round(im.width * s)), max(1, round(im.height * s))), Image.BILINEAR) if s < 1 else im
    w, h = sm.size
    opened = _mask(sm).filter(ImageFilter.MinFilter(3)).filter(ImageFilter.MaxFilter(3))   # 1~2px 구분선 지움
    closed = opened.filter(ImageFilter.MaxFilter(5))                                    # 가까운 조각 잇기
    ob, lb, rgb = opened.tobytes(), sm.convert('L').tobytes(), sm.tobytes()
    out = []
    for x0, y0, x1, y1 in _components(closed):
        if (x1 - x0) * (y1 - y0) < 0.002 * w * h:
            continue
        f = _features(ob, lb, rgb, w, x0, y0, x1, y1)
        f['area'] = (x1 - x0) * (y1 - y0) / (w * h)
        f['wide'] = x0 <= 2 and x1 >= w - 2
        f['box'] = (round(x0 / s), round(y0 / s), min(im.width, round(x1 / s)), min(im.height, round(y1 / s)))
        f['kind'] = kind_of(f)
        out.append(f)
    out.sort(key=lambda b: (b['box'][1] // 40, b['box'][0]))
    return out


def pad(box, size, m=8):
    x0, y0, x1, y1 = box
    return (max(0, x0 - m), max(0, y0 - m), min(size[0], x1 + m), min(size[1], y1 + m))


def drawings(bs, min_rel=0.25):
    """도면 덩어리(plan·iso)만, 가장 큰 것의 min_rel 배 이상"""
    ds = [b for b in bs if b['kind'] in ('plan', 'iso')]
    if not ds:
        return []
    big = max(b['area'] for b in ds)
    return [b for b in ds if b['area'] >= big * min_rel]


def merge_near(bs, gap):
    """붙어 있는 도면 조각(평면 옆의 현관 대안 등)을 하나로. 큰 쪽의 종류를 따른다"""
    bs = [dict(b) for b in bs]
    changed = True
    while changed:
        changed = False
        for i in range(len(bs)):
            for j in range(i + 1, len(bs)):
                a, b = bs[i]['box'], bs[j]['box']
                dx = max(0, max(a[0], b[0]) - min(a[2], b[2]))
                dy = max(0, max(a[1], b[1]) - min(a[3], b[3]))
                ovx = min(a[2], b[2]) - max(a[0], b[0])
                ovy = min(a[3], b[3]) - max(a[1], b[1])
                hmin = min(a[3] - a[1], b[3] - b[1])
                wmin = min(a[2] - a[0], b[2] - b[0])
                if (dx <= gap and ovy >= 0.3 * hmin) or (dy <= gap and ovx >= 0.3 * wmin):
                    big = bs[i] if bs[i]['area'] >= bs[j]['area'] else bs[j]
                    m = dict(big)
                    m['box'] = (min(a[0], b[0]), min(a[1], b[1]), max(a[2], b[2]), max(a[3], b[3]))
                    m['area'] = bs[i]['area'] + bs[j]['area']
                    bs[j] = m
                    del bs[i]
                    changed = True
                    break
            if changed:
                break
    return bs


def plan_parts(img):
    """평면 그림 → {'plans': 2D 평면 [큰 것부터], 'isos': 3D 아이소, 'sheet': 옵션 안내처럼 작은 도면이 여럿인 그림}
    키맵·면적표·표기 단추·사진은 빠진다"""
    bs = blobs(img)
    raw = [b for b in bs if b['kind'] in ('plan', 'iso')]
    sheet = sum(1 for b in raw if b['area'] >= 0.03) >= 5   # 붙이기 전에 센다 (옵션 칸이 서로 붙어 보이므로)
    ds = drawings(merge_near(raw, 0.03 * img.width), 0.25)
    by = lambda k: sorted([b for b in ds if b['kind'] == k], key=lambda b: -b['area'])
    return {'plans': by('plan'), 'isos': by('iso'), 'sheet': sheet}


def siteplan_box(img):
    """배치도 그림 → 위에서 본 배치도 덩어리만 (홍보 문구·범례·사진 빼고). 그림 전체와 거의 같으면 None"""
    bs = [b for b in blobs(img) if b['kind'] not in ('text', 'label') and b['area'] >= 0.12]
    if not bs:
        return None
    b = max(bs, key=lambda b: b['area'])
    return None if b['area'] >= 0.9 else b['box']


def community_boxes(img, limit=4):
    """커뮤니티 그림 → 도면 덩어리들 (사진·키맵·설명 글 빼고), 위→아래 순서. 없으면 []"""
    ds = merge_near([b for b in blobs(img) if b['kind'] in ('plan', 'iso')], 0.02 * img.width)
    ds = [b for b in drawings(ds, 0.2) if b['area'] >= 0.02]
    ds = sorted(sorted(ds, key=lambda b: -b['area'])[:limit], key=lambda b: (b['box'][1], b['box'][0]))
    if len(ds) == 1 and ds[0]['area'] >= 0.9:
        return []
    return [b['box'] for b in ds]


def light_panel(img, thr=200):
    """어두운 배너 속 밝은 네모(지도) 상자"""
    g = img.convert('L')
    s = min(1.0, WORK_W / g.width)
    sm = g.resize((max(1, round(g.width * s)), max(1, round(g.height * s)))) if s < 1 else g
    mask = sm.point(lambda v: 255 if v > thr else 0).filter(ImageFilter.MaxFilter(5))
    w, h = sm.size
    best = None
    data = sm.point(lambda v: 255 if v > thr else 0).tobytes()
    for x0, y0, x1, y1 in _components(mask):
        area = (x1 - x0) * (y1 - y0) / (w * h)
        if area < 0.08 or area > 0.9:
            continue
        lit = sum(len(data[y * w + x0:y * w + x1]) - data[y * w + x0:y * w + x1].count(0) for y in range(y0, y1))
        if lit / ((x1 - x0) * (y1 - y0)) < 0.5:
            continue
        if best is None or area > best[0]:
            best = (area, (round(x0 / s), round(y0 / s), round(x1 / s), round(y1 / s)))
    return best[1] if best else None


def main_box(img):
    """약도 같은 한 장짜리 그림의 본체 상자 (흰 여백·주변 글자 빼고). 어두운 배너면 그 안의 밝은 지도. 자를 게 없으면 None"""
    bs = blobs(img)
    if not bs:
        return None
    b = max(bs, key=lambda b: b['area'])
    if b['area'] >= 0.85:
        lum = sum(img.convert('L').resize((64, 64)).tobytes()) / 4096
        return light_panel(img) if lum < 110 else None
    return b['box'] if b['area'] >= 0.2 else None


def legend_below(img, box, max_frac=0.3, white=243, gap=14):
    """상자 바로 아래에 붙은 범례(옅은 회색 띠·글자)까지 늘린 상자. 빈 줄이 gap 넘게 이어지면 멈춤"""
    g = img.convert('L')
    x0, y0, x1, y1 = box
    w = x1 - x0
    px = g.load()
    step = max(1, w // 300)
    y, blank, last = y1, 0, y1
    limit = min(img.height, y1 + int((y1 - y0) * max_frac))
    while y < limit:
        lit = sum(1 for x in range(x0, x1, step) if px[x, y] < white)
        if lit:
            blank = 0
            last = y + 1
        else:
            blank += 1
            if blank >= gap:
                break
        y += 1
    return (x0, y0, x1, last)
