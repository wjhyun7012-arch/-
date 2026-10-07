"""슬라이드 3 그림(논문 구조도 컬러판) 글자 고침 — v8.
입력 s3_img.png(v7 슬라이드 3의 그림, 3600x2850) → fig8/s3_6.png
"""
from PIL import Image, ImageDraw, ImageFont
import numpy as np, json, sys, os
FD = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'fonts/')
def font(w, s):
    return ImageFont.truetype(FD + {'r': 'NotoSansCJKkr-Regular.otf', 'b': 'NotoSansCJKkr-Bold.otf', 'm': 'NotoSansCJKkr-Medium.otf'}[w], s)
def ink(im, box, bg=None, th=80):
    a = np.asarray(im.crop(box).convert('RGB')).astype(int)
    if bg is None:
        e = np.concatenate([a[0], a[-1], a[:, 0], a[:, -1]]); vals, cnt = np.unique(e, axis=0, return_counts=True); bg = tuple(int(v) for v in vals[cnt.argmax()])
    d = np.abs(a - np.array(bg)).sum(2); m = d > th
    ys, xs = np.where(m)
    col = a[m]; dd = d[m]; col = tuple(int(v) for v in np.median(col[dd >= np.percentile(dd, 70)], axis=0))
    return (box[0] + xs.min(), box[1] + ys.min(), box[0] + xs.max(), box[1] + ys.max()), bg, col
def calib(old, wpx, w):
    best = None
    for s in range(12, 140):
        f = font(w, s); l, t, r, b = f.getbbox(old); e = abs((r - l) - wpx)
        if best is None or e < best[0]: best = (e, s)
    return best[1]
def fit(text, maxw, w, cap):
    s = cap
    while s > 12:
        f = font(w, s); l, t, r, b = f.getbbox(text)
        if r - l <= maxw: break
        s -= 1
    return s
def patch(im, box, old, new, w='r', align='left', bg=None, dy=0, color=None, maxw=None):
    (x0, y0, x1, y1), bg, col = ink(im, box, bg)
    s = calib(old, x1 - x0 + 1, w)
    if maxw: s = fit(new, maxw, w, s)
    f = font(w, s); l, t, r, b = f.getbbox(old)
    D = ImageDraw.Draw(im); D.rectangle((x0 - 4, y0 - 4, x1 + 4, y1 + 4), fill=bg)
    nl, nt, nr, nb = f.getbbox(new)
    x = x0 - nl if align == 'left' else int((x0 + x1) / 2 - (nr - nl) / 2 - nl)
    y = y0 - t + dy
    D.text((x, y), new, font=f, fill=color or col)
    return dict(size=s, ink=(int(x0), int(y0), int(x1), int(y1)), newbox=(x + nl, y + nt, x + nr, y + nb), color=col)

log = {}
im = Image.open(sys.argv[1]).convert('RGB')
# 1) 그림 머리말: 「논문 전체 구조도 — 다섯 장의 큰 흐름」 → [그림 1-4] 캡션
log['head'] = patch(im, (40, 20, 1400, 120), '논문 전체 구조도 — 다섯 장의 큰 흐름', '논문의 구성과 장 간 연결', w='b', align='left')
# 2) 3장 빨간 상자 두 줄 (상자 안쪽 x 150..1710, y 1598..1733)
(ax0, ay0, ax1, ay1), bg3, col3 = ink(im, (300, 1600, 1560, 1660))   # 1줄 잉크
(bx0, by0, bx1, by1), _, _ = ink(im, (300, 1662, 1560, 1725))         # 2줄 잉크
s = calib('연계 성립의 네 지점은 평균으로 보상되지 않음(비보상 판정)', ax1 - ax0 + 1, 'b')
new1 = '네 지점 가운데 하나라도 빠지면 평균이 높아도 ML3 이상으로 가지 못함'
new2 = '모델의 타당성은 두 방향으로 검증'
s = fit(new1, 1480, 'b', s); f = font('b', s)
D = ImageDraw.Draw(im); D.rectangle((160, 1604, 1700, 1727), fill=bg3)
cx = (150 + 1710) / 2
for txt, (y0, _) in [(new1, (ay0, ay1)), (new2, (by0, by1))]:
    l, t, r, b = f.getbbox(txt); x = int(cx - (r - l) / 2 - l); D.text((x, y0 - t), txt, font=f, fill=col3)
log['ch3'] = dict(size=s, color=col3, bg=bg3, ink1=(int(ax0), int(ay0), int(ax1), int(ay1)))
# 3) 4장 상자 1줄: 「낮은 데가 왜 낮은지 — 구조적 미비 3건」 → 「낮은 점수의 원인 — 미비 사항 3건」
log['ch4'] = patch(im, (2000, 1585, 3350, 1650), '낮은 데가 왜 낮은지 — 구조적 미비 3건', '낮은 점수의 원인 — 미비 사항 3건', w='r', align='center')
os.makedirs('fig8', exist_ok=True)
im.save('fig8/s3_6.png')
im.crop((0, 0, 1500, 140)).save('fig8/chk_head.png')
im.crop((80, 1450, 1800, 2000)).save('fig8/chk_ch3.png')
im.crop((1750, 1450, 3550, 2000)).save('fig8/chk_ch4.png')
json.dump(log, open('fig8/log.json', 'w'), ensure_ascii=False, indent=1, default=str)
print(json.dumps(log, ensure_ascii=False, default=str))
