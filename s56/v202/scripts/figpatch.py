"""그림 안 글자 교체 — 같은 글꼴(Noto Sans CJK KR)로 찾아 지우고 다시 그림.
patch(A, old, new, win, align, weights, sizes) : A = RGB float array (in-place 수정), win=(x0,y0,x1,y1)
"""
import numpy as np, cv2
from PIL import Image, ImageDraw, ImageFont

FONTS = {'Black': '/root/.fonts/NotoSansKR-700.ttf', 'Bold': '/root/.fonts/NotoSansKR-700.ttf', 'Medium': '/root/.fonts/NotoSansKR-500.ttf', 'Regular': '/root/.fonts/NotoSansKR-400.ttf', 'DemiLight': '/root/.fonts/NotoSansKR-400.ttf', 'Light': '/root/.fonts/NotoSansKR-400.ttf'}

def render_mask(s, size, weight):
    f = ImageFont.truetype(FONTS[weight], size)
    l, t, r, b = f.getbbox(s)
    im = Image.new('L', (r - l + 8, b - t + 8), 0)
    ImageDraw.Draw(im).text((4 - l, 4 - t), s, font=f, fill=255)
    M = np.array(im).astype(np.float32) / 255.
    ys, xs = np.where(M > 0.2)
    M = M[ys.min():ys.max() + 1, xs.min():xs.max() + 1]
    return M

def find(A, old, win, weights, sizes):
    x0, y0, x1, y1 = win
    G = cv2.cvtColor(A[y0:y1, x0:x1].astype(np.uint8), cv2.COLOR_RGB2GRAY).astype(np.float32)
    best = None
    for w in weights:
        for s in sizes:
            M = render_mask(old, s, w)
            if M.shape[0] >= G.shape[0] or M.shape[1] >= G.shape[1]: continue
            R = cv2.matchTemplate(G, M, cv2.TM_CCOEFF_NORMED)
            for sign in (1, -1):
                v = (sign * R).max()
                if best is None or v > best[0]:
                    yy, xx = np.unravel_index((sign * R).argmax(), R.shape)
                    best = (v, w, s, x0 + xx, y0 + yy, M.shape[1], M.shape[0], sign)
    return best

def patch(A, old, new, win, align='left', weights=('Bold', 'Medium', 'Regular'), sizes=range(14, 90, 1), pad=3, minscore=0.8, erase=None, dx=0, xr=0):
    v, w, s, px, py, tw, th, sign = find(A, old, win, weights, sizes)
    M = render_mask(old, s, w)
    reg = A[py:py + th, px:px + tw]
    ink = M > 0.6; bgm = M < 0.02
    txt = np.median(reg[ink], axis=0); bg = np.median(reg[bgm], axis=0)
    info = dict(old=old, new=new, score=round(float(v), 3), weight=w, size=s, box=(px, py, tw, th))
    assert v >= minscore, info
    # erase
    ex0, ey0, ex1, ey1 = (px - pad, py - pad, px + tw + pad + xr, py + th + pad) if erase is None else erase
    A[ey0:ey1, ex0:ex1] = bg
    if new:
        N = render_mask(new, s, w)
        nh, nw = N.shape
        # 기준선 맞춤: 두 글자열의 첫 글자 높이 차이를 줄이려 위쪽 정렬 대신 아래쪽(글자 몸) 정렬
        ny = py + th - nh if abs(th - nh) <= 3 else py + (th - nh) // 2
        if align == 'left': nx = px
        elif align == 'right': nx = px + tw - nw
        else: nx = px + (tw - nw) // 2
        nx += dx
        sub = A[ny:ny + nh, nx:nx + nw]
        a = N[..., None]
        A[ny:ny + nh, nx:nx + nw] = sub * (1 - a) + txt * a
        info['newbox'] = (nx, ny, nw, nh)
    return info

def load(path, rot=0):
    im = Image.open(path)
    mode = im.mode
    alpha = np.array(im)[:, :, 3] if mode == 'RGBA' else None
    A = np.array(im.convert('RGB')).astype(np.float32)
    if rot: A = np.rot90(A, rot).copy()
    return A, mode, alpha

def save(A, path, mode, alpha, rot=0, quality=95):
    if rot: A = np.rot90(A, -rot).copy()
    B = np.clip(A, 0, 255).astype(np.uint8)
    if mode == 'RGBA':
        Image.fromarray(np.dstack([B, alpha]), 'RGBA').save(path)
    elif path.endswith('.jpg'):
        Image.fromarray(B, 'RGB').save(path, quality=quality, subsampling=0)
    else:
        Image.fromarray(B, 'RGB').save(path)
