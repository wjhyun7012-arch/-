"""v9 → v10: 그림을 화면에서 최대한 크게 (이사님 10/7). 사용: python3 -I fit9.py v9.pptx v10.pptx
규칙: 그림 상자 = (제목·▶ 줄 아래) ~ (쪽번호 위) × (좌우 여백 안), 옆에 다른 도형이 있으면 그 칸 안. 비율 유지.
그림 아래 글은 그림 밑으로 내리고, 안 들어가면 글자를 1pt씩 줄임(최소 11pt).
"""
import sys, json, os, re, unicodedata, math
from pptx import Presentation
from pptx.util import Emu, Pt
from PIL import ImageFont
SRC, OUT = sys.argv[1], sys.argv[2]
prs = Presentation(SRC); W, H = prs.slide_width, prs.slide_height
LOG = []
FD = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'fonts/'); _FC = {}
def _font(bold, pt):
    k = (bold, pt)
    if k not in _FC: _FC[k] = ImageFont.truetype(FD + ('NotoSansCJKkr-Bold.otf' if bold else 'NotoSansCJKkr-Regular.otf'), int(pt * 10))
    return _FC[k]
def wrap_lines(txt, bold, pt, inner_pt):
    f = _font(bold, pt); Wd = inner_pt * 10 / 1.08
    def w(s):
        b = sum(1 for a, c in zip(s, s[1:]) if (unicodedata.east_asian_width(a) in 'WF') != (unicodedata.east_asian_width(c) in 'WF') and a != ' ' and c != ' ')
        return f.getlength(s) + b * pt * 10 * 0.25
    lines = 1; cur = ''
    for word in re.split(r'(?<= )', txt):
        if w(cur + word.rstrip()) <= Wd: cur += word; continue
        if cur: lines += 1; cur = ''
        while w(word.rstrip()) > Wd:
            k = len(word)
            while k > 1 and w(word[:k]) > Wd: k -= 1
            word = word[k:]; lines += 1
        cur = word
    return lines
def tf_need(sh, delta_pt=0):
    tf = sh.text_frame; bp = tf._txBody.bodyPr
    def ins(a, d):
        v = bp.get(a); return int(v) if v is not None else d
    mL, mR, mT, mB = ins('lIns', 91440), ins('rIns', 91440), ins('tIns', 45720), ins('bIns', 45720)
    inner = (sh.width - mL - mR) / 12700.0; h = 0.0
    for p in tf.paragraphs:
        sz = None; bold = False
        for r in p.runs:
            if r.font.size: sz = r.font.size.pt; bold = bool(r.font.bold); break
        if sz is None: sz = 18
        sz = max(11, sz - delta_pt)
        txt = ''.join(r.text for r in p.runs)
        lines = wrap_lines(txt, bold, sz, inner) if txt.strip() else 1
        sa = p.space_after.pt if p.space_after is not None else 0
        sb = p.space_before.pt if p.space_before is not None else 0
        h += lines * sz * 1.36 + sa + sb
    return int(h * 12700 + mT + mB)
def shrink(sh, delta_pt):
    for p in sh.text_frame.paragraphs:
        for r in p.runs:
            if r.font.size: r.font.size = Pt(max(11, r.font.size.pt - delta_pt))


# ---------- (0) 머리 올리기: 제목 161k→80k, 장 표시 190k→110k, 선 837k→690k, ▶ 줄 917k→740k (전 슬라이드, 쪽번호 제외) ----------
UP_T, UP_L, UP_R, UP_B = 161280 - 80000, 190000 - 110000, 836640 - 690000, 917280 - 740000
for i, s in enumerate(prs.slides, 1):
    if i in (1, 29, 30): continue
    for sh in s.shapes:
        if sh.shape_id == 2 and sh.top == 161280: sh.top = Emu(80000)
        elif sh.shape_id == 3 and sh.top in (190000, 262080) and sh.left > 5000000: sh.top = Emu(sh.top - UP_L)
        elif sh.shape_id == 4 and sh.height < 40000 and sh.top == 836640: sh.top = Emu(690000)
        elif sh.has_text_frame and sh.text_frame.text.startswith('▶') and sh.top == 917280: sh.top = Emu(740000)
LOG.append(('전체', '머리 위치', '제목 0.18in·선 0.92in·▶ 1.00in', '제목 0.09in·선 0.75in·▶ 0.81in', '이사님 10/7: 위쪽 선을 올려 화면을 더 크게'))

FOOT = int(H - 610000)     # 쪽번호·근거 줄 위 (A4: 7.56in 기준 6.9in)
MARG = 480000; GAP = 70000
def hov(a, b): return a.left < b.left + b.width and a.left + a.width > b.left
def vov(a, b): return a.top < b.top + b.height and a.top + a.height > b.top

for i, s in enumerate(prs.slides, 1):
    if i in (1, 29, 30): continue
    pics = [sh for sh in s.shapes if sh.shape_type == 13]
    if not pics: continue
    others = [sh for sh in s.shapes if sh.shape_type != 13 and sh.top < FOOT]
    hdr = [o for o in others if o.top + o.height <= min(p.top for p in pics) + 30000 and (o.shape_id in (2, 3, 4) or (o.has_text_frame and o.text_frame.text.startswith('▶')))]
    # 그림 묶음의 경계 상자
    px0 = min(p.left for p in pics); px1 = max(p.left + p.width for p in pics)
    py0 = min(p.top for p in pics); py1 = max(p.top + p.height for p in pics)
    class B: pass
    bb = B(); bb.left, bb.top, bb.width, bb.height = px0, py0, px1 - px0, py1 - py0
    # 옆(세로로 겹치는) 도형 → 좌우 한계
    side = [o for o in others if o not in hdr and vov(o, bb) and not hov(o, bb)]
    left_lim = max([o.left + o.width for o in side if o.left + o.width <= bb.left + 10000] + [MARG]) + GAP if side else MARG
    right_lim = min([o.left for o in side if o.left >= bb.left + bb.width - 10000] + [W - MARG]) - GAP if side else W - MARG
    # 위 한계: 머리(제목·▶ 줄)와 그림 위에 있는 다른 도형
    above = [o for o in others if o.top + o.height <= bb.top + 30000 and hov(o, bb)]
    top_lim = max([o.top + o.height for o in above] + [1080000]) + GAP
    # 아래 글 블록(그림 아래, 가로로 겹침)
    below = sorted([o for o in others if o.top >= bb.top + bb.height - 30000 and hov(o, bb) and o not in side], key=lambda o: o.top)
    # 옆 도형이 그림 아래까지 내려오면(16번 표 등) 아래 글은 옆 도형과도 겹치지 않아야 하므로 그대로 둠
    ar = bb.width / bb.height
    box_w = right_lim - left_lim
    # 아래 글을 줄(row) 단위로 묶음 — 같은 높이에 나란한 상자들은 함께 움직임
    rows = []
    for o in below:
        if rows and o.top < max(x.top + x.height for x in rows[-1]) - 20000: rows[-1].append(o)   # 세로로 겹치면 같은 줄
        else: rows.append([o])
    def oh(o, delta):
        return tf_need(o, delta) if (o.has_text_frame and o.text_frame.text.strip() and not o.has_table and o.shape_type == 17) else o.height
    def layout(delta):
        need_below = sum(max(oh(o, delta) for o in r) for r in rows) + GAP * len(rows)
        avail_h = FOOT - top_lim - need_below
        w = box_w; h = int(w / ar)
        if h > avail_h: h = avail_h; w = int(h * ar)
        return w, h, need_below
    delta = 0; w, h, nb = layout(0)
    # 아래 글을 줄이면 그림이 더 커지는지: 글자 3pt까지 줄여 보고 그림 세로가 10% 이상 커지면 채택
    best = (w, h, 0)
    for d in (1, 2, 3):
        w2, h2, _ = layout(d)
        if below and h2 > best[1] * 1.08 and d <= 2: best = (w2, h2, d)
    w, h, delta = best
    if w <= bb.width * 1.01 and h <= bb.height * 1.01:
        continue   # 더 키울 자리 없음
    k = w / bb.width
    nx0 = left_lim + (box_w - w) // 2 if not side else left_lim
    if side and right_lim - left_lim > w: nx0 = left_lim + (right_lim - left_lim - w) // 2
    for p in pics:
        p.left = Emu(int(nx0 + (p.left - bb.left) * k)); p.top = Emu(int(top_lim + (p.top - bb.top) * k))
        p.width = Emu(int(p.width * k)); p.height = Emu(int(p.height * k))
    # 그림과 같이 움직여야 하는 그림 위 덧글(그림 영역 안에 있는 작은 글상자)은 없음(확인됨) — 아래 글 재배치
    y = top_lim + h + GAP
    for r in rows:
        r0 = min(o.top for o in r)
        for o in r:
            if delta and o.has_text_frame and o.text_frame.text.strip() and not o.has_table and o.shape_type == 17: shrink(o, delta)
            if o.has_text_frame and o.text_frame.text.strip() and not o.has_table and o.shape_type == 17: o.height = Emu(max(tf_need(o), int(o.height * 0.6)))
            o.top = Emu(y + (o.top - r0))
        y = max(o.top + o.height for o in r) + GAP
    LOG.append((i, '그림 확대', f'{bb.width/914400:.2f}×{bb.height/914400:.2f}in', f'{w/914400:.2f}×{h/914400:.2f}in (×{k:.2f})' + (f', 아래 글 {len(below)}개 −{delta}pt' if delta else (f', 아래 글 {len(below)}개 내림' if below else '')), '이사님 10/7: 그림은 화면에서 최대한 크게'))


# ---------- (2) 그림 없는 장: 내용 전체를 ▶ 줄 아래(1.18M)로 당김 ----------
for i, s in enumerate(prs.slides, 1):
    if i in (1, 29, 30): continue
    if any(sh.shape_type == 13 for sh in s.shapes): continue
    content = [sh for sh in s.shapes if sh.shape_id not in (2, 3, 4) and not (sh.has_text_frame and sh.text_frame.text.startswith('▶')) and sh.top < FOOT]
    if not content: continue
    top = min(sh.top for sh in content); up = top - 1180000
    if up > 40000:
        for sh in content: sh.top = Emu(sh.top - up)
        LOG.append((i, '내용 위로', f'{top/914400:.2f}in', f'1.29in ({up/914400:.2f}in 올림)', '머리를 올린 만큼'))

prs.save(OUT)
json.dump(LOG, open(os.path.splitext(OUT)[0] + '_fitlog.json', 'w'), ensure_ascii=False, indent=1)
for l in LOG: print(l)
