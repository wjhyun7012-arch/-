"""v8 A4 넘침 조정: (a) 제목 글꼴 자동 축소(한 줄에 들어가도록, 최소 20pt) + 장 표시 오른쪽 정렬
(b) 표 행 높이를 글자 양에 맞게 미리 늘리고, 표 아래 도형을 그만큼 내림.
사용: python3 -I fit8.py in.pptx out.pptx
"""
import sys, json, unicodedata, math, os, re
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.enum.text import PP_ALIGN
SRC, OUT = sys.argv[1], sys.argv[2]
prs = Presentation(SRC)
LOG = []
def em(s):
    w = 0
    for ch in s:
        if unicodedata.east_asian_width(ch) in 'WF': w += 1.0
        elif ch == ' ': w += 0.3
        elif ch.isdigit() or ch.isupper(): w += 0.62
        else: w += 0.52
    return w



# ---------- (-1) 모든 글자에 한글 글꼴(ea) 「맑은 고딕」 명시 — 표·텍스트 상자 공통 ----------
from pptx.oxml.ns import qn
from lxml import etree
n_ea = 0
for s in prs.slides:
    for rpr in s._element.iter(qn('a:rPr'), qn('a:endParaRPr')):
        if rpr.find(qn('a:ea')) is not None: continue
        lat = rpr.find(qn('a:latin'))
        if lat is None: continue
        ea = etree.SubElement(rpr, qn('a:ea')); ea.set('typeface', lat.get('typeface')); lat.addnext(ea); n_ea += 1
LOG.append(('전체', '한글 글꼴 지정', '라틴 글꼴만 「맑은 고딕」', f'한글(ea)에도 「맑은 고딕」 명시 — {n_ea}곳', '글꼴 대체로 줄바꿈이 달라지는 것 방지'))

# ---------- (0) 본 슬라이드(1~29) 표 글자 14pt 이상 ----------
for i, s in enumerate(prs.slides, 1):
    if i > 29: continue
    for sh in s.shapes:
        if not sh.has_table: continue
        n = 0; mn = 99
        for row in sh.table.rows:
            for cell in row.cells:
                for p in cell.text_frame.paragraphs:
                    for r in p.runs:
                        if r.font.size and r.font.size.pt < 14:
                            mn = min(mn, r.font.size.pt); r.font.size = Pt(14); n += 1
        if n: LOG.append((i, f'표[{sh.shape_id}] 글자', f'{mn:.0f}~13pt', '14pt (이사님: 표 글자 14pt 이상)', ''))


# ---------- (0b) 13번 표 열 너비: 「판정 등급」 -0.15in → 「해당 조항」 +0.15in ----------
for sh in prs.slides[12].shapes:
    if sh.has_table and len(sh.table.columns) == 3:
        c = sh.table.columns; d = 137160
        c[0].width = Emu(c[0].width - d); c[2].width = Emu(c[2].width + d)
        LOG.append((13, '표[9] 열 너비', '판정 등급 / 해당 조항', '판정 등급 -0.15in, 해당 조항 +0.15in', '조항 번호 목록이 14pt에서 덜 접히도록'))


# ---------- (0c) 13번 각주 12 → 11pt ----------
for sh in prs.slides[12].shapes:
    if sh.shape_id == 11 and sh.has_text_frame:
        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                if r.font.size and r.font.size.pt > 11: r.font.size = Pt(11)
        LOG.append((13, '[11] 각주 글자', '12pt', '11pt', '표 14pt로 커진 자리 확보'))


# ---------- (0d) 17번: 각주 12 → 11pt, 아래 분홍 상자 글 14 → 13pt ----------
for sh in prs.slides[16].shapes:
    if sh.shape_id in (18, 19) and sh.has_text_frame:
        tgt = 11 if sh.shape_id == 19 else 13; mx = 0
        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                if r.font.size and r.font.size.pt > tgt: mx = max(mx, r.font.size.pt); r.font.size = Pt(tgt)
        if mx: LOG.append((17, f'[{sh.shape_id}] 글자', f'{mx:.0f}pt', f'{tgt}pt', '표 14pt로 커진 자리 확보'))

# ---------- (a) 제목 ----------
for i, s in enumerate(prs.slides, 1):
    title = label = None
    for sh in s.shapes:
        if sh.shape_id == 2 and sh.has_text_frame and sh.top < 300000 and sh.height < 900000: title = sh
        if sh.shape_id == 3 and sh.has_text_frame and sh.top < 400000 and sh.left > 5000000: label = sh
    if not title: continue
    runs = [r for p in title.text_frame.paragraphs for r in p.runs]
    if not runs or runs[0].font.size is None: continue
    sz = runs[0].font.size.pt
    if sz != 27: continue
    if label:
        for p in label.text_frame.paragraphs: p.alignment = PP_ALIGN.RIGHT
        lab_w = em(label.text_frame.text) * 12 * 12700 + 2 * 91440
        avail = (label.left + label.width - lab_w) - title.left - 2 * 91440 - 91440
    else:
        avail = title.width - 2 * 91440
    t = title.text_frame.text
    need = em(t) * 1.06 * sz * 12700
    if need > avail:
        new = max(20, math.floor(avail / (em(t) * 1.06 * 12700)))
        for r in runs: r.font.size = Pt(new)
        LOG.append((i, '제목 글꼴', f'{sz:.0f}pt', f'{new}pt (한 줄에 맞춤)', t))

# ---------- (b) 표 행 높이 ----------
from PIL import ImageFont
_FD = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'fonts/')
_FC = {}
def _font(bold, pt):
    k = (bold, pt)
    if k not in _FC: _FC[k] = ImageFont.truetype(_FD + ('NotoSansCJKkr-Bold.otf' if bold else 'NotoSansCJKkr-Regular.otf'), int(pt * 10))
    return _FC[k]
def wrap_lines(txt, bold, pt, inner_pt):
    """어절(공백) 단위 줄바꿈, 어절이 한 줄보다 길면 글자 단위. 폭은 Noto Sans CJK KR 실측 + 6% 여유."""
    f = _font(bold, pt); W = inner_pt * 10 / 1.08
    def w(s):
        # 한글↔라틴·숫자 경계의 자동 간격(PowerPoint·LibreOffice 공통) ≈ 0.25em
        b = sum(1 for a, c in zip(s, s[1:]) if (unicodedata.east_asian_width(a) in 'WF') != (unicodedata.east_asian_width(c) in 'WF') and a != ' ' and c != ' ')
        return f.getlength(s) + b * pt * 10 * 0.25
    lines = 1; cur = ''
    for word in re.split(r'(?<= )', txt):
        if w(cur + word.rstrip()) <= W: cur += word; continue
        if cur: lines += 1; cur = ''
        while w(word.rstrip()) > W:   # 긴 어절은 글자 단위
            k = len(word)
            while k > 1 and w(word[:k]) > W: k -= 1
            word = word[k:]; lines += 1
        cur = word
    return lines
def cell_need(cell, colw):
    tf = cell.text_frame
    mL = cell.margin_left if cell.margin_left is not None else 91440
    mR = cell.margin_right if cell.margin_right is not None else 91440
    mT = cell.margin_top if cell.margin_top is not None else 45720
    mB = cell.margin_bottom if cell.margin_bottom is not None else 45720
    inner = (colw - mL - mR) / 12700.0  # pt
    h = 0.0
    for p in tf.paragraphs:
        sz = None; bold = False
        for r in p.runs:
            if r.font.size: sz = r.font.size.pt; bold = bool(r.font.bold); break
        if sz is None: sz = 18
        txt = ''.join(r.text for r in p.runs)
        lines = wrap_lines(txt, bold, sz, inner) if inner > 0 else 1
        sa = p.space_after.pt if p.space_after is not None else 0
        sb = p.space_before.pt if p.space_before is not None else 0
        h += lines * sz * 1.36 + sa + sb
    return int(h * 12700 + mT + mB)

for i, s in enumerate(prs.slides, 1):
    tables = [sh for sh in s.shapes if sh.has_table]
    for sh in tables:
        t = sh.table; colw = [c.width for c in t.columns]
        old_h = sum(r.height for r in t.rows); changed = []
        for ri, row in enumerate(t.rows):
            need = 0
            for ci, cell in enumerate(row.cells):
                if cell.is_spanned: continue
                w = colw[ci]
                if cell.is_merge_origin: w = sum(colw[ci:ci + cell.span_width])
                need = max(need, cell_need(cell, w))
            if need > row.height:
                changed.append((ri + 1, row.height, need)); row.height = Emu(need)
        new_h = sum(r.height for r in t.rows); delta = new_h - old_h
        sh.height = Emu(new_h)   # 표 도형 높이도 행 합계로(아래 도형 배치 계산에 쓰임)
        if delta > 0:
            bottom_old = sh.top + old_h
            moved = []
            below = [o for o in s.shapes if o is not sh and o.top > sh.top + old_h / 2 and o.top < 6900000]
            if below:
                gap = min(o.top for o in below) - bottom_old; d = max(0, delta - max(0, gap - 90000))
                for o in below: o.top = Emu(o.top + d); moved.append(o.shape_id)
                delta = d
            LOG.append((i, f'표[{sh.shape_id}] 행 높이', f'{old_h/914400:.2f}in', f'{new_h/914400:.2f}in (행 {[c[0] for c in changed]}), 아래 도형 {moved} {delta/914400:.2f}in 내림', ''))


# ---------- (c) 본 슬라이드 그림 키우기(세로 여유 활용) + 표지 오른쪽 그림 전체 높이 ----------
W1, H1 = prs.slide_width, prs.slide_height
s1 = prs.slides[0]
for sh in s1.shapes:
    if sh.shape_type == 13:
        ar = sh.width / sh.height; sh.height = Emu(H1); sh.width = Emu(int(H1 * ar)); sh.left = Emu(W1 - sh.width); sh.top = Emu(0); pic_left = sh.left
for sh in s1.shapes:
    if sh.shape_id == 3: sh.width = Emu(pic_left - sh.left - 180000)
LOG.append((1, '표지 그림', '세로 6.0in(가운데)', '세로 전체(7.56in·비율 유지)', 'A4에서 위아래 흰 띠 제거'))
for sn in (11, 12, 23):
    s = prs.slides[sn - 1]
    pics = [sh for sh in s.shapes if sh.shape_type == 13]
    if not pics: continue
    pic = max(pics, key=lambda p: p.width * p.height)
    top_lim = 1300000; bot_lim = H1 - 700000
    for sh in s.shapes:
        if sh is pic or not sh.has_text_frame or not sh.text_frame.text.strip(): continue
        if sh.top >= pic.top + pic.height - 50000 and sh.top < bot_lim: bot_lim = min(bot_lim, sh.top - 80000)
        if sh.top + sh.height <= pic.top + 50000 and sh.top + sh.height > top_lim and sh.left < pic.left + pic.width and sh.left + sh.width > pic.left: top_lim = max(top_lim, sh.top + sh.height + 60000)
    ar = pic.width / pic.height
    h = bot_lim - top_lim; w = int(h * ar); maxw = W1 - 2 * 500000
    if w > maxw: w = maxw; h = int(w / ar)
    if h > pic.height * 1.03:
        cx = pic.left + pic.width / 2
        old = f'{pic.width/914400:.2f}×{pic.height/914400:.2f}in'
        pic.width = Emu(w); pic.height = Emu(h); pic.top = Emu(top_lim + (bot_lim - top_lim - h) // 2); pic.left = Emu(int(cx - w / 2))
        LOG.append((sn, f'[{pic.shape_id}] 그림 크기', old, f'{w/914400:.2f}×{h/914400:.2f}in', '그림 유지, A4 세로 여유만큼 확대'))



def tf_need(sh):
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
        txt = ''.join(r.text for r in p.runs)
        lines = wrap_lines(txt, bold, sz, inner) if txt.strip() else 1
        sa = p.space_after.pt if p.space_after is not None else 0
        sb = p.space_before.pt if p.space_before is not None else 0
        h += lines * sz * 1.36 + sa + sb
    return int(h * 12700 + mT + mB)

LIMIT = 6950000; TOP0 = 1380000
# ---------- (b2) 본 슬라이드: 배경 있는 도형(자동 도형)의 글이 상자보다 길면 세로를 키움 ----------
for i, s in enumerate(prs.slides, 1):
    if i > 29: continue
    for sh in list(s.shapes):
        if sh.shape_type != 1 or not sh.has_text_frame or not sh.text_frame.text.strip(): continue
        if sh.top > 6900000: continue
        need = int(tf_need(sh) * 1.05)
        if need > sh.height:
            d = need - sh.height; bottom_old = sh.top + sh.height; sh.height = Emu(need)
            below = [o for o in s.shapes if o is not sh and o.top >= bottom_old - 30000 and o.top < 6900000 and o.left < sh.left + sh.width and o.left + o.width > sh.left]
            if below:
                gap = min(o.top for o in below) - bottom_old; dd = max(0, d - max(0, gap - 60000))
                for o in below: o.top = Emu(o.top + dd)
            LOG.append((i, f'[{sh.shape_id}] 상자 세로', f'{(need-d)/914400:.2f}in', f'{need/914400:.2f}in (글 양에 맞춤)', 'A4에서 가로가 좁아져 줄이 늘어남'))
# 13번 분홍 상자: 왼쪽 빈 자리(Q1~Q4 아래)로 올림
s13 = prs.slides[12]; _m = {sh.shape_id: sh for sh in s13.shapes}; q = _m[8]; pink = _m[10]
pink.top = Emu(q.top + q.height + 250000)
LOG.append((13, '[10] 분홍 상자 위치', '표 아래 끝 높이', 'Q1~Q4 글 바로 아래', '표가 길어져 각주와 겹치지 않도록'))

# ---------- (d) 본 슬라이드 세로 넘침 압축 ----------
for i, s in enumerate(prs.slides, 1):
    if i > 29: continue
    content = []
    for sh in s.shapes:
        if sh.shape_id in (2, 3, 4): continue
        if sh.has_text_frame and sh.text_frame.text.startswith('▶'): continue
        if sh.top > 6900000: continue
        content.append(sh)
    if not content: continue
    max_bottom = max(sh.top + sh.height for sh in content)
    if max_bottom <= LIMIT: continue
    need_total = max_bottom - LIMIT
    # 세로로 겹치는 도형끼리 묶어 위에서 아래로 띠(cluster) 구성
    content.sort(key=lambda sh: sh.top)
    clusters = []
    for sh in content:
        if clusters and sh.top < clusters[-1]['bottom'] - 20000:
            clusters[-1]['shapes'].append(sh); clusters[-1]['bottom'] = max(clusters[-1]['bottom'], sh.top + sh.height)
        else:
            clusters.append({'shapes': [sh], 'top': sh.top, 'bottom': sh.top + sh.height})
    saved = 0
    # 1) 맨 위 띠를 TOP0까지 올림
    up = max(0, clusters[0]['top'] - TOP0)
    if up:
        for c in clusters:
            for sh in c['shapes']: sh.top = Emu(sh.top - up)
            c['top'] -= up; c['bottom'] -= up
        saved += up
    # 2) 띠 사이 틈을 60k EMU까지 줄임
    for j in range(1, len(clusters)):
        if saved >= need_total: break
        gap = clusters[j]['top'] - clusters[j - 1]['bottom']; red = min(max(0, gap - 60000), need_total - saved)
        if red > 0:
            for c in clusters[j:]:
                for sh in c['shapes']: sh.top = Emu(sh.top - red)
                c['top'] -= red; c['bottom'] -= red
            saved += red
    # 3) 글이 상자를 다 채우지 않는 도형의 세로를 줄임(표·그림 제외)
    for j, c in enumerate(clusters):
        if saved >= need_total: break
        slack = 0
        for sh in c['shapes']:
            if sh.shape_type == 17 and sh.text_frame.text.strip():   # 글상자만(배경 있는 도형은 글이 넘칠 수 있어 제외)
                sl = max(0, sh.height - int(tf_need(sh) * 1.15)); slack = max(slack, sl)
        if slack <= 0: continue
        red = min(slack, need_total - saved)
        for sh in c['shapes']:
            if sh.shape_type == 17 and sh.text_frame.text.strip():
                sh.height = Emu(max(int(tf_need(sh) * 1.15), sh.height - red))
        c['bottom'] = max(sh.top + sh.height for sh in c['shapes'])
        for c2 in clusters[j + 1:]:
            for sh in c2['shapes']: sh.top = Emu(sh.top - red)
            c2['top'] -= red; c2['bottom'] -= red
        saved += red
    max_bottom2 = max(sh.top + sh.height for sh in content)
    LOG.append((i, '세로 압축', f'내용 아래 끝 {max_bottom/914400:.2f}in', f'{max_bottom2/914400:.2f}in (한계 {LIMIT/914400:.2f}in) — 위로 {up/914400:.2f}in, 틈·여백 {(saved-up)/914400:.2f}in', '표 글자 14pt로 커진 만큼 간격을 줄임' + ('' if max_bottom2 <= LIMIT + 20000 else ' — ※ 아직 넘침, 수동 확인')))

prs.save(OUT)
json.dump(LOG, open(os.path.splitext(OUT)[0] + '_fitlog.json', 'w'), ensure_ascii=False, indent=1, default=str)
for l in LOG: print(l)
