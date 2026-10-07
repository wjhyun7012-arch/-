"""v17 → v18 (이사님 10/7 여섯 가지).
1 4번 제목 (linkage) 뺌 / 2 5번 표 4행 → 2행 / 3 13번 위아래 배치(Q 가로 네 칸 → 네 단 막대 가로) / 4 21번 설명 줄·표 내림 /
5 23번 ▶ 「판정은 점수의 크기가 아니라 핵심 문항의 충족 여부로 정해진다」 / 6 6번 상태 B 쉬운 문구.
사용: python3 -I mk_deck18.py v17.pptx v18.pptx
"""
import sys, os, json, copy
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree
from pptlib import _rep_tf
SRC, OUT = sys.argv[1], sys.argv[2]
prs = Presentation(SRC); W, H = prs.slide_width, prs.slide_height
IN = 914400; L = 497190; RW = W - 2 * L; LOG = []
NAVY = RGBColor(0x1F, 0x38, 0x64); LIGHT = RGBColor(0xE8, 0xEE, 0xF7); RED = RGBColor(0xC0, 0x00, 0x00); PINK = RGBColor(0xFB, 0xEC, 0xEC)
INK = RGBColor(0x11, 0x11, 0x11); GREY = RGBColor(0x59, 0x59, 0x59); WHITE = RGBColor(0xFF, 0xFF, 0xFF)
def set_ea(run):
    rpr = run._r.get_or_add_rPr()
    if rpr.find(qn('a:ea')) is None:
        ea = etree.SubElement(rpr, qn('a:ea')); ea.set('typeface', '맑은 고딕')
def par(tf, text, size, bold=False, color=INK, align=None, first=False, after=None, before=None):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    if align is not None: p.alignment = align
    if after is not None: p.space_after = Pt(after)
    if before is not None: p.space_before = Pt(before)
    r = p.add_run(); r.text = text; r.font.size = Pt(size); r.font.bold = bold; r.font.name = '맑은 고딕'; r.font.color.rgb = color; set_ea(r); return p
def box(s, x, y, w, h, fill, line=None, shape=MSO_SHAPE.ROUNDED_RECTANGLE, name='', anchor=MSO_ANCHOR.MIDDLE, margin=90000):
    sh = s.shapes.add_shape(shape, Emu(int(x)), Emu(int(y)), Emu(int(w)), Emu(int(h))); sh.name = name
    if fill is None: sh.fill.background()
    else: sh.fill.solid(); sh.fill.fore_color.rgb = fill
    if line is None: sh.line.fill.background()
    else: sh.line.color.rgb = line; sh.line.width = Pt(1.0)
    sh.shadow.inherit = False
    tf = sh.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Emu(margin); tf.margin_top = tf.margin_bottom = Emu(45000); return sh
def tbox(s, x, y, w, h, name=''):
    tb = s.shapes.add_textbox(Emu(int(x)), Emu(int(y)), Emu(int(w)), Emu(int(h))); tb.name = name
    tf = tb.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_right = Emu(45720); tf.margin_top = tf.margin_bottom = Emu(20000); return tb
def shp(s, sid):
    for sh in s.shapes:
        if sh.shape_id == sid: return sh
    raise KeyError(sid)
def byname(s, name): return [sh for sh in s.shapes if sh.name == name][0]
def remove(s, *shs):
    for sh in shs: sh._element.getparent().remove(sh._element)
def rep(sn, sh, old, new, why):
    n = _rep_tf(sh.text_frame, old, new); assert n == 1, (sn, old, n); LOG.append((sn, sh.name, old, new, why))
def rep_notes(sn, old, new, why):
    n = _rep_tf(prs.slides[sn - 1].notes_slide.notes_text_frame, old, new); assert n == 1, (sn, old, n); LOG.append((sn, '노트', old, new, why))

# ===== 1) 4번 제목 =====
s = prs.slides[3]; rep(4, shp(s, 2), '두 표준과 연계(linkage)', '두 표준과 연계', '이사님 10/7: 제목이 장 표시와 겹침 — 작은 글 첫 줄에만 (linkage)')

# ===== 6) 6번 상태 B =====
s = prs.slides[5]
rep(6, shp(s, 9), '상태 B — 그 결과(LCA 영향평가)를 근거로 중대한 환경측면을 결정하고 환경목표를 정하며, 다음 해에 다시 측정해 확인', '상태 B — LCA 결과를 받아 중대한 환경측면을 정하고, 환경목표를 세우고, 다음 해에 다시 측정해 확인', '이사님 10/7: 너무 어려움 — 쉬운 말(주어는 회사)')
rep_notes(6, '오른쪽 상태 B는 LCA 영향평가 결과를 근거로 중대한 환경측면을 결정하고, 환경목표를 세우고, 다음 해에 다시 측정해 확인하는 회사입니다.', '오른쪽 상태 B는 LCA 결과를 받아 중대한 환경측면을 정하고, 환경목표를 세우고, 다음 해에 다시 측정해 확인하는 회사입니다.', '같은 이유')

# ===== 2) 5번 표 4행 → 2행 =====
s = prs.slides[4]; tb = shp(s, 8); t = tb.table
old_rows = [[c.text_frame.text for c in r.cells] for r in t.rows]
for ri in (3, 1):   # 「경영시스템과의 통합·실행 구조」, 「접근의 성격」 삭제
    tr = t._tbl.tr_lst[ri]; tr.getparent().remove(tr)
for r in list(t.rows)[1:]: r.height = Emu(int(r.height * 1.35))
for r in t.rows:
    for c in r.cells:
        for p in c.text_frame.paragraphs:
            for run in p.runs:
                if run.font.size and run.font.size.pt < 14: run.font.size = Pt(14)
th = sum(r.height for r in t.rows); tb.height = Emu(th)
y = tb.top + th + 250000
for sid in (9, 10): shp(s, sid).top = Emu(y)
shp(s, 11).top = Emu(y + int(1.10 * IN) + 150000)
LOG.append((5, '표', '4행: 접근의 성격 / 정량 평가 / 통합·실행 구조 / 지속성', '2행: 전과정 환경영향의 정량 평가 / 지속성(반복 수행) — ▶ 줄 「한쪽은 방법이 없고 다른 쪽은 주기가 없다」와 1:1', '이사님 10/7: 4번과 겹치고 복잡 — 뺀 두 행은 4번 상자·백업 B-3·B-4에 있음'))

# ===== 3) 13번 위아래 배치 =====
s = prs.slides[12]
remove(s, shp(s, 8), byname(s, 'q note'), *[byname(s, n) for n in ('bar lab0', 'bar0', 'bar lab1', 'bar1', 'bar lab2', 'bar2', 'bar lab3', 'bar3', 'map foot')])
y = int(1.29 * IN)
t = tbox(s, L, y, RW, 320000, 'q head'); par(t.text_frame, '네 가지 판정 질문 — 30개 조항 전부에 같은 순서로', 13.5, True, NAVY, first=True); y += 340000
qs = [('Q1', '전과정 관점을 요구하거나 그 결과를 직접 쓰는가'), ('Q2', 'LCA 없이는 의무(shall)를 다 지킬 수 없는가'), ('Q3', 'LCA 결과가 주요 입력인가'), ('Q4', '보조 정보라도 주는가')]
qg = 150000; qw = (RW - 3 * qg) // 4; qh = int(1.0 * IN)
for i, (n, q) in enumerate(qs):
    b = box(s, L + i * (qw + qg), y, qw, qh, LIGHT, name=f'q{i}', anchor=MSO_ANCHOR.MIDDLE, margin=70000); par(b.text_frame, n, 15, True, NAVY, PP_ALIGN.CENTER, first=True, after=2); par(b.text_frame, q, 12, False, INK, PP_ALIGN.CENTER)
y += qh + 120000
t = tbox(s, L, y, RW, 320000, 'q to'); par(t.text_frame, '▼  「예」가 앞쪽 물음에서 나올수록 연계 가능성이 높다 — 네 등급으로 매김', 12, False, GREY, PP_ALIGN.CENTER, first=True); y += 360000
bars = [('매우 높음', 4, '6.1.2 환경측면 · 6.2.1 환경목표 · 8.1 운용기획 및 관리 · 9.1.1 모니터링  —  뒤의 핵심 문항 4종의 자리', True),
        ('높음', 12, 'LCA 결과가 주요 입력 — 없이도 되지만 품질이 떨어짐', False),
        ('중간', 10, 'LCA는 보조 정보', False),
        ('낮음', 4, '조항 본질이 LCA 영역 밖', False)]
bh = int(0.66 * IN); gap = 80000; lw = int(1.6 * IN); maxw = RW - lw - 100000
for i, (nm, n, desc, hot) in enumerate(bars):
    lab = box(s, L, y, lw, bh, RED if hot else NAVY, name=f'bar lab{i}', margin=60000); par(lab.text_frame, f'{nm}  {n}', 14, True, WHITE, PP_ALIGN.CENTER, first=True)
    bw = int(maxw * (0.5 + 0.5 * n / 12)); b = box(s, L + lw + 100000, y, bw, bh, PINK if hot else LIGHT, name=f'bar{i}', margin=100000)
    par(b.text_frame, desc, 13 if hot else 12.5, hot, RED if hot else INK, PP_ALIGN.LEFT, first=True)
    y += bh + gap
y += 60000
t = tbox(s, L, y, RW, int(0.6 * IN), 'map foot'); par(t.text_frame, 'ISO 14001 30개 조항(7.4.1~7.4.3은 7.4로 통합)을 전부 판정 — 55문항 중 44는 ISO 14001 조항에, 11은 ISO 14044 고유 요건에 대응. 조항별 판정 전문은 백업 B-1(흐름)·B-1a(등급표)·부록 A', 11.5, False, GREY, PP_ALIGN.LEFT, first=True)
LOG.append((13, '배치', '왼쪽 Q 상자(작음) + 오른쪽 세로 막대 넷', '위: Q1~Q4 가로 네 칸 → 가운데: 네 등급 화살표 → 아래: 네 단 막대 가로(매우 높음 행 폭 전체) → 맨 아래 한 줄', '이사님 10/7: 배치 구성'))

# ===== 4) 21번 설명 줄·표 내림 =====
s = prs.slides[20]; d = int(0.3 * IN)
note = byname(s, 'obj note'); note.top = Emu(shp(s, 8).top + shp(s, 8).height + 200000); note.height = Emu(int(0.5 * IN))
tb = shp(s, 10); tb.top = Emu(note.top + note.height + 80000)
for r in tb.table.rows: r.height = Emu(int(r.height * 0.93))
for r in tb.table.rows:
    for c in r.cells:
        for pp in c.text_frame.paragraphs:
            for run in pp.runs:
                if run.font.size and run.font.size.pt > 13: run.font.size = Pt(13); shp(s, 11).top = Emu(tb.top + sum(r.height for r in tb.table.rows) + 30000)
LOG.append((21, '배치', '설명 줄이 상자에 붙음(2.70in), 표 3.30in', f'설명 줄 {note.top/IN:.2f}in, 표 {tb.top/IN:.2f}in', '이사님 10/7: 한 칸 반 아래로'))

# ===== 5) 23번 ▶ 줄 =====
s = prs.slides[22]; rep(23, shp(s, 5), '판정을 가른 것은 점수의 크기가 아니라 필요조건의 충족 여부다.', '판정은 점수의 크기가 아니라 핵심 문항의 충족 여부로 정해진다.', '이사님 10/7: 「가른 것」 표현')
rep_notes(23, '판정을 가른 것은 점수의 크기가 아니라 필요조건의 충족 여부입니다.', '판정은 점수의 크기가 아니라 핵심 문항의 충족 여부로 정해집니다.', '같은 이유')
# 16번 예 문장에도 「판정을 가른 문항」 → 「판정을 정한 문항」
s = prs.slides[15]; rep(16, byname(s, 'sc ex'), '4장에서 판정을 가른 문항', '4장에서 판정을 정한 문항', '같은 이유')
rep_notes(16, '이 문항이 4장에서 판정을 갈랐습니다', '이 문항이 4장에서 판정을 정하였습니다', '같은 이유')
s = prs.slides[20]; rep_notes(21, 'D-1 행이 뒤의 판정을 가릅니다.', 'D-1 행이 뒤의 판정을 정합니다.', '같은 이유')
prs.save(OUT); json.dump(LOG, open('log18.json', 'w'), ensure_ascii=False, indent=1); print(len(LOG), 'changes')
