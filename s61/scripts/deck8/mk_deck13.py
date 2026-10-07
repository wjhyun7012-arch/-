"""v12b(16·17번 사본을 41·43번에 넣은 것) → v13: 본 슬라이드 8장을 「쉽게, 빠지는 것 없이」(이사님 10/7 권고안).
4 문구 논문 표현 / 12 모델링 한 그림 / 14 핵심 연계 지점 4곳 설명 / 16 그림 크게+점수 매기는 법 / 17 여섯 수준 계단 /
18 카드 셋 / 19 평가 한 그림 / 20 상자 둘 세 줄씩 / 41·43 = 옮긴 표(B-10a·B-11a) / 30 백업 문구 / 쪽번호 재부여.
사용: python3 -I mk_deck13.py v12b.pptx v13.pptx
"""
import sys, os, json, copy, re
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree
from pptx.text.text import _Paragraph
from pptlib import _rep_tf
SRC, OUT = sys.argv[1], sys.argv[2]
prs = Presentation(SRC); W, H = prs.slide_width, prs.slide_height
LOG = []
IN = 914400
NAVY = RGBColor(0x1F, 0x38, 0x64); LIGHT = RGBColor(0xE8, 0xEE, 0xF7); RED = RGBColor(0xC0, 0x00, 0x00); PINK = RGBColor(0xFB, 0xEC, 0xEC)
INK = RGBColor(0x11, 0x11, 0x11); GREY = RGBColor(0x59, 0x59, 0x59); GREEN = RGBColor(0x2E, 0x6B, 0x4A); LGREEN = RGBColor(0xEA, 0xF4, 0xEE)
WHITE = RGBColor(0xFF, 0xFF, 0xFF); AMBER = RGBColor(0xB2, 0x6B, 0x00); LAMBER = RGBColor(0xFD, 0xF3, 0xE1); MID = RGBColor(0x5B, 0x7F, 0xB5)
L = 497190; RW = W - 2 * L

def set_ea(run):
    rpr = run._r.get_or_add_rPr()
    if rpr.find(qn('a:ea')) is None:
        ea = etree.SubElement(rpr, qn('a:ea')); ea.set('typeface', '맑은 고딕')
def par(tf, text, size, bold=False, color=INK, align=None, first=False, after=None, before=None):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    if align is not None: p.alignment = align
    if after is not None: p.space_after = Pt(after)
    if before is not None: p.space_before = Pt(before)
    r = p.add_run(); r.text = text; r.font.size = Pt(size); r.font.bold = bold; r.font.name = '맑은 고딕'; r.font.color.rgb = color; set_ea(r)
    return p
def box(s, x, y, w, h, fill, line=None, shape=MSO_SHAPE.ROUNDED_RECTANGLE, name='', anchor=MSO_ANCHOR.MIDDLE, margin=90000):
    sh = s.shapes.add_shape(shape, Emu(int(x)), Emu(int(y)), Emu(int(w)), Emu(int(h))); sh.name = name
    if fill is None: sh.fill.background()
    else: sh.fill.solid(); sh.fill.fore_color.rgb = fill
    if line is None: sh.line.fill.background()
    else: sh.line.color.rgb = line; sh.line.width = Pt(1.0)
    sh.shadow.inherit = False
    tf = sh.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Emu(margin); tf.margin_top = tf.margin_bottom = Emu(45000)
    return sh
def tbox(s, x, y, w, h, name=''):
    tb = s.shapes.add_textbox(Emu(int(x)), Emu(int(y)), Emu(int(w)), Emu(int(h))); tb.name = name
    tf = tb.text_frame; tf.word_wrap = True; tf.margin_left = tf.margin_right = Emu(45720); tf.margin_top = tf.margin_bottom = Emu(20000)
    return tb
def arrow_down(s, cx, y, h=200000, name='arr'):
    a = s.shapes.add_shape(MSO_SHAPE.DOWN_ARROW, Emu(int(cx - 110000)), Emu(int(y)), Emu(220000), Emu(int(h))); a.name = name
    a.fill.solid(); a.fill.fore_color.rgb = RGBColor(0x8A, 0x8A, 0x8A); a.line.fill.background(); a.shadow.inherit = False; return a
def arrow_right(s, x, cy, w=260000, name='arr'):
    a = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Emu(int(x)), Emu(int(cy - 100000)), Emu(int(w)), Emu(200000)); a.name = name
    a.fill.solid(); a.fill.fore_color.rgb = RGBColor(0x8A, 0x8A, 0x8A); a.line.fill.background(); a.shadow.inherit = False; return a
def shp(s, sid):
    for sh in s.shapes:
        if sh.shape_id == sid: return sh
    raise KeyError(sid)
def remove(s, *sids):
    for sid in sids:
        sh = shp(s, sid); sh._element.getparent().remove(sh._element)
def set_text_keep(sh, text):
    tf = sh.text_frame; p = tf.paragraphs[0]; rs = p.runs; rs[0].text = text
    for r in rs[1:]: r._r.getparent().remove(r._r)
    for extra in tf.paragraphs[1:]: extra._p.getparent().remove(extra._p)
def set_hd(s, text):
    hd = [sh for sh in s.shapes if sh.has_text_frame and sh.text_frame.text.startswith('▶')][0]
    p = hd.text_frame.paragraphs[0]; old = hd.text_frame.text; p.runs[-1].text = text
    for r in p.runs[1:-1]: r._r.getparent().remove(r._r)
    return old
def set_notes(sn, txt):
    tf = prs.slides[sn - 1].notes_slide.notes_text_frame; old = tf.text
    if tf.paragraphs and tf.paragraphs[0].runs: tmpl = copy.deepcopy(tf.paragraphs[0]._p)
    else: tmpl = None
    for p in list(tf.paragraphs): p._p.getparent().remove(p._p)
    for line in txt.split('\n'):
        if tmpl is not None:
            np_ = copy.deepcopy(tmpl); tf._txBody.append(np_); para = _Paragraph(np_, tf); rs = para.runs
            rs[0].text = line
            for r in rs[1:]: r._r.getparent().remove(r._r)
        else:
            para = tf.add_paragraph(); r = para.add_run(); r.text = line; r.font.size = Pt(12); r.font.name = '맑은 고딕'; set_ea(r)
    return old
def rep(s, sid, old, new, why, sn):
    n = _rep_tf(shp(s, sid).text_frame, old, new); assert n == 1, (sn, sid, old, n); LOG.append((sn, f'[{sid}]', old, new, why))
def rep_notes(sn, old, new, why):
    n = _rep_tf(prs.slides[sn - 1].notes_slide.notes_text_frame, old, new); assert n == 1, (sn, old, n); LOG.append((sn, '노트', old, new, why))
def label_backup(s, text):
    lb = shp(s, 3); set_text_keep(lb, text); lb.fill.background()
    for p in lb.text_frame.paragraphs:
        for r in p.runs: r.font.size = Pt(12); r.font.bold = False; r.font.color.rgb = RGBColor(0x7F, 0x7F, 0x7F)

# ================= 4번 문구 =================
s = prs.slides[3]
old = set_hd(s, '관리하는 틀과 측정하는 방법, 이 둘의 연계를 조항 단위로 확인하고 성숙도 수준으로 판정한다.'); LOG.append((4, '▶', old, '관리하는 틀과 측정하는 방법, 이 둘의 연계를 조항 단위로 확인하고 성숙도 수준으로 판정한다.', '논문 5장 결론 문장'))
kl = [sh for sh in s.shapes if sh.name == 'key line'][0]
rep(s, kl.shape_id, '둘은 서로를 전제하지 않아 따로 돈다 — 얼마나 이어졌는지를 측정하고 판정하는 모델, 이것이 이 연구가 만든 것이다', '두 표준은 서로를 전제하지 않아 실무에서는 별도로 운영되는 경우가 많다 — 두 표준의 연계를 조항 단위로 확인하고 성숙도 수준으로 판정하는 모델, 이것이 이 연구가 만든 것이다', '논문 1.1·5장 표현', 4)
sl = [sh for sh in s.shapes if sh.name == 'sub lines'][0]
rep(s, sl.shape_id, '이어졌다는 것은 — LCA 결과가 환경측면의 결정·영향평가·환경목표·설계·점검에 실제로 쓰이는 것', '연계란 — LCA 결과가 환경측면의 결정·환경목표·설계·개발·모니터링에 쓰이는 것 (핵심 연계 지점 4곳)', '핵심 연계 지점 이름으로', 4)
rep_notes(4, '이 둘은 서로를 전제하지 않아서 현장에서는 따로 돕니다. LCA 보고서를 받아 놓고 경영시스템에서는 쓰지 않는 식입니다.', '두 표준은 서로를 전제하지 않아, 실무에서는 별도로 운영되는 경우가 많습니다. LCA 보고서를 받아 놓고 경영시스템에서는 쓰지 않는 식입니다.', '논문 표현')
rep_notes(4, '이 연구는 둘이 얼마나 이어졌는지를 측정하고 판정하는 모델을 만든 것입니다. 이어졌다는 것은 LCA 결과가 환경측면의 결정과 영향평가, 환경목표, 설계, 점검에 실제로 쓰인다는 뜻이고, 판정 결과는 여섯 수준으로 나옵니다.', '이 연구는 두 표준의 연계를 조항 단위로 확인하고 성숙도 수준으로 판정하는 모델을 만든 것입니다. 연계란 LCA 결과가 환경측면의 결정, 환경목표, 설계·개발, 모니터링에 쓰인다는 뜻이고, 판정 결과는 여섯 수준으로 나옵니다.', '논문 표현')

# ================= 12번 모델링 한 그림 =================
s = prs.slides[11]; remove(s, 8)
set_text_keep(shp(s, 2), '연계 성숙도 평가 모델(LMA)을 한 그림으로'); set_hd(s, '두 표준을 맞대어 업무 목록 → 채점표 → 수준 규칙의 순서로 만들었다.'); set_text_keep(shp(s, 7), '근거  3.1.2, 3.2~3.5, [그림 3-2]')
y = 1250000; bw = 3100000; bh = 760000; lx = L + 600000; rx = L + RW - 600000 - bw
b = box(s, lx, y, bw, bh, LIGHT, NAVY, name='m std1'); par(b.text_frame, '환경경영시스템(ISO 14001)', 15, True, NAVY, PP_ALIGN.CENTER, first=True); par(b.text_frame, '조항 30개 (4~10절)', 12, False, INK, PP_ALIGN.CENTER)
b = box(s, rx, y, bw, bh, LGREEN, GREEN, name='m std2'); par(b.text_frame, '전과정평가(LCA, ISO 14044)', 15, True, GREEN, PP_ALIGN.CENTER, first=True); par(b.text_frame, '요구사항 (shall·should)', 12, False, INK, PP_ALIGN.CENTER)
mx0 = lx + bw; mx1 = rx; cy = y + bh // 2
a = s.shapes.add_shape(MSO_SHAPE.LEFT_RIGHT_ARROW, Emu(mx0 + 150000), Emu(cy - 130000), Emu(mx1 - mx0 - 300000), Emu(260000)); a.fill.solid(); a.fill.fore_color.rgb = MID; a.line.fill.background(); a.shadow.inherit = False; a.name = 'm map arrow'
t = tbox(s, mx0, y - 330000, mx1 - mx0, 300000, 'm map label'); par(t.text_frame, '매핑 — 조항을 하나하나 대응', 12, True, NAVY, PP_ALIGN.CENTER, first=True)
cx = L + RW // 2; y += bh
t = tbox(s, cx - 2200000, y + 30000, 4400000, 320000, 'm map note'); par(t.text_frame, '맞닿는 곳 = 연계 지점 · 그 가운데 필수인 네 곳 = 핵심 연계 지점', 11.5, False, GREY, PP_ALIGN.CENTER, first=True)
y += 360000; arrow_down(s, cx, y, 220000); y += 260000
rows = [('① 참조 모델', '평가할 업무 목록', '연계 프로세스 16종', '연계 지점을 업무 단위로 묶고, 제목·목적·성과·활동으로 적음'),
        ('② 평가 모델', '채점표', '평가 문항 55 · 문항별 판정 기준 · 증거 목록', '업무마다 「되어 있는가」를 묻고, 확인 항목 개수로 1~5점'),
        ('③ 성숙도 모델', '수준 규칙', 'ML0~ML5 여섯 수준 · 핵심 문항 4종', '점수를 수준으로 바꾸는 규칙 — 핵심 문항 4종은 하나라도 빠지면 ML2에서 멈춤')]
rh = 820000; lw = 2500000; mw = 3000000; rw = RW - lw - mw - 2 * 120000
for i, (nm, easy, scale, expl) in enumerate(rows):
    b = box(s, L, y, lw, rh, NAVY, name=f'm row{i} name'); par(b.text_frame, nm, 15, True, WHITE, PP_ALIGN.CENTER, first=True); par(b.text_frame, easy, 12, False, WHITE, PP_ALIGN.CENTER)
    b = box(s, L + lw + 120000, y, mw, rh, LIGHT, name=f'm row{i} scale'); par(b.text_frame, scale, 13, True, NAVY, PP_ALIGN.CENTER, first=True)
    b = box(s, L + lw + mw + 240000, y, rw, rh, None, name=f'm row{i} expl', anchor=MSO_ANCHOR.MIDDLE); par(b.text_frame, expl, 13, False, INK, PP_ALIGN.LEFT, first=True)
    y += rh
    if i < 2: arrow_down(s, L + lw // 2, y + 20000, 180000); y += 220000
arrow_down(s, L + lw // 2, y + 20000, 180000); y += 240000
b = box(s, L, y, lw, 560000, RED, name='m out'); par(b.text_frame, '판정 — 성숙도 수준', 15, True, WHITE, PP_ALIGN.CENTER, first=True)
b = box(s, L + lw + 120000, y, RW - lw - 120000, 560000, PINK, name='m out expl'); par(b.text_frame, '회사의 연계가 ML0(도입 전)~ML5(혁신화) 가운데 어디에 있는지 — 종합점수는 판정에 쓰지 않는 보조값', 13, True, RED, PP_ALIGN.LEFT, first=True)
y += 560000 + 120000
t = tbox(s, L, y, RW, 380000, 'm verify'); par(t.text_frame, '검증 두 방향 — ① 추적성: 문항마다 어느 조항에서 왔는지 따라감  ② 표준 적합성: ISO/IEC 33004 요구사항 24건 자체점검(충족 22·해당 없음 1·보완 1)', 11.5, False, GREY, PP_ALIGN.LEFT, first=True)
LOG.append((12, '화면', '[그림 3-2] 가로판 그림 한 장', '모델링 한 그림(도형): 두 표준 ↔ 매핑 → ① 참조 모델(업무 목록 16종) → ② 평가 모델(채점표 55문항) → ③ 성숙도 모델(수준 규칙) → 판정 / 검증 두 방향', '이사님 10/7: 모델링에 대한 간단한 그림'))
set_notes(12, """만든 모델을 한 그림으로 말씀드리겠습니다.
위에 두 표준이 있습니다. 환경경영시스템(ISO 14001)의 조항 30개와 전과정평가(LCA)의 요구사항을 하나하나 대응시켰습니다. 이것이 매핑입니다. 맞닿는 곳이 연계 지점이고, 그 가운데 꼭 있어야 할 네 곳이 핵심 연계 지점입니다.
그 아래가 만든 것 세 가지입니다. 첫째, 참조 모델은 평가할 업무 목록입니다. 연계 지점을 업무 단위로 묶어 열여섯 가지로 적었습니다. 둘째, 평가 모델은 채점표입니다. 업무마다 되어 있는가를 묻는 문항 쉰다섯 개와, 확인할 항목, 몇 점으로 볼지의 기준입니다. 셋째, 성숙도 모델은 수준 규칙입니다. 점수를 ML0부터 ML5까지 여섯 수준으로 바꾸는 규칙이고, 핵심 문항 네 가지는 하나라도 빠지면 ML2에서 멈춥니다.
이 셋을 합쳐 연계 성숙도 평가 모델, LMA라고 합니다. 결과는 회사의 연계가 어느 수준에 있는가입니다.
만든 뒤에는 두 방향으로 검증하였습니다. 문항마다 어느 조항에서 왔는지 따라가는 추적성 검증과, 국제표준 요건 24건을 자체점검한 표준 적합성 검증입니다.
이제 하나씩 보겠습니다. 먼저 두 표준을 어떻게 맞대었는지입니다.
〔논문 [그림 3-2] 전체는 백업 B-7. 1분 10초〕""")

# ================= 14번 핵심 연계 지점 4곳 =================
s = prs.slides[13]; pic = shp(s, 8)
pw = int(6.3 * IN); ph = int(pic.height * pw / pic.width); pic.left = Emu(L); pic.top = Emu(1250000); pic.width = Emu(pw); pic.height = Emu(ph)
rx = L + pw + 230000; rw = L + RW - rx
t = tbox(s, rx, 1250000, rw, 330000, 'lp head'); par(t.text_frame, '핵심 연계 지점 네 곳 — 굵은 선 ①~④', 14, True, NAVY, first=True)
pts = [('① 환경측면(6.1.2) ↔ LCA 영향평가', '영향이 가장 큰 곳(Hotspot)을 환경측면 결정의 근거로 쓴다'),
       ('② 환경목표(6.2.1) ↔ LCA 해석', 'Hotspot을 근거로 숫자 목표를 세운다'),
       ('③ 설계·개발(8.1) ↔ LCA 목록분석', '설계할 때 LCA 결과를 반영한다'),
       ('④ 모니터링(9.1.1) ↔ LCA 재수행', '개선 전·후를 LCA로 다시 측정해 비교한다')]
y = 1250000 + 380000; bh = 1000000
for i, (h1, h2) in enumerate(pts):
    b = box(s, rx, y, rw, bh, LIGHT, name=f'lp{i}', anchor=MSO_ANCHOR.MIDDLE); par(b.text_frame, h1, 12.5, True, NAVY, first=True, after=2); par(b.text_frame, h2, 12, False, INK); y += bh + 90000
set_text_keep(shp(s, 9), '왼쪽 — ISO 14001의 관리 주기(계획–실행–점검–조치) / 오른쪽 — ISO 14044의 4단계. 두 주기가 만나는 곳이 연계 지점이고, 네 곳 가운데 하나가 끊기면 순환이 그 단계에서 멈춘다.')
shp(s, 9).top = Emu(1250000 + ph + 80000); shp(s, 9).width = Emu(pw)
LOG.append((14, '화면', '그림 + 아래 한 줄', '그림(왼쪽 6.3in) + 오른쪽 핵심 연계 지점 4곳 설명 상자', '이사님 10/7: 연계 지점 4군데 설명'))
set_notes(14, """왼쪽이 환경경영시스템(ISO 14001)의 관리 주기, 오른쪽이 전과정평가(LCA)의 4단계입니다. 선이 두 표준이 맞닿는 곳, 곧 연계 지점입니다.
그 가운데 굵은 선 네 개가 핵심 연계 지점입니다. 오른쪽에 풀어 두었습니다.
첫째, 환경측면입니다. LCA 영향평가에서 영향이 가장 큰 곳, 핫스팟을 환경측면 결정의 근거로 씁니다. 둘째, 환경목표입니다. 핫스팟을 근거로 숫자 목표를 세웁니다. 셋째, 설계·개발입니다. 설계할 때 LCA 결과를 반영합니다. 넷째, 모니터링입니다. 개선 전과 후를 LCA로 다시 측정해 비교합니다.
이 넷 가운데 하나가 끊기면 순환이 그 단계에서 멈춥니다. 그래서 이 넷이 뒤의 핵심 문항 4종이 됩니다.
이 지점들을 업무 단위로 묶은 것이 다음 장의 참조 모델입니다.
〔1분 10초〕""")

# ================= 16번 그림 크게 + 점수 매기는 법 =================
s = prs.slides[15]; remove(s, 10, 11, 12)
pic = shp(s, 8); pw = int(5.9 * IN); ph = int(pic.height * pw / pic.width); pic.left = Emu(L); pic.top = Emu(1250000); pic.width = Emu(pw); pic.height = Emu(ph)
t9 = shp(s, 9); t9.top = Emu(1250000 + ph + 100000); t9.width = Emu(pw); t9.height = Emu(1500000)
tf = t9.text_frame
for p in list(tf.paragraphs)[1:]: p._p.getparent().remove(p._p)
set_text_keep(t9, '55문항 = ISO 14001 조항 대응 44 + ISO 14044 고유 11 (계획 30 · 실행 6 · 점검 10 · 개선 9)')
par(tf, '문항마다 「무엇을 증거로 볼지」와 「어느 상태를 몇 점으로 볼지」(BARS 앵커)를 미리 정해 둠', 13, False, INK)
par(tf, '그림 — 평가 문항 → 판정 기준 → 평가 기록부 → 평가 보고서의 순서로 쓰이는 평가 모델의 구성 [그림 3-8]', 11.5, False, GREY)
rx = L + pw + 230000; rw = L + RW - rx; y = 1250000
b = box(s, rx, y, rw, 420000, NAVY, name='sc head'); par(b.text_frame, '점수를 매기는 법 — 세 단계', 14, True, WHITE, PP_ALIGN.CENTER, first=True); y += 420000 + 100000
steps = [('①', '확인할 항목을 미리 정한다', '문항마다 필수 항목(shall)과 추가 항목(should)을 적어 둠 — 체크포인트'),
         ('②', '회사 문서를 보고 그 항목이 있는지 센다', '구두 설명은 인정하지 않고 문서·기록만'),
         ('③', '센 개수가 점수', '필수 전부 있으면 3점(기준선) · 추가까지 있으면 4·5점 · 필수가 빠지면 1·2점')]
sh_h = 1000000
for i, (n, h1, h2) in enumerate(steps):
    b = box(s, rx, y, rw, sh_h, LIGHT, name=f'sc{i}', anchor=MSO_ANCHOR.MIDDLE); par(b.text_frame, f'{n} {h1}', 13, True, NAVY, first=True, after=2); par(b.text_frame, h2, 12, False, INK); y += sh_h + 90000
b = box(s, rx, y, rw, 900000, PINK, name='sc ex', anchor=MSO_ANCHOR.MIDDLE)
par(b.text_frame, '예 — ★ D-1 설계·개발 단계의 LCA 결과 반영', 12.5, True, RED, first=True, after=2); par(b.text_frame, '필수 3개 가운데 2개만 있어 2점 — 4장에서 판정을 가른 문항. 점수별 기준 전문은 백업 B-10a', 12, False, INK)
LOG.append((16, '화면', '그림(5.3in) + 오른쪽 BARS 앵커 표(5점 상태 기술) + 체크포인트 각주', '그림(5.9in) + 「점수를 매기는 법 — 세 단계」 상자 + D-1 예. BARS 앵커 표는 백업 B-10a로', '이사님 10/7: 왼쪽 그림 크게, 옆 표가 어려움'))
set_notes(16, """평가 모델은 채점표입니다. 문항은 쉰다섯 개이고, 계획 30, 실행 6, 점검 10, 개선 9입니다.
점수를 매기는 법은 세 단계입니다. 첫째, 문항마다 확인할 항목을 미리 정해 둡니다. 반드시 있어야 할 필수 항목과, 있으면 더 좋은 추가 항목입니다. 이것을 체크포인트라 합니다. 둘째, 회사 문서를 보고 그 항목이 있는지 셉니다. 말로 설명한 것은 인정하지 않고 문서와 기록만 봅니다. 셋째, 센 개수가 점수입니다. 필수가 전부 있으면 3점, 이것이 기준선입니다. 추가까지 있으면 4점이나 5점, 필수가 빠지면 1점이나 2점입니다.
그래서 점수는 평가자의 재량이 아니라 체크포인트 충족 개수로 정해집니다. 평가자가 달라도 같은 문서면 같은 점수가 나오게 한 것입니다.
예를 하나 들면 D-1, 설계·개발 단계의 LCA 결과 반영 문항입니다. A사는 필수 세 개 가운데 두 개만 있어 2점이었고, 이 문항이 4장에서 판정을 갈랐습니다.
점수를 성숙도 수준으로 바꾸는 규칙이 다음 장의 성숙도 모델입니다.
〔점수별 기준 전문(BARS 앵커)은 백업 B-10a, 부록 C. 1분 20초〕""")

# ================= 17번 여섯 수준 계단 =================
s = prs.slides[16]; remove(s, 8, 19)
set_text_keep(shp(s, 9), '핵심 문항 4종 — 연계 성립의 필요조건')
levels = [('ML0', '도입 전', '연계 활동이 없음'),
          ('ML1', '도입', '핵심 문항 4종 모두 2점 이상'),
          ('ML2', '관리화', '+ 나머지 51문항 가운데 절반 이상 3점'),
          ('ML3', '체계화', '+ 핵심 문항 4종 모두 3점 이상'),
          ('ML4', '예측화', '+ 핵심 문항 4종 모두 4점, 점검(C) 영역 전부 3점 이상'),
          ('ML5', '혁신화', '+ 전 문항 3점 이상, 개선(A) 영역 전부 4점 이상')]
area_w = int(7.35 * IN); rh = 600000; gap = 90000; step = 420000; y_top = 1250000
for i in range(5, -1, -1):
    ml, nm, cond = levels[i]; row = 5 - i; y = y_top + row * (rh + gap); x = L + i * step; w = area_w - i * step
    fill = PINK if i == 3 else (LIGHT if i >= 1 else RGBColor(0xF2, 0xF2, 0xF2)); col = RED if i == 3 else NAVY
    b = box(s, x, y, w, rh, fill, name=f'ml{i}', anchor=MSO_ANCHOR.MIDDLE, margin=110000)
    par(b.text_frame, f'{ml} {nm}  —  {cond}', 13, i == 3, col, PP_ALIGN.LEFT, first=True)
    if i == 2:
        m = box(s, L, y_top + 2 * (rh + gap), 3 * step - 60000, rh, None, name='ml mark', anchor=MSO_ANCHOR.MIDDLE, margin=20000)
        par(m.text_frame, 'A사 ▶ ML2', 12, True, RED, PP_ALIGN.LEFT, first=True); par(m.text_frame, '(D-1 2점)', 11, False, RED, PP_ALIGN.LEFT)
y_end = y_top + 6 * (rh + gap)
t = tbox(s, L, y_end - 40000, area_w, 330000, 'ml note'); par(t.text_frame, '조건은 누적 — 아래 수준의 조건을 모두 갖추어야 위로 올라간다. 조건의 어느 항에도 평균은 쓰이지 않는다.', 11.5, False, GREY, first=True)
kb = shp(s, 18); kb.top = Emu(y_end + 330000); kb.height = Emu(760000)
set_text_keep(kb, '핵심 문항 하나라도 3점 미만이면 평균이 높아도 ML2에서 멈춘다 — 연계가 성립하지 않는 지점(핵심 문항 미달)과 연계가 얕은 지점(51문항의 점수 차)은 다르다')
LOG.append((17, '화면', '[표 3-7] 진입 조건·설계 근거 표 + 각주', '여섯 수준 계단(쉬운 말 한 줄씩, A사 멈춘 곳 표시) + 핵심 문항 4종 + 빨간 한 줄. 표·각주는 백업 B-11a로', '이사님 10/7: 그림으로 간단하게'))
set_notes(17, """성숙도 모델은 점수를 수준으로 바꾸는 규칙입니다. 수준은 ML0에서 ML5까지 여섯이고, 계단처럼 아래 조건을 모두 갖추어야 위로 올라갑니다.
ML1은 핵심 문항 네 가지가 모두 2점 이상, ML2는 거기에 나머지 쉰한 문항 가운데 절반 이상이 3점, ML3은 핵심 문항 네 가지가 모두 3점 이상입니다. ML4와 ML5는 점검과 개선 영역까지 갖추어야 합니다.
조건의 어느 항에도 평균은 쓰이지 않습니다. 연계가 성립하지 않는 지점과 연계가 얕은 지점은 다르기 때문에, 연계 성립의 필요조건인 핵심 문항 4종에만 최소충족 판정을 적용하였습니다. 이 원리는 KS X ISO 8000-62에서 가져왔습니다.
오른쪽이 핵심 문항 네 가지입니다. 환경측면 결정, 환경목표, 설계·개발, 성과 모니터링입니다. 핵심 문항 하나라도 3점 미만이면 평균이 높아도 ML2에서 멈춥니다. A사가 바로 그 자리입니다.
만든 도구가 요건을 갖추었는지 한 장으로 정리하겠습니다.
〔진입 조건·설계 근거 표 전문은 백업 B-11a, 수준별 기준은 B-11. 비교주장 조건부 shall(C-7·A-9)은 3.4.3 — 백업 B-Q5. 말로 풀 때만 비유 — 건강검진의 필수 검사(논문에 없음). 1분 20초〕""")

# ================= 18번 카드 셋 =================
s = prs.slides[17]; remove(s, 8)
cards = [('참조 모델', '평가할 업무 목록', '연계 프로세스 16종', '계획 8 · 실행 2 · 점검 3 · 개선 3', '새로 만든 것 — 두 표준의 조항을 대응시켜 도출한 연계 프로세스'),
         ('평가 모델', '채점표', '평가 문항 55', '문항별 판정 기준(BARS 앵커) · 객관적 증거 목록', '새로 만든 것 — 표준이 평가자에게 맡긴 평가지표를 이 분야에 맞게 개발'),
         ('성숙도 모델', '수준 규칙', 'ML0~ML5 여섯 수준', '핵심 문항 4종', '새로 만든 것 — 핵심 문항 4종에만 최소충족 판정을 적용한 규칙')]
cw = int(3.3 * IN); cg = (RW - 3 * cw) // 2; y = 1250000; ch = int(3.25 * IN)
for i, (nm, easy, sz, sz2, new) in enumerate(cards):
    x = L + i * (cw + cg)
    b = box(s, x, y, cw, ch, LIGHT, NAVY, name=f'card{i}', anchor=MSO_ANCHOR.TOP, margin=120000); tf = b.text_frame
    par(tf, nm, 19, True, NAVY, PP_ALIGN.CENTER, first=True, after=0); par(tf, f'— {easy} —', 14, False, GREY, PP_ALIGN.CENTER, after=10)
    par(tf, sz, 17, True, INK, PP_ALIGN.CENTER, after=0); par(tf, sz2, 13, False, INK, PP_ALIGN.CENTER, after=14)
    par(tf, new, 13, False, NAVY, PP_ALIGN.LEFT)
    if i < 2: arrow_right(s, x + cw + 30000, y + ch // 2, cg - 60000, name=f'card arr{i}')
LOG.append((18, '화면', '계층·산출물·규모·따른 표준·새로 만든 것 표(5열)', '카드 셋(참조·평가·성숙도 모델: 쉬운 이름·규모·새로 만든 것). 「따른 표준」 열은 10번·백업 B-2로', '이사님 10/7: 모델 한 장 요약을 그림처럼'))
set_text_keep(shp(s, 7), '근거  3.5.1, 3.5.2, 3.6, [표 3-11], 부록 E')
set_notes(18, """모델을 한 장으로 정리하겠습니다. 세 계층입니다.
참조 모델은 평가할 업무 목록으로, 연계 프로세스 열여섯 가지입니다. 두 표준의 조항을 대응시켜 도출한 것이 새로 만든 것입니다.
평가 모델은 채점표로, 평가 문항 쉰다섯 개와 문항별 판정 기준, 객관적 증거 목록입니다. 표준이 평가자에게 맡긴 평가지표를 이 분야에 맞게 개발한 것이 새로 만든 것입니다.
성숙도 모델은 수준 규칙으로, ML0부터 ML5까지 여섯 수준과 핵심 문항 4종입니다. 핵심 문항 4종에만 최소충족 판정을 적용한 규칙이 새로 만든 것입니다.
검증은 두 방향입니다. 추적성 검증은 쉰다섯 문항을 두 표준의 전 조항 대조에서 나온 근거 조항까지 추적한 것입니다. 표준 적합성 검증은 ISO/IEC 33004 요구사항 24건을 자체 점검한 것으로, 충족 22건, 해당 없음 1건, 보완 1건입니다. 보완 1건은 전문가 합의로, 조사지 설계까지 마쳤고 실시는 향후 연구 과제입니다.
한 줄로 말씀드리면, 표준의 틀을 따르고 표준이 비워 둔 세 곳을 이 연구가 채웠습니다. 이제 이 도구로 한 기업을 평가한 결과입니다.
〔따른 표준 전문은 10번·백업 B-2. 1분 20초〕""")

# ================= 19번 평가 한 그림 =================
s = prs.slides[18]; remove(s, 8, 9, 10, 11, 12, 13, 14, 15, 16)
set_text_keep(shp(s, 2), '평가를 한 그림으로 — 제조기업 A사')
for _p in shp(s, 2).text_frame.paragraphs:
    for _r in _p.runs: _r.font.size = Pt(25); set_text_keep(shp(s, 7), '근거  4.1, 4.2, 4.3.5, 4.5, 4.6')
steps = [('① 문서 받기', 'A사 문서 16종(33건)', '매뉴얼 · 절차서 · 기록 · 보고서'),
         ('② 채점', '55문항마다 확인 항목이 있는지 보고 점수', '잠정 종합 2.89'),
         ('③ 다시 확인', '증빙 문서로 하나씩 대조', '상향 7 · 하향 9 → 확정 2.85'),
         ('④ 판정', '핵심 문항 4종 확인 — D-1이 2점', 'ML2(관리화)'),
         ('⑤ 개선', '낮은 점수의 원인 3건 → 과제 R1~R5', '완수 시 3.15 · ML3 전망')]
n = 5; ag = 230000; sw = (RW - (n - 1) * ag) // n; y = 1250000; sh_h = int(1.75 * IN)
for i, (h1, h2, h3) in enumerate(steps):
    x = L + i * (sw + ag); fill = PINK if i == 3 else LIGHT; col = RED if i == 3 else NAVY
    b = box(s, x, y, sw, sh_h, fill, name=f'ev{i}', anchor=MSO_ANCHOR.TOP, margin=70000); tf = b.text_frame
    par(tf, h1, 14, True, col, PP_ALIGN.CENTER, first=True, after=4); par(tf, h2, 11.5, False, INK, PP_ALIGN.CENTER, after=4); par(tf, h3, 13, True, col, PP_ALIGN.CENTER)
    if i < n - 1: arrow_right(s, x + sw + 20000, y + sh_h // 2, ag - 40000, name=f'ev arr{i}')
y += sh_h + 200000
b = box(s, L, y, RW, int(1.55 * IN), None, GREY, name='ev info', anchor=MSO_ANCHOR.MIDDLE, margin=120000); tf = b.text_frame
par(tf, '대상 — 제조기업 A사의 환경경영시스템(ISO 14001)과 2025년 제품 전과정평가(LCA)', 12.5, False, INK, PP_ALIGN.LEFT, first=True, after=3)
par(tf, '방법 — 문서 심사만(면담·현장 관찰 없음), 평가자 1명  ·  기간 — 2026년 6~7월', 12.5, False, INK, PP_ALIGN.LEFT, after=3)
par(tf, '지위 — 연구 목적의 시범 적용(ISO/IEC 33002 평가 클래스 3 상당 · 독립성 범주 B 상당) — 회사에 공식 성숙도 등급을 주는 평가가 아님', 12.5, False, INK, PP_ALIGN.LEFT)
y += int(1.55 * IN) + 160000
b = box(s, L, y, RW, 640000, PINK, name='ev key'); par(b.text_frame, '목적은 모델의 판정 논리가 실제 자료에서 작동하는지 확인하는 것 — 이 다섯 단계를 차례로 말씀드린다', 14, True, RED, PP_ALIGN.CENTER, first=True)
LOG.append((19, '화면', '대상·방법·기간·지위 4줄 + 빨간 줄', '평가 한 그림: 문서 받기 → 채점 → 다시 확인 → 판정 → 개선(숫자 하나씩) + 대상·방법·기간·지위 쉬운 말 + 빨간 줄', '이사님 10/7: 평가에 대한 그림, 19·20이 너무 어려움'))
set_notes(19, """4부는 이 모델을 한 회사에 적용한 결과입니다. 평가가 어떻게 흘러갔는지 한 그림으로 먼저 보겠습니다.
첫째, 문서를 받았습니다. 제조기업 A사의 환경경영시스템(ISO 14001)과 2025년 제품 전과정평가(LCA)에 관한 문서 열여섯 종, 서른세 건입니다. 매뉴얼, 절차서, 기록, 보고서입니다.
둘째, 채점입니다. 쉰다섯 문항마다 확인 항목이 있는지 보고 점수를 매겼습니다. 잠정 종합점수는 2.89였습니다.
셋째, 다시 확인입니다. 잠정 점수를 증빙 문서로 하나씩 대조하였습니다. 올라간 것 일곱, 내려간 것 아홉으로 확정 종합점수는 2.85입니다.
넷째, 판정입니다. 핵심 문항 네 가지를 확인하니 D-1이 2점이었고, 판정은 ML2, 관리화입니다.
다섯째, 개선입니다. 낮은 점수의 원인 세 가지에서 개선 과제 다섯 건을 도출하였고, 모두 마치면 3.15, ML3이 전망됩니다.
평가는 문서 심사만으로, 평가자 한 명이, 2026년 6월과 7월에 수행하였습니다. ISO/IEC 33002의 평가 클래스 3에 상당하는 연구 목적의 시범 적용이며, 회사에 공식 성숙도 등급을 주는 평가가 아닙니다. 목적은 모델의 판정 논리가 실제 자료에서 작동하는지 확인하는 것입니다.
평가자가 한 사람이므로 객관성을 어떻게 확보하였는지 먼저 말씀드리겠습니다.
〔「회사가 쓰겠다던가」 → 백업 B-Q9. 1분 20초〕""")

# ================= 20번 상자 둘 세 줄씩 =================
s = prs.slides[19]
for sid, lines, head in ((8, ['미리 정한 규칙 (판정 관례)', '· 말로만 한 것은 인정하지 않음', '· 양식만 있고 기입이 없으면 미충족', '· 판단이 어려우면 미충족 — 증빙은 체크포인트 단위로 적어 제3자가 재현 가능'], 1),
                         (9, ['다시 확인한 결과 (재검증)', '· 잠정 점수를 증빙 문서로 하나씩 대조', '· 올라간 것 7건, 내려간 것 9건 (순변동 −2점)', '· 종합 2.89 → 2.85 — 재검증이 점수가 후해지는 것을 막는 방향으로 작동'], 1)):
    b = shp(s, sid); tf = b.text_frame; old = tf.text
    ps = list(tf.paragraphs)
    for p in ps[1:]: p._p.getparent().remove(p._p)
    set_text_keep(b, lines[0])
    for p in tf.paragraphs:
        for r in p.runs: r.font.size = Pt(13.5); r.font.bold = True; r.font.color.rgb = NAVY
    for ln in lines[1:]: par(tf, ln, 12.5, False, RED if '2.89' in ln else INK, before=3)
    b.height = Emu(int(1.95 * IN)); LOG.append((20, f'[{sid}]', old, '\n'.join(lines), '이사님 10/7: 20번 쉽게 — 세 줄씩'))
t = tbox(s, L, 1250000 + int(1.95 * IN) + 60000, RW, 400000, 'obj note')
par(t.text_frame, '평가자가 이 회사의 환경경영 업무를 지원한 이력은 공평성 고려 사항으로 밝혔고, 판정은 회사가 생성·기입한 기록의 실재만을 근거로 하였다. 채점 기준은 평가 중 바꾸지 않았다(C-7·A-9 처리만 평가 뒤 3.4.3으로 확정, 점수 불변).', 11, False, GREY, first=True)
set_notes(20, """평가자가 한 사람이라 객관성을 두 가지로 확보하였습니다.
첫째, 규칙을 미리 정해 두었습니다. 말로만 한 것은 인정하지 않고, 양식만 있고 기입이 없으면 미충족, 판단이 어려우면 미충족입니다. 증빙은 체크포인트 단위로 적어 두어 제3자가 다시 볼 수 있게 하였습니다.
둘째, 잠정 점수를 증빙 문서로 하나씩 다시 확인하였습니다. 올라간 것이 일곱 건, 내려간 것이 아홉 건이고, 종합점수는 2.89에서 2.85로 낮아졌습니다. 재검증이 점수가 후해지는 것을 막는 방향으로 작동한 것입니다.
평가자가 이 회사의 환경경영 업무를 지원한 이력은 공평성 고려 사항으로 밝혔고, 판정은 회사가 만든 기록의 실재만을 근거로 하였습니다.
아래는 평가 기록부의 실제 기입으로, 핵심 문항 네 가지만 발췌한 것입니다. D-1 행이 뒤의 판정을 가릅니다.
〔1분 10초〕""")

# ================= 41번(B-10a) · 43번(B-11a) 옮긴 표 =================
s = prs.slides[40]; remove(s, 5, 8, 9, 13)
set_text_keep(shp(s, 2), 'D-1 채점 기준 — 점수별 상태(BARS 앵커) 전문'); label_backup(s, '백업  B-10a'); set_text_keep(shp(s, 7), '근거  3.3.2, [표 3-6], 부록 C')
ex = shp(s, 10); tb = shp(s, 11); ft = shp(s, 12)
ex.left = Emu(L); ex.top = Emu(1250000); ex.width = Emu(RW); ex.height = Emu(600000)
tb.left = Emu(L); tb.top = Emu(1950000); tb.width = Emu(RW)
for c in tb.table.columns: c.width = Emu(int(c.width * RW / (4.91 * IN)))
for row in tb.table.rows:
    row.height = Emu(int(row.height * 1.15))
    for cell in row.cells:
        for p in cell.text_frame.paragraphs:
            for r in p.runs:
                if r.font.size: r.font.size = Pt(r.font.size.pt + 2)
ft.left = Emu(L); ft.top = Emu(1950000 + sum(r.height for r in tb.table.rows) + 150000); ft.width = Emu(RW)
for p in ft.text_frame.paragraphs:
    for r in p.runs:
        if r.font.size: r.font.size = Pt(r.font.size.pt + 2)
set_notes(41, '백업 — 질문 때만. 16번에서 뺀 BARS 앵커 표 전문. 「점수 기준을 어떻게 정했나」가 나오면 이 장.')
LOG.append((41, '새 백업 B-10a', '(16번 사본)', 'D-1 채점 기준 BARS 앵커 표 전문', '16번에서 옮김'))
s = prs.slides[42]
for sid in (9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 20): remove(s, sid)
remove(s, 5)
set_text_keep(shp(s, 2), '성숙도 수준 진입 조건과 설계 근거 — [표 3-7]'); label_backup(s, '백업  B-11a'); set_text_keep(shp(s, 7), '근거  3.4.1~3.4.4, [표 3-7]')
tb = shp(s, 8); tb.left = Emu(L); tb.top = Emu(1250000); tb.width = Emu(RW)
for c in tb.table.columns: c.width = Emu(int(c.width * RW / (7.37 * IN)))
for row in tb.table.rows:
    row.height = Emu(int(row.height * 1.1))
    for cell in row.cells:
        for p in cell.text_frame.paragraphs:
            for r in p.runs:
                if r.font.size: r.font.size = Pt(r.font.size.pt + 1)
ft = shp(s, 19); ft.top = Emu(1250000 + sum(r.height for r in tb.table.rows) + 150000); ft.width = Emu(RW)
for p in ft.text_frame.paragraphs:
    for r in p.runs:
        if r.font.size: r.font.size = Pt(12)
set_notes(43, '백업 — 질문 때만. 17번에서 뺀 [표 3-7] 진입 조건·설계 근거 표 전문. 「PA.1.1 상당」 같은 근거를 물으면 이 장.')
LOG.append((43, '새 백업 B-11a', '(17번 사본)', '[표 3-7] 진입 조건·설계 근거 표 전문 + 최소충족·비교주장 각주', '17번에서 옮김'))

# ================= 30번 백업 문구 =================
s = prs.slides[29]; rep(s, 2, 'B-1~B-16 그림·표', 'B-1~B-16 그림·표(B-10a 채점 기준 표 · B-11a 진입 조건 표 포함)', '백업 두 장 추가', 30)

# ================= 쪽번호 재부여 (31번부터 밀림) =================
n_fix = 0
for i, s in enumerate(prs.slides, 1):
    for sh in s.shapes:
        if sh.has_text_frame and sh.top > int(7.5 * IN) and sh.left > int(9.5 * IN) and re.fullmatch(r'\d{1,2}', sh.text_frame.text.strip()):
            if sh.text_frame.text.strip() != str(i): set_text_keep(sh, str(i)); n_fix += 1
LOG.append(('41~72', '쪽번호', '옛 번호', '새 번호', f'{n_fix}곳'))
prs.save(OUT)
json.dump(LOG, open('log13.json', 'w'), ensure_ascii=False, indent=1)
print(len(LOG), 'changes;', len(prs.slides), 'slides; page numbers fixed', n_fix)
