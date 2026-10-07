"""v10 → v11: (1) 용어 풀이 줄(본 슬라이드 13곳, 근거 줄 위 회색 11pt) (2) 그림 장 노트 쉽게(9장).
사용: python3 -I mk_deck11.py v10.pptx v11.pptx
"""
import sys, os, json, copy
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR
from terms11 import T
from notes11 import N
SRC, OUT = sys.argv[1], sys.argv[2]
prs = Presentation(SRC); W, H = prs.slide_width, prs.slide_height
LOG = []
FOOT = 7076160   # 쪽번호·근거 줄 top (A4 변환 후)
for sn, txt in T.items():
    s = prs.slides[sn - 1]
    content = [sh for sh in s.shapes if sh.top < FOOT - 10000 and sh.shape_id not in (2, 3, 4)]
    bottom = max(sh.top + sh.height for sh in content)
    # 용어 줄 자리: 근거 줄 바로 위 2줄(약 0.42in). 내용이 그 자리까지 내려와 있으면 내용을 위로 밀 수 있는 만큼 밀고, 안 되면 내용 아래 글 글자 -1pt
    box_h = 380000; top = FOOT - 10000 - box_h; limit = top - 50000
    if bottom > limit:
        need = bottom - limit
        top_content = min(sh.top for sh in content); room = top_content - 1180000
        shift = min(need, max(0, room))
        for sh in content: sh.top = Emu(sh.top - shift)
        bottom -= shift
        if bottom > limit:
            rest = bottom - limit
            pics = [sh for sh in content if sh.shape_type == 13]
            if pics:   # 그림이 있으면 그림을 줄이고(비율 유지) 그 아래 도형을 그만큼 올림
                pic = max(pics, key=lambda p: p.width * p.height)
                nh = pic.height - rest; k = nh / pic.height; nw = int(pic.width * k)
                pic.left = Emu(pic.left + (pic.width - nw) // 2); pic.width = Emu(nw); pic.height = Emu(int(nh))
                for sh in content:
                    if sh is not pic and sh.top >= pic.top + pic.height: sh.top = Emu(sh.top - rest)
            else:      # 글 도형의 높이를 줄임(아래쪽에 걸친 것만)
                for sh in content:
                    if sh.top + sh.height > limit and not sh.has_table:
                        sh.height = Emu(max(sh.height - rest, int(sh.height * 0.8)))
    tb = s.shapes.add_textbox(Emu(566928 * 0.877), Emu(top), Emu(W - 2 * int(566928 * 0.877)), Emu(box_h))
    tf = tb.text_frame; tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.BOTTOM
    tf.margin_left = tf.margin_right = Emu(45720); tf.margin_top = tf.margin_bottom = Emu(0)
    p = tf.paragraphs[0]; r = p.add_run(); r.text = txt
    r.font.size = Pt(10.5); r.font.name = '맑은 고딕'; r.font.color.rgb = RGBColor(0x59, 0x59, 0x59)
    rpr = r._r.get_or_add_rPr(); from pptx.oxml.ns import qn; from lxml import etree
    ea = etree.SubElement(rpr, qn('a:ea')); ea.set('typeface', '맑은 고딕')
    # 「용어」 머리 굵게
    r.text = ''; r1 = r; r1.text = '용어  '; r1.font.bold = True
    r2 = copy.deepcopy(r1._r); r1._r.addnext(r2)
    from pptx.text.text import _Run
    rr = _Run(r2, p); rr.text = txt[len('용어  '):]; rr.font.bold = False
    tb.name = f'Terms {sn}'
    LOG.append((sn, '용어 줄 추가', '', txt, '처음 나오는 장'))
for sn, txt in N.items():
    tf = prs.slides[sn - 1].notes_slide.notes_text_frame; old = tf.text
    p0 = tf.paragraphs[0]; tmpl = copy.deepcopy(p0._p)
    for p in list(tf.paragraphs): p._p.getparent().remove(p._p)
    from pptx.text.text import _Paragraph
    for line in txt.split('\n'):
        np_ = copy.deepcopy(tmpl); tf._txBody.append(np_); para = _Paragraph(np_, tf)
        rs = para.runs
        if rs:
            rs[0].text = line
            for r in rs[1:]: r._r.getparent().remove(r._r)
        else: para.text = line
    LOG.append((sn, '노트 쉽게', old, txt, '그림 장 — 그림 어디 → 뜻 → 다음'))
prs.save(OUT)
json.dump(LOG, open('log11.json', 'w'), ensure_ascii=False, indent=1)
print(len(T), 'term lines;', len(N), 'notes')
