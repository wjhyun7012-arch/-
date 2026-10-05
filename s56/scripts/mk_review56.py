# -*- coding: utf-8 -*-
"""세션 56 — 적합성검토표_v201d_20261005_v1.docx/.pdf + 판정_*.md 생성 + 「현행」 실재 검증.
사용: python3 mk_review56.py   (작업 경로 /home/claude/s54)
"""
import os, re, sys, json, shutil, zipfile, html, subprocess, hashlib, importlib.util
from datetime import date

BASE = '/home/claude/s54'
OUT = os.path.join(BASE, 'out')
TPL = os.path.join(BASE, '적합성검토표_v201_20260930_v1.docx')
THESIS = os.path.join(BASE, 'phd_v201d_20260930.docx')
NAME = '적합성검토표_v201d_20261005_v1'
DOCX = os.path.join(OUT, NAME + '.docx')

spec = importlib.util.spec_from_file_location('rows', os.path.join(OUT, 'rows_s56.py'))
rows = importlib.util.module_from_spec(spec); spec.loader.exec_module(rows)
ROWS, EXCLUDED = rows.ROWS, rows.EXCLUDED
OKLIST = getattr(rows, 'OKLIST', [])
REVIEWERS = getattr(rows, 'REVIEWERS', [])

KIND_ORDER = {'필수': 0, '결정 필요': 1, '권고': 2, '선택': 3}
SECS = ['그림', '표', '내용']

# ───────────────────────── 검증: 「현행」 전수 실재 ─────────────────────────
dump = open(os.path.join(BASE, 'kit', 'dump_all.txt'), encoding='utf8').read()
dump_flat = dump.replace('⏎', ' ')

def quotes(s):
    return re.findall(r'「([^「」]+)」', s)

def exists(q):
    q2 = q.strip()
    if q2 in dump: return True
    # 「… 앞부분 생략 …」 꼴 — 생략 기호로 나눈 조각마다 확인
    if '…' in q2:
        parts = [p.strip(' \t') for p in q2.split('…') if p.strip(' \t')]
        if parts and all(exists(p) for p in parts): return True
    # 표 행 인용: 「A | B | C」 → 셀 단위 확인
    if ' | ' in q2:
        return all(c.strip() == '' or c.strip() in dump for c in q2.split(' | '))
    # 줄바꿈(⏎·/) 인용 → 조각 단위
    parts = [p.strip() for p in re.split(r' ⏎ | / |⏎', q2) if p.strip()]
    if len(parts) > 1 and all(p in dump or p in dump_flat for p in parts): return True
    # 쪽 번호 목차 줄 「… 59」 → 탭 유무 무시
    q3 = re.sub(r'\s+', '', q2)
    if q3 and q3 in re.sub(r'\s+', '', dump): return True
    return False

def verify():
    bad = []
    for r in ROWS:
        if r.get('fig'):
            continue  # 그림 안 글자 — 세션이 그림 파일로 확인(✓그림)
        for q in quotes(r['cur']):
            if not exists(q):
                bad.append((r['id'], q))
    return bad

# ───────────────────────── docx XML ─────────────────────────
FONT = '<w:rFonts w:ascii="맑은 고딕" w:cs="맑은 고딕" w:eastAsia="맑은 고딕" w:hAnsi="맑은 고딕"/>'
def esc(t): return html.escape(t, quote=False)
def run(t, sz=17, b=False, color=None):
    rpr = FONT + ('<w:b/><w:bCs/>' if b else '<w:b w:val="false"/><w:bCs w:val="false"/>') + \
          (f'<w:color w:val="{color}"/>' if color else '') + f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>'
    return f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{esc(t)}</w:t></w:r>'
def para(t, sz=17, b=False, color=None, before=0, after=80, keep=False):
    ppr = ('<w:keepNext/>' if keep else '') + f'<w:spacing w:after="{after}" w:before="{before}"/>'
    return f'<w:p><w:pPr>{ppr}</w:pPr>{run(t, sz, b, color)}</w:p>'
def h1(t): return para(t, 24, True, '1F3864', before=240, after=120, keep=True)
def cell_paras(t, sz=16, b=False):
    ps = [p for p in t.split('\n')] or ['']
    return ''.join(f'<w:p>{run(p, sz, b)}</w:p>' for p in ps)
def tc(t, w, head=False):
    shd = '<w:shd w:fill="DCE3EC" w:color="auto" w:val="clear"/>' if head else ''
    bd = ''.join(f'<w:{s} w:val="single" w:color="9AA5B1" w:sz="4"/>' for s in ('top','left','bottom','right'))
    mar = '<w:tcMar><w:top w:type="dxa" w:w="40"/><w:left w:type="dxa" w:w="80"/><w:bottom w:type="dxa" w:w="40"/><w:right w:type="dxa" w:w="80"/></w:tcMar>'
    return f'<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="{w}"/><w:tcBorders>{bd}</w:tcBorders>{shd}{mar}</w:tcPr>{cell_paras(t, 17 if head else 16, head)}</w:tc>'
def table(widths, header, data):
    tw = sum(widths)
    bd = ''.join(f'<w:{s} w:val="single" w:color="auto" w:sz="4"/>' for s in ('top','left','bottom','right','insideH','insideV'))
    grid = ''.join(f'<w:gridCol w:w="{w}"/>' for w in widths)
    rows_x = [f'<w:tr><w:trPr><w:cantSplit/><w:tblHeader/></w:trPr>' + ''.join(tc(h, w, True) for h, w in zip(header, widths)) + '</w:tr>']
    for d in data:
        rows_x.append('<w:tr><w:trPr><w:cantSplit/></w:trPr>' + ''.join(tc(c, w) for c, w in zip(d, widths)) + '</w:tr>')
    return f'<w:tbl><w:tblPr><w:tblW w:type="dxa" w:w="{tw}"/><w:tblBorders>{bd}</w:tblBorders></w:tblPr><w:tblGrid>{grid}</w:tblGrid>{"".join(rows_x)}</w:tbl>'

W7 = [640, 1500, 700, 3500, 3500, 4000, 1000]
H7 = ['번호', '위치', '구분', '현행(요지)', '수정안', '이유·근거', '출처·확인']

def counts(sec=None):
    rs = [r for r in ROWS if sec is None or r['sec'] == sec]
    c = {k: sum(1 for r in rs if r['kind'] == k) for k in KIND_ORDER}
    return c

def md5(p):
    return hashlib.md5(open(p, 'rb').read()).hexdigest()

def build_document_xml(summary_paras, method_paras):
    body = []
    body.append(para(f'적합성 검토표 — phd_v201d_20260930 전체 검토 (세션 54~56)', 30, True, '1F3864', after=120))
    body.append(para(f'2026-10-05 · v1 · 대상 판 phd_v201d_20260930.docx(md5 {md5(THESIS)}, LibreOffice 203쪽, 참고문헌 61건) · 지침·자료는 검토 묶음(session56_outputs.zip: kit/지침_*.md·out/*.md)'))
    body.append(para('논문은 고치지 않았다. 10/7 출력본(v200 + 면담자료 Ver4.2)과 무관하며, 번호를 골라 승인해 주시면 10/7 뒤 판(v201e 또는 v202)에 반영한다.'))
    body.append(h1('1. 요약'))
    for p in summary_paras: body.append(para(p))
    body.append(h1('2. 검토 방법'))
    for p in method_paras: body.append(para(p))
    n = 3
    sec_title = {'그림': '그림 (그림 35장 안의 글자)', '표': '표 (표 119개의 셀)', '내용': '내용 (본문·주석·목차·참고문헌)'}
    for s in SECS:
        c = counts(s)
        body.append(h1(f'{n}. {sec_title[s]} — 필수 {c["필수"]} · 결정 필요 {c["결정 필요"]} · 권고 {c["권고"]} · 선택 {c["선택"]}'))
        data = sorted([r for r in ROWS if r['sec'] == s], key=lambda r: (KIND_ORDER[r['kind']], r['id']))
        body.append(table(W7, H7, [[r['id'], r['loc'], r['kind'], r['cur'], r['fix'], r['why'], r['src']] for r in data]))
        n += 1
    body.append(h1(f'{n}. 반영하지 않는 후보와 근거 ({len(EXCLUDED)}건)')); n += 1
    body.append(para('검토자가 틀렸거나 이미 결정된 것 — 결정기록·9/30 검토표와 대조.'))
    body.append(table([1800, 5600, 7500], ['후보', '내용', '반영하지 않는 근거'], [list(e) for e in EXCLUDED]))
    body.append(h1(f'{n}. 이상 없음으로 확인한 것 (1차 검토 요약)')); n += 1
    body.append(table([2200, 12700], ['검토', '확인한 것'], [list(o) for o in OKLIST]))
    body.append(h1(f'{n}. 검토자 기록')); n += 1
    body.append(table([2600, 2200, 10100], ['검토', '모델·시각', '범위·산출'], [list(v) for v in REVIEWERS]))
    sect = '<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="800" w:right="700" w:bottom="800" w:left="700" w:header="708" w:footer="708" w:gutter="0"/><w:pgNumType/><w:docGrid w:linePitch="360"/></w:sectPr>'
    tpl_xml = zipfile.ZipFile(TPL).read('word/document.xml').decode('utf8')
    head = tpl_xml.split('<w:body>')[0]
    return head + '<w:body>' + ''.join(body) + sect + '</w:body></w:document>'

def write_docx(doc_xml):
    zin = zipfile.ZipFile(TPL)
    zout = zipfile.ZipFile(DOCX, 'w', zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == 'word/document.xml':
            data = doc_xml.encode('utf8')
        elif item.filename == 'word/comments.xml':
            continue
        elif item.filename == 'word/_rels/document.xml.rels':
            data = re.sub(r'<Relationship Id="rId6"[^>]*comments[^>]*/>', '', data.decode('utf8')).encode('utf8')
        elif item.filename == '[Content_Types].xml':
            data = re.sub(r'<Override [^>]*comments\+xml[^>]*/>', '', data.decode('utf8')).encode('utf8')
        elif item.filename == 'docProps/core.xml':
            d = data.decode('utf8')
            d = re.sub(r'<dc:title>.*?</dc:title>', '', d)
            data = d.encode('utf8')
        zout.writestr(item, data)
    zout.close()

# ───────────────────────── HTML (PDF 인쇄용) ─────────────────────────
def build_html(summary_paras, method_paras):
    def e(t): return esc(t).replace('\n', '<br>')
    css = """@page{size:A4 landscape;margin:12mm 12mm 14mm 12mm}
    body{font-family:'Noto Sans KR','맑은 고딕','WenQuanYi Zen Hei',sans-serif;font-size:8.5pt;line-height:1.35;color:#111}
    h1{font-size:15pt;color:#1F3864;margin:0 0 6pt 0} h2{font-size:12pt;color:#1F3864;margin:14pt 0 6pt 0;page-break-after:avoid}
    p{margin:0 0 4pt 0} table{border-collapse:collapse;width:100%;table-layout:fixed;margin-bottom:6pt}
    th,td{border:0.5pt solid #9AA5B1;padding:2pt 4pt;vertical-align:top;font-size:8pt;word-break:break-all}
    th{background:#DCE3EC;font-weight:700;font-size:8.5pt} tr{page-break-inside:avoid} thead{display:table-header-group}
    .meta{font-size:8.5pt}"""
    def tbl(widths, header, data):
        tot = sum(widths)
        cols = ''.join(f'<col style="width:{w*100/tot:.1f}%">' for w in widths)
        h = ''.join(f'<th>{e(x)}</th>' for x in header)
        b = ''.join('<tr>' + ''.join(f'<td>{e(c)}</td>' for c in row) + '</tr>' for row in data)
        return f'<table><colgroup>{cols}</colgroup><thead><tr>{h}</tr></thead><tbody>{b}</tbody></table>'
    out = [f'<!doctype html><html lang="ko"><head><meta charset="utf-8"><title>{NAME}</title><style>{css}</style></head><body>']
    out.append('<h1>적합성 검토표 — phd_v201d_20260930 전체 검토 (세션 54~56)</h1>')
    out.append(f'<p class="meta">2026-10-05 · v1 · 대상 판 phd_v201d_20260930.docx(md5 {md5(THESIS)}, LibreOffice 203쪽, 참고문헌 61건) · 지침·자료는 검토 묶음(session56_outputs.zip: kit/지침_*.md·out/*.md)</p>')
    out.append('<p class="meta">논문은 고치지 않았다. 10/7 출력본(v200 + 면담자료 Ver4.2)과 무관하며, 번호를 골라 승인해 주시면 10/7 뒤 판(v201e 또는 v202)에 반영한다.</p>')
    out.append('<h2>1. 요약</h2>' + ''.join(f'<p>{e(p)}</p>' for p in summary_paras))
    out.append('<h2>2. 검토 방법</h2>' + ''.join(f'<p>{e(p)}</p>' for p in method_paras))
    n = 3
    sec_title = {'그림': '그림 (그림 35장 안의 글자)', '표': '표 (표 119개의 셀)', '내용': '내용 (본문·주석·목차·참고문헌)'}
    for s_ in SECS:
        c = counts(s_)
        out.append(f'<h2>{n}. {e(sec_title[s_])} — 필수 {c["필수"]} · 결정 필요 {c["결정 필요"]} · 권고 {c["권고"]} · 선택 {c["선택"]}</h2>')
        data = sorted([r for r in ROWS if r['sec'] == s_], key=lambda r: (KIND_ORDER[r['kind']], r['id']))
        out.append(tbl(W7, H7, [[r['id'], r['loc'], r['kind'], r['cur'], r['fix'], r['why'], r['src']] for r in data])); n += 1
    out.append(f'<h2>{n}. 반영하지 않는 후보와 근거 ({len(EXCLUDED)}건)</h2><p>검토자가 틀렸거나 이미 결정된 것 — 결정기록·9/30 검토표와 대조.</p>'); n += 1
    out.append(tbl([1800, 5600, 7500], ['후보', '내용', '반영하지 않는 근거'], [list(x) for x in EXCLUDED]))
    out.append(f'<h2>{n}. 이상 없음으로 확인한 것 (1차 검토 요약)</h2>'); n += 1
    out.append(tbl([2200, 12700], ['검토', '확인한 것'], [list(o) for o in OKLIST]))
    out.append(f'<h2>{n}. 검토자 기록</h2>'); n += 1
    out.append(tbl([2600, 2200, 10100], ['검토', '모델·시각', '범위·산출'], [list(v) for v in REVIEWERS]))
    out.append('</body></html>')
    return ''.join(out)

# ───────────────────────── 판정 md ─────────────────────────
def write_md():
    for s in SECS:
        fn = os.path.join(OUT, f'판정_{s}.md')
        data = sorted([r for r in ROWS if r['sec'] == s], key=lambda r: (KIND_ORDER[r['kind']], r['id']))
        c = counts(s)
        L = [f'# 2차 판정 — {s} (세션 56, 2026-10-05) · 필수 {c["필수"]} · 결정 필요 {c["결정 필요"]} · 권고 {c["권고"]} · 선택 {c["선택"]}', '',
             '판정 방법: 1차 검토 후보마다 kit/dump_all.txt(Grep)·그림 파일·표 파일로 「현행」의 실재와 근거를 다시 확인하고, 중복을 합치고, 지침_공통의 「이미 결정된 것·결정 대기」와 결정기록_메모리·9/30 검토표(적합성검토표_v201.txt)와 대조해 제외할 것을 가렸다. 판정자: 주 세션(Claude Fable 5.1). 출처 칸의 ✓는 이 세션이 직접 확인, △는 검토자 근거만, ✓그림은 그림 파일로 확인, ○원문은 표준 원문 확인이 남은 것.', '']
        for r in data:
            L += [f'### {r["id"]} [{r["kind"]}] {r["loc"]}', f'- 현행: {r["cur"]}', f'- 수정안: {r["fix"]}', f'- 이유·근거: {r["why"]}', f'- 출처·확인: {r["src"]}', '']
        ex = [e for e in EXCLUDED if any(k in e[0] for k in ({'그림': ['G1', 'G2'], '표': ['T1', 'T2', 'T3', 'T4'], '내용': ['N-', 'R3', 'R4']}[s]))]
        if ex:
            L += ['## 반영하지 않는 후보', '']
            for e in ex: L += [f'- {e[0]} — {e[1]} → {e[2]}', '']
        open(fn, 'w', encoding='utf8').write('\n'.join(L))

if __name__ == '__main__':
    bad = verify()
    print('「현행」 실재 검증:', '전부 실재' if not bad else f'미실재 {len(bad)}건')
    for b in bad: print('   MISSING', b[0], '::', b[1][:120])
    tot = counts()
    print('집계:', tot, '제외', len(EXCLUDED))
    write_md()
    from summary56 import SUMMARY, METHOD
    xml = build_document_xml(SUMMARY(tot, ROWS, EXCLUDED), METHOD)
    write_docx(xml)
    print('docx:', DOCX, os.path.getsize(DOCX))
    # PDF — 이 컨테이너의 LibreOffice가 docx를 열지 못해(「source file could not be loaded」, 본보기 파일도 동일) 같은 데이터로 HTML을 만들어 Chromium으로 인쇄한다.
    HTML = os.path.join(OUT, NAME + '.html'); PDF = os.path.join(OUT, NAME + '.pdf')
    open(HTML, 'w', encoding='utf8').write(build_html(SUMMARY(tot, ROWS, EXCLUDED), METHOD))
    r = subprocess.run(['/opt/pw-browsers/chromium', '--headless=new', '--no-sandbox', '--disable-gpu', '--no-pdf-header-footer',
                        '--print-to-pdf=' + PDF, 'file://' + HTML], capture_output=True, text=True, timeout=300)
    print('pdf:', os.path.exists(PDF) and os.path.getsize(PDF), r.stderr.strip().splitlines()[-1:] )
