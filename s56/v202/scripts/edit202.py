# v202 (2026-10-05, 세션 56): v201d(md5 e5ee08b4…) + 적합성검토표_v201d_20261005_v1의 필수 4·권고(첫째 안)·결정 필요 (가) + 세션 53 확정분(ISO/IEC TS 33030[21] 삭제·참고문헌 재부여 [22]~[61]→[21]~[60]) + 목차 필드 갱신 플래그
# 선택(3G-06·3T-19)과 3T-09(원문 확인 뒤)는 반영하지 않음. 그림 4장은 fig202.py.
import sys, json, re, copy
sys.path.insert(0, '/home/claude/s56w')
from lib191 import Doc, q
from lxml import etree
X = '/home/claude/s56w/x/word/'
d = Doc(X + 'document.xml')
LOG = []
XMLNS = 'http://www.w3.org/XML/1998/namespace'

def rep_first(p, old, new):
    ts = list(p.iter(q('t'))); full = ''.join(t.text or '' for t in ts)
    k = full.find(old); assert k >= 0, old
    st, en = k, k + len(old); pos = 0; first = True
    for t in ts:
        tx = t.text or ''; a = pos; b = pos + len(tx); pos = b
        if b <= st or a >= en: continue
        i0 = max(st, a) - a; i1 = min(en, b) - a
        if first: t.text = tx[:i0] + new + tx[i1:]; first = False
        else: t.text = tx[:i0] + tx[i1:]
        t.set('{%s}space' % XMLNS, 'preserve')

def one(old, new, tag, where, n=1):
    ps = [p for p in d.body.iter(q('p')) if old in d.text(p)]
    cnt = sum(d.text(p).count(old) for p in ps)
    assert cnt == n, (tag, old, cnt)
    for p in ps:
        for _ in range(d.text(p).count(old)): rep_first(p, old, new)
    LOG.append(dict(tag=tag, where=where, old=old, new=new, n=n))

def one_exact(txt, new, tag, where, count=1, idx=None):
    ps = [p for p in d.body.iter(q('p')) if d.text(p) == txt]
    assert len(ps) == count, (tag, txt, len(ps))
    for p in ([ps[idx]] if idx is not None else ps):
        rep_first(p, txt, new)
    LOG.append(dict(tag=tag, where=where, old=txt, new=new, n=1 if idx is not None else count))

def row_cell(first, cell_txt, new, tag, where):
    """첫 칸이 first인 표 행에서 글자가 cell_txt인 칸 하나를 new로."""
    hit = []
    for tr in d.body.iter(q('tr')):
        tcs = tr.findall(q('tc'))
        if not tcs: continue
        if ''.join(d.text(p) for p in tcs[0].findall(q('p'))).strip() != first: continue
        for tc in tcs[1:]:
            ps = tc.findall(q('p'))
            if len(ps) == 1 and d.text(ps[0]) == cell_txt: hit.append(ps[0])
    assert len(hit) == 1, (tag, first, cell_txt, len(hit))
    rep_first(hit[0], cell_txt, new)
    LOG.append(dict(tag=tag, where=where, old=f'{first} 행 「{cell_txt}」', new=new, n=1))

# ═════ 필수 4
one_exact('2006년 개정 시 ISO 14044로 통합된 구판 — 2.2.2 LCA 표준 체계 연혁 서술', '2006년 개정 시 ISO 14044로 통합된 구판 — 2.2.1 LCA 표준 체계 연혁 서술', '3T-01', '1.3.3 [표 1-1] ISO 14041~14043 행 (T1.r23.c4)')
one_exact('Before/After 검토 평가', '개선 전·후 검토 평가', '3T-02', '부록 D 표 D-1 A-2 행 (T104.r49.c4)')
one('일관성 검사을 수행하며', '일관성 검사를 수행하며', '3N-01', '4.6.1 R5 설명 문단 (B966)')
# 3G-01 [그림 4-6] — fig202.py

# ═════ 권고 (첫째 안)
one('프로세스 능력 수준(0~5)으로 올리는 틀이다', '프로세스 능력도 수준(0~5)으로 올리는 틀이다', '3G-02', '2.4.3 B460 (그림 2-6은 fig202.py)')
one_exact('핵심 평가 질문', '평가 질문', '3T-03', '3.3.2 [표 3-6] 머리글 (T12.r1.c2)')
row_cell('P-10', '14001 6.1.2 ↔ 14044 4.4.2', '14001 6.1.2 ↔ 14044 4.4.2·4.4.3(추가④)', '3T-04', '4.3.1 [표 4-9] P-10 행 조항 칸 (T26.r14.c4)')
one_exact('14044 4.2.3.8·4.2.3.1 (shall)', '14044 4.2.3.8·4.2.3.1 (14001 —)', '3T-05', '4.3.1 [표 4-9] P-21 행 조항 칸 (T26.r27.c4)')
one('문서와 기록의 실재를 기준으로 판정하고 면담과 현장 확인을 포함하지 않으므로', '문서와 기록의 실재를 기준으로 판정하고 면담과 실행 상태의 현장 관찰을 포함하지 않으므로', '3T-06', '5.2 [표 5-1] 문서 심사 방식 행 (T36.r13.c3)')
one('※ 본문 [표 3-3]에서는 9.3(경영검토)을 문항 배정(A-1·A-2, LA.1)에 맞추어 A 행으로 옮겨 실었다.',
    '※ 본문 [표 3-3]에서는 9.3(경영검토)을 문항 배정(A-1·A-2, LA.1)에 맞추어 A 행으로 옮겨 싣고, P 행의 LCA 대응 단계에 6.2.1 환경목표가 대응하는 해석 §4.5를, A 행에 9.3이 대응하는 보고 §5.1을 더하였으며, P 행의 연계 가능성은 표 B-1대로 4.3을 「높음」으로 구분하여 적었다.',
    '3T-07', '부록 B.5 표 B-5 뒤 주석 (B1160)')
one_exact('KS Q ISO 14031:2013, ISO 14006, ISO 14025, KS I ISO 14067:2018, ISO 9001, ISO 9004, ISO 45001', 'KS Q ISO 14031:2013, ISO 14006:2020, ISO 14025:2006, KS I ISO 14067:2018, ISO 9001, ISO 9004:2018, ISO 45001', '3T-08', '1.3.3 [표 1-1] 배경·보완 언급 행 (T1.r22.c2)')
# 3T-09 — 원문 확인 뒤(미반영)
# 3T-11 표 B-3 ML4 칸: 한 문단 안의 <w:br/> 4개(【】 뒤 1 + 끝 3)를 항목 사이로 옮김
_hit = [p for p in d.body.iter(q('p')) if d.text(p).startswith('【shall 완전 충족 + should 일부 충족】• 4.4.3')]
assert len(_hit) == 1
_p = _hit[0]
_full = d.text(_p)
_segs = ['【shall 완전 충족 + should 일부 충족】', '• 4.4.3 선택요소(가중평가) 수행(조건부 shall)', '• 4.5.3 신뢰도 평가 체계적 수행', '• 6.2 내부 전문가 검토 도입(조건부 shall)', '• 6.2 외부 전문가 검토 도입(조건부 shall)']
assert ''.join(_segs) == _full, _full
_runs = [r for r in _p if r.tag == q('r')]
_rpr = copy.deepcopy(_runs[0].find(q('rPr'))) if _runs[0].find(q('rPr')) is not None else None
for r in _runs: _p.remove(r)
for i, sg in enumerate(_segs):
    r = etree.SubElement(_p, q('r'))
    if _rpr is not None: r.append(copy.deepcopy(_rpr))
    if i > 0: etree.SubElement(r, q('br'))
    t = etree.SubElement(r, q('t')); t.text = sg; t.set('{%s}space' % XMLNS, 'preserve')
assert d.text(_p) == _full and len(_p.findall('.//' + q('br'))) == 4
LOG.append(dict(tag='3T-11', where='부록 B.3 표 B-3 ML4 행 ISO 14044 충족 요건 칸 (T43.r5.c3)', old='【…】⏎• 4.4.3 …• 4.5.3 …• 6.2 …• 6.2 …⏎⏎⏎', new='【…】⏎• 4.4.3 …⏎• 4.5.3 …⏎• 6.2 …⏎• 6.2 … (끝 빈 줄 3개 삭제)', n=1))
one('A-1·A-5의 초기 판정 포인트 행 첫머리에 단독으로 붙은 ★는 전사 원본 시트의 「일반 핵심」 표기(부록 H 표 H-1 범례)로서 층위가 다르다',
    '핵심 문항 4종과 A-1·A-5의 초기 판정 포인트 행 첫머리에 단독으로 붙은 ★는 전사 원본 시트의 「일반 핵심」 표기(부록 H 표 H-1 범례)로서 ◆★와 층위가 다르다', '3T-12', '부록 C 머리 주석 (B1170)')
one_exact('필수 2/2 — 비교주장 정책이 결정·문서화되고 해당 시 패널 검토 수행', '필수 2/2 — 비교주장 정책이 결정·문서화되고 해당 시 패널 검토 수행 (shall 충족 기준)', '3T-13', '부록 C C-7 카드 3점 행 (T91.r9)')
one_exact('필수 2/2 — 비교주장 정책 결정·문서화, 해당 시 §5.3 요건 충족', '필수 2/2 — 비교주장 정책 결정·문서화, 해당 시 §5.3 요건 충족 (shall 충족 기준)', '3T-13', '부록 C A-9 카드 3점 행 (T103.r9)')
one('재검증 워크시트 ? 전량 해소', '재검증 워크시트 ⏸ 전량 해소', '3T-15', '부록 G.2 표 G-2 r8 (T110.r8) — 깨진 기호를 ⏸(잔여 확인)로 복원(추정)')
one_exact('③ 나머지 보유분 (E7~E11): 12문항', '③ 나머지 보유분 (E7~E11): 15문항', '3T-16', '부록 G.2 표 G-2 r5 (T110.r5)')
one_exact('ISO/IEC Directives, Part 2 · KS A 0001:2023 부속서 I', 'KS A 0001:2023[24] 부속서 I', '3T-17', '부록 K shall/should 행 출처 (T119.r14.c4) — [24]는 재부여 뒤 [23]')
one_exact('연구에서 제외되는 물질·에너지 흐름의 양 또는 환경적 중대성에 대한 규정.', '연구에서 제외되는 물질·에너지 흐름의 양 또는 환경적 중대성에 대한 규정', '3T-18', '부록 K 제외 기준 행 뜻 칸 (T119.r20.c3)')
one('성숙도 모델은 ML0~ML5의 판정 규칙과 핵심 문항 4종으로 이루어진다.', '성숙도 모델은 성숙도 수준별 기준과 ML0~ML5의 판정 규칙, 핵심 문항 4종으로 이루어진다.', '3N-02', '1.3.1 (B296)')
one('성숙도 모델(ML0~ML5 판정 규칙과 핵심 문항 4종)의 3계층', '성숙도 모델(수준별 기준·ML0~ML5 판정 규칙·핵심 문항 4종)의 3계층', '3N-02', '부록 K LMA 행 (T119.r50.c3)')
one('긴 제목(4.4.5·5.3)', '긴 제목(4.4.5·5.3·6.3)', '3N-03', '부록 A 머리 주석 (B1104)')
one('※ 표 J-1(전수 데이터)·표 J-2(영역 요약)는 본문 4.3의 [표 4-9]~[표 4-13]과 중복되므로 생략하고', '※ 원자료의 전수 데이터 표와 영역 요약 표는 본문 4.3의 [표 4-9]~[표 4-13]과 중복되므로 생략하고', '3N-04', '부록 J 머리 주석 (B1381)')

# ═════ 결정 필요 — 안 (가)
# 3G-05 (가) 현행 유지 — 무변경
one('9.1.1 행의 6.3(외부 전문가 검토)은 6.2로 정정하여 옮겼다.', '9.1.1 행의 6.3(외부 전문가 검토)은 6.2로 정정하여 옮겼다(이 6.2는 [표 3-1]에서는 9.2.1 쪽 근거로 정리하여 9.1.1 행에 적지 않았다).', '3T-10(가)', '부록 A 머리 주석 (B1104)')
one('◆ 판정 원본은 「평가 기록부 (확정)」 — 불일치 시 기록부 우선.', '◆ 판정 원본은 「평가 기록부 (확정)」 — 불일치 시 기록부 우선. 【평가 결론】 블록의 평균·ML은 잠정 채점 시점의 값이며, 종합 「ML3」 표기는 영역별 판정을 묶기 전의 시트 기재로 확정 판정(ML2, [표 4-13])과 다르다.', '3T-14(가)', '부록 H.1 표 H-1 머리 주석 (B1348)')

# ═════ 세션 53 확정분: ISO/IEC TS 33030[21] 삭제 + 참고문헌 재부여 [22]~[61] → [21]~[60]
one('이는 KS X ISO 8000-62 4.7 비고 2가 참조하는 ISO/IEC TS 33030:2017[21]의 평가자 역량 검증 체계와 같은 취지이며, 자격 인정의 일반 요건은 ISO/IEC 17024를 따른다.', '자격 인정의 일반 요건은 ISO/IEC 17024를 따른다.', 'TS33030', '3.3.4 (B742)')
_r21 = [p for p in d.body.iter(q('p')) if d.text(p).startswith('[21] ISO/IEC TS 33030:2017')]
assert len(_r21) == 1
_r21[0].getparent().remove(_r21[0])
LOG.append(dict(tag='TS33030', where='참고문헌 [21]', old='[21] ISO/IEC TS 33030:2017, Information technology — Process assessment — An exemplar documented assessment process.', new='(삭제)', n=1))
assert not any('[21]' in d.text(p) for p in d.body.iter(q('p')))
def rep_at(p, st, en, new):
    ts = list(p.iter(q('t'))); pos = 0; first = True
    for t in ts:
        tx = t.text or ''; a = pos; b = pos + len(tx); pos = b
        if b <= st or a >= en: continue
        i0 = max(st, a) - a; i1 = min(en, b) - a
        if first: t.text = tx[:i0] + new + tx[i1:]; first = False
        else: t.text = tx[:i0] + tx[i1:]
        t.set('{%s}space' % XMLNS, 'preserve')
_n = 0
for p in d.body.iter(q('p')):
    ms = [m for m in re.finditer(r'\[(\d+)\]', d.text(p)) if 22 <= int(m.group(1)) <= 61]
    for m in reversed(ms):
        rep_at(p, m.start(1), m.end(1), str(int(m.group(1)) - 1)); _n += 1
LOG.append(dict(tag='TS33030', where='본문·표·부록·참고문헌 전체(번호 재부여)', old='[22]~[61]', new='[21]~[60]', n=_n))
print('renumbered', _n)

d.save()
# ═════ 3N-05 목차 필드 갱신 — Word가 열 때 필드 갱신을 묻도록 settings.xml에 updateFields
S = X + 'settings.xml'
st = etree.parse(S); sr = st.getroot()
if sr.find(q('updateFields')) is None:
    uf = etree.Element(q('updateFields')); uf.set(q('val'), 'true')
    anchor = None
    for tag in ('hdrShapeDefaults', 'footnotePr', 'endnotePr', 'compat', 'rsids'):
        anchor = sr.find(q(tag))
        if anchor is not None: break
    if anchor is not None: anchor.addprevious(uf)
    else: sr.append(uf)
    st.write(S, xml_declaration=True, encoding='UTF-8', standalone=True)
LOG.append(dict(tag='3N-05', where='word/settings.xml', old='(없음)', new='<w:updateFields w:val="true"/> — Word가 열 때 목차·표목차·그림목차 쪽 번호 갱신을 물음(한 번 갱신·저장 뒤 사라져도 됨)', n=1))
json.dump(LOG, open('/home/claude/s56w/log_202.json', 'w'), ensure_ascii=False, indent=1)
print(len(LOG), 'changes')
