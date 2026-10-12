# 면담자료 Ver5(간결판) — Ver4.2의 패키지·서식을 그대로 쓰고 body만 다시 짬: 9/16 지시 → 수정 결과 + 심사 일정 (+ 출력본 뒤 변경 참고, 그 밖에 여쭐 것)
import copy, re, zipfile, shutil, os
from lxml import etree
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'; XML='http://www.w3.org/XML/1998/namespace'
q=lambda t:'{%s}%s'%(W,t)
SRC='x'; OUT='면담자료_20261012_Ver5.1.docx'
tree=etree.parse(SRC+'/word/document.xml'); root=tree.getroot(); body=root.find(q('body')); K=list(body)
def txt(e): return ''.join(t.text or '' for t in e.iter(q('t')))
def set_text(p,text):
    runs=[r for r in p if r.tag==q('r')]
    rpr=copy.deepcopy(runs[0].find(q('rPr'))) if runs and runs[0].find(q('rPr')) is not None else None
    for r in runs: p.remove(r)
    for c in list(p):
        if c.tag in (q('hyperlink'),q('commentRangeStart'),q('commentRangeEnd')): p.remove(c)
    r=etree.SubElement(p,q('r'))
    if rpr is not None: r.append(rpr)
    t=etree.SubElement(r,q('t')); t.text=text; t.set('{%s}space'%XML,'preserve'); return p
def para(tmpl,text): return set_text(copy.deepcopy(tmpl),text)
def table(tmpl,rows,keep_header=True,header=None):
    t=copy.deepcopy(tmpl); trs=t.findall(q('tr'))
    hdr=trs[0]; data_tmpl=trs[1]
    if header:
        for tc,val in zip(hdr.findall(q('tc')),header): set_text(tc.findall(q('p'))[0],val)
    for tr in trs[1:]: t.remove(tr)
    for row in rows:
        tr=copy.deepcopy(data_tmpl); tcs=tr.findall(q('tc'))
        assert len(tcs)==len(row),(len(tcs),row)
        for tc,val in zip(tcs,row):
            ps=tc.findall(q('p'))
            for p in ps[1:]: tc.remove(p)
            set_text(ps[0],val)
        t.append(tr)
    return t
# 템플릿
P_TITLE=K[1]; P_META=K[6]; H1=K[9]; P_LIST=K[10]; P_NORM=K[20]; P_INTRO=K[22]
T_TOC=K[8]; T_SUM=K[17]; T_SCHED=K[19]; T_A=K[23]
# A-표의 상태 칸 보정(별도 자료 절 번호 참조)
TA=copy.deepcopy(T_A)
for tc in TA.iter(q('tc')):
    s=txt(tc).strip()
    if s=='확인 (5.2 나-1)': set_text(tc.findall(q('p'))[0],'확정 (3절)')
    elif s=='진행 중': set_text(tc.findall(q('p'))[0],'준비 완료 (v22·v7)')
    elif s.startswith('발표 자료 v3(본 29장·백업 37장)'): set_text(tc.findall(q('p'))[0],'발표 자료 v22(본 30장·백업 45장, 노트 대본). 예상 질문 v7 48문 — 답의 뼈대와 논문 속 근거 위치, 답할 백업 번호')
    elif s=='반영 (문안 확인 4.1)': set_text(tc.findall(q('p'))[0],'반영 (문안은 Ver4.2 4.1)')
    elif s=='반영 (문안 확인 4.2)': set_text(tc.findall(q('p'))[0],'반영 (문안은 Ver4.2 4.2)')
    elif s=='v199 반영 내용': set_text(tc.findall(q('p'))[0],'반영 내용 (v200 출력본)')
    elif s.startswith('2절. 심사위원 후보와 날짜는 5.2 나-1'): set_text(tc.findall(q('p'))[0],'심사 신청·서류 완료. 심사위원 5인(위원장 김동욱·부위원장 이진원)과 심사 3회 일정(10/22·11/12·11/26) 확정 — 3절')
    elif '6절' in s or '5.2' in s or '5.1' in s or '4.1' in s or '4.2' in s:
        ps=tc.findall(q('p'))
        if len(ps)==1:
            t2=s.replace('예상 질문 43문은 6절','예상 질문 43문은 별도 자료 Ver4.2 6절').replace('(5.2 나-1)','(3절)').replace('5.2 나-','Ver4.2 5.2 나-').replace('문안 확인 4.1','문안은 Ver4.2 4.1').replace('문안 확인 4.2','문안은 Ver4.2 4.2').replace('v199에서 고침(5.1)','v199에서 고침(Ver4.2 5.1)')
            if t2!=s: set_text(ps[0],t2)
import re as _re
for tc in TA.iter(q('tc')):
    s0=txt(tc)
    if s0.strip()=='9/16 지적 (녹취 시각)': set_text(tc.findall(q('p'))[0],'9/16 지적'); continue
    s1=_re.sub(r'\s*\(\d{1,2}:\d{2}(?:[~·]\d{1,2}:\d{2})*\)','',s0)
    if s1!=s0:
        ps=tc.findall(q('p'))
        for p in ps[1:]: tc.remove(p)
        set_text(ps[0],s1.strip())
# 새 body
NB=[]
NB.append(copy.deepcopy(K[0]))
NB.append(para(P_TITLE,'2026년 10월 지도교수 면담 준비자료 (예비심사 10/22 대비)'))
for i in (2,3,4,5): NB.append(copy.deepcopy(K[i]))
NB.append(para(P_META,'대상 판: 현행 phd_v203_20261007.docx(10/8 제출, 참고문헌 61건) — 교수님 검토 요청 판. 교수님께서 마지막으로 보신 판은 9/16 면담의 v98(2026-09-16)이며, 그 뒤의 변경은 2절(지시 사항 반영)과 4절(그 밖의 변경)에 정리'))
NB.append(para(P_META,'작성일: 2026-10-12 · 면담자료 Ver5.1 (Ver5(10/6)를 v203·확정 일정 기준으로 갱신 — 9/16 지시 사항의 수정 결과, 심사 일정과 심사위원, 여쭐 것 세 가지. 국문 초록 전문·3.6 결론·경량화 메모는 별도 자료 Ver4.2, 예상 질문은 v7 48문 별도)'))
NB.append(table(T_TOC,[('1','한눈에 보기'),('2','9/16 지시 사항과 수정 결과 (A-01~A-39)'),('3','심사 일정과 심사위원 — 확정 사항'),('4','9/16 보신 판(v98) 뒤에 고친 것 — v203까지'),('5','여쭐 것 세 가지와 송부·제본 계획')]))
NB.append(para(H1,'1. 한눈에 보기'))
NB.append(para(P_LIST,'9/16 면담에서 주신 지적을 메모와 빨간 펜 표시로 정리하여 39건(A-01~A-39)으로 나누었고, 논문에 관한 33건(중복 3건 제외)은 국문 초록부터 부록·영문 초록까지 모두 반영하였습니다(2절). 논문 외 3건은 심사 신청·학과 서류 제출 완료, 심사위원 5인과 심사 3회 일정(10/22·11/12·11/26)은 확정(3절), 발표 자료는 v22까지 만들어 교수님 확인을 기다립니다.'))
NB.append(para(P_LIST,'교수님 검토를 부탁드리는 판은 v203(10/7)입니다. 9/16에 보신 v98과 견주면 지시 사항 39건 반영(2절) 외에 2장·3장의 절 구성 재편, 그림 22장·표 7개 추가, 용어 통일(표기 규칙표), 참고문헌 45→61건, 개선 과제 완수 시 전망(3.15점·ML3) 추가가 있습니다(4절). 사례 평가의 수치와 판정(55문항 종합 2.85점·ML2, 영역 평점)은 v98과 같습니다.'))
NB.append(para(P_LIST,'오늘 여쭐 것은 세 가지입니다 — 예비심사 발표 시간, 본심사에서 보고할 범위, 생성형 AI 활용의 논문 표기 여부(5절). 교수님 확인이 끝나면 소프트 제본을 넣고 심사위원께 논문과 발표자료를 같은 날 드리겠습니다.'))
NB.append(table(T_SUM,[('9/16 지적 — 논문','33건','전부 반영. A-37~A-39는 A-18·A-20·A-23과 같은 지적이라 중복으로 셈'),
                       ('9/16 지적 — 논문 외','3건','심사 신청·학과 서류 제출 완료(A-32) · 심사위원 5인·심사 3회 일정 확정(A-31, 3절) · 발표 자료 v22·예상 질문 v7 48문 준비(A-33)'),
                       ('v98 뒤 그 밖의 변경','구성·그림·표·용어·참고문헌','4절 — 2·3장 절 재편, 그림 12→34장, 표 31→36개, 참고문헌 45→61건, 본문 약 19% 증가(LibreOffice 기준 179→202쪽). 사례 수치·판정은 동일'),
                       ('여쭐 것','3건','5절 — 발표 시간 · 본심사 범위 · AI 표기')]))
NB.append(para(H1,'2. 9/16 지시 사항과 수정 결과 (A-01~A-39)'))
NB.append(para(P_INTRO,'「9/16 지적」은 면담 메모와 빨간 펜 표시에서 뽑은 요지이고, 「위치」는 v203의 절 번호입니다(v98과 절 번호가 다른 곳은 4절의 구성 표 참조). 초록의 [1]~[9]는 국문 초록의 문단 번호입니다. 상태 칸의 「Ver4.2 4.1·4.2」는 별도 자료의 절 번호입니다.'))
NB.append(TA)
NB.append(para(H1,'3. 심사 일정과 심사위원 — 확정 사항'))
NB.append(para(P_INTRO,'심사위원 5인(위원장·부위원장 포함)과 세 차례 심사 일시는 확정되었습니다. 장소는 확정되는 대로 채우겠습니다.'))
NB.append(table(T_SUM,[('위원장','김동욱 교수 — 강원대학교 산업경영공학과','본과'),
                       ('부위원장','이진원 교수 — 강원대학교 산업경영공학과','본과'),
                       ('위원(지도교수)','이창수 교수 — 강원대학교 산업경영공학과','본과'),
                       ('위원','김계수 교수 — 세명대학교 경영학과','외부위원 (생산관리·품질경영, 계량경영/MIS)'),
                       ('위원','서병석 교수 — 상지대학교 전자·보안공학부','외부위원 (정보시스템 보안, 머신러닝·딥러닝, 컴퓨터시스템)')],header=('역할','성함 · 소속','비고')))
NB.append(table(T_SUM,[('예비심사','2026년 10월 22일(목) 18:00','단독. 장소는 확정 뒤 기재. 발표 약 30분 분량 준비(시간이 20분이면 축소)'),
                       ('본심사 1차 + 공개발표','2026년 11월 12일(목) 18:00','석사과정 박현재 학생과 같은 자리 — 석사 발표 먼저, 박사 발표 뒤. 예비심사 지적 사항의 수정 결과 보고'),
                       ('본심사 2차','2026년 11월 26일(목) 18:00','단독. 12/9까지 3회 완료 요건 충족'),
                       ('결과 보고 → 온라인 제출','~12/10 → dCollection 2027.1.7~1.13','학위수여 사정 2027년 2월 초')],header=('일정','날짜 · 시간','비고')))
NB.append(copy.deepcopy(T_SCHED))
NB.append(copy.deepcopy(K[20]))
NB.append(para(P_LIST,'① 심사용 논문: v203을 A4 소프트커버 5부(+여유 2부)로 제본하여 예비심사 전에 심사위원께 직접 전달 — 교수님 검토가 끝난 뒤 인쇄합니다(수정 지시가 있으면 v204로 제본). 심사판정지는 이름을 비워 둔 채로 두고 본심사 통과 뒤 채웁니다.'))
NB.append(para(P_LIST,'② 발표자료: 본 슬라이드 30장 PDF(A4 가로)를 논문 제본본과 같은 날 심사위원께 드립니다. 백업 45장과 발표 노트는 배포하지 않습니다.'))
NB.append(para(P_LIST,'③ 사전 송부: ○○○ 교수님께서 논문을 미리 요청하셔서, 교수님 확인 뒤 v203 PDF를 보내 드리려 합니다.'))
NB.append(para(H1,'4. 9/16 보신 판(v98) 뒤에 고친 것 — v203까지'))
NB.append(para(P_LIST,'9/16 지시 사항 39건의 반영은 2절에 있습니다. 여기에는 그 밖의 변경을 종류별로 적었습니다. 사례 평가의 수치와 판정은 바뀌지 않았습니다.'))
NB.append(table(T_SUM,[('2장 구성 재편','절 이동·신설','v98 2.1.4~2.1.6(연계 타당성·조항별 연계 가능성·EMS Process Mapping)을 2.3.1~2.3.2(연계의 근거·프로세스 대응)로 옮기고, 2.4 「프로세스 성숙도 평가의 표준과 기법」(2.4.1 성숙도 모델의 정의와 세 계층 · 2.4.2 판정 체계 표준 · 2.4.3 측정 프레임워크·행동기술척도·집계 · 2.4.4 운영 준거 · 2.4.5 개발 절차와 내용타당도)을 신설. 선행연구는 2.4 → 2.5'),
                       ('3장 구성 재편','검증 절 재구성','제목 「연구 방법 및 연계 모델 개발」. 3.5를 「모델의 검증」으로 — 3.5.1 추적성 검증(신설), 3.5.2 ISO/IEC 33004 자체점검, 3.5.3 CVI의 설계(실시가 아니라 설계임을 명시). v98 3.4.5 「판정 규칙의 회귀 검증」은 4.6.2 「판정 규칙의 적용 확인」으로, v98 3.5.3 「평가 클래스와 평가자 독립성」은 2.4.4로'),
                       ('4장','절 제목·내용 보강','4.1 적용 개요 · 4.2.3 판정의 신뢰성(구 「객관성의 통제」) · 4.5.2 데이터 가용성의 구별과 판정의 민감도 · 4.6.2 기대 효과 시뮬레이션과 판정 규칙의 적용 확인 — 개선 과제 R1~R5 완수 시 종합 3.15점·ML3 진입 전망([표 4-16], [그림 4-3])은 v98에 없던 내용'),
                       ('그림·표','그림 12 → 34장, 표 31 → 36개','추가 그림: 1-2~1-4(논문의 구성), 2-6·2-7, 3-5~3-10(모델·판정 규칙·BARS 카드), 4-3~4-10(영역 평점, 판정 경로, 개선 과제 카드 R1~R5), 5-1~5-3(결론 요약). 추가 표: 3-8~3-11, 4-17·4-18, 5-1(장점·단점). 삭제 표: 2-6·2-7(그림으로 대체)'),
                       ('용어·표기 통일','표기 규칙표 v18(금지 72건)','「환경경영시스템(ISO 14001)」·「전과정평가(LCA)」 고정형, 「재는·잇는다·갈린다」 제거, 「통합」→「연계」, 「Clause」→「조항」, 「3층」→「3계층」, 「BARS 판정척도」→「행동기술척도」, 절 제목의 영문 병기 제거(연구 문제·연구 방법), 「연계 성숙도 평가 모델(LMA)」 명칭 정착'),
                       ('초록·결론','9/16 지시대로 재작성','국문 초록 [1]~[9] 문단 순서(성과 → 의의 → 한계 → 극복 방안 → 추가 연구), 초록 ↔ 본문 1:1 대응, Abstract는 부록 뒤. 5장 절 구성(5.1 결과 요약 · 5.2 기여 · 5.3 한계와 향후 과제)은 그대로'),
                       ('부록·참고문헌','부록 I·J 제목, 참고문헌 45 → 61건','부록 I 「평가 기록부 — 문항별 판정 요약」, 부록 J 「Building Chart — 55문항 점수 구조」. 참고문헌에 ISO/IEC 33003·33020, KS X ISO 8000-61·62, Polit 2007(S-CVI), ISO 14006·14025, 교수님 2017년 논문 인용 정정 등'),
                       ('검토 라운드','v199~v203','9/30 1차·2차 검토(v201·v201d), 10/5~10/6 전수 대조(v202: 그림 35장·표 119개·수치·인용), 10/7 쪽 여백 조정(v203, 글자 무변경). 표절검사 v198 1%(카피킬러, 2026-09-29)'),
                       ('바뀌지 않은 것','—','55문항 점수, 영역 평점(P 3.23·D 3.00·C 2.10·A 2.33), 종합 2.85·ML2, 미비 사항 3건·개선 과제 5건, 연계 프로세스 16종·핵심 문항 4종, 장 구성(5장 + 부록 A~K)')],header=('구분','규모','내용')))
NB.append(para(H1,'5. 여쭐 것 세 가지와 송부·제본 계획'))
NB.append(para(P_LIST,'① 예비심사 발표 시간 — 본 슬라이드 30장, 노트 대본 기준 약 30분 분량으로 준비하였습니다. 20분이면 1장(개요·용어)과 백업으로 돌릴 장을 정해 줄이겠습니다.'))
NB.append(para(P_LIST,'② 본심사에서 보고할 범위 — 예비심사 지적 사항의 수정 결과만 보고하는지, 전체를 다시 발표하는지.'))
NB.append(para(P_LIST,'③ 생성형 AI 활용의 표기 — 문장 교정·대조 보조·형식 점검에 썼고, 대학원 지침(2026.2.2)·심사 안내·학위수여규정·연구윤리준수확인서(서식 6)에 관련 조항이 없음을 확인하여(9/29) 논문에 별도 표기는 두지 않았습니다. 이대로 두어도 괜찮으신지. 범위는 발표자료 백업 75번(B-Q13)에 정리하였고 물으면 그대로 답하겠습니다.'))
NB.append(para(P_LIST,'참고 — 발표자료 v22(75장 = 본 30 + 백업 45, 노트 대본 포함)와 예상 질문 v7(48문)은 별도 파일입니다. 본 슬라이드는 쉬운 말로, 세부 표와 조항 번호는 백업으로 옮겼습니다. 지난 면담자료의 「그 밖에 여쭐 것」(경량화 5건·장절 번호·문구 8건)은 v203 제출로 일단락되어 본심사 수정 때 다시 여쭙겠습니다.'))
NB.append(copy.deepcopy(K[-1]))
for c in list(body): body.remove(c)
for e in NB: body.append(e)
tree.write(SRC+'/word/document.xml',xml_declaration=True,encoding='UTF-8',standalone=True)
fp=SRC+'/word/footer1.xml'; f=open(fp,encoding='utf8').read().replace('10/7 면담 준비자료 Ver4.1 · ','면담 준비자료 Ver5.1 (2026-10-12) · '); open(fp,'w',encoding='utf8').write(f)
if os.path.exists(OUT): os.remove(OUT)
z=zipfile.ZipFile(OUT,'w',zipfile.ZIP_DEFLATED)
for dp,dn,fn in os.walk(SRC):
    for f in fn:
        full=os.path.join(dp,f); z.write(full,os.path.relpath(full,SRC))
z.close(); print('written',OUT,os.path.getsize(OUT))
