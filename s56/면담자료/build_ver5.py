# 면담자료 Ver5(간결판) — Ver4.2의 패키지·서식을 그대로 쓰고 body만 다시 짬: 9/16 지시 → 수정 결과 + 심사 일정 (+ 출력본 뒤 변경 참고, 그 밖에 여쭐 것)
import copy, re, zipfile, shutil, os
from lxml import etree
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'; XML='http://www.w3.org/XML/1998/namespace'
q=lambda t:'{%s}%s'%(W,t)
SRC='/home/claude/mt/x'; OUT='/home/claude/mt/면담자료_20261007_Ver5.docx'
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
def table(tmpl,rows,keep_header=True):
    t=copy.deepcopy(tmpl); trs=t.findall(q('tr'))
    hdr=trs[0]; data_tmpl=trs[1]
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
    if s=='확인 (5.2 나-1)': set_text(tc.findall(q('p'))[0],'확인 (3절에서 여쭘)')
    elif s=='반영 (문안 확인 4.1)': set_text(tc.findall(q('p'))[0],'반영 (문안은 Ver4.2 4.1)')
    elif s=='반영 (문안 확인 4.2)': set_text(tc.findall(q('p'))[0],'반영 (문안은 Ver4.2 4.2)')
    elif s=='v199 반영 내용': set_text(tc.findall(q('p'))[0],'반영 내용 (v200 출력본)')
    elif s.startswith('2절. 심사위원 후보와 날짜는 5.2 나-1'): set_text(tc.findall(q('p'))[0],'심사 신청·서류 완료. 심사위원 5인 확정, 날짜는 3절에서 확정')
    elif '6절' in s or '5.2' in s or '5.1' in s or '4.1' in s or '4.2' in s:
        ps=tc.findall(q('p'))
        if len(ps)==1:
            t2=s.replace('예상 질문 43문은 6절','예상 질문 43문은 별도 자료 Ver4.2 6절').replace('(5.2 나-1)','(3절)').replace('5.2 나-','Ver4.2 5.2 나-').replace('문안 확인 4.1','문안은 Ver4.2 4.1').replace('문안 확인 4.2','문안은 Ver4.2 4.2').replace('v199에서 고침(5.1)','v199에서 고침(Ver4.2 5.1)')
            if t2!=s: set_text(ps[0],t2)
# 새 body
NB=[]
NB.append(copy.deepcopy(K[0]))
for i in (1,2,3,4,5): NB.append(copy.deepcopy(K[i]))
NB.append(para(P_META,'대상 판: 출력본 phd_v200_20260930.docx(10/7 지참 · LibreOffice 기준 200쪽 · 참고문헌 58건) · 현행 phd_v202_20261005.docx(9/30·10/5·10/6 검토 반영 — 내용·수치·판정 동일, 참고문헌 61건)'))
NB.append(para(P_META,'작성일: 2026-10-06 · 면담자료 Ver5 (간결판 — 9/16 지시 사항의 수정 결과와 심사 일정. 국문 초록 전문·3.6 결론·논문 문구 8건·예상 질문 43문·경량화 메모는 별도 자료 Ver4.2)'))
NB.append(table(T_TOC,[('1','한눈에 보기'),('2','9/16 지시 사항과 수정 결과 (A-01~A-39)'),('3','심사 일정과 서류 — 날짜 확정'),('4','출력본(v200) 뒤에 고친 것 — 참고'),('5','그 밖에 여쭐 것 (시간이 되면)')]))
NB.append(para(H1,'1. 한눈에 보기'))
NB.append(para(P_LIST,'9/16 면담에서 주신 지적을 녹취(37분 51초)와 빨간 펜 사진 14장으로 정리하여 39건(A-01~A-39)으로 나누었고, 논문에 관한 33건(중복 3건 제외)은 국문 초록부터 부록·영문 초록까지 모두 반영하였습니다(2절). 논문 외 3건은 심사 신청·학과 서류 제출 완료, 심사위원·일정은 3절에서 여쭙고, 발표 자료는 10/7 말씀 뒤 확정합니다.'))
NB.append(para(P_LIST,'지참한 출력본은 v200(9/30)입니다. 그 뒤 세 차례 검토(9/30·10/5·10/6)로 표기·정합 500여 곳과 참고문헌을 고쳤으나 내용·수치·판정(A사 2.85점·ML2, 개선 과제 완수 시 3.15점·ML3)과 장·절 구성은 같습니다(4절).'))
NB.append(para(P_LIST,'심사위원 5인은 정해졌으므로 오늘 정할 것은 예비심사·본심사 날짜와, 10/8 제출 서식의 심사 일시·장소 칸입니다(3절).'))
NB.append(table(T_SUM,[('9/16 지적 — 논문','33건','전부 반영. A-37~A-39는 A-18·A-20·A-23과 같은 지적이라 중복으로 셈'),
                       ('9/16 지적 — 논문 외','3건','심사 신청·학과 서류 제출 완료(A-32) · 심사위원 5인 확정, 날짜는 3절에서 확정(A-31) · 발표 자료와 예상 질문은 10/7 뒤 확정(A-33)'),
                       ('출력본 뒤 변경','표기·정합 500여 곳','4절 — 내용·수치 동일. 교수님 2017년 논문 인용 1곳 정정, 참고문헌 3건 추가 포함'),
                       ('정할 것','날짜 확정 등 3건 + 선택 3건','3절·5절')]))
NB.append(para(H1,'2. 9/16 지시 사항과 수정 결과 (A-01~A-39)'))
NB.append(para(P_INTRO,'「9/16 지적」은 녹취와 사진에서 뽑은 요지이고, 「위치」는 출력본 v200의 절 번호입니다(v202도 같음). 초록의 [1]~[9]는 국문 초록의 문단 번호입니다. 상태 칸의 「Ver4.2 4.1·4.2」는 별도 자료의 절 번호입니다.'))
NB.append(TA)
NB.append(para(H1,'3. 심사 일정과 서류 — 날짜 확정'))
NB.append(copy.deepcopy(T_SCHED))
NB.append(copy.deepcopy(K[20]))
NB.append(para(P_LIST,'① 심사위원 5인은 정해졌습니다 — 김계수·김동욱·이진원·서병석·이창수 교수. 날짜만 확정하면 됩니다. 10/1 일정안: 박사 예비심사 10/22(목) 18시(김계수 교수 시간 조정 요청), 본심사 1차 11/12(목) 18시, 본심사 2차 11/26(목) 18시 — 11/12·11/26은 석사 심사를 먼저 하고 박사를 이어서. 서병석 교수의 11월 일정은 10/1 기준 확인 중이었습니다. 대학원 안내(원주)의 틀(위촉 ~11/10, 심사 11/11~12/9, 박사 3회 이상)에 맞습니다.'))
NB.append(para(P_LIST,'② 10/8까지 학과사무실에 낼 서식 4(심사위원 추천서)·4-1(외부위원 개인정보 동의서)·5(심사 계획서) — 심사위원 명단은 위 5인으로 적고, 심사 일시·장소는 오늘 확정되는 대로 적어 제출하겠습니다.'))
NB.append(para(P_LIST,'③ 심사용 논문(A4 소프트커버 5부, 심사위원에게 직접 전달)의 인쇄 시점 — 오늘 말씀 주시는 내용을 반영한 판으로, 예비심사(10/22안) 전에 전달되도록 인쇄하겠습니다.'))
NB.append(para(H1,'4. 출력본(v200) 뒤에 고친 것 — 참고'))
NB.append(para(P_LIST,'출력본을 그대로 읽으셔도 되는 범위의 변경입니다. 전체 목록은 대조표(v200→v201, v201→v201d, v201d→v202)에 있습니다.'))
NB.append(table(T_SUM,[('9/30 1차 검토 반영(v201)','183건·그림 11장','조사·참조·용어 정합, [표 3-8] 주석, 부록 주석 보강'),
                       ('9/30 2차 검토 반영(v201d)','252건','표준 원문과 달랐던 서술 4건 정정(KS I ISO 14044 가중치 합산·6.2/6.3, ISO/IEC 33020 판 표기 2건) · 교수님 2017년 논문 인용 정정([표 2-4]·2.5.4 「3개 공공기관 시범 적용」은 2015년 모형의 것) · 4장 사실 정리(D-1 「절차서가 없음」, C-7·A-9 처리를 평가 뒤 3.4.3 규칙으로 확정했다고 명시) · 참고문헌 3건 추가(Polit 2007 — S-CVI/Ave 0.90의 출처, ISO 14006, ISO 14025)'),
                       ('10/5~10/6 전체 검토 반영(v202)','41건·그림 5곳','그림 35장·표 119개·수치·인용 전수 대조 — 그림 안 잔존 문구, [표 1-1] 절 번호, 조사 1건, 주석 보완, 4.2.3 간접 확인 5건(P-24 포함, 31/51) · 3.3.4의 ISO/IEC TS 33030 인용 문장을 원문(4.1·4.2.2.3·4.2.2.4) 확인 뒤 다시 씀'),
                       ('바뀌지 않은 것','—','55문항 점수, 영역 평점(P 3.23·D 3.00·C 2.10·A 2.33), 종합 2.85·ML2, 완수 시 3.15·ML3, 장·절 구성, 국문 초록·3.6 결론의 문안')]))
NB.append(para(H1,'5. 그 밖에 여쭐 것 (시간이 되면)'))
NB.append(para(P_LIST,'본문 경량화 — 3.2.3 16종 기술 블록, 4.3의 55문항 결과 표 4개, [표 4-4], 3.5.2 항목별 서술, 3.4.4 대안 서술을 부록으로 옮길지(5건, 별도 자료 Ver4.2 별첨). 승인해 주시는 범위대로 다음 판에서 정리합니다.'))
NB.append(para(P_LIST,'장·절 번호 체계는 현행 「제1장 / 1.1 / 1.1.1」 유지로 보았습니다(대학원 별지 4의 「Ⅰ. / 1. / 1)」은 예시로 봄). 「ABSTRACT/Abstract」 표제도 현행대로입니다.'))
NB.append(para(P_LIST,'논문 문구·형식 8건(「LCA 없이 shall 이행 불완전」 표현, 사용 단계 제외 사유, 8000-6x 현행 구성, R4 내부 검토자 배치, 부록 캡션 형식, 참조 모델 정당화 문단 등)과 예상 질문 43문은 별도 자료 Ver4.2의 5.2·6절에 있습니다.'))
NB.append(copy.deepcopy(K[-1]))
for c in list(body): body.remove(c)
for e in NB: body.append(e)
tree.write(SRC+'/word/document.xml',xml_declaration=True,encoding='UTF-8',standalone=True)
if os.path.exists(OUT): os.remove(OUT)
z=zipfile.ZipFile(OUT,'w',zipfile.ZIP_DEFLATED)
for dp,dn,fn in os.walk(SRC):
    for f in fn:
        full=os.path.join(dp,f); z.write(full,os.path.relpath(full,SRC))
z.close(); print('written',OUT,os.path.getsize(OUT))
