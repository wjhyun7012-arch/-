// 심사위원용 한 쪽 — 연구 요지와 논문 안내
const fs=require('fs');
const {Document,Packer,Paragraph,TextRun,Table,TableRow,TableCell,WidthType,BorderStyle,ShadingType,AlignmentType}=require('docx');
const F='맑은 고딕';
const t=(s,o={})=>new TextRun({text:String(s),font:F,size:o.size||17,bold:o.bold,color:o.color});
const P=(s,o={})=>new Paragraph({children:Array.isArray(s)?s:[t(s,o)],spacing:{after:o.after??40,line:o.line||240,before:o.before||0},alignment:o.align});
const bd={style:BorderStyle.SINGLE,size:4,color:'A6B0C3'};const borders={top:bd,bottom:bd,left:bd,right:bd};
const cell=(children,w,fill)=>new TableCell({borders,width:{size:w,type:WidthType.DXA},shading:fill?{type:ShadingType.CLEAR,fill,color:'auto'}:undefined,margins:{top:40,bottom:40,left:80,right:80},children});
const row=(k,v,loc)=>new TableRow({children:[cell([P(k,{size:17,bold:true})],1500,'F2F4F8'),cell(v.map(x=>P(x,{size:16,after:20})),7000),cell([P(loc,{size:14,color:'595959',after:20})],2100)]});
const kids=[];
kids.push(P('박사학위 논문 예비심사 — 연구 요지와 논문 안내',{size:26,bold:true,color:'1F3864',after:60}));
kids.push(P([t('논문 제목  ',{size:16,bold:true}),t('환경경영시스템(ISO 14001)과 전과정평가(LCA)의 프로세스 연계 모델 개발 및 성숙도 평가에 대한 연구',{size:16})],{after:10}));
kids.push(P([t('영문 제목  ',{size:16,bold:true}),t('A Study on the Development and Maturity Assessment of a Process Linkage Model between Environmental Management Systems (ISO 14001) and Life Cycle Assessment (LCA)',{size:15})],{after:10}));
kids.push(P([t('발표자  ',{size:16,bold:true}),t('현욱진 (강원대학교대학원 산업경영공학과 박사과정)   ',{size:16}),t('지도교수  ',{size:16,bold:true}),t('이창수 교수   ',{size:16}),t('일시  ',{size:16,bold:true}),t('2026년 10월 22일(목) 18:00   ',{size:16}),t('논문  ',{size:16,bold:true}),t('v203(2026.10.7), 5장 + 부록 A~K',{size:16})],{after:120}));
const rows=[new TableRow({tableHeader:true,children:['항목','요지','논문 위치'].map((h,i)=>new TableCell({borders,width:{size:[1500,7000,2100][i],type:WidthType.DXA},shading:{type:ShadingType.CLEAR,fill:'1F3864',color:'auto'},children:[P(h,{size:16,bold:true,color:'FFFFFF',after:0})]}))}),
 row('연구 배경',['제품 전과정의 환경정보를 요구하는 규제(CBAM·ESG 공시)가 늘어 환경경영시스템(ISO 14001)과 전과정평가(LCA)를 함께 운영해야 하나, 실무에서는 두 표준이 별도로 운영된다. 두 표준을 연계하는 프로세스와 그 연계 수준을 측정하는 기준은 찾아보기 어렵다.'],'1.1, 2.3'),
 row('연구 목적',['① 두 표준의 조항을 양방향으로 대조하여 연계 지점을 규명 ② 참조 모델·평가 모델·성숙도 모델의 세 계층으로 연계 성숙도 평가 모델(LMA)을 개발 ③ 제조기업 A사에 적용하여 판정 규칙의 작동을 확인하고 영역별 성숙도를 평가 ④ 결과를 개선 과제와 로드맵으로 연결'],'1.2.1'),
 row('연구 문제',['RQ1 두 표준의 조항은 서로 어떻게 매핑되는가 · RQ2 연계 프로세스와 평가지표는 어떻게 구성되는가 · RQ3 성숙도는 어떤 규칙으로 판정되는가 · RQ4 판정 결과가 실제 조직에서 어떤 의미를 갖는가'],'1.2.2, [그림 1-2]'),
 row('연구 방법',['문헌 연구 → 비교분석(ISO 14001 4~10절 ↔ ISO 14044 4~6절 전조항 매핑, 네 가지 판정 질문·네 등급) → 모델 개발(세 계층, 반복 개발) → 사례 연구(A사 2025년 제품 LCA와 ISO 14001 기록의 문서 심사, 평가자 1인, 2026년 6~7월)'],'1.3.1, 3.1, 4.1~4.2'),
 row('개발한 모델',['참조 모델: 연계 프로세스 16종(계획 8·실행 2·점검 3·개선 3, KS X ISO 8000-61 5.3 형식) · 평가 모델: 55문항(P 30·D 6·C 10·A 9)과 문항별 5점 채점 기준(체크포인트·BARS 앵커) · 성숙도 모델: ML0~ML5 여섯 수준의 판정 규칙, 핵심 문항 4종(P-9·P-14·D-1·C-1)에만 최소충족 판정(KS X ISO 8000-62 4.5.1 준용) — 하나라도 3점 미만이면 평균과 무관하게 ML3 이상 불가'],'3.2~3.4, [표 3-4]·[표 3-5]·[표 3-7], 부록 C'),
 row('검증',['추적성 검증(55문항 → 두 표준의 근거 조항) · 표준 적합성 검증(ISO/IEC 33004 24개 요구사항 자체점검: 충족 22·해당 없음 1·보완 1) · 전문가 내용타당도(CVI) 조사는 조사지 설계까지'],'3.5, [표 3-9], 부록 E·F'),
 row('적용 결과',['A사 증빙 문서 16종(33건) 문서 심사 → 잠정 2.89, 재검증(상승 7·하락 9) 후 확정 종합점수 2.85 · 영역 평점 P 3.23 / D 3.00 / C 2.10 / A 2.33 · 핵심 문항 D-1(설계·개발 단계의 LCA 결과 반영) 2점 미충족 → ML2(관리화) · 구조적 미비 3건(설계·개발 절차의 환경 고려 단계 부재, 정밀검토 체계 부재, LCA 재수행 사이클 미도래) → 개선 과제 R1~R5(+장기 R6), 완수 시 3.15·ML3 전망'],'4.3~4.6, [표 4-13]·[표 4-15]·[표 4-16], [그림 4-2]'),
 row('독창성·기여',['① 측정 대상 — 단일 분야 안의 프로세스가 아니라 두 표준 사이의 연계를 평가 대상으로 삼은 성숙도 모델 ② 판정 구조 — 연계 성립의 필요조건(핵심 문항)에만 비보상 최소충족 판정을 적용하여 종합점수에 묻히는 미비를 드러냄 ③ 평가–개선–재평가 순환 — 개선 과제의 목표 상태를 채점 기준과 같은 문구로 써서 차기 평가에서 그대로 판정 · 실무: 55문항 채점 기준·증거 목록·기록부·보고서 양식의 실물 도구'],'5.2, [그림 5-2]'),
 row('한계·향후 연구 과제',['단일 사례·단일 시점·단일 평가자 · 문항 간 종속 · 상위 구간(ML4·ML5) 사례 미검증 · ISO 14001:2015 기준(개정판은 향후 연구 과제) · 전문가 조사(CVI)와 다사례 적용으로 타당화 확장 · 진단·개선 고도화 · 미측정 구간 · 표준화 경로 검토'],'5.3, [그림 5-3]')];
kids.push(new Table({width:{size:10600,type:WidthType.DXA},columnWidths:[1500,7000,2100],rows}));
kids.push(P('자주 찾는 표·그림 — 연계 가능성 판정 기준 [표 3-2](3.2.2) · 연계 지점 맵 [그림 3-4](3.2.2) · 연계 프로세스 16종 [표 3-4](3.2.3) · 핵심 문항 4종의 채점 기준 [표 3-6](3.3.2) · 성숙도 수준 판정 규칙 [표 3-7](3.4.3) · 적용 개요 [표 4-1](4.1) · 영역별 평가 결과 및 판정 [표 4-13](4.3.5) · 개선 과제 [표 4-15](4.6.1) · 완수 시 기대 효과 [표 4-16](4.6.2) · 장점과 단점 [표 5-1](5.2) · 전 조항 분석 원자료 부록 A · 55문항 채점 기준 부록 C · 평가 기록부 부록 I · 용어집 부록 K',{size:14,color:'595959',before:100,line:230}));
kids.push(P('발표 구성(약 15~20분, 논문 장 순서) — 논문 구성 → 1장 개요·두 표준(ISO 14001, LCA)이 어떤 표준인가·연계의 뜻 → 2장 두 표준에서 비어 있는 곳·따른 표준 → 3장 조항 매핑·연계 지점·LMA 구성·16종·55문항·판정 규칙 → 4장 A사 적용·영역별 평점과 핵심 문항의 판정·개선 과제 → 5장 장점과 단점·한계 → 결론(네 가지 연구 문제의 답). 세부 그림·표와 예상 질문의 근거는 백업 슬라이드에 두었습니다.',{size:14,color:'595959',line:230}));
const doc=new Document({styles:{default:{document:{run:{font:F,size:17}}}},sections:[{properties:{page:{size:{width:11906,height:16838},margin:{top:560,bottom:500,left:650,right:650}}},children:kids}]});
Packer.toBuffer(doc).then(b=>{fs.writeFileSync('심사위원용_연구요지와논문안내_20261022.docx',b);console.log('ok')});
