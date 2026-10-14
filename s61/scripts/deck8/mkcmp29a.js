const fs=require('fs');
const {Document,Packer,Paragraph,TextRun,Table,TableRow,TableCell,WidthType,ShadingType,ImageRun,HeadingLevel,PageOrientation,BorderStyle}=require('docx');
const F='맑은 고딕';
const t=(s,o={})=>new TextRun({text:String(s),font:F,size:o.size||17,bold:o.bold,color:o.color});
const P=(s,o={})=>new Paragraph({children:Array.isArray(s)?s:[t(s,o)],spacing:{after:o.after??60}});
const H=(s,l)=>new Paragraph({heading:l,children:[new TextRun({text:s,font:F,bold:true,size:l===HeadingLevel.HEADING_1?28:22})],spacing:{before:200,after:120}});
const bd={style:BorderStyle.SINGLE,size:4,color:'A6B0C3'};const borders={top:bd,bottom:bd,left:bd,right:bd};
const cell=(s,w,o={})=>new TableCell({width:{size:w,type:WidthType.DXA},borders,shading:o.fill?{type:ShadingType.CLEAR,fill:o.fill,color:'auto'}:undefined,margins:{top:40,bottom:40,left:70,right:70},children:String(s).split('\n').map(x=>P(x,{size:o.size||16,bold:o.bold,color:o.color,after:30}))});
const log=JSON.parse(fs.readFileSync('log29a_list.json','utf8'));
const W=[600,1300,5000,5600,2898];
const rows=[new TableRow({tableHeader:true,children:['#','슬라이드(옛 번호)','항목','v29 → v29a','이유'].map((h,i)=>cell(h,W[i],{fill:'1F3864',bold:true,color:'FFFFFF'}))}),...log.map((r,i)=>new TableRow({cantSplit:true,children:[cell(i+1,W[0]),cell(r[0]+' '+r[1],W[1]),cell(r[2],W[2],{color:'7F7F7F',size:15}),cell(r[3],W[3]),cell(r[4],W[4],{size:15})]}))];
const imgs=fs.readdirSync('cmp29a').filter(f=>f.endsWith('.png')).sort(); const imgParas=[];
for(const f of imgs){const buf=fs.readFileSync('cmp29a/'+f);const [w0,h0]=require('child_process').execSync(`python3 -I -c "from PIL import Image;im=Image.open('cmp29a/${f}');print(im.width,im.height)"`).toString().trim().split(' ').map(Number);let w=940,h=Math.round(h0*w/w0);const n=parseInt(f.slice(1,3));imgParas.push(P([t(`v29 슬라이드 ${n} — 왼쪽 v29, 오른쪽 v29a`,{bold:true,size:18})],{after:40}));imgParas.push(new Paragraph({children:[new ImageRun({type:'png',data:buf,transformation:{width:w,height:h}})],spacing:{after:160}}));}
const rules=['검토 방식 — 설계를 모르는 검토자에게 v29 발표자료(본 19장 화면·노트 전수)·노트북 참고와 물으면(cheat29)·논문 덤프(v203)·규칙표 v18만 주고 ① 숫자·용어·그림 번호의 논문 일치 ② 규칙 문맥 ③ 심사위원 다섯 분 관점에서 빠진 질문을 보게 함. 숫자·셈(30조항, 28/25/3, 44/11, 16종 8·2·3·3, 55문항 30·6·10·9, 핵심 4종 배치, ML 조건, 2.89→2.85, 36/51, 영역 평점, R1~R5 효과와 3.15, 24건 22·1·1)은 모두 일치.','고친 것(논문 일치) — 16번 단점은 7건(장점 6건·단점 7건), 넷은 보완·셋은 설계 제약(5.2) · 노트북 16 예상질문 37 답을 5.2의 세 가지(측정 대상·표준 정합성·선행연구와의 관계)로 · 17번 한계 여덟 항목 중 ⑦ 미측정 구간 보강, 대응 경로 정정(다사례·다평가자·재평가 / CVI는 타당화) · 7번 측정 틀 ISO/IEC 33020·점수별 상태 기술 BARS, 선례는 논문 표현 「국내 공공데이터 품질 연구(PPDQM3·APDQM3)」 · 9번 「조항 25개(전조항 28개 가운데)」 · 11번 「모두 4점 이상」 · 12·15번 「구조적 미비 세 가지와 공통 원인 두 가지에서 과제 다섯 건」 · 12번 「문서 심사 — 면담·실행 상태의 현장 관찰은 포함하지 않음」 · 노트북 8 참고에 A-8 주석.','고친 것(표현) — 6번 논문 고정 문구 「그 결과를 되풀이하여 관리하는 주기는 정하지 않았다」(화면·노트) · 10번 「추가 항목(should 또는 체계화·지속 상태)」 · 5번 「중대한 환경측면을 결정하고」 · 11번 「연계가 성립하지 않은 것과 성숙도가 낮은 것」 · 14번 「LCA 연 1회 재수행 규정은 있으나 아직 이행 전」. 두지 않은 것 — 15번 그림 안 「외부 정밀검토」(논문 그림 그대로) · 2번 목차 로마자 Ⅴ(구조 변경이라 보류).','노트북 예상 질문 54~63 추가(검토자 제안, 이사님 54번 답 방향 반영) — 경영검토 9.3 점검인가 개선인가 · 역방향 매핑 기준 · 형성형 구성개념과 CVI · ML2 과반 조건의 P 편중 · 기계학습 축약의 표본 수 · 평가자 전환교육 · 갱신심사 부적합 0인데 왜 ML2 · D-2의 필수 체크포인트 · ML4 「예측화」 명칭 · 34행→30 셈. 노트 합계 15분 55초, 규칙 0건.'];
const checks=['61번(D-2에 필수 체크포인트가 있는 이유)은 검토자가 약점으로 짚은 자리 — 부록 C D-2의 필수 두 항목을 한 번 읽어 두실 것.','7번 노트에서 PPDQM3·APDQM3는 입으로는 「국내 공공데이터 품질 연구」까지만 말해도 됨.'];
const doc=new Document({styles:{default:{document:{run:{font:F,size:18}}}},sections:[{properties:{page:{size:{width:11906,height:16838,orientation:PageOrientation.LANDSCAPE},margin:{top:720,bottom:720,left:720,right:720}}},children:[
 new Paragraph({children:[new TextRun({text:'대조표 — 발표 자료 v29 → v29a (검토자 지적 반영)',font:F,bold:true,size:36})],spacing:{after:80}}),
 P('발표자료_20261007_v29_15분판.pptx → v29a_15분판.pptx (81장 그대로). 2026-10-14 세션 61 — 독립 검토자 1명의 논문 대조 결과 반영.',{size:18,after:120}),
 H('1. 원칙',HeadingLevel.HEADING_1),...rules.map(s=>P('· '+s,{size:17})),
 H('2. 바뀐 곳',HeadingLevel.HEADING_1),new Table({width:{size:15398,type:WidthType.DXA},columnWidths:W,rows}),
 
 H('3. 바뀐 장 비교 그림',HeadingLevel.HEADING_1),...imgParas,
 H('4. 이사님 확인 사항',HeadingLevel.HEADING_1),...checks.map(s=>P('· '+s,{size:17}))]}]});
Packer.toBuffer(doc).then(b=>{fs.writeFileSync('대조표_발표자료_v29_v29a_20261014.docx',b);console.log('ok',log.length,imgs.length)});
