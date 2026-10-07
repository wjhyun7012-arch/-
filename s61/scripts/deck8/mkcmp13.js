const fs=require('fs');
const {Document,Packer,Paragraph,TextRun,Table,TableRow,TableCell,WidthType,ShadingType,ImageRun,HeadingLevel,PageOrientation,BorderStyle}=require('docx');
const F='맑은 고딕';
const t=(s,o={})=>new TextRun({text:String(s),font:F,size:o.size||17,bold:o.bold,color:o.color});
const P=(s,o={})=>new Paragraph({children:Array.isArray(s)?s:[t(s,o)],spacing:{after:o.after??60}});
const H=(s,l)=>new Paragraph({heading:l,children:[new TextRun({text:s,font:F,bold:true,size:l===HeadingLevel.HEADING_1?28:22})],spacing:{before:200,after:120}});
const bd={style:BorderStyle.SINGLE,size:4,color:'A6B0C3'};const borders={top:bd,bottom:bd,left:bd,right:bd};
const cell=(s,w,o={})=>new TableCell({width:{size:w,type:WidthType.DXA},borders,shading:o.fill?{type:ShadingType.CLEAR,fill:o.fill,color:'auto'}:undefined,margins:{top:40,bottom:40,left:70,right:70},children:String(s).split('\n').map(x=>P(x,{size:o.size||16,bold:o.bold,color:o.color,after:30}))});
const log=JSON.parse(fs.readFileSync('log13.json','utf8'));
const W=[600,1300,5000,5600,2898];
const rows=[new TableRow({tableHeader:true,children:['#','슬라이드','항목','v12 → v13','이유'].map((h,i)=>cell(h,W[i],{fill:'1F3864',bold:true,color:'FFFFFF'}))}),...log.map((r,i)=>new TableRow({cantSplit:true,children:[cell(i+1,W[0]),cell(r[0]+' '+r[1],W[1]),cell(r[2],W[2],{color:'7F7F7F',size:15}),cell(r[3],W[3]),cell(r[4],W[4],{size:15})]}))];
const imgs=fs.readdirSync('cmp13').filter(f=>f.endsWith('.png')).sort(); const imgParas=[];
for(const f of imgs){const buf=fs.readFileSync('cmp13/'+f);const [w0,h0]=require('child_process').execSync(`python3 -I -c "from PIL import Image;im=Image.open('cmp13/${f}');print(im.width,im.height)"`).toString().trim().split(' ').map(Number);let w=Math.min(940,w0),h=Math.round(h0*w/w0);const n=parseInt(f.slice(1,3));imgParas.push(P([t(n>29?`슬라이드 ${n} — 새 백업(옮긴 표)`:`슬라이드 ${n} — 왼쪽 v12, 오른쪽 v13`,{bold:true,size:18})],{after:40}));imgParas.push(new Paragraph({children:[new ImageRun({type:'png',data:buf,transformation:{width:w,height:h}})],spacing:{after:160}}));}
const rules=['이사님 10/7 권고안: 본 슬라이드만 「쉽게, 그러나 들어갈 것은 다 들어가게」. 세부 표는 백업으로 옮김(삭제 아님).','4번 문구는 논문 표현으로(「서로를 전제하지 않는다」 5장, 「별도로 운영」 1.1, 「조항 단위로 확인하고 성숙도 수준으로 판정」 5장). 「따로 돈다」「얼마나 이어졌는지를 측정」은 뺌.','12번 = 모델링 한 그림(도형): 두 표준 ↔ 매핑 → ① 참조 모델(업무 목록 16종) → ② 평가 모델(채점표 55문항) → ③ 성숙도 모델(수준 규칙) → 판정 / 검증 두 방향. [그림 3-2] 전체는 백업 B-7.','14번 = 그림 + 핵심 연계 지점 네 곳 설명([표 3-5]). 16번 = 그림 크게 + 「점수를 매기는 법 세 단계」 + D-1 예(BARS 앵커 표 → 백업 B-10a). 17번 = 여섯 수준 계단 + A사 멈춘 곳 + 핵심 문항 4종([표 3-7] → 백업 B-11a). 18번 = 카드 셋(참조·평가·성숙도) + 검증 상자. 19번 = 평가 한 그림(문서 받기 → 채점 → 다시 확인 → 판정 → 개선, 숫자 하나씩) + 대상·방법·기간·지위. 20번 = 상자 둘 세 줄씩.','화면 글은 논문 용어를 앞에 두고 쉬운 말은 괄호·노트·용어 줄로. 규칙표 v18 점검 0건(「ML3이」, PDCA 「계획–실행–점검–조치」, 「재검증이 점수가 후해지는 것을 막는 방향으로 작동」(T41)).','백업 두 장 추가로 70 → 72장. 31번 이후 쪽번호 32곳 재부여. 30번 백업 구분 문구에 B-10a·B-11a 추가.','노트 8장(4·12·14·16·17·18·19·20) 새 화면에 맞춰 다시 씀(존댓말, 논문 용어, 〔메모〕).'];
const checks=['12·19번 그림의 글자 양 — 더 깎을지.','17번 계단의 쉬운 말 조건이 [표 3-7]과 어긋나지 않는지(ML4·ML5는 요약).','18번 카드에서 뺀 「따른 표준」 열은 10번과 백업 B-2에 있음 — 괜찮은지.','백업 40장 ▶ 한 줄(2차)은 이 판 보신 뒤 결정.'];
const doc=new Document({styles:{default:{document:{run:{font:F,size:18}}}},sections:[{properties:{page:{size:{width:11906,height:16838,orientation:PageOrientation.LANDSCAPE},margin:{top:720,bottom:720,left:720,right:720}}},children:[
 new Paragraph({children:[new TextRun({text:'대조표 — 발표 자료 v12 → v13 (본 슬라이드 8장 쉽게 · 백업 2장 추가)',font:F,bold:true,size:36})],spacing:{after:80}}),
 P('발표자료_20261007_v12.pptx(70장) → v13.pptx(72장, A4). 2026-10-07 세션 61.',{size:18,after:120}),
 H('1. 원칙',HeadingLevel.HEADING_1),...rules.map(s=>P('· '+s,{size:17})),
 H('2. 바뀐 곳',HeadingLevel.HEADING_1),new Table({width:{size:15398,type:WidthType.DXA},columnWidths:W,rows}),
 H('3. 화면 전·후',HeadingLevel.HEADING_1),...imgParas,
 H('4. 이사님 확인 사항',HeadingLevel.HEADING_1),...checks.map(s=>P('· '+s,{size:17}))]}]});
Packer.toBuffer(doc).then(b=>{fs.writeFileSync('대조표_발표자료_v12_v13_20261007.docx',b);console.log('ok',log.length,imgs.length)});
