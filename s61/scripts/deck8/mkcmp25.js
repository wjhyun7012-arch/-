const fs=require('fs');
const {Document,Packer,Paragraph,TextRun,Table,TableRow,TableCell,WidthType,ShadingType,ImageRun,HeadingLevel,PageOrientation,BorderStyle}=require('docx');
const F='맑은 고딕';
const t=(s,o={})=>new TextRun({text:String(s),font:F,size:o.size||17,bold:o.bold,color:o.color});
const P=(s,o={})=>new Paragraph({children:Array.isArray(s)?s:[t(s,o)],spacing:{after:o.after??60}});
const H=(s,l)=>new Paragraph({heading:l,children:[new TextRun({text:s,font:F,bold:true,size:l===HeadingLevel.HEADING_1?28:22})],spacing:{before:200,after:120}});
const bd={style:BorderStyle.SINGLE,size:4,color:'A6B0C3'};const borders={top:bd,bottom:bd,left:bd,right:bd};
const cell=(s,w,o={})=>new TableCell({width:{size:w,type:WidthType.DXA},borders,shading:o.fill?{type:ShadingType.CLEAR,fill:o.fill,color:'auto'}:undefined,margins:{top:40,bottom:40,left:70,right:70},children:String(s).split('\n').map(x=>P(x,{size:o.size||16,bold:o.bold,color:o.color,after:30}))});
const log=JSON.parse(fs.readFileSync('log25.json','utf8'));
const W=[600,1300,5000,5600,2898];
const rows=[new TableRow({tableHeader:true,children:['#','슬라이드','항목','v24 → v25','이유'].map((h,i)=>cell(h,W[i],{fill:'1F3864',bold:true,color:'FFFFFF'}))}),...log.map((r,i)=>new TableRow({cantSplit:true,children:[cell(i+1,W[0]),cell(r[0]+' '+r[1],W[1]),cell(r[2],W[2],{color:'7F7F7F',size:15}),cell(r[3],W[3]),cell(r[4],W[4],{size:15})]}))];
const imgs=fs.readdirSync('cmp25').filter(f=>f.endsWith('.png')).sort(); const imgParas=[];
for(const f of imgs){const buf=fs.readFileSync('cmp25/'+f);const [w0,h0]=require('child_process').execSync(`python3 -I -c "from PIL import Image;im=Image.open('cmp25/${f}');print(im.width,im.height)"`).toString().trim().split(' ').map(Number);let w=940,h=Math.round(h0*w/w0);const n=parseInt(f.slice(1,3));imgParas.push(P([t(`슬라이드 ${n} — 왼쪽 v24, 오른쪽 v25`,{bold:true,size:18})],{after:40}));imgParas.push(new Paragraph({children:[new ImageRun({type:'png',data:buf,transformation:{width:w,height:h}})],spacing:{after:160}}));}
const rules=['이사님 10/14 장별 검토 13건. ① 3번 노트 — 논문 한 장 요약 1분 30초 판 ② 5번 제목 「두 표준에서 서로 비어 있는 곳」, 5번 노트 첫 문장, 2번 지시문, 58번 ▶ ③ 1·5·8번 「전과정평가(LCA) 표준인 ISO 14044」 ④⑤ 노트북 참고·물으면 상자(별도 파일) ⑥ 13번 제목 「두 표준의 조항 매핑」, 「맞대」 전부 정리(12번 ▶·노트, 18번 「비워 둔」→「정하지 않은」) ⑦ 13번 노트 끝 — 매핑의 결과 한 묶음, 연결 문장 ⑧ 15번 노트 — 16종을 다 읽는 1분 20초 판 ⑨ 채점 표현 — 16번 「채점 방법 — 세 단계」·②「충족되었는지 확인」·③「충족한 체크포인트 개수가 점수」(화면·노트), 10번 「채점의 틀」, 13번 ▶ 「네 등급으로 판정」, 20번 「채점하였습니다」, 용어 줄 ⑩ 18번 「참조 모델, 평가 모델, 성숙도 모델의 세 계층」 ⑪ 20번 「4장은」, 12·20번 「그림으로」 ⑫ 27번 — ▶와 노트를 화면(장점 6·단점 6·한계 3)에 맞춰 다시 씀, 「갈라 둔다」 제거 ⑬ 노트북 예상 질문 전문 장별·번호별(별도 파일).','규칙표 v18 0건, validate 통과, 75장 그대로. 본 슬라이드 노트 합계 30분 55초 → 32분 15초(3번 +30초, 13번 +20초, 15번 +20초, 27번 +10초).'];
const checks=['27번은 노트가 「장점 넷·단점 다섯」이라 화면(여섯·여섯)과 어긋나 있었음 — 이번에 맞춤. 발표 시간이 20분이면 생략 순서를 따로 정해야 함.','발표노트북_v25_연단용(47쪽): 참고 상자를 넣어 12·15·16·17·18·20·27번은 같은 번호의 둘째 쪽으로 이어짐(이사님 「복잡해도 넣자」).'];
const doc=new Document({styles:{default:{document:{run:{font:F,size:18}}}},sections:[{properties:{page:{size:{width:11906,height:16838,orientation:PageOrientation.LANDSCAPE},margin:{top:720,bottom:720,left:720,right:720}}},children:[
 new Paragraph({children:[new TextRun({text:'대조표 — 발표 자료 v24 → v25 (장별 검토 13건)',font:F,bold:true,size:36})],spacing:{after:80}}),
 P('발표자료_20261007_v24.pptx → v25.pptx (75장, 본 슬라이드 30장). 2026-10-14 세션 61.',{size:18,after:120}),
 H('1. 원칙',HeadingLevel.HEADING_1),...rules.map(s=>P('· '+s,{size:17})),
 H('2. 바뀐 곳',HeadingLevel.HEADING_1),new Table({width:{size:15398,type:WidthType.DXA},columnWidths:W,rows}),
 H('3. 화면 전·후',HeadingLevel.HEADING_1),...imgParas,
 H('4. 이사님 확인 사항',HeadingLevel.HEADING_1),...checks.map(s=>P('· '+s,{size:17}))]}]});
Packer.toBuffer(doc).then(b=>{fs.writeFileSync('대조표_발표자료_v24_v25_20261014.docx',b);console.log('ok',log.length,imgs.length)});
