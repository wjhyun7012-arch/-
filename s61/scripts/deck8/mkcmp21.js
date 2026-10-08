const fs=require('fs');
const {Document,Packer,Paragraph,TextRun,Table,TableRow,TableCell,WidthType,ShadingType,ImageRun,HeadingLevel,PageOrientation,BorderStyle}=require('docx');
const F='맑은 고딕';
const t=(s,o={})=>new TextRun({text:String(s),font:F,size:o.size||17,bold:o.bold,color:o.color});
const P=(s,o={})=>new Paragraph({children:Array.isArray(s)?s:[t(s,o)],spacing:{after:o.after??60}});
const H=(s,l)=>new Paragraph({heading:l,children:[new TextRun({text:s,font:F,bold:true,size:l===HeadingLevel.HEADING_1?28:22})],spacing:{before:200,after:120}});
const bd={style:BorderStyle.SINGLE,size:4,color:'A6B0C3'};const borders={top:bd,bottom:bd,left:bd,right:bd};
const cell=(s,w,o={})=>new TableCell({width:{size:w,type:WidthType.DXA},borders,shading:o.fill?{type:ShadingType.CLEAR,fill:o.fill,color:'auto'}:undefined,margins:{top:40,bottom:40,left:70,right:70},children:String(s).split('\n').map(x=>P(x,{size:o.size||16,bold:o.bold,color:o.color,after:30}))});
const log=JSON.parse(fs.readFileSync('log21.json','utf8'));
const W=[600,1300,5000,5600,2898];
const rows=[new TableRow({tableHeader:true,children:['#','슬라이드','항목','v21에 넣은 것','이유'].map((h,i)=>cell(h,W[i],{fill:'1F3864',bold:true,color:'FFFFFF'}))}),...log.map((r,i)=>new TableRow({cantSplit:true,children:[cell(i+1,W[0]),cell(r[0]+' '+r[1],W[1]),cell(r[2],W[2],{color:'7F7F7F',size:15}),cell(r[3],W[3]),cell(r[4],W[4],{size:15})]}))];
const imgs=fs.readdirSync('cmp21').filter(f=>f.endsWith('.png')).sort(); const imgParas=[];
for(const f of imgs){const buf=fs.readFileSync('cmp21/'+f);const [w0,h0]=require('child_process').execSync(`python3 -I -c "from PIL import Image;im=Image.open('cmp21/${f}');print(im.width,im.height)"`).toString().trim().split(' ').map(Number);let w=940,h=Math.round(h0*w/w0);const n=parseInt(f.slice(1,3));imgParas.push(P([t(`슬라이드 ${n} — 새 장(v21)`,{bold:true,size:18})],{after:40}));imgParas.push(new Paragraph({children:[new ImageRun({type:'png',data:buf,transformation:{width:w,height:h}})],spacing:{after:160}}));}
const rules=['이사님 10/8: 심사위원이 AI 활용을 물으면 답할 한 장 — 「백업으로 가자」. 백업 75번 B-Q13 「생성형 AI 활용 범위와 저자 책임」을 B-Q12 뒤에 새로 넣음(75장). 본 슬라이드 30장·다른 백업은 무변경.','내용은 예상 질문 43번 답(2026-09-29 확인분) 기준. 이사님 정정 두 가지 반영 — ① 「개발 33건」이 아니라 「기록물 33건」 ② 논문 전체 설계·구성안·로드맵은 연구자가 했고 AI는 그 검토·수정.','세 상자: 연구자가 한 것(파랑) / AI를 도구로 쓴 것(회색) / 원칙과 확인(빨강). ▶ 길잡이 한 줄, 아래 근거 줄, 노트 존댓말 대본 다섯 문장. 규칙표 v18 0건, validate 통과.','예상 질문 v6: 43번 답의 「구상 단계(2025년 10월)의 논문 구성안·연구 로드맵 정리」 → 「연구자가 세운 논문 구성안·연구 로드맵의 검토·수정(2025년 10월)」, 위치에 → B-Q13 추가. 문항 수 48 그대로.'];
const checks=['왼쪽 상자 「연구자가 한 것」 다섯 줄과 가운데 「AI를 도구로 쓴 것」 다섯 줄이 실제와 맞는지 — 특히 「그림 작성, 검토표 작성, 발표자료 편집」.','근거 줄의 「면담 자료 5.2 나-3」은 43번 답에서 가져온 것 — 자료 이름이 맞는지.'];
const doc=new Document({styles:{default:{document:{run:{font:F,size:18}}}},sections:[{properties:{page:{size:{width:11906,height:16838,orientation:PageOrientation.LANDSCAPE},margin:{top:720,bottom:720,left:720,right:720}}},children:[
 new Paragraph({children:[new TextRun({text:'대조표 — 발표 자료 v20 → v21 (백업 75번 B-Q13 생성형 AI 활용)',font:F,bold:true,size:36})],spacing:{after:80}}),
 P('발표자료_20261007_v20.pptx → v21.pptx (75장, 본 슬라이드 30장). 2026-10-08 세션 61.',{size:18,after:120}),
 H('1. 원칙',HeadingLevel.HEADING_1),...rules.map(s=>P('· '+s,{size:17})),
 H('2. 바뀐 곳',HeadingLevel.HEADING_1),new Table({width:{size:15398,type:WidthType.DXA},columnWidths:W,rows}),
 H('3. 새 장 화면',HeadingLevel.HEADING_1),...imgParas,
 H('4. 이사님 확인 사항',HeadingLevel.HEADING_1),...checks.map(s=>P('· '+s,{size:17}))]}]});
Packer.toBuffer(doc).then(b=>{fs.writeFileSync('대조표_발표자료_v20_v21_20261008.docx',b);console.log('ok',log.length,imgs.length)});
