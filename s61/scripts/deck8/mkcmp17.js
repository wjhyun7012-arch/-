const fs=require('fs');
const {Document,Packer,Paragraph,TextRun,Table,TableRow,TableCell,WidthType,ShadingType,ImageRun,HeadingLevel,PageOrientation,BorderStyle}=require('docx');
const F='맑은 고딕';
const t=(s,o={})=>new TextRun({text:String(s),font:F,size:o.size||17,bold:o.bold,color:o.color});
const P=(s,o={})=>new Paragraph({children:Array.isArray(s)?s:[t(s,o)],spacing:{after:o.after??60}});
const H=(s,l)=>new Paragraph({heading:l,children:[new TextRun({text:s,font:F,bold:true,size:l===HeadingLevel.HEADING_1?28:22})],spacing:{before:200,after:120}});
const bd={style:BorderStyle.SINGLE,size:4,color:'A6B0C3'};const borders={top:bd,bottom:bd,left:bd,right:bd};
const cell=(s,w,o={})=>new TableCell({width:{size:w,type:WidthType.DXA},borders,shading:o.fill?{type:ShadingType.CLEAR,fill:o.fill,color:'auto'}:undefined,margins:{top:40,bottom:40,left:70,right:70},children:String(s).split('\n').map(x=>P(x,{size:o.size||16,bold:o.bold,color:o.color,after:30}))});
const log=JSON.parse(fs.readFileSync('log17.json','utf8'));
const W=[600,1300,5000,5600,2898];
const rows=[new TableRow({tableHeader:true,children:['#','슬라이드(v15 번호)','항목','v14 → v15','이유'].map((h,i)=>cell(h,W[i],{fill:'1F3864',bold:true,color:'FFFFFF'}))}),...log.map((r,i)=>new TableRow({cantSplit:true,children:[cell(i+1,W[0]),cell(r[0]+' '+r[1],W[1]),cell(r[2],W[2],{color:'7F7F7F',size:15}),cell(r[3],W[3]),cell(r[4],W[4],{size:15})]}))];
const imgs=fs.readdirSync('cmp17').filter(f=>f.endsWith('.png')).sort(); const imgParas=[];
for(const f of imgs){const buf=fs.readFileSync('cmp17/'+f);const [w0,h0]=require('child_process').execSync(`python3 -I -c "from PIL import Image;im=Image.open('cmp17/${f}');print(im.width,im.height)"`).toString().trim().split(' ').map(Number);let w=940,h=Math.round(h0*w/w0);const n=parseInt(f.slice(1,3));imgParas.push(P([t(n>30?`슬라이드 ${n} — 새 백업(옮긴 표·그림)`:`슬라이드 ${n} — 왼쪽 v16, 오른쪽 v17`,{bold:true,size:18})],{after:40}));imgParas.push(new Paragraph({children:[new ImageRun({type:'png',data:buf,transformation:{width:w,height:h}})],spacing:{after:160}}));}
const rules=['이사님 10/7: 10번 조항 번호 → 「역할 한 줄」(번호는 백업 B-2) / 13번 Q1~Q4 짧게 + 네 단 막대(매우 높음 4개 조항만 번호) — [표 3-2] 표는 백업 B-1a(33번) / 19번 제목 「반복 개발」 / 21번 표 0.15in 위로 / 23번 판정 경로를 도형 4칸으로 크게 — [그림 4-3] 막대그래프는 백업 B-13a(35번).','74장 = 본 30 + 구분 1 + 백업 43. 쪽번호 42곳 재부여. 규칙표 v18 0건, validate 통과. 노트 10·13·23번 다시 씀.'];
const checks=['13번 네 단 막대의 설명 문구(높음·중간·낮음)가 [표 3-2] 요지와 맞는지.','23번 판정 경로 네 칸 — 「ML2 진입 36/51 = 70.6%」 수치는 4.3.5.','남은 것: 26번(로드맵) 아래 글, 백업 43장 ▶ 길잡이 한 줄(2차).'];
const doc=new Document({styles:{default:{document:{run:{font:F,size:18}}}},sections:[{properties:{page:{size:{width:11906,height:16838,orientation:PageOrientation.LANDSCAPE},margin:{top:720,bottom:720,left:720,right:720}}},children:[
 new Paragraph({children:[new TextRun({text:'대조표 — 발표 자료 v16 → v17 (10·13·23번 쉽게 · 백업 B-1a·B-13a)',font:F,bold:true,size:36})],spacing:{after:80}}),
 P('발표자료_20261007_v16.pptx(72장) → v17.pptx(74장, 본 슬라이드 30장). 2026-10-07 세션 61.',{size:18,after:120}),
 H('1. 원칙',HeadingLevel.HEADING_1),...rules.map(s=>P('· '+s,{size:17})),
 H('2. 바뀐 곳',HeadingLevel.HEADING_1),new Table({width:{size:15398,type:WidthType.DXA},columnWidths:W,rows}),
 H('3. 화면 전·후',HeadingLevel.HEADING_1),...imgParas,
 H('4. 이사님 확인 사항',HeadingLevel.HEADING_1),...checks.map(s=>P('· '+s,{size:17}))]}]});
Packer.toBuffer(doc).then(b=>{fs.writeFileSync('대조표_발표자료_v16_v17_20261007.docx',b);console.log('ok',log.length,imgs.length)});
