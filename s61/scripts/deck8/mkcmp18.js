const fs=require('fs');
const {Document,Packer,Paragraph,TextRun,Table,TableRow,TableCell,WidthType,ShadingType,ImageRun,HeadingLevel,PageOrientation,BorderStyle}=require('docx');
const F='맑은 고딕';
const t=(s,o={})=>new TextRun({text:String(s),font:F,size:o.size||17,bold:o.bold,color:o.color});
const P=(s,o={})=>new Paragraph({children:Array.isArray(s)?s:[t(s,o)],spacing:{after:o.after??60}});
const H=(s,l)=>new Paragraph({heading:l,children:[new TextRun({text:s,font:F,bold:true,size:l===HeadingLevel.HEADING_1?28:22})],spacing:{before:200,after:120}});
const bd={style:BorderStyle.SINGLE,size:4,color:'A6B0C3'};const borders={top:bd,bottom:bd,left:bd,right:bd};
const cell=(s,w,o={})=>new TableCell({width:{size:w,type:WidthType.DXA},borders,shading:o.fill?{type:ShadingType.CLEAR,fill:o.fill,color:'auto'}:undefined,margins:{top:40,bottom:40,left:70,right:70},children:String(s).split('\n').map(x=>P(x,{size:o.size||16,bold:o.bold,color:o.color,after:30}))});
const log=JSON.parse(fs.readFileSync('log18.json','utf8'));
const W=[600,1300,5000,5600,2898];
const rows=[new TableRow({tableHeader:true,children:['#','슬라이드(v15 번호)','항목','v14 → v15','이유'].map((h,i)=>cell(h,W[i],{fill:'1F3864',bold:true,color:'FFFFFF'}))}),...log.map((r,i)=>new TableRow({cantSplit:true,children:[cell(i+1,W[0]),cell(r[0]+' '+r[1],W[1]),cell(r[2],W[2],{color:'7F7F7F',size:15}),cell(r[3],W[3]),cell(r[4],W[4],{size:15})]}))];
const imgs=fs.readdirSync('cmp18').filter(f=>f.endsWith('.png')).sort(); const imgParas=[];
for(const f of imgs){const buf=fs.readFileSync('cmp18/'+f);const [w0,h0]=require('child_process').execSync(`python3 -I -c "from PIL import Image;im=Image.open('cmp18/${f}');print(im.width,im.height)"`).toString().trim().split(' ').map(Number);let w=940,h=Math.round(h0*w/w0);const n=parseInt(f.slice(1,3));imgParas.push(P([t(n>30?`슬라이드 ${n} — 새 백업(옮긴 표·그림)`:`슬라이드 ${n} — 왼쪽 v17, 오른쪽 v18`,{bold:true,size:18})],{after:40}));imgParas.push(new Paragraph({children:[new ImageRun({type:'png',data:buf,transformation:{width:w,height:h}})],spacing:{after:160}}));}
const rules=['이사님 10/7 여섯 가지: ① 4번 제목 (linkage) 뺌(작은 글 첫 줄에만) ② 5번 표 4행 → 2행(정량 평가 / 지속성 — ▶ 줄과 1:1, 14pt) ③ 13번 위아래 배치(Q1~Q4 가로 네 칸 → 네 등급 화살표 → 네 단 막대 가로) ④ 21번 설명 줄·표 내림(표 글자 13pt·행 높이 0.93) ⑤ 「판정을 가른 것」 → 「판정은 … 핵심 문항의 충족 여부로 정해진다」(23번 ▶·노트, 16·21번 「가른/갈랐/가릅」 → 「정한/정하였/정합」) ⑥ 6번 상태 B 「LCA 결과를 받아 중대한 환경측면을 정하고, 환경목표를 세우고, 다음 해에 다시 측정해 확인」.','74장 그대로. 규칙표 v18 0건, validate 통과.'];
const checks=['5번에서 뺀 두 행(접근의 성격·통합·실행 구조)은 4번 상자·백업 B-3·B-4에 있음 — 괜찮은지.','남은 것: 26번(로드맵) 아래 글, 백업 43장 ▶ 길잡이(2차).'];
const doc=new Document({styles:{default:{document:{run:{font:F,size:18}}}},sections:[{properties:{page:{size:{width:11906,height:16838,orientation:PageOrientation.LANDSCAPE},margin:{top:720,bottom:720,left:720,right:720}}},children:[
 new Paragraph({children:[new TextRun({text:'대조표 — 발표 자료 v17 → v18 (여섯 가지 수정)',font:F,bold:true,size:36})],spacing:{after:80}}),
 P('발표자료_20261007_v17.pptx → v18.pptx (74장, 본 슬라이드 30장). 2026-10-07 세션 61.',{size:18,after:120}),
 H('1. 원칙',HeadingLevel.HEADING_1),...rules.map(s=>P('· '+s,{size:17})),
 H('2. 바뀐 곳',HeadingLevel.HEADING_1),new Table({width:{size:15398,type:WidthType.DXA},columnWidths:W,rows}),
 H('3. 화면 전·후',HeadingLevel.HEADING_1),...imgParas,
 H('4. 이사님 확인 사항',HeadingLevel.HEADING_1),...checks.map(s=>P('· '+s,{size:17}))]}]});
Packer.toBuffer(doc).then(b=>{fs.writeFileSync('대조표_발표자료_v17_v18_20261007.docx',b);console.log('ok',log.length,imgs.length)});
