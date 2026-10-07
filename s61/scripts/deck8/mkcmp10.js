const fs=require('fs');
const {Document,Packer,Paragraph,TextRun,Table,TableRow,TableCell,WidthType,ShadingType,ImageRun,HeadingLevel,PageOrientation,BorderStyle}=require('docx');
const F='맑은 고딕';
const t=(s,o={})=>new TextRun({text:String(s),font:F,size:o.size||17,bold:o.bold,color:o.color});
const P=(s,o={})=>new Paragraph({children:Array.isArray(s)?s:[t(s,o)],spacing:{after:o.after??60}});
const H=(s,l)=>new Paragraph({heading:l,children:[new TextRun({text:s,font:F,bold:true,size:l===HeadingLevel.HEADING_1?28:22})],spacing:{before:200,after:120}});
const bd={style:BorderStyle.SINGLE,size:4,color:'A6B0C3'};const borders={top:bd,bottom:bd,left:bd,right:bd};
const cell=(s,w,o={})=>new TableCell({width:{size:w,type:WidthType.DXA},borders,shading:o.fill?{type:ShadingType.CLEAR,fill:o.fill,color:'auto'}:undefined,margins:{top:40,bottom:40,left:70,right:70},children:[P(s,{size:o.size||16,bold:o.bold,color:o.color,after:0})]});
const log=JSON.parse(fs.readFileSync('v10_fitlog.json','utf8'));
const W=[600,1500,3600,5200,4498];
const rows=[new TableRow({tableHeader:true,children:['#','슬라이드','항목','전 → 후','이유'].map((h,i)=>cell(h,W[i],{fill:'1F3864',bold:true,color:'FFFFFF'}))}),
 ...log.map((r,i)=>new TableRow({cantSplit:true,children:[cell(i+1,W[0]),cell(r[0],W[1]),cell(r[1],W[2]),cell(r[2]+'  →  '+r[3],W[3]),cell(r[4]||'',W[4],{size:15})]}))];
const imgs=fs.readdirSync('cmp10').filter(f=>f.endsWith('.png')).sort();
const imgParas=[];
for(const f of imgs){const buf=fs.readFileSync('cmp10/'+f);const [w0,h0]=require('child_process').execSync(`python3 -I -c "from PIL import Image;im=Image.open('cmp10/${f}');print(im.width,im.height)"`).toString().trim().split(' ').map(Number);let w=940,h=Math.round(h0*w/w0);
 imgParas.push(P([t(`슬라이드 ${parseInt(f.slice(1,3))} — 왼쪽 v9, 오른쪽 v10`,{bold:true,size:18})],{after:40}));imgParas.push(new Paragraph({children:[new ImageRun({type:'png',data:buf,transformation:{width:w,height:h}})],spacing:{after:160}}));}
const rules=['머리(제목·장 표시·선·▶ 줄)를 전 슬라이드에서 위로 약 0.17in 올림 — 이사님 10/7 「위쪽 선도 조금 더 올리자」.','그림은 ▶ 줄 아래부터 쪽번호 위까지, 좌우 여백 안에서 비율을 지켜 최대로. 옆에 표·상자가 있는 장(16·21·22번)은 그 칸 안에서만.','그림 아래 글은 그림 바로 밑으로 내림. 글자 크기는 줄이지 않았음(줄여도 그림이 8% 이상 커지는 장이 없었음).','그림 없는 장은 내용 전체를 ▶ 줄 바로 아래로 당김(14장).','백업 31~70의 그림 장도 같은 규칙(B-1~B-9, B-13·B-14, B-0a~c).','화면 글·노트는 손대지 않음.'];
const checks=['PowerPoint에서 3·6·8·9·14·15·25번(가장 많이 커진 장)과 백업 34~48을 넘겨 보고, 그림 아래 글이 쪽번호와 닿는 곳이 없는지.','10번 슬라이드: 삭제된 것이 아니라 v7 그대로 있음(「무엇을 어떻게 측정하는가 — 따른 표준」). 삭제한 것은 71~80번(예상 질문 43문 표)뿐 — 예상 질문 v5 파일로 대체.'];
const doc=new Document({styles:{default:{document:{run:{font:F,size:18}}}},sections:[{properties:{page:{size:{width:11906,height:16838,orientation:PageOrientation.LANDSCAPE},margin:{top:720,bottom:720,left:720,right:720}}},children:[
 new Paragraph({children:[new TextRun({text:'대조표 — 발표 자료 v9 → v10 (그림 최대화)',font:F,bold:true,size:36})],spacing:{after:80}}),
 P('발표자료_20261007_v9.pptx → 발표자료_20261007_v10.pptx (70장, A4). 2026-10-07 세션 61. 이사님 10/7: 「그림이 화면에서 최대한 크게 보이도록, 위쪽 선도 조금 올리자. 크게 잘 보여야 질문도 줄어든다」.',{size:18,after:120}),
 H('1. 원칙',HeadingLevel.HEADING_1),...rules.map(s=>P('· '+s,{size:17})),
 H('2. 바뀐 곳',HeadingLevel.HEADING_1),new Table({width:{size:15398,type:WidthType.DXA},columnWidths:W,rows}),
 H('3. 전·후 화면 (10% 이상 커진 장)',HeadingLevel.HEADING_1),...imgParas,
 H('4. 이사님 확인 사항',HeadingLevel.HEADING_1),...checks.map(s=>P('· '+s,{size:17}))]}]});
Packer.toBuffer(doc).then(b=>{fs.writeFileSync('대조표_발표자료_v9_v10_20261007.docx',b);console.log('ok',log.length,imgs.length)});
