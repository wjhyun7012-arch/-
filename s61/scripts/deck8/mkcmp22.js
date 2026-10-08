const fs=require('fs');
const {Document,Packer,Paragraph,TextRun,Table,TableRow,TableCell,WidthType,ShadingType,ImageRun,HeadingLevel,PageOrientation,BorderStyle}=require('docx');
const F='맑은 고딕';
const t=(s,o={})=>new TextRun({text:String(s),font:F,size:o.size||17,bold:o.bold,color:o.color});
const P=(s,o={})=>new Paragraph({children:Array.isArray(s)?s:[t(s,o)],spacing:{after:o.after??60}});
const H=(s,l)=>new Paragraph({heading:l,children:[new TextRun({text:s,font:F,bold:true,size:l===HeadingLevel.HEADING_1?28:22})],spacing:{before:200,after:120}});
const bd={style:BorderStyle.SINGLE,size:4,color:'A6B0C3'};const borders={top:bd,bottom:bd,left:bd,right:bd};
const cell=(s,w,o={})=>new TableCell({width:{size:w,type:WidthType.DXA},borders,shading:o.fill?{type:ShadingType.CLEAR,fill:o.fill,color:'auto'}:undefined,margins:{top:40,bottom:40,left:70,right:70},children:String(s).split('\n').map(x=>P(x,{size:o.size||16,bold:o.bold,color:o.color,after:30}))});
const log=JSON.parse(fs.readFileSync('log22.json','utf8'));
const W=[600,1300,5000,5600,2898];
const rows=[new TableRow({tableHeader:true,children:['#','슬라이드','항목','v21 → v22','이유'].map((h,i)=>cell(h,W[i],{fill:'1F3864',bold:true,color:'FFFFFF'}))}),...log.map((r,i)=>new TableRow({cantSplit:true,children:[cell(i+1,W[0]),cell(r[0]+' '+r[1],W[1]),cell(r[2],W[2],{color:'7F7F7F',size:15}),cell(r[3],W[3]),cell(r[4],W[4],{size:15})]}))];
const imgs=fs.readdirSync('cmp22').filter(f=>f.endsWith('.png')).sort(); const imgParas=[];
for(const f of imgs){const buf=fs.readFileSync('cmp22/'+f);const [w0,h0]=require('child_process').execSync(`python3 -I -c "from PIL import Image;im=Image.open('cmp22/${f}');print(im.width,im.height)"`).toString().trim().split(' ').map(Number);let w=940,h=Math.round(h0*w/w0);const n=parseInt(f.slice(1,3));imgParas.push(P([t(`슬라이드 ${n} — 왼쪽 v21, 오른쪽 v22`,{bold:true,size:18})],{after:40}));imgParas.push(new Paragraph({children:[new ImageRun({type:'png',data:buf,transformation:{width:w,height:h}})],spacing:{after:160}}));}
const rules=['이사님 10/8: 「원문은 다 보지는 않고 찾아 준 곳만 확인」, 전 조항 대조도 「전체는 아니고 일부분」 AI 보조 — 75번 B-Q13의 문구를 사실대로 정확하게. 「원문을 확인한 것만」 → 「인용한 조항·문헌은 모두 원문에서 확인 — 찾는 것은 AI, 확인과 판단은 연구자」. 전 조항 대조 줄에 「일부 AI 보조, 연구자가 검토·확정」.','▶ 길잡이 끝 「인용한 조항·수치는 연구자가 원문으로 확인」. 노트 두 문장 같은 취지로. 다른 장 무변경(75장). 규칙표 v18 0건, validate 통과.','예상 질문 v7: 43번 답을 같은 표현으로 — 전 조항 대조에 「일부 AI 보조, 연구자가 검토·확정」, 원칙 문장을 「조항을 찾는 데는 AI를 썼고, 인용한 조항과 문헌은 연구자가 모두 원문에서 확인」으로.','더 물으면 할 답(노트에는 넣지 않음): 「표준 전체를 통독하지는 않았습니다. 조항을 찾는 데 AI를 썼고, 논문에 인용한 조항은 모두 원문을 펴서 확인했습니다. 어느 조항을 쓸지, 어떻게 해석할지는 제가 정했습니다.」'];
const checks=['「일부 AI 보조」의 범위를 물으면 어디까지였는지(검토표 단계인지, 전조항 분석 단계인지) 한 문장으로 준비해 두시면 좋습니다.'];
const doc=new Document({styles:{default:{document:{run:{font:F,size:18}}}},sections:[{properties:{page:{size:{width:11906,height:16838,orientation:PageOrientation.LANDSCAPE},margin:{top:720,bottom:720,left:720,right:720}}},children:[
 new Paragraph({children:[new TextRun({text:'대조표 — 발표 자료 v21 → v22 (B-Q13 문구 정정 — 원문 확인의 범위)',font:F,bold:true,size:36})],spacing:{after:80}}),
 P('발표자료_20261007_v21.pptx → v22.pptx (75장, 본 슬라이드 30장). 2026-10-08 세션 61.',{size:18,after:120}),
 H('1. 원칙',HeadingLevel.HEADING_1),...rules.map(s=>P('· '+s,{size:17})),
 H('2. 바뀐 곳',HeadingLevel.HEADING_1),new Table({width:{size:15398,type:WidthType.DXA},columnWidths:W,rows}),
 H('3. 화면 전·후',HeadingLevel.HEADING_1),...imgParas,
 H('4. 이사님 확인 사항',HeadingLevel.HEADING_1),...checks.map(s=>P('· '+s,{size:17}))]}]});
Packer.toBuffer(doc).then(b=>{fs.writeFileSync('대조표_발표자료_v21_v22_20261008.docx',b);console.log('ok',log.length,imgs.length)});
