const fs=require('fs');
const {Document,Packer,Paragraph,TextRun,Table,TableRow,TableCell,WidthType,ShadingType,ImageRun,HeadingLevel,PageOrientation,BorderStyle}=require('docx');
const F='맑은 고딕';
const t=(s,o={})=>new TextRun({text:String(s),font:F,size:o.size||17,bold:o.bold,color:o.color});
const P=(s,o={})=>new Paragraph({children:Array.isArray(s)?s:[t(s,o)],spacing:{after:o.after??60}});
const H=(s,l)=>new Paragraph({heading:l,children:[new TextRun({text:s,font:F,bold:true,size:l===HeadingLevel.HEADING_1?28:22})],spacing:{before:200,after:120}});
const bd={style:BorderStyle.SINGLE,size:4,color:'A6B0C3'};const borders={top:bd,bottom:bd,left:bd,right:bd};
const cell=(s,w,o={})=>new TableCell({width:{size:w,type:WidthType.DXA},borders,shading:o.fill?{type:ShadingType.CLEAR,fill:o.fill,color:'auto'}:undefined,margins:{top:40,bottom:40,left:70,right:70},children:String(s).split('\n').map(x=>P(x,{size:o.size||16,bold:o.bold,color:o.color,after:30}))});
const log=JSON.parse(fs.readFileSync('log11.json','utf8'));
const terms=log.filter(r=>r[1]==='용어 줄 추가'), notes=log.filter(r=>r[1]==='노트 쉽게');
const W1=[700,14698]; const W2=[700,7100,7598];
const rowsT=[new TableRow({tableHeader:true,children:['장','용어 줄 (해당 장 맨 아래, 10.5pt 회색)'].map((h,i)=>cell(h,W1[i],{fill:'1F3864',bold:true,color:'FFFFFF'}))}),...terms.map(r=>new TableRow({cantSplit:true,children:[cell(r[0],W1[0]),cell(r[3],W1[1])]}))];
const rowsN=[new TableRow({tableHeader:true,children:['장','v10 노트(논문 문장 기준)','v11 노트(쉬운 말, 그림 어디 → 뜻 → 다음)'].map((h,i)=>cell(h,W2[i],{fill:'1F3864',bold:true,color:'FFFFFF'}))}),...notes.map(r=>new TableRow({children:[cell(r[0],W2[0]),cell(r[2],W2[1],{color:'7F7F7F',size:15}),cell(r[3],W2[2])]}))];
const imgs=fs.readdirSync('cmp11').filter(f=>f.endsWith('.png')).sort(); const imgParas=[];
for(const f of imgs){const buf=fs.readFileSync('cmp11/'+f);const [w0,h0]=require('child_process').execSync(`python3 -I -c "from PIL import Image;im=Image.open('cmp11/${f}');print(im.width,im.height)"`).toString().trim().split(' ').map(Number);let w=940,h=Math.round(h0*w/w0);imgParas.push(P([t(`슬라이드 ${parseInt(f.slice(1,3))} — 왼쪽 v10, 오른쪽 v11`,{bold:true,size:18})],{after:40}));imgParas.push(new Paragraph({children:[new ImageRun({type:'png',data:buf,transformation:{width:w,height:h}})],spacing:{after:160}}));}
const rules=['용어 풀이 줄: 용어가 본 슬라이드에서 처음 나오는 장 14곳의 맨 아래(근거 줄 위)에 「용어 …」 한두 줄, 10.5pt 회색. 풀이는 논문 본문·부록 K 정의를 쉬운 말로(외부 위원 기준). 자리가 모자라면 내용을 위로 밀고, 그래도 모자라면 그림을 조금 줄임(6번 ×0.93).','전과정 관점은 단계 나열로(이사님 10/7): 원료 채취 → 설계 → 생산 → 운송·배송 → 사용 → 폐기 (KS I ISO 14001 3.3.3 순서). 4번 노트에도 같은 문장 추가.','노트 쉽게: 그림 장 10곳(3·6·8·9·11·12·14·15·23·25)을 「그림 어디를 보라 → 무슨 뜻 → 그래서 다음」으로 줄임. 판정·숫자·용어는 논문 그대로. 나머지 19장은 v9 대본 그대로.','화면 글(용어 줄 외)·백업 31~70은 손대지 않음.'];
const checks=['용어 풀이 문안이 논문 정의와 어긋나는 곳이 있는지(특히 17번 비교주장, 19번 평가 클래스 3, 10번 참조·평가·성숙도 모델의 「업무 목록·채점표·등급 규칙」).','노트 쉽게 쓴 10장을 소리 내어 읽어 보고 너무 가벼운 곳이 있는지.','발표 시간: 노트 축약으로 본 슬라이드 합계 약 30분(v9 32분 30초).'];
const doc=new Document({styles:{default:{document:{run:{font:F,size:18}}}},sections:[{properties:{page:{size:{width:11906,height:16838,orientation:PageOrientation.LANDSCAPE},margin:{top:720,bottom:720,left:720,right:720}}},children:[
 new Paragraph({children:[new TextRun({text:'대조표 — 발표 자료 v10 → v11 (용어 풀이 줄 · 노트 쉽게)',font:F,bold:true,size:36})],spacing:{after:80}}),
 P('발표자료_20261007_v10.pptx → v11.pptx (70장, A4). 2026-10-07 세션 61. 이사님 10/7: 「중요한 용어는 해당 슬라이드에 설명을 조금」, 「그림 장 노트가 너무 어려운 데가 있어 — 쉽고 이해도 쉬운 방향으로」, 「전과정 관점은 원료 채취, 그리고 다음 다음 … 이렇게」.',{size:18,after:120}),
 H('1. 원칙',HeadingLevel.HEADING_1),...rules.map(s=>P('· '+s,{size:17})),
 H('2. 용어 풀이 줄 (14장)',HeadingLevel.HEADING_1),new Table({width:{size:15398,type:WidthType.DXA},columnWidths:W1,rows:rowsT}),
 H('3. 노트 전 → 후 (10장)',HeadingLevel.HEADING_1),new Table({width:{size:15398,type:WidthType.DXA},columnWidths:W2,rows:rowsN}),
 H('4. 화면 전·후 (용어 줄이 들어간 장)',HeadingLevel.HEADING_1),...imgParas,
 H('5. 이사님 확인 사항',HeadingLevel.HEADING_1),...checks.map(s=>P('· '+s,{size:17}))]}]});
Packer.toBuffer(doc).then(b=>{fs.writeFileSync('대조표_발표자료_v10_v11_20261007.docx',b);console.log('ok',terms.length,notes.length,imgs.length)});
