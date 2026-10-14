const fs=require('fs');
const {Document,Packer,Paragraph,TextRun,Table,TableRow,TableCell,WidthType,ShadingType,ImageRun,HeadingLevel,PageOrientation,BorderStyle}=require('docx');
const F='맑은 고딕';
const t=(s,o={})=>new TextRun({text:String(s),font:F,size:o.size||17,bold:o.bold,color:o.color});
const P=(s,o={})=>new Paragraph({children:Array.isArray(s)?s:[t(s,o)],spacing:{after:o.after??60}});
const H=(s,l)=>new Paragraph({heading:l,children:[new TextRun({text:s,font:F,bold:true,size:l===HeadingLevel.HEADING_1?28:22})],spacing:{before:200,after:120}});
const bd={style:BorderStyle.SINGLE,size:4,color:'A6B0C3'};const borders={top:bd,bottom:bd,left:bd,right:bd};
const cell=(s,w,o={})=>new TableCell({width:{size:w,type:WidthType.DXA},borders,shading:o.fill?{type:ShadingType.CLEAR,fill:o.fill,color:'auto'}:undefined,margins:{top:40,bottom:40,left:70,right:70},children:String(s).split('\n').map(x=>P(x,{size:o.size||16,bold:o.bold,color:o.color,after:30}))});
const log=JSON.parse(fs.readFileSync('log28_list.json','utf8'));
const W=[600,1300,5000,5600,2898];
const rows=[new TableRow({tableHeader:true,children:['#','슬라이드(옛 번호)','항목','v27 → v28','이유'].map((h,i)=>cell(h,W[i],{fill:'1F3864',bold:true,color:'FFFFFF'}))}),...log.map((r,i)=>new TableRow({cantSplit:true,children:[cell(i+1,W[0]),cell(r[0]+' '+r[1],W[1]),cell(r[2],W[2],{color:'7F7F7F',size:15}),cell(r[3],W[3]),cell(r[4],W[4],{size:15})]}))];
const imgs=fs.readdirSync('cmp28').filter(f=>f.endsWith('.png')).sort(); const imgParas=[];
for(const f of imgs){const buf=fs.readFileSync('cmp28/'+f);const [w0,h0]=require('child_process').execSync(`python3 -I -c "from PIL import Image;im=Image.open('cmp28/${f}');print(im.width,im.height)"`).toString().trim().split(' ').map(Number);let w=940,h=Math.round(h0*w/w0);const n=parseInt(f.slice(1,3));imgParas.push(P([t(`슬라이드 ${n} — 왼쪽 v27, 오른쪽 v28`,{bold:true,size:18})],{after:40}));imgParas.push(new Paragraph({children:[new ImageRun({type:'png',data:buf,transformation:{width:w,height:h}})],spacing:{after:160}}));}
const rules=['이사님 10/14 지시 다섯 가지 — ① 3번 노트 첫 문장 「이 논문은 다섯 장으로 구성되어 있고, 이 그림은 장 사이의 연결을 보여 줍니다.」 ② 8번 제목 「연계 지점 맵 — 관리 조항과 평가 단계가 만나는 곳」 ③ 논문 그림·표를 그대로 보여 주는 아홉 장(5·6·8·10·11·13·14·15·16)의 제목을 논문 캡션 그대로 쓰고 끝에 [그림·표 번호]를 작은 회색 글자로 ④ 표지 제목을 논문 제목 그대로 「…성숙도 평가에 대한 연구」(네 줄) ⑤ 표지에 영문 제목(심사지 기재용)을 작은 글자로.','슬라이드 본문·노트는 ①을 빼고 손대지 않음. 쪽번호·백업·구분 슬라이드 변동 없음(75장 그대로). 본심사판은 15분판 번호를 본심사 번호(3·6·12·14·15·16·17·23·24·25·27)로 옮겨 같은 내용 반영.','함께 만든 보조 자료 — 심사위원용 한 쪽 「연구 요지와 논문 안내」(국문·영문 제목, 심사지 기록용), 발표자료 유인물(A4 한 쪽에 슬라이드 2장 + 메모 줄, 9쪽), 발표 노트북 v28(제목만 갱신, 27쪽). 규칙표 v18 0건(요지 포함), validate 통과.'];
const checks=['표지 네 줄 제목(24pt)과 영문 제목(10.5pt 회색)이 겹치지 않는 것을 렌더로 확인. 영문 제목이 표지에 없는 편이 깔끔하면 빼도 됨(심사지 기재는 요지 한 쪽으로 충분).','아홉 장 제목이 길어진 곳(8·13·14번)은 23pt로 줄였음. 투영 화면에서 두 줄로 꺾이지 않는지 리허설 때 확인.'];
const doc=new Document({styles:{default:{document:{run:{font:F,size:18}}}},sections:[{properties:{page:{size:{width:11906,height:16838,orientation:PageOrientation.LANDSCAPE},margin:{top:720,bottom:720,left:720,right:720}}},children:[
 new Paragraph({children:[new TextRun({text:'대조표 — 발표 자료 v27 → v28 (15분판·본심사판 공통)',font:F,bold:true,size:36})],spacing:{after:80}}),
 P('발표자료_20261007_v27_15분판.pptx → v28_15분판.pptx (75장, 본 18장). 본심사판 30장(v26 → v28_본심사판30장)에도 같은 열한 곳을 번호만 바꿔 반영. 2026-10-14 세션 61.',{size:18,after:120}),
 H('1. 원칙',HeadingLevel.HEADING_1),...rules.map(s=>P('· '+s,{size:17})),
 H('2. 바뀐 곳',HeadingLevel.HEADING_1),new Table({width:{size:15398,type:WidthType.DXA},columnWidths:W,rows}),
 
 H('3. 바뀐 장 비교 그림',HeadingLevel.HEADING_1),...imgParas,
 H('4. 이사님 확인 사항',HeadingLevel.HEADING_1),...checks.map(s=>P('· '+s,{size:17}))]}]});
Packer.toBuffer(doc).then(b=>{fs.writeFileSync('대조표_발표자료_v27_v28_20261014.docx',b);console.log('ok',log.length,imgs.length)});
