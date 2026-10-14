const fs=require('fs');
const {Document,Packer,Paragraph,TextRun,Table,TableRow,TableCell,WidthType,ShadingType,ImageRun,HeadingLevel,PageOrientation,BorderStyle}=require('docx');
const F='맑은 고딕';
const t=(s,o={})=>new TextRun({text:String(s),font:F,size:o.size||17,bold:o.bold,color:o.color});
const P=(s,o={})=>new Paragraph({children:Array.isArray(s)?s:[t(s,o)],spacing:{after:o.after??60}});
const H=(s,l)=>new Paragraph({heading:l,children:[new TextRun({text:s,font:F,bold:true,size:l===HeadingLevel.HEADING_1?28:22})],spacing:{before:200,after:120}});
const bd={style:BorderStyle.SINGLE,size:4,color:'A6B0C3'};const borders={top:bd,bottom:bd,left:bd,right:bd};
const cell=(s,w,o={})=>new TableCell({width:{size:w,type:WidthType.DXA},borders,shading:o.fill?{type:ShadingType.CLEAR,fill:o.fill,color:'auto'}:undefined,margins:{top:40,bottom:40,left:70,right:70},children:String(s).split('\n').map(x=>P(x,{size:o.size||16,bold:o.bold,color:o.color,after:30}))});
const log=JSON.parse(fs.readFileSync('log24.json','utf8'));
const W=[600,1300,5000,5600,2898];
const rows=[new TableRow({tableHeader:true,children:['#','슬라이드','항목','v23 → v24','이유'].map((h,i)=>cell(h,W[i],{fill:'1F3864',bold:true,color:'FFFFFF'}))}),...log.map((r,i)=>new TableRow({cantSplit:true,children:[cell(i+1,W[0]),cell(r[0]+' '+r[1],W[1]),cell(r[2],W[2],{color:'7F7F7F',size:15}),cell(r[3],W[3]),cell(r[4],W[4],{size:15})]}))];
const imgs=fs.readdirSync('cmp24').filter(f=>f.endsWith('.png')).sort(); const imgParas=[];
for(const f of imgs){const buf=fs.readFileSync('cmp24/'+f);const [w0,h0]=require('child_process').execSync(`python3 -I -c "from PIL import Image;im=Image.open('cmp24/${f}');print(im.width,im.height)"`).toString().trim().split(' ').map(Number);let w=940,h=Math.round(h0*w/w0);const n=parseInt(f.slice(1,3));imgParas.push(P([t(`슬라이드 ${n} — 왼쪽 v23, 오른쪽 v24`,{bold:true,size:18})],{after:40}));imgParas.push(new Paragraph({children:[new ImageRun({type:'png',data:buf,transformation:{width:w,height:h}})],spacing:{after:160}}));}
const rules=['이사님 10/14 검토 결과 반영. ① 1번 노트: 「산업경영공학과 박사과정 현욱진」, 「…성숙도 평가에 대한 논문을 발표하겠습니다」, 요지는 「서로를 전제하지 않는」을 빼고 초록 [1] 기반 (나)안(T38 괄호 고정형). ② 2번 목차 노트: 「다섯 부와 맺음」 → 논문 장 순서(1장~5장 결론), 「3장이 기준, 4장이 적용 결과와 개선 과제」, 40초 → 15초, 숫자(16종·55문항·2.85·ML2·3.15)는 4번으로, 지시문 세부 목록을 현재 제목에 맞춤(T15 세 계층, 19번 반복 개발). ③ 화면 「맺음」 → 「결론」(논문 5장 「결론 및 시사점」): 2번 목차 줄, 28번 장 표시, 29번 제목·장 표시.','규칙표 v18 0건, validate 통과. 75장 그대로. 본 슬라이드 노트 합계 31분 20초 → 30분 55초.'];
const checks=['노트 쉽게 쓰기 후보 11장(5·7·10·12·17·18·19·20·21·27·28)은 발표 노트북 v24 검토용으로 장별 결정 대기.'];
const doc=new Document({styles:{default:{document:{run:{font:F,size:18}}}},sections:[{properties:{page:{size:{width:11906,height:16838,orientation:PageOrientation.LANDSCAPE},margin:{top:720,bottom:720,left:720,right:720}}},children:[
 new Paragraph({children:[new TextRun({text:'대조표 — 발표 자료 v23 → v24 (1·2번 노트, 「맺음」→「결론」)',font:F,bold:true,size:36})],spacing:{after:80}}),
 P('발표자료_20261007_v23.pptx → v24.pptx (75장, 본 슬라이드 30장). 2026-10-14 세션 61.',{size:18,after:120}),
 H('1. 원칙',HeadingLevel.HEADING_1),...rules.map(s=>P('· '+s,{size:17})),
 H('2. 바뀐 곳',HeadingLevel.HEADING_1),new Table({width:{size:15398,type:WidthType.DXA},columnWidths:W,rows}),
 H('3. 화면 전·후',HeadingLevel.HEADING_1),...imgParas,
 H('4. 이사님 확인 사항',HeadingLevel.HEADING_1),...checks.map(s=>P('· '+s,{size:17}))]}]});
Packer.toBuffer(doc).then(b=>{fs.writeFileSync('대조표_발표자료_v23_v24_20261014.docx',b);console.log('ok',log.length,imgs.length)});
