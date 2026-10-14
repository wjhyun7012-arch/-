const fs=require('fs');
const {Document,Packer,Paragraph,TextRun,Table,TableRow,TableCell,WidthType,ShadingType,ImageRun,HeadingLevel,PageOrientation,BorderStyle}=require('docx');
const F='맑은 고딕';
const t=(s,o={})=>new TextRun({text:String(s),font:F,size:o.size||17,bold:o.bold,color:o.color});
const P=(s,o={})=>new Paragraph({children:Array.isArray(s)?s:[t(s,o)],spacing:{after:o.after??60}});
const H=(s,l)=>new Paragraph({heading:l,children:[new TextRun({text:s,font:F,bold:true,size:l===HeadingLevel.HEADING_1?28:22})],spacing:{before:200,after:120}});
const bd={style:BorderStyle.SINGLE,size:4,color:'A6B0C3'};const borders={top:bd,bottom:bd,left:bd,right:bd};
const cell=(s,w,o={})=>new TableCell({width:{size:w,type:WidthType.DXA},borders,shading:o.fill?{type:ShadingType.CLEAR,fill:o.fill,color:'auto'}:undefined,margins:{top:40,bottom:40,left:70,right:70},children:String(s).split('\n').map(x=>P(x,{size:o.size||16,bold:o.bold,color:o.color,after:30}))});
const log=JSON.parse(fs.readFileSync('log27_list.json','utf8'));
const W=[600,1300,5000,5600,2898];
const rows=[new TableRow({tableHeader:true,children:['#','슬라이드(옛 번호)','항목','v26 → v27 15분판','이유'].map((h,i)=>cell(h,W[i],{fill:'1F3864',bold:true,color:'FFFFFF'}))}),...log.map((r,i)=>new TableRow({cantSplit:true,children:[cell(i+1,W[0]),cell(r[0]+' '+r[1],W[1]),cell(r[2],W[2],{color:'7F7F7F',size:15}),cell(r[3],W[3]),cell(r[4],W[4],{size:15})]}))];
const imgs=fs.readdirSync('cmp27').filter(f=>f.endsWith('.png')).sort(); const imgParas=[];
for(const f of imgs){const buf=fs.readFileSync('cmp27/'+f);const [w0,h0]=require('child_process').execSync(`python3 -I -c "from PIL import Image;im=Image.open('cmp27/${f}');print(im.width,im.height)"`).toString().trim().split(' ').map(Number);let w=940,h=Math.round(h0*w/w0);const n=parseInt(f.slice(1,3));imgParas.push(P([t(`슬라이드 ${n} — 왼쪽 v24, 오른쪽 v25`,{bold:true,size:18})],{after:40}));imgParas.push(new Paragraph({children:[new ImageRun({type:'png',data:buf,transformation:{width:w,height:h}})],spacing:{after:160}}));}
const rules=['교수님 10/14: 발표 15분 내외 + 질의응답 15분 내외. 이사님 「15분용으로 만들자」. 본 슬라이드 30장 → 18장(표지·목차·논문 한 장 요약·이 연구를 한 장으로·두 상태·LMA 그림·조항 매핑·연계 지점·16종·55문항·판정 규칙·A사 그림·종합 판정·결과 요약·개선 과제·장단점·결론·감사). 뺀 12장(옛 5·7·8·9·10·11·18·19·21·22·26·28)은 삭제하지 않고 구분 슬라이드 바로 뒤에 「백업 S-n (본심사판 n번)」으로 옮김 — 전체 75장 그대로.','남는 장의 노트는 15분에 맞게 줄이고, 뺀 장의 핵심 한 문장을 이웃 장에 흡수(객관성 두 가지 → 12번, 영역 평점 → 13번, 로드맵·3.15 → 15번). 본 18장 노트 합계 16분 10초. 목차의 Ⅱ(학술적 배경)는 본 장이 없어 2번 노트에 한 문장. 쪽번호 재부여, 노트의 백업 번호 치환, 구분 슬라이드 문구 갱신.','11/12 본심사(논문 전체 보고)는 v26 30장 체계를 그대로 씀(S-n 장을 되돌리면 됨). 규칙표 v18 0건, validate 통과.'];
const checks=['16분 10초 — 「15분 내외」 안이지만 리허설에서 재 보고 3번(1분 30초)이나 13번(1분 20초)에서 더 줄일 수 있음.','배포용 PDF는 본 18장만. 노트북 v27은 참고·물으면을 새 번호로 다시 매김(뺀 장의 질문·용어는 이웃 장에 합침).'];
const doc=new Document({styles:{default:{document:{run:{font:F,size:18}}}},sections:[{properties:{page:{size:{width:11906,height:16838,orientation:PageOrientation.LANDSCAPE},margin:{top:720,bottom:720,left:720,right:720}}},children:[
 new Paragraph({children:[new TextRun({text:'대조표 — 발표 자료 v26 → v27 (15분판 — 본 18장 + 백업 S-n)',font:F,bold:true,size:36})],spacing:{after:80}}),
 P('발표자료_20261007_v26.pptx → v27_15분판.pptx (75장, 본 슬라이드 18장). 2026-10-14 세션 61.',{size:18,after:120}),
 H('1. 원칙',HeadingLevel.HEADING_1),...rules.map(s=>P('· '+s,{size:17})),
 H('2. 바뀐 곳',HeadingLevel.HEADING_1),new Table({width:{size:15398,type:WidthType.DXA},columnWidths:W,rows}),
 
 H('4. 이사님 확인 사항',HeadingLevel.HEADING_1),...checks.map(s=>P('· '+s,{size:17}))]}]});
Packer.toBuffer(doc).then(b=>{fs.writeFileSync('대조표_발표자료_v26_v27_15분판_20261014.docx',b);console.log('ok',log.length,imgs.length)});
