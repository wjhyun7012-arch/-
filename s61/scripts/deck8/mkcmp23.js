const fs=require('fs');
const {Document,Packer,Paragraph,TextRun,Table,TableRow,TableCell,WidthType,ShadingType,ImageRun,HeadingLevel,PageOrientation,BorderStyle}=require('docx');
const F='맑은 고딕';
const t=(s,o={})=>new TextRun({text:String(s),font:F,size:o.size||17,bold:o.bold,color:o.color});
const P=(s,o={})=>new Paragraph({children:Array.isArray(s)?s:[t(s,o)],spacing:{after:o.after??60}});
const H=(s,l)=>new Paragraph({heading:l,children:[new TextRun({text:s,font:F,bold:true,size:l===HeadingLevel.HEADING_1?28:22})],spacing:{before:200,after:120}});
const bd={style:BorderStyle.SINGLE,size:4,color:'A6B0C3'};const borders={top:bd,bottom:bd,left:bd,right:bd};
const cell=(s,w,o={})=>new TableCell({width:{size:w,type:WidthType.DXA},borders,shading:o.fill?{type:ShadingType.CLEAR,fill:o.fill,color:'auto'}:undefined,margins:{top:40,bottom:40,left:70,right:70},children:String(s).split('\n').map(x=>P(x,{size:o.size||16,bold:o.bold,color:o.color,after:30}))});
const log=JSON.parse(fs.readFileSync('log23.json','utf8'));
const W=[600,1300,5000,5600,2898];
const rows=[new TableRow({tableHeader:true,children:['#','슬라이드','항목','v22 → v23 (노트)','규칙·이유'].map((h,i)=>cell(h,W[i],{fill:'1F3864',bold:true,color:'FFFFFF'}))}),...log.map((r,i)=>new TableRow({cantSplit:true,children:[cell(i+1,W[0]),cell(r[0]+' '+r[1],W[1]),cell(r[2],W[2],{color:'7F7F7F',size:15}),cell(r[3],W[3]),cell(r[4],W[4],{size:15})]}))];
const imgs=fs.readdirSync('cmp23').filter(f=>f.endsWith('.png')).sort(); const imgParas=[];
for(const f of imgs){const buf=fs.readFileSync('cmp23/'+f);const [w0,h0]=require('child_process').execSync(`python3 -I -c "from PIL import Image;im=Image.open('cmp23/${f}');print(im.width,im.height)"`).toString().trim().split(' ').map(Number);let w=940,h=Math.round(h0*w/w0);const n=parseInt(f.slice(1,3));imgParas.push(P([t(`슬라이드 ${n} — 왼쪽 v21, 오른쪽 v22`,{bold:true,size:18})],{after:40}));imgParas.push(new Paragraph({children:[new ImageRun({type:'png',data:buf,transformation:{width:w,height:h}})],spacing:{after:160}}));}
const rules=['이사님 10/14 「발표노트를 규칙표에 따라 수정」. 노트 75장을 규칙표 v18(금지 72·경고 3·약어·숫자)로 다시 훑음 — 정규식 점검기 0건이었으나 패턴 밖 변형과 약어 첫 등장, 발표 대본 존댓말, 지시문 참조를 손봄. 화면·표·그림은 무변경(v22와 동일 확인).','고친 것: 9번 「잇고·재려면·빌려야」 → 「연계하고·측정하려면·적용해야」(T04·W05, 이사님 「적용」) / 10번·23번 「…인가. 다음 장입니다」 → 「…인지, 다음 장에서 보겠습니다」, 23번 판정 경로 「2점 이상인가. 통과입니다」 → 「2점 이상이어서 통과입니다」(이사님 문안) / 12번 LMA 첫 등장 영문 풀이 / 19번 「행동기술척도, 곧 BARS」 / 30번 지시문 B-Q13·예상 질문 v7 / 백업 44·49~53 지시문 「행동기술척도(BARS) 앵커」.','확인하고 둔 것: 「이어진다」(22·25~28)는 논문 본문이 16곳에서 쓰는 말이라 유지. 숫자(2.85·2.89·3.15·영역 평점) 규칙표와 일치. 〔 〕 지시문의 반말은 읽지 않는 부분이라 유지. 규칙표 v18 0건, validate 통과.'];
const checks=['23번 ML3 줄은 「3점 이상이어야 하는데 D-1이 2점이라 미충족입니다」로 — 「이어서 통과」 꼴을 쓸 수 없어 이렇게 맞춤.','본 슬라이드 노트 시간은 바뀌지 않음(글자 수 거의 같음).'];
const doc=new Document({styles:{default:{document:{run:{font:F,size:18}}}},sections:[{properties:{page:{size:{width:11906,height:16838,orientation:PageOrientation.LANDSCAPE},margin:{top:720,bottom:720,left:720,right:720}}},children:[
 new Paragraph({children:[new TextRun({text:'대조표 — 발표 자료 v22 → v23 (발표 노트 규칙표 수정)',font:F,bold:true,size:36})],spacing:{after:80}}),
 P('발표자료_20261007_v22.pptx → v23.pptx (75장, 본 슬라이드 30장, 노트만 변경). 2026-10-14 세션 61.',{size:18,after:120}),
 H('1. 원칙',HeadingLevel.HEADING_1),...rules.map(s=>P('· '+s,{size:17})),
 H('2. 바뀐 곳',HeadingLevel.HEADING_1),new Table({width:{size:15398,type:WidthType.DXA},columnWidths:W,rows}),
 
 H('4. 이사님 확인 사항',HeadingLevel.HEADING_1),...checks.map(s=>P('· '+s,{size:17}))]}]});
Packer.toBuffer(doc).then(b=>{fs.writeFileSync('대조표_발표자료_v22_v23_20261014.docx',b);console.log('ok',log.length,imgs.length)});
