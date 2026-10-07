const fs=require('fs');
const {Document,Packer,Paragraph,TextRun,Table,TableRow,TableCell,WidthType,ShadingType,ImageRun,HeadingLevel,PageOrientation,BorderStyle}=require('docx');
const F='맑은 고딕';
const t=(s,o={})=>new TextRun({text:String(s),font:F,size:o.size||17,bold:o.bold,color:o.color});
const P=(s,o={})=>new Paragraph({children:Array.isArray(s)?s:[t(s,o)],spacing:{after:o.after??60}});
const H=(s,l)=>new Paragraph({heading:l,children:[new TextRun({text:s,font:F,bold:true,size:l===HeadingLevel.HEADING_1?28:22})],spacing:{before:200,after:120}});
const bd={style:BorderStyle.SINGLE,size:4,color:'A6B0C3'};const borders={top:bd,bottom:bd,left:bd,right:bd};
const cell=(s,w,o={})=>new TableCell({width:{size:w,type:WidthType.DXA},borders,shading:o.fill?{type:ShadingType.CLEAR,fill:o.fill,color:'auto'}:undefined,margins:{top:40,bottom:40,left:70,right:70},children:String(s).split('\n').map(x=>P(x,{size:o.size||16,bold:o.bold,color:o.color,after:30}))});
const log=JSON.parse(fs.readFileSync('log15.json','utf8')).concat(JSON.parse(fs.readFileSync('log14.json','utf8')));
const W=[600,1300,5000,5600,2898];
const rows=[new TableRow({tableHeader:true,children:['#','슬라이드(v15 번호)','항목','v14 → v15','이유'].map((h,i)=>cell(h,W[i],{fill:'1F3864',bold:true,color:'FFFFFF'}))}),...log.map((r,i)=>new TableRow({cantSplit:true,children:[cell(i+1,W[0]),cell(r[0]+' '+r[1],W[1]),cell(r[2],W[2],{color:'7F7F7F',size:15}),cell(r[3],W[3]),cell(r[4],W[4],{size:15})]}))];
const imgs=fs.readdirSync('cmp15').filter(f=>f.endsWith('.png')).sort(); const imgParas=[];
for(const f of imgs){const buf=fs.readFileSync('cmp15/'+f);const [w0,h0]=require('child_process').execSync(`python3 -I -c "from PIL import Image;im=Image.open('cmp15/${f}');print(im.width,im.height)"`).toString().trim().split(' ').map(Number);let w=940,h=Math.round(h0*w/w0);const m=f.match(/s(\d+)_from(\d+)/);imgParas.push(P([t(`슬라이드 ${parseInt(m[1])} (v14에서는 ${parseInt(m[2])}번) — 왼쪽 v14, 오른쪽 v15`,{bold:true,size:18})],{after:40}));imgParas.push(new Paragraph({children:[new ImageRun({type:'png',data:buf,transformation:{width:w,height:h}})],spacing:{after:160}}));}
const rules=['이사님 10/7 결정 8건: ① 16번 ③에 (BARS 앵커) ② 적용 개요 상자를 대상·방법·기간·비고 네 줄로 ③ 객관성 장 상자 줄이고 표 올림 ④ 로드맵 장 그림 줄이고 표·설명 올림 ⑤ 장점·단점 풀이 짧게 14pt([표 5-1] 항목 이름 그대로) ⑥ 한 장 요약 표 0.2in 내림 ⑦ 백업 B-17 개발 경위 타임라인을 본문 19번으로(권장안) — 본 슬라이드 29 → 30장, 백업 원본 삭제 ⑧ 6번 상태 B 「다시 결정」 → 「그 결과(LCA 영향평가)를 근거로 … 결정」(권고안, 그림 안 글자는 [그림 1-1] 그대로).','4번 「서로를 전제하지 않는다」 → 「별개의 표준이라 한쪽이 다른 쪽을 요구하지 않는다」는 4번만(v14 정정). 1번·28번 노트는 논문 문장 그대로.','6번 노트의 「다시 재서」(규칙 W05 「재다」) → 「다시 측정해」.','전체 72장(본 30 + 구분 1 + 백업 41). 쪽번호 34곳 재부여. 규칙표 v18 점검 0건, validate 통과.'];
const checks=['19번(개발 경위) 노트 1분 — 발표 시간 합계 약 31분. 교수님 시간 나오면 생략 후보에 19번도 포함할지.','27번 장점·단점 풀이가 너무 짧아진 곳이 있는지.','장점·단점·기회 재정리(학문적 관점·기후변화·ESG 기준)는 별도 결정 — 아래 의견 참조(인수인계 2절).'];
const doc=new Document({styles:{default:{document:{run:{font:F,size:18}}}},sections:[{properties:{page:{size:{width:11906,height:16838,orientation:PageOrientation.LANDSCAPE},margin:{top:720,bottom:720,left:720,right:720}}},children:[
 new Paragraph({children:[new TextRun({text:'대조표 — 발표 자료 v14 → v15 (결정 8건 · 개발 경위 본문으로)',font:F,bold:true,size:36})],spacing:{after:80}}),
 P('발표자료_20261007_v14.pptx(72장) → v15.pptx(72장, 본 슬라이드 30장). 2026-10-07 세션 61.',{size:18,after:120}),
 H('1. 원칙',HeadingLevel.HEADING_1),...rules.map(s=>P('· '+s,{size:17})),
 H('2. 바뀐 곳',HeadingLevel.HEADING_1),new Table({width:{size:15398,type:WidthType.DXA},columnWidths:W,rows}),
 H('3. 화면 전·후',HeadingLevel.HEADING_1),...imgParas,
 H('4. 이사님 확인 사항',HeadingLevel.HEADING_1),...checks.map(s=>P('· '+s,{size:17}))]}]});
Packer.toBuffer(doc).then(b=>{fs.writeFileSync('대조표_발표자료_v14_v15_20261007.docx',b);console.log('ok',log.length,imgs.length)});
