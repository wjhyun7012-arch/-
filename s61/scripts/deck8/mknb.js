// 발표 노트북(연단용) — 사용: node mknb.js data.json out.docx "제목" "판 설명"
const fs=require('fs');
const {Document,Packer,Paragraph,TextRun,ImageRun,Table,TableRow,TableCell,WidthType,BorderStyle,PageBreak,AlignmentType,ShadingType}=require('docx');
const [,,dataPath,outPath,TITLE,SUB]=process.argv;
const F='맑은 고딕';
const data=JSON.parse(fs.readFileSync(dataPath,'utf8'));
const t=(s,o={})=>new TextRun({text:String(s),font:F,size:o.size||28,bold:o.bold,color:o.color,italics:o.italics});
const P=(s,o={})=>new Paragraph({children:Array.isArray(s)?s:[t(s,o)],spacing:{after:o.after??120,line:o.line||300},alignment:o.align});
const fmt=s=>s==null?'—':(s>=60?`${Math.floor(s/60)}분${s%60?' '+s%60+'초':''}`:`${s}초`);
const bd={style:BorderStyle.SINGLE,size:4,color:'A6B0C3'};const borders={top:bd,bottom:bd,left:bd,right:bd};
const kids=[];
// 표지
kids.push(P(TITLE,{size:40,bold:true,after:200}));
kids.push(P(SUB,{size:20,color:'595959',after:200}));
kids.push(P('쓰는 법: 한 쪽이 슬라이드 하나. 큰 글자가 읽는 대본, 아래 회색 글이 지시문(읽지 않음). 오른쪽 위 시간은 이 장의 목표 시간과 누적 시간. 질문이 오면 마지막 쪽의 백업 색인에서 번호를 찾아 띄운다.',{size:20,after:200}));
const rows=[new TableRow({tableHeader:true,children:['장','제목','목표','누적'].map(h=>new TableCell({borders,width:{size:[700,7500,1200,1200][['장','제목','목표','누적'].indexOf(h)],type:WidthType.DXA},shading:{type:ShadingType.CLEAR,fill:'1F3864',color:'auto'},children:[P(h,{size:18,bold:true,color:'FFFFFF',after:0})]}))})];
for(const d of data.filter(x=>x.n<=30)) rows.push(new TableRow({children:[[d.n,700],[d.title,7500],[fmt(d.secs),1200],[fmt(d.cum),1200]].map(([v,w])=>new TableCell({borders,width:{size:w,type:WidthType.DXA},margins:{top:20,bottom:20,left:60,right:60},children:[P(v,{size:17,after:0})]}))}));
kids.push(new Table({width:{size:10600,type:WidthType.DXA},columnWidths:[700,7500,1200,1200],rows}));
kids.push(new Paragraph({children:[new PageBreak()]}));
// 장별
for(const d of data.filter(x=>x.n<=30)){
  kids.push(new Table({width:{size:10600,type:WidthType.DXA},columnWidths:[7600,3000],rows:[new TableRow({children:[
    new TableCell({borders:{top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.SINGLE,size:12,color:'1F3864'},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}},width:{size:7600,type:WidthType.DXA},children:[P([t(`${d.n}  `,{size:44,bold:true,color:'C00000'}),t(d.title,{size:26,bold:true,color:'1F3864'})],{after:40}),P(d.sec||' ',{size:17,color:'7F7F7F',after:40})]}),
    new TableCell({borders:{top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.SINGLE,size:12,color:'1F3864'},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}},width:{size:3000,type:WidthType.DXA},children:[P(`목표 ${fmt(d.secs)}`,{size:20,bold:true,align:AlignmentType.RIGHT,after:0}),P(`누적 ${fmt(d.cum)} / 31분 20초`,{size:17,color:'7F7F7F',align:AlignmentType.RIGHT,after:40})]})]})]}));
  const buf=fs.readFileSync(d.img); const w=330,h=Math.round(w*210/297);
  kids.push(new Paragraph({children:[new ImageRun({type:'png',data:buf,transformation:{width:w,height:h}})],spacing:{before:120,after:160}}));
  for(const s of d.sents) kids.push(P(s,{size:28,after:140,line:340}));
  if(d.cues.length){kids.push(P('—',{size:16,color:'A6A6A6',after:40}));for(const c of d.cues) kids.push(P('〔'+c+'〕',{size:17,color:'7F7F7F',after:40,line:260}));}
  kids.push(new Paragraph({children:[new PageBreak()]}));
}
// 백업 색인
kids.push(P('백업 색인 — 질문이 오면 이 번호를 띄운다',{size:30,bold:true,color:'1F3864',after:160}));
const brows=[new TableRow({tableHeader:true,children:['장','백업 번호','제목'].map((h,i)=>new TableCell({borders,width:{size:[700,1700,8200][i],type:WidthType.DXA},shading:{type:ShadingType.CLEAR,fill:'1F3864',color:'auto'},children:[P(h,{size:18,bold:true,color:'FFFFFF',after:0})]}))})];
for(const d of data.filter(x=>x.n>30)) brows.push(new TableRow({children:[[d.n,700],[(d.label||'').replace(/^백업\s*/,''),1700],[d.title,8200]].map(([v,w])=>new TableCell({borders,width:{size:w,type:WidthType.DXA},margins:{top:20,bottom:20,left:60,right:60},children:[P(v,{size:17,after:0})]}))}));
kids.push(new Table({width:{size:10600,type:WidthType.DXA},columnWidths:[700,1700,8200],rows:brows}));
const doc=new Document({styles:{default:{document:{run:{font:F,size:24}}}},sections:[{properties:{page:{size:{width:11906,height:16838},margin:{top:650,bottom:650,left:650,right:650}}},children:kids}]});
Packer.toBuffer(doc).then(b=>{fs.writeFileSync(outPath,b);console.log('ok',outPath,b.length)});
