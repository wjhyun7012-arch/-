// 발표 노트북(연단용 커닝 페이퍼) — node mknb2.js data.json cheat.json out.docx "제목" "설명"
const fs=require('fs');
const {Document,Packer,Paragraph,TextRun,ImageRun,Table,TableRow,TableCell,WidthType,BorderStyle,PageBreak,AlignmentType,ShadingType}=require('docx');
const [,,dataPath,cheatPath,outPath,TITLE,SUB]=process.argv;
const F='맑은 고딕';
const data=JSON.parse(fs.readFileSync(dataPath,'utf8')); const CH=JSON.parse(fs.readFileSync(cheatPath,'utf8'));
const t=(s,o={})=>new TextRun({text:String(s),font:F,size:o.size||28,bold:o.bold,color:o.color});
const P=(s,o={})=>new Paragraph({children:Array.isArray(s)?s:[t(s,o)],spacing:{after:o.after??120,line:o.line||300,before:o.before||0},alignment:o.align});
const fmt=s=>s==null?'—':(s>=60?`${Math.floor(s/60)}분${s%60?' '+s%60+'초':''}`:`${s}초`);
const last=data.filter(x=>x.n<=30).slice(-1)[0]; const TOTAL=fmt(last.cum);
const bd={style:BorderStyle.SINGLE,size:4,color:'A6B0C3'};const borders={top:bd,bottom:bd,left:bd,right:bd};
const none={style:BorderStyle.NONE};
const box=(title,paras,fill,w=10600)=>new Table({width:{size:w,type:WidthType.DXA},columnWidths:[w],rows:[new TableRow({children:[new TableCell({borders:{top:bd,bottom:bd,left:bd,right:bd},width:{size:w,type:WidthType.DXA},shading:{type:ShadingType.CLEAR,fill,color:'auto'},margins:{top:60,bottom:60,left:100,right:100},children:[P(title,{size:17,bold:true,color:'1F3864',after:30}),...paras]})]})]});
const kids=[];
kids.push(P(TITLE,{size:40,bold:true,after:200}));
kids.push(P(SUB,{size:20,color:'595959',after:200}));
kids.push(P('쓰는 법: 한 쪽이 슬라이드 하나. 큰 글자가 읽는 대본, 회색 글이 지시문(읽지 않음). 아래 「용어」는 논문 부록 K 용어집의 문장 그대로(한글·영문·뜻·출처), 「물으면」은 예상 질문의 한 줄 답과 전문 번호·띄울 백업 번호. 질문 전문은 맨 뒤 부록, 백업 색인은 그 앞.',{size:20,after:200}));
const rows=[new TableRow({tableHeader:true,children:['장','제목','목표','누적'].map((h,i)=>new TableCell({borders,width:{size:[700,7500,1200,1200][i],type:WidthType.DXA},shading:{type:ShadingType.CLEAR,fill:'1F3864',color:'auto'},children:[P(h,{size:18,bold:true,color:'FFFFFF',after:0})]}))})];
for(const d of data.filter(x=>x.n<=30)) rows.push(new TableRow({children:[[d.n,700],[d.title,7500],[fmt(d.secs),1200],[fmt(d.cum),1200]].map(([v,w])=>new TableCell({borders,width:{size:w,type:WidthType.DXA},margins:{top:20,bottom:20,left:60,right:60},children:[P(v,{size:17,after:0})]}))}));
kids.push(new Table({width:{size:10600,type:WidthType.DXA},columnWidths:[700,7500,1200,1200],rows}));
kids.push(new Paragraph({children:[new PageBreak()]}));
for(const d of data.filter(x=>x.n<=30)){
  const c=CH.cheat[String(d.n)]||{terms:[],qs:[]};
  kids.push(new Table({width:{size:10600,type:WidthType.DXA},columnWidths:[7600,3000],rows:[new TableRow({children:[
    new TableCell({borders:{top:none,bottom:{style:BorderStyle.SINGLE,size:12,color:'1F3864'},left:none,right:none},width:{size:7600,type:WidthType.DXA},children:[P([t(`${d.n}  `,{size:44,bold:true,color:'C00000'}),t(d.title,{size:26,bold:true,color:'1F3864'})],{after:40}),P(d.sec||' ',{size:17,color:'7F7F7F',after:40})]}),
    new TableCell({borders:{top:none,bottom:{style:BorderStyle.SINGLE,size:12,color:'1F3864'},left:none,right:none},width:{size:3000,type:WidthType.DXA},children:[P(`목표 ${fmt(d.secs)}`,{size:20,bold:true,align:AlignmentType.RIGHT,after:0}),P(`누적 ${fmt(d.cum)} / ${TOTAL}`,{size:17,color:'7F7F7F',align:AlignmentType.RIGHT,after:40})]})]})]}));
  const buf=fs.readFileSync(d.img); const w=(c.thumb||300),h=Math.round(w*210/297);
  kids.push(new Paragraph({children:[new ImageRun({type:'png',data:buf,transformation:{width:w,height:h}})],spacing:{before:100,after:120}}));
  for(const s of d.sents) kids.push(P(s,{size:27,after:110,line:320}));
  if(d.cues.length){for(const cu of d.cues) kids.push(P('〔'+cu+'〕',{size:16,color:'7F7F7F',after:30,line:240}));}
  if(c.terms.length){kids.push(P(' ',{size:8,after:0}));kids.push(box('용어 — 논문 부록 K 그대로',c.terms.map(x=>P([t(x.ko,{size:17,bold:true}),t(x.en?`(${x.en})`:'',{size:16,color:'595959'}),t(' — '+x.def,{size:16}),t('  · '+x.src,{size:14,color:'7F7F7F'})],{after:30,line:230})),'F2F4F8'));}
  if(c.qs.length){kids.push(P(' ',{size:8,after:0}));kids.push(box('물으면',c.qs.map(x=>P([t('Q. '+x.q+'  ',{size:17,bold:true}),t('→ '+x.a,{size:16}),t('  · 전문 '+x.no+'번'+(x.ref?' · '+x.ref:'')+(x.loc?' · 논문 '+x.loc:''),{size:14,color:'7F7F7F'})],{after:30,line:230})),'FFF8E1'));}
  kids.push(new Paragraph({children:[new PageBreak()]}));
}
kids.push(P('백업 색인 — 질문이 오면 이 번호를 띄운다',{size:30,bold:true,color:'1F3864',after:160}));
const brows=[new TableRow({tableHeader:true,children:['장','백업 번호','제목'].map((h,i)=>new TableCell({borders,width:{size:[700,1700,8200][i],type:WidthType.DXA},shading:{type:ShadingType.CLEAR,fill:'1F3864',color:'auto'},children:[P(h,{size:18,bold:true,color:'FFFFFF',after:0})]}))})];
for(const d of data.filter(x=>x.n>30)) brows.push(new TableRow({children:[[d.n,700],[(d.label||'').replace(/^백업\s*/,''),1700],[d.title,8200]].map(([v,w])=>new TableCell({borders,width:{size:w,type:WidthType.DXA},margins:{top:20,bottom:20,left:60,right:60},children:[P(v,{size:17,after:0})]}))}));
kids.push(new Table({width:{size:10600,type:WidthType.DXA},columnWidths:[700,1700,8200],rows:brows}));
kids.push(new Paragraph({children:[new PageBreak()]}));
// 예상질문 전문
kids.push(P('예상 질문 전문 — 52문 (v7 48문 + ★49~52)',{size:30,bold:true,color:'1F3864',after:120}));
const inv={}; for(const [n,arr] of Object.entries(CH.M)) for(const q of arr) (inv[q]=inv[q]||[]).push(n);
for(const no of Object.keys(CH.Q).map(Number).sort((a,b)=>a-b)){const q=CH.Q[String(no)];
  kids.push(P([t(`${no}. `,{size:19,bold:true,color:'C00000'}),t(q.q,{size:19,bold:true}),t(inv[no]?`  〔장 ${inv[no].join('·')}〕`:'',{size:15,color:'7F7F7F'})],{after:30,before:120,line:260}));
  kids.push(P(q.a,{size:17,after:20,line:250}));
  kids.push(P('논문 '+q.loc,{size:14,color:'7F7F7F',after:60}));}
const doc=new Document({styles:{default:{document:{run:{font:F,size:24}}}},sections:[{properties:{page:{size:{width:11906,height:16838},margin:{top:600,bottom:600,left:650,right:650}}},children:kids}]});
Packer.toBuffer(doc).then(b=>{fs.writeFileSync(outPath,b);console.log('ok',outPath,b.length)});
