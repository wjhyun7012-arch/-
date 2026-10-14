// 발표 노트북 v3 — 대본 + 용어(부록 K) + 참고(두 단) + 물으면(위치·답·전문) + 백업 색인 + 장별·번호별 예상 질문 전문
const fs=require('fs');
const {Document,Packer,Paragraph,TextRun,ImageRun,Table,TableRow,TableCell,WidthType,BorderStyle,PageBreak,AlignmentType,ShadingType}=require('docx');
const [,,dataPath,cheatPath,outPath,TITLE,SUB]=process.argv;
const F='맑은 고딕';
const data=JSON.parse(fs.readFileSync(dataPath,'utf8')); const CH=JSON.parse(fs.readFileSync(cheatPath,'utf8'));
const t=(s,o={})=>new TextRun({text:String(s),font:F,size:o.size||28,bold:o.bold,color:o.color});
const P=(s,o={})=>new Paragraph({children:Array.isArray(s)?s:[t(s,o)],spacing:{after:o.after??120,line:o.line||300,before:o.before||0},alignment:o.align});
const fmt=s=>s==null?'—':(s>=60?`${Math.floor(s/60)}분${s%60?' '+s%60+'초':''}`:`${s}초`);
const last=data.filter(x=>x.img).slice(-1)[0]; const TOTAL=fmt(last.cum);
const bd={style:BorderStyle.SINGLE,size:4,color:'A6B0C3'};const borders={top:bd,bottom:bd,left:bd,right:bd};const none={style:BorderStyle.NONE};
const cell=(children,w,fill)=>new TableCell({borders,width:{size:w,type:WidthType.DXA},shading:fill?{type:ShadingType.CLEAR,fill,color:'auto'}:undefined,margins:{top:50,bottom:50,left:90,right:90},children});
const box1=(title,paras,fill)=>new Table({width:{size:10600,type:WidthType.DXA},columnWidths:[10600],rows:[new TableRow({children:[cell([P(title,{size:17,bold:true,color:'1F3864',after:30}),...paras],10600,fill)]})]});
const box2=(title,items,fill)=>{const L=[],R=[];items.forEach((it,i)=>(i%2?R:L).push(it));const mk=arr=>arr.map(x=>P(x,{size:14,after:30,line:205}));
  return new Table({width:{size:10600,type:WidthType.DXA},columnWidths:[5300,5300],rows:[
    new TableRow({children:[new TableCell({borders:{top:bd,left:bd,right:none,bottom:none},width:{size:5300,type:WidthType.DXA},shading:{type:ShadingType.CLEAR,fill,color:'auto'},margins:{top:50,bottom:0,left:90,right:90},children:[P(title,{size:17,bold:true,color:'1F3864',after:30})]}),new TableCell({borders:{top:bd,left:none,right:bd,bottom:none},width:{size:5300,type:WidthType.DXA},shading:{type:ShadingType.CLEAR,fill,color:'auto'},children:[P(' ',{size:8,after:0})]})]}),
    new TableRow({children:[new TableCell({borders:{top:none,left:bd,right:none,bottom:bd},width:{size:5300,type:WidthType.DXA},shading:{type:ShadingType.CLEAR,fill,color:'auto'},margins:{top:0,bottom:50,left:90,right:60},children:mk(L).length?mk(L):[P(' ',{size:8})]}),new TableCell({borders:{top:none,left:none,right:bd,bottom:bd},width:{size:5300,type:WidthType.DXA},shading:{type:ShadingType.CLEAR,fill,color:'auto'},margins:{top:0,bottom:50,left:60,right:90},children:mk(R).length?mk(R):[P(' ',{size:8})]})]})]});};
const kids=[];
kids.push(P(TITLE,{size:40,bold:true,after:200}));
kids.push(P(SUB,{size:20,color:'595959',after:160}));
kids.push(P('쓰는 법 — 한 쪽이 슬라이드 하나. 큰 글자가 읽는 대본, 회색 글이 지시문(읽지 않음). 「용어」는 논문 부록 K 용어집 문장 그대로(한글·영문·뜻·출처). 「참고」는 대본에 나온 숫자·표·그림이 논문 어디(절·표·그림 번호)에 있고 무엇인지, 그리고 뒤 몇 쪽에서 자세히 다루는지. 「물으면」은 논문 위치 → 답 요지 → 전문 번호·띄울 백업 번호. 뒤 장에서 다루는 질문이 오면 「n번 장에서 말씀드리겠습니다」 하고 넘어간다. 백업 색인과 장별·번호별 예상 질문 전문은 맨 뒤. 「짧게」 상자가 있는 장은 시간이 밀리면 그 상자만 읽고 넘어간다(심사위원은 배포용 PDF로 전부 본다). 그 장의 대본에서 굵은 문장이 꼭 말할 것.',{size:19,after:200,line:280}));
const rows=[new TableRow({tableHeader:true,children:['장','제목','목표','누적'].map((h,i)=>new TableCell({borders,width:{size:[700,7500,1200,1200][i],type:WidthType.DXA},shading:{type:ShadingType.CLEAR,fill:'1F3864',color:'auto'},children:[P(h,{size:18,bold:true,color:'FFFFFF',after:0})]}))})];
for(const d of data.filter(x=>x.img)) rows.push(new TableRow({children:[[d.n,700],[d.title,7500],[fmt(d.secs),1200],[fmt(d.cum),1200]].map(([v,w])=>new TableCell({borders,width:{size:w,type:WidthType.DXA},margins:{top:20,bottom:20,left:60,right:60},children:[P(v,{size:17,after:0})]}))}));
kids.push(new Table({width:{size:10600,type:WidthType.DXA},columnWidths:[700,7500,1200,1200],rows}));
kids.push(new Paragraph({children:[new PageBreak()]}));
for(const d of data.filter(x=>x.img)){
  const c=CH.cheat[String(d.n)]||{terms:[],qs:[],ref:[]};
  kids.push(new Paragraph({children:[t(' ',{size:2})],pageBreakBefore:true,spacing:{after:0,line:120}}));
  kids.push(new Table({width:{size:10600,type:WidthType.DXA},columnWidths:[7600,3000],rows:[new TableRow({children:[
    new TableCell({borders:{top:none,bottom:{style:BorderStyle.SINGLE,size:12,color:'1F3864'},left:none,right:none},width:{size:7600,type:WidthType.DXA},children:[P([t(`${d.n}  `,{size:44,bold:true,color:'C00000'}),t(d.title,{size:26,bold:true,color:'1F3864'})],{after:40}),P(d.sec||' ',{size:17,color:'7F7F7F',after:40})]}),
    new TableCell({borders:{top:none,bottom:{style:BorderStyle.SINGLE,size:12,color:'1F3864'},left:none,right:none},width:{size:3000,type:WidthType.DXA},children:[P(`목표 ${fmt(d.secs)}`,{size:20,bold:true,align:AlignmentType.RIGHT,after:0}),P(`누적 ${fmt(d.cum)} / ${TOTAL}`,{size:17,color:'7F7F7F',align:AlignmentType.RIGHT,after:40})]})]})]}));
  const buf=fs.readFileSync(d.img); const w=(c.thumb||260)-20,h=Math.round(w*210/297);
  kids.push(new Paragraph({children:[new ImageRun({type:'png',data:buf,transformation:{width:w,height:h}})],spacing:{before:80,after:100}}));
  const BOLD=/\d|ML|핵심|매핑|LMA|연계 지점|열여섯|쉰다섯|서른|스물|네 가지|넷/;
  for(const s of d.sents) kids.push(P(s,{size:c.small?24:26,after:c.small?60:90,line:c.small?270:300,bold:(c.bold&&BOLD.test(s))?true:undefined}));
  if(c.short){kids.push(P(' ',{size:6,after:0}));kids.push(box1('짧게 — 시간이 밀리면 이것만 읽고 넘어간다',[P(c.short,{size:20,bold:true,after:20,line:260})],'FDECEC'));}
  for(const cu of d.cues) kids.push(P('〔'+cu+'〕',{size:16,color:'7F7F7F',after:30,line:240}));
  if(c.terms.length){kids.push(P(' ',{size:6,after:0}));kids.push(box1('용어 — 논문 부록 K 그대로',c.terms.map(x=>P([t(x.ko,{size:16,bold:true}),t(x.en?`(${x.en})`:'',{size:15,color:'595959'}),t(' — '+x.def,{size:15}),t('  · '+x.src,{size:13,color:'7F7F7F'})],{after:30,line:215})),'F2F4F8'));}
  if(c.ref&&c.ref.length){kids.push(P(' ',{size:6,after:0}));kids.push(box2('참고 — 논문 어디에, 무엇이 (→ 뒤 쪽)',c.ref,'EEF6EE'));}
  if(c.qs.length){kids.push(P(' ',{size:6,after:0}));kids.push(box1('물으면 — 논문 위치 → 답 요지 → 전문·백업',c.qs.map(x=>P([t('Q. '+x.q+'  ',{size:16,bold:true}),t('→ 논문 '+x.loc+'. ',{size:15,bold:true,color:'1F3864'}),t(x.a,{size:15}),t('  · 전문 '+x.no+'번'+(x.ref?' · '+x.ref:''),{size:13,color:'7F7F7F'})],{after:40,line:215})),'FFF8E1'));}
}
kids.push(new Paragraph({children:[t(' ',{size:2})],pageBreakBefore:true,spacing:{after:0,line:120}}));
kids.push(P('백업 색인 — 질문이 오면 이 번호를 띄운다',{size:30,bold:true,color:'1F3864',after:160}));
const brows=[new TableRow({tableHeader:true,children:['장','백업 번호','제목'].map((h,i)=>new TableCell({borders,width:{size:[700,1700,8200][i],type:WidthType.DXA},shading:{type:ShadingType.CLEAR,fill:'1F3864',color:'auto'},children:[P(h,{size:18,bold:true,color:'FFFFFF',after:0})]}))})];
for(const d of data.filter(x=>!x.img)) brows.push(new TableRow({children:[[d.n,700],[(d.label||'').replace(/^백업\s*/,''),1700],[d.title,8200]].map(([v,w])=>new TableCell({borders,width:{size:w,type:WidthType.DXA},margins:{top:20,bottom:20,left:60,right:60},children:[P(v,{size:17,after:0})]}))}));
kids.push(new Table({width:{size:10600,type:WidthType.DXA},columnWidths:[700,1700,8200],rows:brows}));
kids.push(new Paragraph({children:[new PageBreak()]}));
// 예상 질문 전문 — 장별·번호별 (첫 배정 장에 전문, 다른 장엔 참조)
kids.push(P('예상 질문 전문 — 장별 · 번호별 (53문)',{size:30,bold:true,color:'1F3864',after:60}));
kids.push(P('쪽의 「물으면」에 적힌 「전문 n번」을 그 장 묶음에서 찾는다. 여러 장에 걸친 질문은 처음 나오는 장에 전문, 다른 장에는 「→ 참조」.',{size:16,color:'595959',after:160}));
const first={}; const order=Object.keys(CH.M).map(Number).sort((a,b)=>a-b);
for(const n of order) for(const q of CH.M[String(n)]) if(first[q]==null) first[q]=n;
const titleOf=n=>(data.find(x=>x.n===n)||{}).title||'';
for(const n of order){const arr=[...CH.M[String(n)]].sort((a,b)=>a-b); if(!arr.length) continue;
  kids.push(P([t(`장 ${n} — `,{size:22,bold:true,color:'C00000'}),t(titleOf(n),{size:22,bold:true,color:'1F3864'})],{before:200,after:80}));
  for(const no of arr){const q=CH.Q[String(no)];
    if(first[no]!==n){kids.push(P([t(`${no}. `,{size:17,bold:true}),t(q.q.replace('★ ',''),{size:17}),t(`  → 장 ${first[no]} 참조`,{size:15,color:'7F7F7F'})],{after:40,line:240}));continue;}
    kids.push(P([t(`${no}. `,{size:19,bold:true,color:'C00000'}),t(q.q.replace('★ ',''),{size:19,bold:true})],{after:30,before:80,line:260}));
    kids.push(P(q.a,{size:17,after:20,line:250}));
    kids.push(P('논문 '+q.loc,{size:14,color:'7F7F7F',after:60}));}}
const doc=new Document({styles:{default:{document:{run:{font:F,size:24}}}},sections:[{properties:{page:{size:{width:11906,height:16838},margin:{top:600,bottom:600,left:650,right:650}}},children:kids}]});
Packer.toBuffer(doc).then(b=>{fs.writeFileSync(outPath,b);console.log('ok',outPath,b.length)});
