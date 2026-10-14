const fs=require('fs');
const {Document,Packer,Paragraph,TextRun,Table,TableRow,TableCell,WidthType,ShadingType,ImageRun,HeadingLevel,PageOrientation,BorderStyle}=require('docx');
const F='맑은 고딕';
const t=(s,o={})=>new TextRun({text:String(s),font:F,size:o.size||17,bold:o.bold,color:o.color});
const P=(s,o={})=>new Paragraph({children:Array.isArray(s)?s:[t(s,o)],spacing:{after:o.after??60}});
const H=(s,l)=>new Paragraph({heading:l,children:[new TextRun({text:s,font:F,bold:true,size:l===HeadingLevel.HEADING_1?28:22})],spacing:{before:200,after:120}});
const bd={style:BorderStyle.SINGLE,size:4,color:'A6B0C3'};const borders={top:bd,bottom:bd,left:bd,right:bd};
const cell=(s,w,o={})=>new TableCell({width:{size:w,type:WidthType.DXA},borders,shading:o.fill?{type:ShadingType.CLEAR,fill:o.fill,color:'auto'}:undefined,margins:{top:40,bottom:40,left:70,right:70},children:String(s).split('\n').map(x=>P(x,{size:o.size||16,bold:o.bold,color:o.color,after:30}))});
const log=JSON.parse(fs.readFileSync('log30_list.json','utf8'));
const W=[600,1300,5000,5600,2898];
const rows=[new TableRow({tableHeader:true,children:['#','슬라이드(옛 번호)','항목','v28 → v30','이유'].map((h,i)=>cell(h,W[i],{fill:'1F3864',bold:true,color:'FFFFFF'}))}),...log.map((r,i)=>new TableRow({cantSplit:true,children:[cell(i+1,W[0]),cell(r[0]+' '+r[1],W[1]),cell(r[2],W[2],{color:'7F7F7F',size:15}),cell(r[3],W[3]),cell(r[4],W[4],{size:15})]}))];
const imgs=fs.readdirSync('cmp30').filter(f=>f.endsWith('.png')).sort(); const imgParas=[];
for(const f of imgs){const buf=fs.readFileSync('cmp30/'+f);const [w0,h0]=require('child_process').execSync(`python3 -I -c "from PIL import Image;im=Image.open('cmp30/${f}');print(im.width,im.height)"`).toString().trim().split(' ').map(Number);let w=940,h=Math.round(h0*w/w0);const n=parseInt(f.slice(1,3));imgParas.push(P([t(`v30 슬라이드 ${n} — 왼쪽 v28(옛 번호), 오른쪽 v30`,{bold:true,size:18})],{after:40}));imgParas.push(new Paragraph({children:[new ImageRun({type:'png',data:buf,transformation:{width:w,height:h}})],spacing:{after:160}}));}
const rules=['이사님 10/14 — 「원래 것(v28)에 부족한 것을 조금 더 보태자」. v29 계열(장마다 논문 그림, 본 19장 재구성)은 개요·16종·연계 지점 맵·장단점·한계·결론이 빠져 더 어려워졌다는 판단으로 쓰지 않음. v28의 18장은 전부 그대로 두고 아래만 보탬.','보탠 것 ① 2장 한 장 — 백업 S-5 「두 표준에서 서로 비어 있는 곳」을 본 6번으로 복귀(장 표시 Ⅱ, 노트에 「가져온 방법·선행연구의 빈 자리」 한 문장 추가) ② 매핑(7)과 LMA 구성(8)의 순서를 바꾸고 매핑에 30개 조항·25개의 한 줄, LMA 장은 세 모델에 집중 ③ 2번 목차의 진행 안내 상자 삭제, 노트의 「다음 장에서 보겠습니다」류 전부 삭제.','검토자(1명) 지적 반영 — 장단점 장의 단점을 논문대로 7건으로 나누고 「넷은 보완·셋은 설계 제약」(5.2) · 「ML4 핵심 문항 모두 4점 이상」 · 「조항 25개(전조항 28개 가운데)」 · 「추가 항목(should 또는 체계화·지속 상태)」 · 「문서 심사 — 면담·실행 상태의 현장 관찰은 포함하지 않음」 · 「구조적 미비 세 가지와 공통 원인 두 가지에서 과제 다섯 건」 · 「중대한 환경측면을 결정하고」 · 「연계가 성립하지 않은 것과 성숙도가 낮은 것」 · 「LCA 연 1회 재수행 규정은 있으나 아직 이행 전」 · 2장 노트 「되풀이하여 관리하는 주기」(논문 고정 문구). 노트북 예상 질문 54~63 추가, Q37·Q39 답 정정.','본 19장 노트 합계 17분 50초(v28 16분 10초 + 2장 1분 10초 + 매핑 20초 + 개요 10초) — 「15분 내외」보다 길어 리허설에서 3번(1분 30초)·7번(1분 30초)·14번(1분 20초)에서 10초씩 줄일 후보. 규칙표 v18 0건, 전체 75장 그대로(S-5가 본으로 가서 S 블록은 11장).'];
const checks=['시간 17분 50초 — 줄일지 결정. 줄이면 3·7·14번 노트에서 한 문장씩.','3번을 [그림 1-2] 연구 절차 흐름도로 바꾸는 것은 보류(원래 [그림 1-4] 그대로). 연구 절차·연구 문제는 백업 S-7.','본심사판(30장)은 예비심사 뒤 v30 구성으로 다시 만든다.'];
const doc=new Document({styles:{default:{document:{run:{font:F,size:18}}}},sections:[{properties:{page:{size:{width:11906,height:16838,orientation:PageOrientation.LANDSCAPE},margin:{top:720,bottom:720,left:720,right:720}}},children:[
 new Paragraph({children:[new TextRun({text:'대조표 — 발표 자료 v28 → v30 (15분판 — 원래 구성에 보태기, 본 19장)',font:F,bold:true,size:36})],spacing:{after:80}}),
 P('발표자료_20261007_v28_15분판.pptx → v30_15분판.pptx (75장, 본 슬라이드 19장). 2026-10-14 세션 61. v29·v29a·v29b(논문 그림 중심 재구성)는 이사님 판단으로 폐기 — 「원래 것에 부족한 것을 보태는」 방식으로 v28에서 다시 만듦.',{size:18,after:120}),
 H('1. 원칙',HeadingLevel.HEADING_1),...rules.map(s=>P('· '+s,{size:17})),
 H('2. 바뀐 곳',HeadingLevel.HEADING_1),new Table({width:{size:15398,type:WidthType.DXA},columnWidths:W,rows}),
 
 H('3. 바뀐 장 비교 그림',HeadingLevel.HEADING_1),...imgParas,
 H('4. 이사님 확인 사항',HeadingLevel.HEADING_1),...checks.map(s=>P('· '+s,{size:17}))]}]});
Packer.toBuffer(doc).then(b=>{fs.writeFileSync('대조표_발표자료_v28_v30_20261014.docx',b);console.log('ok',log.length,imgs.length)});
