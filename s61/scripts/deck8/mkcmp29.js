const fs=require('fs');
const {Document,Packer,Paragraph,TextRun,Table,TableRow,TableCell,WidthType,ShadingType,ImageRun,HeadingLevel,PageOrientation,BorderStyle}=require('docx');
const F='맑은 고딕';
const t=(s,o={})=>new TextRun({text:String(s),font:F,size:o.size||17,bold:o.bold,color:o.color});
const P=(s,o={})=>new Paragraph({children:Array.isArray(s)?s:[t(s,o)],spacing:{after:o.after??60}});
const H=(s,l)=>new Paragraph({heading:l,children:[new TextRun({text:s,font:F,bold:true,size:l===HeadingLevel.HEADING_1?28:22})],spacing:{before:200,after:120}});
const bd={style:BorderStyle.SINGLE,size:4,color:'A6B0C3'};const borders={top:bd,bottom:bd,left:bd,right:bd};
const cell=(s,w,o={})=>new TableCell({width:{size:w,type:WidthType.DXA},borders,shading:o.fill?{type:ShadingType.CLEAR,fill:o.fill,color:'auto'}:undefined,margins:{top:40,bottom:40,left:70,right:70},children:String(s).split('\n').map(x=>P(x,{size:o.size||16,bold:o.bold,color:o.color,after:30}))});
const log=JSON.parse(fs.readFileSync('log29_list.json','utf8'));
const W=[600,1300,5000,5600,2898];
const rows=[new TableRow({tableHeader:true,children:['#','슬라이드(옛 번호)','항목','v28 → v29','이유'].map((h,i)=>cell(h,W[i],{fill:'1F3864',bold:true,color:'FFFFFF'}))}),...log.map((r,i)=>new TableRow({cantSplit:true,children:[cell(i+1,W[0]),cell(r[0]+' '+r[1],W[1]),cell(r[2],W[2],{color:'7F7F7F',size:15}),cell(r[3],W[3]),cell(r[4],W[4],{size:15})]}))];
const imgs=fs.readdirSync('cmp29').filter(f=>f.endsWith('.png')).sort(); const imgParas=[];
for(const f of imgs){const buf=fs.readFileSync('cmp29/'+f);const [w0,h0]=require('child_process').execSync(`python3 -I -c "from PIL import Image;im=Image.open('cmp29/${f}');print(im.width,im.height)"`).toString().trim().split(' ').map(Number);let w=940,h=Math.round(h0*w/w0);const n=parseInt(f.slice(1,3));imgParas.push(P([t(`v29 슬라이드 ${n} — 왼쪽 v28(옛 번호), 오른쪽 v29 — 새 장은 오른쪽만`,{bold:true,size:18})],{after:40}));imgParas.push(new Paragraph({children:[new ImageRun({type:'png',data:buf,transformation:{width:w,height:h}})],spacing:{after:160}}));}
const rules=['이사님 10/14 결정 — 「논문을 간략하게 리뷰한다」는 구성. 장마다 논문의 그림·표 하나 + 요지 한 줄, 숫자는 결과에 꼭 필요한 것만, 셈법·근거·예외는 노트 지시문·노트북·백업으로. 노트는 장마다 네 문장 안팎이고 「다음 장에서 보겠습니다」류 진행 안내는 모두 뺌.','바뀐 구성(본 19장) — 3 연구 절차와 논문 구성 [그림 1-2](장 대응 띠) · 4 연구 배경과 필요성(축소) · 6 두 표준의 연계 [그림 2-5](새) · 7 성숙도 평가의 방법과 선행연구 [그림 2-7](새) · 8 조항 매핑(옛 7, 30개 조항과 25개 셈·핵심 4 추가) · 9 LMA 구성(옛 6, 세 모델에 집중) · 10 평가 모델(D-1 예시 상자 → 지시문) · 15 개선 로드맵 [그림 4-10](새) · 16 기여와 장단점 [그림 5-2](새) · 17 한계와 향후 연구 과제 [그림 5-3](새) · 18 결론 — 네 가지 연구 문제의 답 [그림 5-1](새). 그대로인 장 — 1·2·5·11·12·13·14·19(노트만 손봄).','백업으로 내려간 다섯 장 — 연계 지점 맵(S-14)·16종(S-15)·개선 과제 표(S-25)·장점과 단점(S-27)·결론 글 상자(S-29). 삭제 없음, 전체 75장 → 81장. 본 19장 노트 합계 15분 50초. 규칙표 v18 0건, 재열기 통과.','노트 문안은 이사님 10/14 수정 반영 — 「본 논문은 다섯 장으로 구성되어 있고, 연구는 여섯 단계로 진행했습니다」 · 「~째 연구 문제에 해당됩니다」 · 6번은 논문 2.3.2·2.3.3 문장 그대로 · 「스물여덟 개 가운데 스물다섯 개」 · 「네 가지 질문·네 가지 등급」 · 9번 「조항 매핑의 결과로」 · 16번 첫 문장 삭제 · 18번 「네 가지 연구 문제의 답입니다」(꼬리 물기 없음).'];
const checks=['[그림 2-5]·[그림 5-1]·[그림 5-3]은 글이 많은 그림 — 화면에서는 상자 제목만 읽히므로 발표에서는 제목 줄만 짚는다. 배포용 PDF·유인물에서는 다 읽힘.','8번 매핑이 1분 30초로 가장 길다 — 30·25·네 가지 질문·네 가지 등급·핵심 넷의 순서를 입에 익힐 것. 리허설에서 길면 4번(40초)을 3번에 합칠 수 있음.','본심사판(30장)은 예비심사 지적을 받은 뒤 v29 구성으로 다시 만든다 — 지금은 v28_본심사판30장.pptx 그대로.'];
const doc=new Document({styles:{default:{document:{run:{font:F,size:18}}}},sections:[{properties:{page:{size:{width:11906,height:16838,orientation:PageOrientation.LANDSCAPE},margin:{top:720,bottom:720,left:720,right:720}}},children:[
 new Paragraph({children:[new TextRun({text:'대조표 — 발표 자료 v28 → v29 (15분판 — 논문 리뷰 구성, 본 19장)',font:F,bold:true,size:36})],spacing:{after:80}}),
 P('발표자료_20261007_v28_15분판.pptx → v29_15분판.pptx (81장, 본 슬라이드 19장 + 구분 1 + 백업 61). 2026-10-14 세션 61. 본심사판은 예비심사 뒤에 다시 만든다(이사님).',{size:18,after:120}),
 H('1. 원칙',HeadingLevel.HEADING_1),...rules.map(s=>P('· '+s,{size:17})),
 H('2. 바뀐 곳',HeadingLevel.HEADING_1),new Table({width:{size:15398,type:WidthType.DXA},columnWidths:W,rows}),
 
 H('3. 바뀐 장 비교 그림',HeadingLevel.HEADING_1),...imgParas,
 H('4. 이사님 확인 사항',HeadingLevel.HEADING_1),...checks.map(s=>P('· '+s,{size:17}))]}]});
Packer.toBuffer(doc).then(b=>{fs.writeFileSync('대조표_발표자료_v28_v29_20261014.docx',b);console.log('ok',log.length,imgs.length)});
