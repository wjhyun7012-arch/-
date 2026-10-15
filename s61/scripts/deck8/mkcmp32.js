const fs=require('fs');
const {Document,Packer,Paragraph,TextRun,Table,TableRow,TableCell,WidthType,ShadingType,ImageRun,HeadingLevel,PageOrientation,BorderStyle}=require('docx');
const F='맑은 고딕';
const t=(s,o={})=>new TextRun({text:String(s),font:F,size:o.size||17,bold:o.bold,color:o.color});
const P=(s,o={})=>new Paragraph({children:Array.isArray(s)?s:[t(s,o)],spacing:{after:o.after??60}});
const H=(s,l)=>new Paragraph({heading:l,children:[new TextRun({text:s,font:F,bold:true,size:l===HeadingLevel.HEADING_1?28:22})],spacing:{before:200,after:120}});
const bd={style:BorderStyle.SINGLE,size:4,color:'A6B0C3'};const borders={top:bd,bottom:bd,left:bd,right:bd};
const cell=(s,w,o={})=>new TableCell({width:{size:w,type:WidthType.DXA},borders,shading:o.fill?{type:ShadingType.CLEAR,fill:o.fill,color:'auto'}:undefined,margins:{top:40,bottom:40,left:70,right:70},children:String(s).split('\n').map(x=>P(x,{size:o.size||16,bold:o.bold,color:o.color,after:30}))});
const log=JSON.parse(fs.readFileSync('log32_list.json','utf8'));
const W=[600,1300,5000,5600,2898];
const rows=[new TableRow({tableHeader:true,children:['#','슬라이드(옛 번호)','항목','v30 → v32','이유'].map((h,i)=>cell(h,W[i],{fill:'1F3864',bold:true,color:'FFFFFF'}))}),...log.map((r,i)=>new TableRow({cantSplit:true,children:[cell(i+1,W[0]),cell(r[0]+' '+r[1],W[1]),cell(r[2],W[2],{color:'7F7F7F',size:15}),cell(r[3],W[3]),cell(r[4],W[4],{size:15})]}))];
const imgs=fs.readdirSync('cmp32').filter(f=>f.endsWith('.png')).sort(); const imgParas=[];
for(const f of imgs){const buf=fs.readFileSync('cmp32/'+f);const [w0,h0]=require('child_process').execSync(`python3 -I -c "from PIL import Image;im=Image.open('cmp32/${f}');print(im.width,im.height)"`).toString().trim().split(' ').map(Number);let w=940,h=Math.round(h0*w/w0);const n=parseInt(f.slice(1,3));imgParas.push(P([t(`v32 슬라이드 ${n} — 왼쪽 v30(옛 번호), 오른쪽 v32 — 새 장은 오른쪽만`,{bold:true,size:18})],{after:40}));imgParas.push(new Paragraph({children:[new ImageRun({type:'png',data:buf,transformation:{width:w,height:h}})],spacing:{after:160}}));}
const rules=['구성(본 21장) — 1 표지(영문 제목 두 줄) · 2 목차(진행 안내 상자 삭제) · 3 논문 구성 · 4 개요(「이것이 이 연구가 만든 것이다」 아랫줄) · 5 ISO 14001은 어떤 표준인가 [그림 2-2](새, 상위 문서 구조(HLS), 정식 명칭) · 6 LCA는 어떤 표준인가 [그림 2-4](S-9 복귀, 정식 명칭, 4.1·5절·6절) · 7 두 상태 · 8 두 표준의 빈 곳 · 9 따른 표준(S-10 복귀, 행동기술척도(BARS)) · 10 두 표준의 조항 매핑(쉬운 장 유지, 「두세 번 물으면」) · 11 연계 지점 맵 · 12 LMA 구성(상자 두 줄) · 13 16종 · 14 55문항 · 15 판정 규칙(A사 라벨 ML2 줄, 쉬운 노트) · 16 A사 평가(문서 접수, 상향·하향) · 17 영역별 평점과 핵심 문항의 판정 [그림 4-1](새) · 18 개선 과제 · 19 장단점·한계 · 20 결론 — 네 가지 연구 문제의 답 [그림 5-1](새) · 21 감사합니다 + Q & A.','백업으로 간 것 — 판정 경로 상자 장 S-23, 연계 상태 [그림 4-4] S-24, 결론 글 상자 S-29, [그림 3-3] 판정 질문 흐름 B-1, 새 B-1b 「조항별 판정 질문 경로 — 부록 A.1 등급에서 재구성」(높음 12개의 (a)·(b) 나눔은 이사님 확인). 삭제 없음, 75 → 79장.','노트 — 「다음 장에서 보겠습니다」류 진행 안내 전부 삭제, 15·17번 쉬운 말, 6번 LCA 정식 명칭과 4.1·5·6절, 10번 매핑 세 문장(「조항마다 두 번 또는 세 번 물으면 등급」), 19번 ML4·ML5는 기록이 없어 미확인, 20번 네 답. 합계 19분 15초 — 시간은 노트북 「짧게」 상자로 조절(이사님 「시간은 신경 쓰지 말자, 구성이 중요」).','검토 과정에서 폐기한 것 — v29·v29a·v29b(장마다 논문 그림, 본 19장 재구성): 「원래 것보다 못하다」 → v30에서 v28 기준으로 되돌림. v32에서도 [그림 3-3]을 10번에 넣었다가 「매핑이 원래 이렇게 어려웠나」 → 쉬운 장으로 복귀. 규칙표 v18 0건(발표자료·노트북 전체 — 노트북 전체에서 걸린 9건은 부록 K 인용·범주어·8절 운용의 예외).'];
const checks=['백업 B-1b 높음 12개의 경로 나눔 — (a) Q1 예→Q2 아니오: 6.1.1·6.1.4·6.2.2·9.3·10.1·10.3 / (b) Q1 아니오→Q3 예: 4.2·4.3·5.2·7.2·7.4·7.5.1. A.1 근거 문구로 가른 초안이므로 바꿀 조항이 있으면 지정.','v204(본심사용 논문) — [그림 1-4] 화살표(3장 → 5.1, 4장 → 5.2), 부록 A.1에 「판정 경로(Q1~Q4 응답)」 열, [그림 3-7]·부록 C P-9 카드의 환경측면 용어를 「결정」으로 통일, Keywords 괄호, 인준지 이름, (선택) 생성형 AI 한 문장.','본심사판(30장)은 예비심사 뒤 v32 구성으로 다시 만든다.'];
const doc=new Document({styles:{default:{document:{run:{font:F,size:18}}}},sections:[{properties:{page:{size:{width:11906,height:16838,orientation:PageOrientation.LANDSCAPE},margin:{top:720,bottom:720,left:720,right:720}}},children:[
 new Paragraph({children:[new TextRun({text:'대조표 — 발표 자료 v30 → v32 (15분판 — 이사님 장별 검토, 본 21장)',font:F,bold:true,size:36})],spacing:{after:80}}),
 P('발표자료_20261007_v30_15분판.pptx → v32_15분판.pptx (79장, 본 슬라이드 21장). 2026-10-14~15 세션 61. 이사님이 장별로 보시며 지적한 것을 v31·v31b·v32·v32b·v32c·v32d 여섯 단계로 반영 — 표의 슬라이드 번호는 각 단계 당시 번호.',{size:18,after:120}),
 H('1. 원칙',HeadingLevel.HEADING_1),...rules.map(s=>P('· '+s,{size:17})),
 H('2. 바뀐 곳',HeadingLevel.HEADING_1),new Table({width:{size:15398,type:WidthType.DXA},columnWidths:W,rows}),
 
 H('3. 바뀐 장 비교 그림',HeadingLevel.HEADING_1),...imgParas,
 H('4. 이사님 확인 사항',HeadingLevel.HEADING_1),...checks.map(s=>P('· '+s,{size:17}))]}]});
Packer.toBuffer(doc).then(b=>{fs.writeFileSync('대조표_발표자료_v30_v32_20261015.docx',b);console.log('ok',log.length,imgs.length)});
