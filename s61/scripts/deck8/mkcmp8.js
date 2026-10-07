// 대조표 — 발표 자료 v7 → v8 (A4). 사용: node mkcmp8.js   (log8.json, v8b_fitlog.json, cmp/*.png 필요)
const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType, ImageRun, HeadingLevel, PageOrientation, BorderStyle } = require('docx');
const F = '맑은 고딕';
const t = (s, o = {}) => new TextRun({ text: String(s), font: F, size: o.size || 17, bold: o.bold, color: o.color });
const P = (s, o = {}) => new Paragraph({ children: Array.isArray(s) ? s : [t(s, o)], spacing: { after: o.after ?? 60 } });
const H = (s, l) => new Paragraph({ heading: l, children: [new TextRun({ text: s, font: F, bold: true, size: l === HeadingLevel.HEADING_1 ? 28 : 22 })], spacing: { before: 200, after: 120 } });
const bd = { style: BorderStyle.SINGLE, size: 4, color: 'A6B0C3' }; const borders = { top: bd, bottom: bd, left: bd, right: bd };
const cell = (s, w, o = {}) => new TableCell({ width: { size: w, type: WidthType.DXA }, borders, shading: o.fill ? { type: ShadingType.CLEAR, fill: o.fill, color: 'auto' } : undefined, margins: { top: 40, bottom: 40, left: 70, right: 70 }, children: String(s).split('\n').map(x => P(x, { size: o.size || 16, bold: o.bold, color: o.color, after: 0 })) });
const mkTable = (head, rows, W) => new Table({ width: { size: W.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: W, rows: [new TableRow({ tableHeader: true, children: head.map((h, i) => cell(h, W[i], { fill: '1F3864', bold: true, color: 'FFFFFF' })) }), ...rows.map(r => new TableRow({ cantSplit: true, children: r.map((c, i) => cell(c, W[i], i === 2 ? { color: '7F7F7F' } : i === 4 ? { size: 15 } : {})) }))] });
const img = (file, maxW = 900, maxH = 560) => { const buf = fs.readFileSync(file); const [w0, h0] = require('child_process').execSync(`python3 -I -c "from PIL import Image;im=Image.open('${file}');print(im.width,im.height)"`).toString().trim().split(' ').map(Number); let w = Math.min(w0, maxW), h = Math.round(h0 * w / w0); if (h > maxH) { h = maxH; w = Math.round(w0 * h / h0); } return new Paragraph({ children: [new ImageRun({ type: 'png', data: buf, transformation: { width: w, height: h } })], spacing: { after: 160 } }); };

const log1 = JSON.parse(fs.readFileSync('log8.json', 'utf8'));      // 내용 수정 + 삭제 + A4
const log2 = JSON.parse(fs.readFileSync('v8b_fitlog.json', 'utf8')); // 넘침 조정
const W = [520, 760, 1300, 4300, 4700, 3818];
const rowsA = log1.map((r, i) => [i + 1, r[0], r[1], String(r[2]).replace(/\.(png|jpg)$/, ' (그림)'), String(r[3]).replace(/^.*\/(s3_6\.png)$/, '그림 글자 고침 (2절)'), r[4]]);
const titleRows = log2.filter(r => r[1] === '제목 글꼴'); const otherRows = log2.filter(r => r[1] !== '제목 글꼴');
const rowsB = otherRows.map((r, i) => [i + 1, r[0], r[1], r[2], r[3], r[4] || '']);
const rowsT = titleRows.map((r, i) => [i + 1, r[0], r[4], r[2], r[3], '한 줄에 들어가도록(가로가 좁아짐)']);

// 쉽게 쓰기 제안 (승인 대기) — 인수인계 4-10
const prop = [
  ['10', '참조 모델 상자 오른쪽', 'ISO/IEC 33004 5.2·5.4 (형식의 요건)\nKS X ISO 8000-61 5.1·5.3 (기술 형식 · PDCA 편성)\n김선호·이창수(2013) 참조 모델의 원형', '연계 업무 목록 — 「무엇을 해야 하는가」를 업무 16종으로 적은 것\n형식은 국제표준(ISO/IEC 33004 · KS X ISO 8000-61)대로, 원형은 김선호·이창수(2013)', '조항 번호는 백업 B-2 [그림 1-3]·[표 1-1]에 이미 있음'],
  ['10', '평가 모델 상자 오른쪽', 'ISO/IEC 33004 6.1·6.3.2 (참조 모델 + 측정 틀)\nKS X ISO 8000-62 4.3·4.7·4.9 (평가 산출물)\nISO/IEC 33020 5.3·PA (N·P·L·F 부분 충족) · 33001 용어', '채점표 — 업무마다 「되어 있는가」를 55문항으로 묻고, 증거를 보고 점수를 매기는 기준\n점수 기준은 ISO/IEC 33020의 4등급 대신 5점 상태 기술(BARS)', '같음'],
  ['10', '성숙도 모델 상자 오른쪽', 'ISO/IEC 33004 7.1·7.3.5·7.3.6 (수준의 요건)\nKS X ISO 8000-62 4.5.1·부속서 C (최소충족 판정)\nISO/IEC 33003 3.7·3.12 (형성형 · 비보상) · 33020 5.5', '등급 규칙 — 점수를 ML0~ML5 여섯 등급으로 바꾸는 규칙\n핵심 문항 4종은 필수 — 하나라도 빠지면 평균이 높아도 ML3 이상으로 못 감(KS X ISO 8000-62의 최소충족 원리)', '같음'],
  ['10', '아래 띠', '평가의 수행 — ISO/IEC 33002 (평가 클래스) · 운영 준거 — KS Q ISO 19011 · ISO/IEC 17024 · KS Q ISO 10015 · KS I ISO 14004 · 문장 판별 — KS A 0001 · 「핵심 문항」 용어 — KS Q 8001·8002', '평가 수행·평가자 자격·개선 계획의 운영 준거도 모두 표준에서 — 전체 목록은 백업 B-2 [그림 1-3]', '목록 자체는 B-2에 있음'],
  ['12', '▶ 한 줄', '참조·평가·성숙도의 세 계층으로 만들었다.', '연계 업무 목록(참조 모델) · 채점표(평가 모델) · 등급 규칙(성숙도 모델)의 세 계층으로 만들었다.', '용어 첫 등장 때 풀어 쓰기'],
  ['13', '왼쪽 Q2', 'Q2 LCA 산출물 없이는 shall 이행이 불완전한가?', 'Q2 LCA 결과가 없으면 그 조항의 의무(shall)를 다 지킬 수 없는가?', '풀어 쓰기'],
  ['13', '아래 각주 첫 문단', 'ISO 14001 30개 조항 단위(7.4.1~7.4.3은 7.4로 통합)의 전 조항 판정 — 조항별 판정 전문은 부록 A·B.1. 판정 질문의 적용 흐름은 [그림 3-3](백업 B-1).', '(삭제 — 백업 B-1·부록 A·B.1로)', '둘째 문단(55문항 중 44·11)만 남김'],
  ['17', '표 「설계 근거」 열', '연계 결정 요인의 활동 부재 = 구성개념 불성립 / 필요조건 활동의 수행 증거 (PA.1.1 상당) / 과반이 문서화된 정보로 관리 (PA.2.2 상당) / 연계 절차 전부 shall 충족 기준 정의 (PA.3.1 — 8000-62 4.5.1 최소충족) / 핵심 문항 통제 수준 + 성과 측정 전면 충족 (PA.4 상당) / shall 전 항목 충족 + 개선 체계 통제 수준 (PA.5 상당)', '열 이름 「뜻」 — 연계 활동이 없음 / 네 핵심이 있기는 함 / 나머지 절반 이상이 문서로 관리됨 / 네 핵심이 모두 갖추어짐 — 체계가 섬 / 핵심이 통제되고 점검 영역이 모두 갖추어짐 — 숫자로 확인 / 전 문항이 갖추어지고 개선까지 돎', 'PA 번호·8000-62 4.5.1은 백업 B-11 수준별 기준 전문에 있음'],
  ['17', '아래 각주', '최소충족 원리는 KS X ISO 8000-62 4.5.1에서 가져오고 적용 범위만 네 문항으로 한정 — 그 밖의 51문항은 BARS 5점으로 종합점수에 반영. 비교주장 조건부 shall(C-7·A-9)은 3.4.3의 기준 — 비교주장 미수행 조직은 최저점(1점)·N/A 민감도 병기, 미수행 정책을 문서화하면 3점.', '비보상 = 필수 항목이 빠지면 평균이 높아도 안 올라감 — 이 원리만 KS X ISO 8000-62에서 가져오고 적용 범위는 네 문항으로 한정. 비교주장 조건부 문항(C-7·A-9)의 처리는 백업 B-Q5.', '풀어 쓰기 + 세부는 백업'],
  ['18', '표 「따른 표준(구조·방법)」 열', 'ISO/IEC 33004 5절 · KS X ISO 8000-61 5.3 / ISO/IEC 33004 6절 · ISO/IEC 33020(측정 프레임워크 — N–P–L–F 대신 BARS 5점 척도) / ISO/IEC 33004 7절 · KS X ISO 8000-62 4.5.1 · ISO/IEC 33003', 'ISO/IEC 33004 · KS X ISO 8000-61 / ISO/IEC 33004 · ISO/IEC 33020(BARS 5점) / ISO/IEC 33004 · KS X ISO 8000-62 · ISO/IEC 33003', '절 번호만 뺌(표가 짧아져 14pt에서도 들어감)'],
];
const WP = [520, 1700, 5200, 5200, 2778];

const kept = [
  '시간: 본 슬라이드 29장·노트 합 약 32분 30초 그대로(이사님 10/7: 「교수님이 말씀하시면 그때 빼자」). 20분판 생략 순서 후보는 인수인계서에.',
  '백업 31~70번: 본문과 겹쳐 보이는 장(B-13 강점·미비, B-15 [표 5-1] 전문, B-0 장별 요약, B-Q1)도 그대로 둠 — 질문 때 꺼낼 전문이므로(이사님 10/7).',
  '그림 장(11·12·23번): 그림을 유지하고 A4 세로 여유만큼 키움(이사님 10/7: 「그림으로 보여주는 게 더 효과적」). 16번 그림은 오른쪽 표와 나란하므로 그대로.',
  '30번 백업 구분 슬라이드: 「B-Q표 43문 전문」만 뺌. 인덱스(번호·제목 목록)로 바꾸는 것은 제안만.',
  '10·13·17·18번 조항 번호·용어: 손대지 않음 — 아래 4절 제안표(승인 대기).',
  '3번 제목은 권장 (가) 「논문의 구성과 장 간 연결」로. (나) 「— 다섯 장과 부록」은 쓰지 않음.',
  '27번 한 장 요약에 뼈대 한 줄(16종 → 55문항 → 핵심 4종 → ML0~ML5 / 2.85 → ML2 → 3.15 → ML3)을 넣을지는 미정 그대로(2번에서 뺀 것은 노트로).',
  '발표 노트 시간 표기·대본화(인수인계 4절의 (4))는 이번에 하지 않음 — A4 확정 뒤.',
];
const checks = [
  'PowerPoint(맑은 고딕)에서 열어 표 칸 넘침을 확인 — 13·17·18·20번(14pt로 키운 표). 이 작업 환경 렌더는 Noto Sans CJK 대체 글꼴.',
  '화면 비율: 강의실 프로젝터가 16:9이면 A4 가로는 양옆에 검은 띠가 생기고 글자가 약 20% 작게 보임. 발표 전 강의실에서 한 번 띄워 볼 것.',
  '26번 장점·단점 풀이(짧은 판) 문안이 괜찮은지.',
  '2번 목차 글자 28pt — 더 키울지.',
  '예상 질문 v5의 ★44~48 답변(정식 문장) 확인.',
];
const imgs = [['s3_fig', '3번 그림 — 머리말·3장 상자·4장 상자 글자(위 v7, 아래 v8)'], ['s01', '1번 표지(왼쪽 v7 16:9, 오른쪽 v8 A4)'], ['s02', '2번 목차'], ['s03', '3번 논문의 구성과 장 간 연결'], ['s04', '4번 — 물음 줄 삭제 · 오른쪽 위 장 표시(Ⅰ~Ⅴ·맺음) 14pt 남색 상자'], ['s12', '12번 — 물음 줄 삭제·그림 확대'], ['s13', '13번 — 표 14pt·행 높이'], ['s18', '18번 — 표 14pt·제목 축소'], ['s20', '20번 — 표 14pt·세로 압축'], ['s26', '26번 — 장점·단점 쉬운 말(짧은 판)'], ['s30', '30번 — 백업 구분']];
const imgParas = [];
for (const [f, cap] of imgs) { if (!fs.existsSync(`cmp/${f}.png`)) continue; imgParas.push(P([t(cap, { bold: true, size: 18 })], { after: 40 })); imgParas.push(img(`cmp/${f}.png`, 940, 600)); }

const doc = new Document({ styles: { default: { document: { run: { font: F, size: 18 } } } }, sections: [{ properties: { page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }, margin: { top: 720, bottom: 720, left: 720, right: 720 } } }, children: [
  new Paragraph({ children: [new TextRun({ text: '대조표 — 발표 자료 v7 → v8 (A4)', font: F, bold: true, size: 36 })], spacing: { after: 80 } }),
  P('발표자료_20261007_v7.pptx(80장, 16:9, 논문 v203 기준) → 발표자료_20261007_v8.pptx(70장, A4 가로 297×210mm, 논문 v203 기준). 2026-10-07 세션 61.', { size: 18 }),
  P('순서: (1) 인수인계 4절의 내용 수정 → (2) 백업 71~80 삭제 → (3) A4 가로 변환 → (4) 넘침 조정(제목 글꼴 자동 축소, 표 글자 14pt·행 높이, 그림 확대, 세로 압축). 이사님 10/7 결정: 시간은 그대로, 표 글자 14pt, 26번은 항목 전부 두되 짧게, 백업은 그대로, 그림 장은 그림으로.', { size: 18, after: 160 }),
  H('1. 바뀐 곳 — 내용·구성·크기', HeadingLevel.HEADING_1),
  mkTable(['#', '슬라이드', '위치', 'v7 (10/7 오전)', 'v8 (10/7)', '근거·이유'], rowsA, W),
  H('2. 그림·화면 전후', HeadingLevel.HEADING_1), ...imgParas,
  H('3. 넘침 조정 (A4에서 생긴 것)', HeadingLevel.HEADING_1),
  P('3.1 표 글자·행 높이·그림·세로 압축', { bold: true, size: 18 }),
  mkTable(['#', '슬라이드', '항목', '전', '후', '비고'], rowsB, W),
  P('3.2 제목 글꼴 자동 축소 — 27pt 기준, 오른쪽 위 장 표시와 겹치지 않고 한 줄에 들어가는 크기(최소 20pt)', { bold: true, size: 18 }),
  mkTable(['#', '슬라이드', '제목', '전', '후', '비고'], rowsT, W),
  H('4. 승인 대기 — 본 슬라이드 쉽게 쓰기 제안 (10·12·13·17·18번)', HeadingLevel.HEADING_1),
  P('인수인계 4-10: 「슬라이드별 지금 → 바꿀 문안 표를 먼저 드리고 승인 뒤 반영」. 외부 위원(경영학과·컴퓨터공학과) 기준으로 조항 번호는 백업으로 내리고 용어는 첫 등장 때 풀어 씀. v8에는 반영하지 않았음.', { size: 17 }),
  mkTable(['슬라이드', '위치', '지금(v8)', '바꿀 문안(안)', '비고'], prop, WP),
  H('5. 확인했으나 바꾸지 않은 것', HeadingLevel.HEADING_1), ...kept.map(s => P('· ' + s, { size: 17 })),
  H('6. 이사님 확인 사항', HeadingLevel.HEADING_1), ...checks.map(s => P('· ' + s, { size: 17 })),
] }] });
Packer.toBuffer(doc).then(b => { fs.writeFileSync('대조표_발표자료_v7_v8_20261007.docx', b); console.log('ok rows', rowsA.length, rowsB.length, rowsT.length); });
