// 대조표 — 발표 노트 v8 → v9 (존댓말 대본). 사용: node mkcmp9.js  (log9.json 필요)
const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType, HeadingLevel, PageOrientation, BorderStyle } = require('docx');
const F = '맑은 고딕';
const t = (s, o = {}) => new TextRun({ text: String(s), font: F, size: o.size || 17, bold: o.bold, color: o.color });
const P = (s, o = {}) => new Paragraph({ children: Array.isArray(s) ? s : [t(s, o)], spacing: { after: o.after ?? 60 } });
const H = (s, l) => new Paragraph({ heading: l, children: [new TextRun({ text: s, font: F, bold: true, size: l === HeadingLevel.HEADING_1 ? 28 : 22 })], spacing: { before: 200, after: 120 } });
const bd = { style: BorderStyle.SINGLE, size: 4, color: 'A6B0C3' }; const borders = { top: bd, bottom: bd, left: bd, right: bd };
const cell = (s, w, o = {}) => new TableCell({ width: { size: w, type: WidthType.DXA }, borders, shading: o.fill ? { type: ShadingType.CLEAR, fill: o.fill, color: 'auto' } : undefined, margins: { top: 40, bottom: 40, left: 70, right: 70 }, children: String(s).split('\n').map(x => P(x, { size: o.size || 16, bold: o.bold, color: o.color, after: 40 })) });
const log = JSON.parse(fs.readFileSync('log9.json', 'utf8'));
const W = [700, 7100, 7598];
const rows = [new TableRow({ tableHeader: true, children: ['장', 'v8 노트 (메모 투)', 'v9 대본 (읽는 문장 + 〔메모〕)'].map((h, i) => cell(h, W[i], { fill: '1F3864', bold: true, color: 'FFFFFF' })) }),
  ...log.map(([n, o, nw]) => new TableRow({ children: [cell(n, W[0]), cell(o, W[1], { color: '7F7F7F', size: 15 }), cell(nw, W[2])] }))];
const rules = [
  '읽는 문장은 논문 v203 본문 문장을 가져와 「~하였다 → ~하였습니다」, 「~이다 → ~입니다」로만 바꿈. 숫자·용어·판정 문구는 논문 그대로.',
  '〔 〕 안은 발표자 메모 — 비유, 백업 번호, 「물음이 나오면」 대비, 시간. 읽지 않음.',
  '화면 글은 손대지 않음(논문과 같은 글말 「~다」). 청중에게 직접 말하는 줄(2번 아래 「…두었습니다」)만 존댓말 — 이미 그러함.',
  '1번에 인사·소속·제목, 28번에 「이상으로 발표를 마치겠습니다. 감사합니다.」를 넣음. 29번(감사합니다)은 질의응답 메모만.',
  '시간 표기는 v8 그대로(합 약 32분 30초). 교수님이 발표 시간을 말씀하시면 그때 줄임.',
];
const checks = ['17번 비유(건강검진)는 메모로만 — 쓸지 결정.', '26번 단점 다섯째(환경성과와의 관계 미검증)를 읽는 문장에 넣었음 — 빼도 됨.', '10번은 조항 번호를 읽지 않도록 썼음(화면에는 있음) — 쉽게 쓰기 제안표(대조표 v7→v8 4절)와 함께 결정.'];
const doc = new Document({ styles: { default: { document: { run: { font: F, size: 18 } } } }, sections: [{ properties: { page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }, margin: { top: 720, bottom: 720, left: 720, right: 720 } } }, children: [
  new Paragraph({ children: [new TextRun({ text: '대조표 — 발표 노트 v8 → v9 (존댓말 대본)', font: F, bold: true, size: 36 })], spacing: { after: 80 } }),
  P('발표자료_20261007_v8.pptx → 발표자료_20261007_v9.pptx. 본 슬라이드 1~29의 발표 노트만 바꿈(화면 글·백업 31~70 불변). 2026-10-07 세션 61.', { size: 18, after: 120 }),
  H('1. 원칙', HeadingLevel.HEADING_1), ...rules.map(s => P('· ' + s, { size: 17 })),
  H('2. 장별 전 → 후', HeadingLevel.HEADING_1),
  new Table({ width: { size: 15398, type: WidthType.DXA }, columnWidths: W, rows }),
  H('3. 이사님 확인 사항', HeadingLevel.HEADING_1), ...checks.map(s => P('· ' + s, { size: 17 })),
] }] });
Packer.toBuffer(doc).then(b => { fs.writeFileSync('대조표_발표노트_v8_v9_20261007.docx', b); console.log('ok', log.length); });
