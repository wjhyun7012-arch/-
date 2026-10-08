// 예상 질문 v5 (48문) — v4 43문(발표자료 v7 백업 71~80번 표에서 추출, 논문 위치 v203 기준) + ★5문 추가
// 사용: node mkq5.js  (q43.json 필요)
const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType, HeadingLevel, PageOrientation, BorderStyle } = require('docx');
const F = '맑은 고딕';
const t = (s, o = {}) => new TextRun({ text: String(s), font: F, size: o.size || 18, bold: o.bold, color: o.color });
const P = (s, o = {}) => new Paragraph({ children: Array.isArray(s) ? s : [t(s, o)], spacing: { after: o.after ?? 80 } });
const H = (s, l) => new Paragraph({ heading: l, children: [new TextRun({ text: s, font: F, bold: true, size: l === HeadingLevel.HEADING_1 ? 26 : 22 })], spacing: { before: 240, after: 120 } });
const bd = { style: BorderStyle.SINGLE, size: 4, color: 'A6B0C3' }; const borders = { top: bd, bottom: bd, left: bd, right: bd };
const W = [600, 3300, 9100, 2398];
const cell = (s, w, o = {}) => new TableCell({ width: { size: w, type: WidthType.DXA }, borders, shading: o.fill ? { type: ShadingType.CLEAR, fill: o.fill, color: 'auto' } : undefined, margins: { top: 50, bottom: 50, left: 80, right: 80 }, children: (Array.isArray(s) ? s : [s]).map(x => P(x, { size: o.size || 17, bold: o.bold, color: o.color, after: 40 })) });
const q43 = JSON.parse(fs.readFileSync('q43.json', 'utf8')); // [title, no, q, a, loc]
// v6: 43번 답 정정(이사님 10/8) — 구성안·로드맵은 연구자가 설계, AI는 검토·수정; 위치에 B-Q13
for (const r of q43) { if (r[1] === '43') { const a0 = r[3]; r[3] = r[3].replace('구상 단계(2025년 10월)의 논문 구성안·연구 로드맵 정리', '연구자가 세운 논문 구성안·연구 로드맵의 검토·수정(2025년 10월)'); if (r[3] === a0) throw new Error('43 unchanged'); r[4] = r[4] + ' → B-Q13';
    // v7 (10/8): 원문 확인의 범위, 전 조항 대조 일부 AI 보조
    const a1 = r[3]; r[3] = r[3].replace('두 표준의 전 조항 대조(2025년 10월 검토표 → 2026년 3월 전조항 분석)', '두 표준의 전 조항 대조(2025년 10월 검토표 → 2026년 3월 전조항 분석 — 일부 AI 보조, 연구자가 검토·확정)').replace('원칙 — 표준 조항과 인용은 원문을 확인한 것만 넣었고,', '원칙 — 조항을 찾는 데는 AI를 썼고, 인용한 조항과 문헌은 연구자가 모두 원문에서 확인하였으며,'); if (r[3] === a1) throw new Error('43 v7 unchanged'); } }
const sec = s => s.replace(/^예상 질문 43문 — /, '').replace(/\s*\(\d+\/10\)\s*$/, '');
const groups = {};
for (const [title, no, q, a, loc] of q43) { const g = sec(title); (groups[g] = groups[g] || []).push({ no, q, a, loc }); }
// ★ 추가 5문 — 근거 v203 (10/7 대화 요지를 정식 답변으로)
const star = [
  { no: '44', q: '★ 문항에 가중치는 왜 주지 않았는가. AHP 같은 방법으로 중요도를 반영해야 하지 않는가.',
    a: '검토하였고, 일부러 주지 않았다. ISO/IEC 33003 부속서 D는 형성형 구성개념의 집계 방법으로 보상·비보상 모델과 AHP 등을 들고 있어 핵심 조항의 문항에 가중치를 주는 방식도 검토하였다(3.4.4). 그러나 가중치를 쓰는 보상 모델은 값이 높은 문항이 핵심 지점의 부재를 가중치에 비례하여 메워 주므로 — 이 연구가 가장 드러내려는 「연계의 미성립」이 가려진다. 또 같은 표준은 가중치를 쓰면 부여 방법의 명시와 가중치 민감도 분석을 요구한다(4.7.1 c), 4.7.2). 그래서 필요조건인 핵심 문항 4종에는 비보상 모델(최소충족)을 적용하고, 종합점수는 무가중 평균의 보조값으로만 둔다 — 종합점수는 판정에 쓰지 않으므로 가중치 민감도 요건은 해당하지 않는다. 바꾸어 말하면 「중요한 문항」의 중요성은 가중치가 아니라 「빠지면 수준이 올라가지 못한다」는 규칙으로 반영하였다.',
    loc: '3.4.1, 3.4.2, 3.4.4, 2.4.3, 4.3.5 → B-Q1' },
  { no: '45', q: '★ 통계적 신뢰도·타당도 검증은 왜 없는가 — 요인분석이나 평가자 간 일치도는?',
    a: '이 연구는 모델 개발 연구이고, 검증은 그 단계에 맞는 것을 하였다. 타당성은 두 방향으로 검증하였다 — 55문항을 두 표준의 전 조항 대조에서 나온 근거 조항까지 추적한 추적성 검증(3.5.1)과, 세 모델의 요건을 정한 ISO/IEC 33004 요구사항 24건의 자체점검(3.5.2). 전문가 내용타당도 조사(CVI)는 조사지 설계까지 마쳤고 실시는 향후 과제다(3.5.3). 요인분석은 반영형 구성개념에 맞는 방법이고, 본 모델의 문항은 활동·행동을 측정하는 형성형 구성개념이라 문항끼리 서로 바꿀 수 없어 요인 구조를 전제하지 않는다(2.4.3, ISO/IEC 33003 부속서 B). 평가자 간 일치도와 모델 적용의 효과성은 단일 평가자·단일 사례로는 잴 수 없으므로 다사례·다평가자 평가와 개선 이행 뒤의 재평가로 확인할 과제로 두었다(3.5 머리 문단, 5.3). 대신 한 평가자 안에서의 재현성은 판정 관례의 사전 명문화와 체크포인트 단위 전수 재검증(상향 7·하향 9, 2.89→2.85)으로 확보하였다(4.2.2·4.2.3).',
    loc: '3.5(머리 문단), 3.5.1~3.5.3, 2.4.3, 4.2.2, 4.2.3, 5.3 → B-Q6·B-Q10' },
  { no: '46', q: '★ 성숙도가 올라가면 실제로 비용이나 탄소가 줄어드는가.',
    a: '이 연구에서는 확인하지 않았고, 그렇게 쓰지도 않았다. 성숙도 수준과 실제 환경성과(환경영향의 저감, 비용)의 관계는 [표 5-1]의 단점에 「환경성과와의 관계 미검증」으로 적었고, 5.3 다섯째 과제로 두었다. 이 연구에서 확인된 것은 판정 논리가 실제 자료에서 작동한다는 것까지다 — 종합점수가 높아도 핵심 문항 하나가 빠지면 ML2에 머무른다는 것을 A사에서 확인하였다. 기대 효과(LCA 결과가 환경측면·목표·설계·모니터링에 반영되면 탄소 저감이 비용 절감으로 이어진다)는 모델의 구조에서 예상되는 것이며, 실증은 ML4 이상(재수행에 의한 정량 검증)에 이른 조직을 추적해야 가능하다.',
    loc: '[표 5-1] 단점 「환경성과와의 관계 미검증」, 5.3 다섯째, 4.7' },
  { no: '47', q: '★ EcoVadis·CDP 같은 ESG 평가와는 무엇이 다른가.',
    a: '보는 것이 다르다. ESG 평가는 공시된 정보와 성과의 결과(배출량·정책·점수)를 밖에서 평가한다. 본 모델은 조직 안에서 두 표준 — 환경경영시스템(ISO 14001)과 전과정평가(LCA) — 이 프로세스로 이어져 돌고 있는지를 조항 단위로 본다. 즉 결과가 아니라 결과를 만드는 체계의 연계 수준을 재는 것이고, 그래서 평가 결과가 곧 개선 과제가 된다. 논문에서 EcoVadis는 1.1과 2.2.1에서 「가치사슬 전반의 정량 환경정보를 요구하는 규제·시장 환경」의 예로만 언급하였고, ESG 평가와의 비교는 연구 범위 밖이다 — 물으면 위와 같이 말로 답한다.',
    loc: '1.1, 2.2.1(규제·시장 환경 예시), 5.2 — 비교 자체는 논문 밖' },
  { no: '48', q: '★ 평가를 도구화·자동화할 수 있는가.',
    a: '점수 산정은 이미 규칙이라 자동화할 수 있고, 증빙 판단은 평가자가 한다. 문항 점수는 체크포인트 충족 개수로 정해지므로(필수 3/3 + 추가 2/2 …), 평가 기록부는 체크포인트 개별 판정을 입력하면 점수가 자동 산정되는 형식이다([표 4-8], 4.2.3). 판정 규칙(핵심 문항 4종의 최소충족, 과반 기준)도 규칙 기반이라 소프트웨어로 구현할 수 있다. 다만 「이 문서가 이 체크포인트를 충족하는가」의 판단은 사내표준과 기록을 읽는 평가자의 몫이고, 그 판단 기준을 BARS 앵커와 객관적 증거 목록으로 미리 적어 둔 것이 이 모델이다. 시스템 구현은 연구 범위 밖이며, 도구화는 다사례 적용과 함께 향후 가능한 과제로 본다.',
    loc: '[표 4-8], 4.2.3, 3.3.2, 부록 C·I — 구현은 논문 밖' },
];
groups['바. 추가 ★ (10/7 — 외부 위원 대비)'] = star;
const rowsOf = list => [new TableRow({ tableHeader: true, children: ['번호', '질문', '답', '논문 위치'].map((h, i) => cell(h, W[i], { fill: '1F3864', bold: true, color: 'FFFFFF' })) }),
  ...list.map(r => new TableRow({ cantSplit: true, children: [cell(r.no, W[0]), cell(r.q, W[1], { bold: true }), cell(r.a, W[2]), cell(r.loc, W[3], { size: 15 })] }))];
const children = [
  new Paragraph({ children: [new TextRun({ text: '예상 질문 v7 — 48문', font: F, bold: true, size: 34 })], spacing: { after: 80 } }),
  P('2026-10-08 세션 61. v5(2026-10-07, 48문)에서 43번(AI 활용) 답을 정정 — 논문 구성안·로드맵은 연구자가 설계하고 AI는 검토·수정, 전 조항 대조는 일부 AI 보조·연구자 확정, 인용한 조항·문헌은 연구자가 모두 원문에서 확인(찾는 것은 AI), 답할 장은 백업 B-Q13(75번). v4(2026-09-29, 43문)의 논문 위치를 v203 기준으로 대조한 것에 ★5문(44~48)을 더함. 발표 자료 v8에서는 43문 표(백업 71~80)를 뺐으므로 이 파일이 예상 질문의 유일한 전문.', { size: 17 }),
  P('답은 먼저 하고 논문 위치를 덧붙인다. 「→ B-Qn」은 같은 물음의 그림·표 백업(발표 자료 v8 58~70번). 모르는 것은 「확인해서 보완하겠습니다」.', { size: 17, after: 160 }),
  P('심사위원 구성 — 교내 산업경영공학과 3명(방법론: 핵심 문항 선정, 판정 파라미터, CVI, 단일 사례), 교외 경영학과 1명(기업에 무슨 도움, 한 문장으로, ESG 평가와의 차이, 비용·성과), 컴퓨터공학과 1명(왜 33000 계열, 척도, 재현성, 자동화). ★5문은 교외 두 분 관점에서 보탠 것.', { size: 17, after: 200 }),
];
for (const g of Object.keys(groups)) { children.push(H(g, HeadingLevel.HEADING_1)); children.push(new Table({ width: { size: 15398, type: WidthType.DXA }, columnWidths: W, rows: rowsOf(groups[g]) })); }
const doc = new Document({ styles: { default: { document: { run: { font: F, size: 18 } } } }, sections: [{ properties: { page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }, margin: { top: 720, bottom: 720, left: 720, right: 720 } } }, children }] });
Packer.toBuffer(doc).then(b => { fs.writeFileSync('예상질문_v7_48문_20261008.docx', b); console.log('ok', Object.keys(groups).map(k => k + ':' + groups[k].length).join(' ')); });
