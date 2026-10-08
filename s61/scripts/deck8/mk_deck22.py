# v21 -> v22: 75번 B-Q13 문구 정정 — 원문 확인의 범위, 전 조항 대조 일부 AI 보조
import sys, json
sys.path.insert(0, '.')
from pptx import Presentation
from pptlib import rep, rep_notes, LOG
src, dst = sys.argv[1], sys.argv[2]
prs = Presentation(src); S = 75
WHY = '이사님 10/8 — 원문은 전부가 아니라 AI가 찾아 준 곳을 확인; 전 조항 대조도 일부 AI 보조'
rep(prs, S, 6, '조항·수치는 원문과 연구자가 확인', '인용한 조항·수치는 연구자가 원문으로 확인', WHY)
rep(prs, S, 7, '두 표준 전 조항 대조(2025.10 검토표 → 2026.3 전조항 분석)', '두 표준 전 조항 대조(2025.10 검토표 → 2026.3 전조항 분석) — 일부 AI 보조, 연구자가 검토·확정', WHY)
rep(prs, S, 9, '표준 조항·인용은 원문을 확인한 것만 넣음', '인용한 조항·문헌은 모두 원문에서 확인 — 찾는 것은 AI, 확인과 판단은 연구자', WHY)
rep_notes(prs, S, '논문 전체 설계와 두 표준의 전 조항 대조, 55문항과 채점 기준, A사 판정과 개선 과제, 그리고 모든 결정은 제가 했습니다.',
          '논문 전체 설계, 55문항과 채점 기준, A사 판정과 개선 과제, 그리고 모든 결정은 제가 했습니다. 두 표준의 전 조항 대조는 일부 AI의 보조를 받아 제가 검토하고 확정했습니다.', WHY)
rep_notes(prs, S, '표준 조항과 인용은 원문을 확인한 것만 넣었고, 수치와 판정은 제가 산출하여',
          '조항을 찾는 데는 AI를 썼고, 논문에 넣은 조항과 인용은 제가 모두 원문에서 확인했습니다. 수치와 판정은 제가 산출하여', WHY)
prs.save(dst); json.dump(LOG, open('log22.json', 'w', encoding='utf8'), ensure_ascii=False, indent=1); print('ok', len(LOG))
