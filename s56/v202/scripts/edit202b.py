# v202 추가 반영 — 반영 문단 검토자(out_review202.md) 권고 V-01~V-03
import sys, json
sys.path.insert(0,'/home/claude/s56w')
exec(open('/home/claude/s56w/edit202.py',encoding='utf8').read().split('# ═════ 필수 4')[0])  # Doc·one·one_exact·LOG 정의만
LOG.clear()
one_exact('14044 §4.2.3.8·§4.2.3.1 (shall)','14044 §4.2.3.8·§4.2.3.1 (14001 — )','V-01','부록 C P-21 카드 근거 조항 (T69.r3)·부록 F.2 조사지 P-21 행 (T106.r22) — P-20과 같은 꼴',count=2)
one('자격 인정의 일반 요건은 ISO/IEC 17024를 따른다.','자격 인정의 일반 요건은 ISO/IEC 17024[14]를 따른다.','V-02','3.3.4 (B742) — 33030 문장 삭제로 인용 번호가 빠진 자리')
one('(이 6.2는 [표 3-1]에서는 9.2.1 쪽 근거로 정리하여 9.1.1 행에 적지 않았다)','(이 6.2는 본 모델에서 정밀검토 문항 C-5·C-6·C-9의 근거이며 [표 3-1] 9.1.1 행에는 적지 않았다)','V-03','부록 A 머리 주석 (B1104) — 3T-10(가) 문안 정정')
d.save()
L=json.load(open('/home/claude/s56w/log_202.json')); L+=LOG
json.dump(L,open('/home/claude/s56w/log_202.json','w'),ensure_ascii=False,indent=1); print(len(LOG),'added; total',len(L))
