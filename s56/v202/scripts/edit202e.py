# v202 추가 반영 4 — ISO/IEC TS 33030:2017 원문 확인(10/6 이사님 업로드, 미리보기 1~7쪽: 4.1 역할 정의·4.2.2.3·4.2.2.4) → 삭제했던 [21] 복원, 3.3.4 문장을 원문대로 다시 씀
import sys, json, re
sys.path.insert(0,'/home/claude/s56w')
exec(open('/home/claude/s56w/edit202.py',encoding='utf8').read().split('# ═════ 필수 4')[0])
LOG.clear()
# 1) 번호 되돌림 [21]~[60] → [22]~[61]
def rep_at(p, st, en, new):
    ts=list(p.iter(q('t'))); pos=0; first=True
    for t in ts:
        tx=t.text or ''; a=pos; b=pos+len(tx); pos=b
        if b<=st or a>=en: continue
        i0=max(st,a)-a; i1=min(en,b)-a
        if first: t.text=tx[:i0]+new+tx[i1:]; first=False
        else: t.text=tx[:i0]+tx[i1:]
        t.set('{%s}space'%XMLNS,'preserve')
_n=0
for p in d.body.iter(q('p')):
    ms=[m for m in re.finditer(r'\[(\d+)\]', d.text(p)) if 21<=int(m.group(1))<=60]
    for m in reversed(ms):
        rep_at(p, m.start(1), m.end(1), str(int(m.group(1))+1)); _n+=1
print('renumbered back', _n)
# 2) 참고문헌 [21] 복원 — [20] 다음
_r20=[p for p in d.body.iter(q('p')) if d.text(p).startswith('[20] ISO/IEC 33020:2019')]; assert len(_r20)==1
_new='[21] ISO/IEC TS 33030:2017, Information technology — Process assessment — An exemplar documented assessment process.'
_r20[0].addnext(d.clone(_r20[0], _new))
# 3) 3.3.4 문장 — 원문 내용으로
_b742=[p for p in d.body.iter(q('p')) if d.text(p).startswith('평가자 요건은 KS X ISO 8000-62:2018 4.7이 규정하는')]; assert len(_b742)==1
def one(old,new,tag,where,n=1):
    assert d.text(_b742[0]).count(old)==1,(tag,old); rep_first(_b742[0],old,new); LOG.append(dict(tag=tag,where=where,old=old,new=new,n=1))
one('자격 인정의 일반 요건은 ISO/IEC 17024[14]를 따른다.',
    'KS X ISO 8000-62 4.7의 세 요소는 ISO/IEC TS 33030:2017[21]이 예시하는 평가 프로세스가 평가자에게 요구하는 교육·훈련, 평가 경험, 도메인 경험과 같고(4.1 역할 정의), 같은 프로세스는 착수 단계에서 적용 분야의 요구를 반영한 평가자 역량 기준을 정해 그 기준을 갖춘 평가자를 선정하도록 하므로(4.2.2.3·4.2.2.4), 본 모델의 전환교육은 그 기준을 환경경영과 LCA 분야에서 채우는 방법에 해당한다. 자격 인정의 일반 요건은 ISO/IEC 17024[14]를 따른다.',
    'TS33030(원문 확인)','3.3.4 (B742)')
d.save()
# 4) 로그 정리 — v201d 기준으로 다시 씀: 삭제·재부여 항목 제거, 복원 항목 1건
L=json.load(open('/home/claude/s56w/log_202.json'))
L=[e for e in L if e['tag']!='TS33030']
for e in L:
    if e['tag']=='V-02': e['where']='3.3.4 (B742) — 「ISO/IEC 17024」에 인용 번호 [14] 보충(다른 3곳과 같게)'
    if e['tag']=='3T-17': e['where']='부록 K shall/should 행 출처 (T119.r14.c4)'
L.append(dict(tag='TS33030(원문 확인)',where='3.3.4 (B742) · 참고문헌 [21] 유지 — 세션 53 「원문 미확인이면 삭제」 결정은 10/6 원문(ISO/IEC TS 33030:2017 미리보기 1~7쪽) 확인으로 철회',
   old='이는 KS X ISO 8000-62 4.7 비고 2가 참조하는 ISO/IEC TS 33030:2017[21]의 평가자 역량 검증 체계와 같은 취지이며, 자격 인정의 일반 요건은 ISO/IEC 17024를 따른다.',
   new='KS X ISO 8000-62 4.7의 세 요소는 ISO/IEC TS 33030:2017[21]이 예시하는 평가 프로세스가 평가자에게 요구하는 교육·훈련, 평가 경험, 도메인 경험과 같고(4.1 역할 정의), 같은 프로세스는 착수 단계에서 적용 분야의 요구를 반영한 평가자 역량 기준을 정해 그 기준을 갖춘 평가자를 선정하도록 하므로(4.2.2.3·4.2.2.4), 본 모델의 전환교육은 그 기준을 환경경영과 LCA 분야에서 채우는 방법에 해당한다. 자격 인정의 일반 요건은 ISO/IEC 17024[14]를 따른다. — 「8000-62 4.7 비고 2가 참조하는」은 8000-62 발췌집에서 확인되지 않아 뺌(33030은 8000-62 4.9 비고·부속서 C 비고 1에 나옴)',n=1))
json.dump(L,open('/home/claude/s56w/log_202.json','w'),ensure_ascii=False,indent=1); print('log entries',len(L))
