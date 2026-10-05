import sys, json, shutil, os
sys.path.insert(0,'/home/claude/s56w')
from figpatch import load, save, patch
M='/home/claude/s56w/x/word/media/'; O='/home/claude/s56w/fig/'; os.makedirs(O,exist_ok=True)
FLOG=[]
def do(fn, fig, edits):
    A,mode,alpha=load(M+fn)
    if not os.path.exists(O+'before_'+fn): shutil.copyfile(M+fn, O+'before_'+fn)
    for e in edits:
        old,new,win=e[:3]; kw=e[3] if len(e)>3 else {}
        info=patch(A,old,new,win,**kw); print(fig,info['score'],info['size'],info['weight'],info['box'],old,'→',new)
        FLOG.append(dict(fig=fig,file=fn,old=old,new=new,score=info['score'],box=info['box']))
    save(A,M+fn,mode,alpha); shutil.copyfile(M+fn, O+'after_'+fn)
J=[
 ('image17.png','[그림 4-6]',[
    ('사전 확정 기준에 따라 최저점(1점)으로 확정됨','3.4.3의 기준에 따라 최저점(1점)으로 확정됨',(480,270,2100,350),dict(align='left',weights=('Regular','Medium'),sizes=range(16,44),minscore=0.5)),
    ('(정책 결정이 없으면 사전 확정 기준대로 1점 — 4.5.2)','(정책 결정이 없으면 3.4.3의 기준대로 1점 — 4.5.2)',(480,870,1600,960),dict(align='left',weights=('Regular','Medium'),sizes=range(16,44),minscore=0.5))]),
 ('image50.png','[그림 2-6]',[('프로세스 능력 수준','프로세스 능력도 수준',(2800,195,3480,272),dict(align='center',weights=('Regular','Medium'),sizes=range(16,60),minscore=0.5))]),
 ('image41.png','[그림 2-7]',[('형식(목적·성과·활동)과 P-D-C-A 편성 ← 8000-61, 2013 참조모델','형식(제목·목적·성과·활동)과 P-D-C-A 편성 ← 8000-61, 2013 참조모델',(100,1490,1700,1590),dict(align='left',weights=('Regular','Medium'),sizes=range(16,50),minscore=0.5))]),
 ('image5.png','[그림 2-5]',[('(구성 불성립)','(구성개념 불성립)',(120,860,650,980),dict(align='center',weights=('Regular','Medium'),sizes=range(16,50),minscore=0.5))]),
]
only=sys.argv[1:]
for fn,fig,ed in J:
    if only and fn not in only: continue
    do(fn,fig,ed)
json.dump([{k:(int(v) if hasattr(v,"__int__") and not isinstance(v,(str,float)) else (tuple(int(t) for t in v) if isinstance(v,tuple) else v)) for k,v in r.items()} for r in FLOG],open('/home/claude/s56w/log_fig202.json','w'),ensure_ascii=False,indent=1)
