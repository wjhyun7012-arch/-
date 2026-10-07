#!/usr/bin/env python3
"""논문 표기 규칙·상호 참조 점검 — usage: python3 check_rules.py <docx> [rules.json] > report.txt"""
import sys,re,json,zipfile,html,bisect,os
from collections import defaultdict,Counter
docx=sys.argv[1]; rules=json.load(open(sys.argv[2] if len(sys.argv)>2 else os.path.join(os.path.dirname(__file__),'rules.json'),encoding='utf-8'))
xml=zipfile.ZipFile(docx).read('word/document.xml').decode('utf-8')
b=xml.find('<w:body>')
S=[(m.start()+b,m.end()+b) for m in re.finditer(r'<w:p[ >].*?</w:p>',xml[b:],flags=re.S)]
def ptext(p): return ''.join(html.unescape(t) for t in re.findall(r'<w:t[^>]*>([^<]*)</w:t>',p))
def pstyle(p):
    m=re.search(r'<w:pStyle w:val="([^"]+)"',p); return m.group(1) if m else ''
P=[]
for i,(s,e) in enumerate(S):
    p=xml[s:e]; P.append(dict(i=i,t=ptext(p),sty=pstyle(p),tbl=xml.rfind('<w:tbl>',0,s)>xml.rfind('</w:tbl>',0,s)))
N=len(P)
# ---- regions
def first(pred,start=0):
    for q in P[start:]:
        if pred(q): return q['i']
    return None
i_key=first(lambda q:q['t'].startswith('핵심주제어'))
i_abs_h=None
if i_key is not None:
    j=i_key
    while j>0 and '강원대학교 대학원' not in P[j-1]['t'] and i_key-j<20: j-=1
    i_abs_h=j
chap=[q['i'] for q in P if q['sty']=='1' and re.match(r'제\d장',q['t'])]
i_ref=first(lambda q:q['sty']=='1' and re.sub(r'\s','',q['t']).startswith('참고문헌'))
i_abst=first(lambda q:q['t'].strip().upper()=='ABSTRACT' and q['i']>(i_ref or 0))
i_app=first(lambda q:q['sty']=='1' and re.sub(r'\s','',q['t']).startswith('부록'),i_ref or 0)
region={}
for q in P:
    i=q['i']
    if i_abs_h is not None and i_abs_h<=i<=(i_key or i_abs_h): region[i]='초록'
    elif chap and chap[0]<=i<(i_ref or N): region[i]=f'{bisect.bisect_right(chap,i)}장'
    elif i_ref is not None and i_ref<=i<(i_app or i_abst or N): region[i]='참고문헌'
    elif i_app is not None and i_app<=i<(i_abst or N): region[i]='부록'
    elif i_abst is not None and i>=i_abst: region[i]='ABSTRACT'
    else: region[i]='앞부분'
body=lambda i: region[i]=='초록' or region[i].endswith('장')
scopes={'body':body,'body_appendix':lambda i: body(i) or region[i]=='부록','all':lambda i:True,
        'body_headings':lambda i: body(i) or P[i]['sty'] in ('1','2','3','11','20','30')}
out=[]; viol=0; warn=0
def loc(i): return f"[{i}] {region[i]}{' 표' if P[i]['tbl'] else ''}: "
# ---- 1 forbidden
out.append('## 1. 용어·표기 규칙 위반')
for r in rules['forbidden']:
    hits=[]
    for q in P:
        i=q['i']
        if not scopes[r['scope']](i): continue
        for m in re.finditer(r['pattern'],q['t']):
            ctx=q['t'][max(0,m.start()-25):m.end()+25]
            if any(a in q['t'][max(0,m.start()-30):m.end()+30] for a in r.get('allow',[])): continue
            hits.append(loc(i)+'…'+ctx+'…')
    viol+=len(hits)
    out.append(f"- {r['id']} {'위반 %d건'%len(hits) if hits else '0건'} — {r['rule']}")
    out+= ['    '+h for h in hits[:20]]
# ---- 2 abbreviations
out.append('\n## 2. 약어 첫 등장(장별)')
ab=0
for reg in [f'{k}장' for k in range(1,len(chap)+1)]:
    for a in rules['abbr']['check']:
        if a in rules['abbr']['exempt']: continue
        pat=re.compile(r'(?<![A-Za-z])'+re.escape(a)+r'(?![A-Za-z])')
        ign=rules['abbr'].get('ignore_tokens',[])
        def tx(t):
            for g in ign: t=t.replace(g,'')
            return t
        firsts=[q for q in P if region[q['i']]==reg and not q['tbl'] and pat.search(tx(q['t']))]
        if not firsts: continue
        q=firsts[0]
        ok=re.search(r'\('+re.escape(a)+r'\s*[:：]',q['t']) is not None or re.search(r'(?<![A-Za-z])'+re.escape(a)+r'\((?:[A-Z]|[a-z]+ )',q['t']) is not None
        if not ok:
            ab+=1; out.append(f"- {reg} {a}: 첫 등장 {loc(q['i'])}…{q['t'][max(0,pat.search(q['t']).start()-30):pat.search(q['t']).end()+30]}…")
out.append(f"약어 규칙 위반 {ab}건 — {rules['abbr']['rule']}"); viol+=ab
# ---- 3 warnings
out.append('\n## 3. 경고(확인만)')
for r in rules['warnings']:
    hits=[]
    for q in P:
        i=q['i']
        if not scopes[r['scope']](i): continue
        for m in re.finditer(r['pattern'],q['t']):
            if any(a in q['t'][max(0,m.start()-30):m.end()+30] for a in r.get('allow',[])): continue
            hits.append(loc(i)+'…'+q['t'][max(0,m.start()-25):m.end()+25]+'…')
    warn+=len(hits)
    out.append(f"- {r['id']} {len(hits)}건 — {r['rule']}"); out+=['    '+h for h in hits[:12]]
# ---- 4 cross references
out.append('\n## 4. 상호 참조')
heads=[(q['i'],q['sty'],re.sub(r'\s+',' ',q['t']).strip()) for q in P if q['sty'] in('1','2','3')]
toc=[(q['i'],q['sty'],re.sub(r'\s+',' ',re.sub(r'\d+$','',q['t'])).strip()) for q in P if q['sty'] in('11','20','30')]
hb=[t for i,s,t in heads]; tb=[t for i,s,t in toc]
import difflib
tocbad=[(op,hb[op[1]:op[2]],tb[op[3]:op[4]]) for op in difflib.SequenceMatcher(None,hb,tb).get_opcodes() if op[0]!='equal' and not (op[0]=='insert' and [z.upper() for z in tb[op[3]:op[4]]]==['ABSTRACT']) and not (op[0]=='replace' and all(a.startswith(b_) for a,b_ in zip(hb[op[1]:op[2]],tb[op[3]:op[4]])))]
out.append(f"- 목차 {len(toc)}줄 ↔ 제목 {len(heads)}개: {'일치' if not tocbad else 'DIFF '+str(tocbad)}")
starts=[s for s,e in S]
bmpos={m.group(1):m.start() for m in re.finditer(r'<w:bookmarkStart [^>]*w:name="([^"]+)"',xml)}
def para_of(pos):
    k=bisect.bisect_right(starts,pos)-1
    if pos>S[k][1]: k+=1
    return k
nrm=lambda s:re.sub(r'\s+',' ',s).strip()
bad=[];npr=0
for q in P:
    if region[q['i']]!='앞부분': continue
    p=xml[S[q['i']][0]:S[q['i']][1]]
    ins=''.join(re.findall(r'<w:instrText[^>]*>([^<]*)</w:instrText>',p)); m=re.search(r'PAGEREF\s+(\S+)',ins)
    if not m: continue
    npr+=1
    if m.group(1) not in bmpos: bad.append(('책갈피 없음',q['i'])); continue
    pre=p[:p.find('<w:fldChar w:fldCharType="begin"')]; lt=nrm(ptext(pre)); k=para_of(bmpos[m.group(1)]); tt=nrm(P[k]['t'])
    if lt and lt!=tt: bad.append(('불일치',q['i'],lt[:30],tt[:30]))
out.append(f"- 목차·표목차·그림목차 PAGEREF {npr}개 ↔ 책갈피: {'모두 정상' if not bad else bad}")
caps={}
for q in P:
    if not body(q['i']) or q['tbl']: continue
    m=re.match(r'\s*\[(표|그림) (\d+-\d+)\] ',q['t'])
    if m and not re.match(r'\s*\[(표|그림) \d+-\d+\](에서|의|과|은|는)',q['t']): caps[(m.group(1),m.group(2))]=q['i']
refs=defaultdict(list)
for q in P:
    if not body(q['i']): continue
    for m in re.finditer(r'(표|그림) ?(\d+)-(\d+)',q['t']):
        k=(m.group(1),m.group(2)+'-'+m.group(3))
        if caps.get(k)==q['i'] and m.start()<3: continue
        refs[k].append(q['i'])
miss=[k for k in refs if k not in caps]; unref=[k for k in caps if not refs.get(k)]
out.append(f"- 표·그림 캡션 {len(caps)}개, 인용 대상 없음 {miss or '없음'}, 인용되지 않은 캡션 {unref or '없음'}(범위 인용 [그림 4-5]~[그림 4-9]는 4-6·4-8을 덮음)")
secs=set()
for i,s,t in heads:
    m=re.match(r'(\d+\.\d+(?:\.\d+)?)\s',t) or re.match(r'([A-K]\.\d+)\s',t)
    if m: secs.add(m.group(1))
std=re.compile(r'(ISO|KS|IEC|표준|조항|§|부속서|8000|33\d\d\d|14\d\d\d|19011|17024|10015|9004|9001|45001|9000|17000)')
secbad=[]
for q in P:
    if not body(q['i']): continue
    for m in re.finditer(r'(?<![\d.§])([1-5]\.\d{1,2}(?:\.\d{1,2})?)(?![\d.]*\d)',q['t']):
        num=m.group(1)
        if num in secs or re.fullmatch(r'[1-5]\.\d\d',num): continue
        if std.search(q['t']): continue
        secbad.append((q['i'],num))
out.append(f"- (정보) 논문 절 번호와 겹치지 않는 x.y.z 번호 — 표준 조항일 가능성이 큼, 확인용: {secbad[:10] or '없음'}")
# citations
refstart=first(lambda q:q['t'].startswith('[1] ') and region[q['i']]=='참고문헌')
entries=[q for q in P if region[q['i']]=='참고문헌' and re.match(r'\[\d+\] ',q['t'])]
nref=len(entries)
cites=Counter(); oor=[]
for q in P:
    if region[q['i']]=='참고문헌': continue
    for m in re.finditer(r'\[(\d{1,2})\]',q['t']):
        n=int(m.group(1)); cites[n]+=1
        if n<1 or n>nref: oor.append((q['i'],n))
uncited=[k for k in range(1,nref+1) if cites[k]==0]
seq=[int(re.match(r'\[(\d+)\]',q['t']).group(1)) for q in entries]
out.append(f"- 참고문헌 {nref}건(번호 연속 {'정상' if seq==list(range(1,nref+1)) else '이상 '+str(seq)}), 범위 밖 인용 {oor or '없음'}, 인용되지 않은 문헌 {uncited or '없음'}, 본문 인용 {sum(cites.values())}곳")
# numbers
out.append('\n## 5. 핵심 수치 출현 횟수')
for k,v in rules['numbers'].items():
    c=sum(len(re.findall(r'(?<![\d.])'+re.escape(v)+r'(?![\d])',q['t'])) for q in P)
    out.append(f"- {k} {v}: {c}회")
out.insert(0,f"# 점검 결과 — {os.path.basename(docx)} · 규칙표 {rules['version']} · 문단 {N}개 (초록 {sum(1 for i in region if region[i]=='초록')}·본문 {sum(1 for i in region if region[i].endswith('장'))}·참고문헌·부록 {sum(1 for i in region if region[i]=='부록')})\n**규칙 위반 {viol}건 · 경고 {warn}건 · 상호 참조 {'이상 없음' if not (tocbad or bad or miss or oor or uncited) else '확인 필요'}**\n")
print('\n'.join(out))
