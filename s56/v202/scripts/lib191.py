import copy, random, re
from lxml import etree
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
W14='http://schemas.microsoft.com/office/word/2010/wordml'
XML='http://www.w3.org/XML/1998/namespace'
def q(t): return '{%s}%s'%(W,t)
class Doc:
    def __init__(s,path):
        s.path=path
        s.tree=etree.parse(path); s.root=s.tree.getroot(); s.body=s.root.find(q('body'))
        s.pids=set(e.get('{%s}paraId'%W14) for e in s.root.iter() if e.get('{%s}paraId'%W14))
        s.bmid=max(int(b.get(q('id'))) for b in s.root.iter(q('bookmarkStart')))+100
        s.log=[]
    def save(s):
        s.tree.write(s.path,xml_declaration=True,encoding='UTF-8',standalone=True)
    # ---------- text
    @staticmethod
    def text(p): return ''.join(t.text or '' for t in p.iter(q('t')))
    def allp(s): return list(s.body.iter(q('p')))
    def findp(s,sub,start=None,n=1,top=False):
        ps=[p for p in (s.body if top else s.body.iter(q('p'))) if p.tag==q('p') and sub in s.text(p)]
        if start is not None:
            ps=[p for p in ps if s.after(p,start)]
        assert len(ps)==n,(sub,len(ps))
        return ps[0] if n==1 else ps
    def after(s,a,b):
        # is a after b in document order
        order=getattr(s,'_order',None)
        allp=list(s.body.iter())
        return allp.index(a)>allp.index(b)
    def rep(s,p,old,new,tag=''):
        ts=list(p.iter(q('t')))
        full=''.join(t.text or '' for t in ts)
        assert full.count(old)==1,(tag,old,full.count(old),full[:80])
        k=full.find(old); st,en=k,k+len(old); pos=0; first=True
        for t in ts:
            tx=t.text or ''; a=pos; b=pos+len(tx); pos=b
            if b<=st or a>=en:
                if not(b==st and a==b and first and False): continue
            i0=max(st,a)-a; i1=min(en,b)-a
            if first:
                t.text=tx[:i0]+new+tx[i1:]; first=False
            else:
                t.text=tx[:i0]+tx[i1:]
            t.set('{%s}space'%XML,'preserve')
        s.log.append((tag,old,new))
    def rep_find(s,old,new,tag='',n=1):
        ps=[p for p in s.body.iter(q('p')) if old in s.text(p)]
        cnt=sum(s.text(p).count(old) for p in ps)
        assert cnt==n,(tag,old,cnt)
        for p in ps:
            while old in s.text(p): s.rep(p,old,new,tag)
    def newpid(s):
        while True:
            v='%08X'%random.randint(0x10000000,0x7FFFFFFF)
            if v not in s.pids: s.pids.add(v); return v
    def set_text(s,p,text):
        runs=[r for r in p if r.tag==q('r') and r.find(q('t')) is not None]
        tmpl=copy.deepcopy(runs[0]) if runs else None
        for r in list(p):
            if r.tag in (q('r'),q('hyperlink')): p.remove(r)
        r=tmpl if tmpl is not None else etree.SubElement(p,q('r'))
        for c in list(r):
            if c.tag!=q('rPr'): r.remove(c)
        t=etree.SubElement(r,q('t')); t.text=text; t.set('{%s}space'%XML,'preserve')
        # insert run before trailing bookmarkEnd if any
        be=[c for c in p if c.tag==q('bookmarkEnd')]
        if be: be[0].addprevious(r)
        else: p.append(r)
        return p
    def clone(s,tmpl,text=None,keep_bm=False):
        p=copy.deepcopy(tmpl)
        if not keep_bm:
            for b in list(p.iter(q('bookmarkStart')))+list(p.iter(q('bookmarkEnd'))): b.getparent().remove(b)
        for e in p.iter():
            if e.get('{%s}paraId'%W14): e.set('{%s}paraId'%W14,s.newpid())
        if text is not None: s.set_text(p,text)
        return p
    def heading(s,tmpl,text,bmname):
        p=s.clone(tmpl,None)
        s.bmid+=1
        bs=etree.Element(q('bookmarkStart')); bs.set(q('id'),str(s.bmid)); bs.set(q('name'),bmname)
        be=etree.Element(q('bookmarkEnd')); be.set(q('id'),str(s.bmid))
        ppr=p.find(q('pPr')); ppr.addnext(bs)
        s.set_text(p,text); p.append(be)
        return p
    def tocline(s,tmpl,text,anchor):
        p=copy.deepcopy(tmpl)
        for e in p.iter():
            if e.get('{%s}paraId'%W14): e.set('{%s}paraId'%W14,s.newpid())
        h=p.find(q('hyperlink')); h.set(q('anchor'),anchor)
        ts=list(h.iter(q('t'))); ts[0].text=text
        for it in h.iter(q('instrText')): it.text=' PAGEREF %s \\h '%anchor
        return p
    def cap_bm(s,p,name):
        s.bmid+=1
        bs=etree.Element(q('bookmarkStart')); bs.set(q('id'),str(s.bmid)); bs.set(q('name'),name)
        be=etree.Element(q('bookmarkEnd')); be.set(q('id'),str(s.bmid))
        ppr=p.find(q('pPr')); ppr.addnext(bs); bs.addnext(be)
    @staticmethod
    def top_of(e):
        while e.getparent().tag!=q('body'): e=e.getparent()
        return e
