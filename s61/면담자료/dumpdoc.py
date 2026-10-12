import sys,zipfile,re,json
W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
from lxml import etree
def dump(path):
    z=zipfile.ZipFile(path); root=etree.fromstring(z.read('word/document.xml'))
    body=root.find(W+'body'); out=[]
    for el in body:
        if el.tag==W+'p':
            sty=el.find(W+'pPr/'+W+'pStyle'); sty=sty.get(W+'val') if sty is not None else ''
            t=''.join(x.text or '' for x in el.iter(W+'t'))
            out.append(('p',sty,t))
        elif el.tag==W+'tbl':
            rows=el.findall(W+'tr'); first=''.join(x.text or '' for x in rows[0].iter(W+'t')) if rows else ''
            out.append(('tbl','',f'[표 {len(rows)}행] '+first[:80]))
    media=[n for n in z.namelist() if n.startswith('word/media/')]
    return out,media
out,media=dump(sys.argv[1])
paras=[o for o in out if o[0]=='p' and o[2].strip()]
heads=[o for o in paras if o[1].lower().startswith('heading') or re.match(r'^(제\s*\d+\s*장|\d+\.\d+(\.\d+)?\s)',o[2])]
caps_t=[o[2] for o in paras if re.match(r'^\[?표\s*\d+-\d+\]?',o[2])]
caps_f=[o[2] for o in paras if re.match(r'^\[?그림\s*\d+-\d+\]?',o[2])]
refs=[o[2] for o in paras if re.match(r'^\[\d+\]',o[2])]
chars=sum(len(o[2]) for o in paras)
print(json.dumps({'paras':len(paras),'chars':chars,'tables':sum(1 for o in out if o[0]=='tbl'),'media':len(media),'table_caps':len(caps_t),'fig_caps':len(caps_f),'refs':len(refs)},ensure_ascii=False))
json.dump({'heads':[h[2] for h in heads],'caps_t':caps_t,'caps_f':caps_f,'paras':[o[2] for o in paras]},open(sys.argv[2],'w',encoding='utf8'),ensure_ascii=False)
