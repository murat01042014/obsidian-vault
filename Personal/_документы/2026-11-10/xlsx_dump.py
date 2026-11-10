import zipfile,re,sys,io,html
def load(f):
    z=zipfile.ZipFile(f)
    wb=z.read("xl/workbook.xml").decode("utf8")
    rels=z.read("xl/_rels/workbook.xml.rels").decode("utf8")
    rmap=dict(re.findall(r'<Relationship [^>]*?Id="([^"]+)"[^>]*?Target="([^"]+)"',rels))
    rmap.update({a:b for b,a in re.findall(r'<Relationship [^>]*?Target="([^"]+)"[^>]*?Id="([^"]+)"',rels)})
    ss=z.read("xl/sharedStrings.xml").decode("utf8") if "xl/sharedStrings.xml" in z.namelist() else ""
    strs=[html.unescape("".join(re.findall(r"<t[^>]*>(.*?)</t>",s,re.S))) for s in re.findall(r"<si>(.*?)</si>",ss,re.S)]
    out={}
    for name,rid in re.findall(r'<sheet name="([^"]+)"[^>]*?r:id="([^"]+)"',wb):
        t=rmap[rid]; t=t.lstrip("/"); t=t if t.startswith("xl/") else "xl/"+t
        x=z.read(t).decode("utf8")
        grid={}
        for rm in re.finditer(r'<row [^>]*?r="(\d+)"[^>]*?>(.*?)</row>',x,re.S):
            r=int(rm.group(1))
            for cm in re.finditer(r'<c r="([A-Z]+)\d+"([^>]*?)(?:/>|>(.*?)</c>)',rm.group(2),re.S):
                col,attrs,body=cm.groups()
                if not body: continue
                v=re.search(r"<v>(.*?)</v>",body,re.S)
                if "t=\"s\"" in attrs and v: val=strs[int(v.group(1))]
                elif "t=\"inlineStr\"" in attrs: val=html.unescape("".join(re.findall(r"<t[^>]*>(.*?)</t>",body,re.S)))
                elif v: val=html.unescape(v.group(1))
                else: continue
                if val!="": grid[(r,col)]=val
        merges=re.findall(r'<mergeCell ref="([^"]+)"',x)
        out[name]=(grid,merges)
    return out
if __name__=="__main__":
    sys.stdout=io.TextIOWrapper(sys.stdout.buffer,encoding="utf-8")
    f,sheet=sys.argv[1],sys.argv[2]
    d=load(f)[sheet][0]
    rows=sorted(set(r for r,c in d))
    for r in rows[:int(sys.argv[3]) if len(sys.argv)>3 else 60]:
        print(r,{c:v for (rr,c),v in sorted(d.items(),key=lambda kv:(len(kv[0][1]),kv[0][1])) if rr==r})
