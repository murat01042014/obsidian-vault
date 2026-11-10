import sys,io,re,json,datetime
sys.path.insert(0,"D:/Кадры/_исходники")
from xlsx_dump import load
MON={"ЯНВАРЬ":1,"ФЕВРАЛЬ":2,"МАРТ":3,"АПРЕЛЬ":4,"МАЙ":5,"ИЮНЬ":6,"ИЮЛЬ":7,"АВГУСТ":8,"СЕНТЯБРЬ":9,"ОКТЯБРЬ":10,"НОЯБРЬ":11,"ДЕКАБРЬ":12}
def colnum(c):
    n=0
    for ch in c: n=n*26+ord(ch)-64
    return n
def parse(f):
    res={}  # name -> {date: code}
    for sheet,(grid,merges) in load(f).items():
        m=MON.get(sheet.split()[0].upper())
        if not m: continue
        # day header row: row with D=1,E=2
        hdr=[r for (r,c),v in grid.items() if c=="D" and v=="1" and (r,"E") in grid and grid[(r,"E")]=="2"]
        if not hdr: print("нет шапки",sheet); continue
        hr=hdr[0]
        days={}
        for (r,c),v in grid.items():
            if r==hr and v.isdigit() and 4<=colnum(c)<=34: days[c]=int(v)
        for (r,c),v in grid.items():
            if c=="B" and r>hr and (r,"A") in grid and grid[(r,"A")].isdigit():
                name=v.strip()
                for col,d in days.items():
                    code=grid.get((r,col))
                    if code is None: continue
                    # дни > число дней месяца не считаем
                    try: dt=datetime.date(2026,m,d)
                    except: continue
                    res.setdefault(name,{})[dt]=code.strip()
    return res
if __name__=="__main__":
    sys.stdout=io.TextIOWrapper(sys.stdout.buffer,encoding="utf-8")
    for f in sys.argv[1:]:
        r=parse(f)
        print("=====",f)
        for name,dd in r.items():
            codes={}
            for dt,c in dd.items():
                if not c.isdigit(): codes.setdefault(c,[]).append(dt)
            print(name, {k:len(v) for k,v in codes.items()})

def ranges(dates, hol=None):
    dates=sorted(dates); out=[]
    for d in dates:
        if out and (d-out[-1][1]).days<=1: out[-1][1]=d
        else: out.append([d,d])
    return out
