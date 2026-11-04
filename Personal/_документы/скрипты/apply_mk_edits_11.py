# Одиннадцатый этап (04.11.2026): 1) реквизиты Фонда в блоках подписей (строки 163, 169, 198) - без жирного шрифта;
# 2) строки 196-197 Приложения 2 связаны с блоком подписей (keepNext), чтобы подписи не уходили на отдельный лист.
import zipfile, re, os, sys, io
from xml.dom import minidom
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
P = r"D:\Шаблоны\ИТД\Шаблон ИТД ЦАКФ - эталон (проект 2026-11-03).docx"
zi = zipfile.ZipFile(P); items = [(i, zi.read(i.filename)) for i in zi.infolist()]; zi.close()
x = dict((i.filename, d) for i, d in items)["word/document.xml"].decode("utf8")
rows = list(re.finditer(r"<w:tr[ >].*?</w:tr>", x, re.S)); assert len(rows) == 199
PAR = re.compile(r"<w:p[ >].*?</w:p>", re.S)
def T(p): return re.sub(r"<[^>]+>", "", p)
def unbold(p):
    return p.replace("<w:b/><w:bCs/>", "")
def keep(p):
    if "<w:keepNext" in p: return p
    return p.replace("<w:pPr>", "<w:pPr><w:keepNext/>", 1)
new = {}
for i in (163, 169, 198):
    r = rows[i].group(0); n = 0
    def f(m):
        global n
        p = m.group(0); t = T(p)
        if t.startswith(("Central Asia Climate Foundation", "Центрально-Азиатский Климатический Фонд")):
            n += 1; return unbold(p)
        return p
    new[i] = PAR.sub(f, r); assert n == 2, (i, n)
for i in (196, 197):
    r = rows[i].group(0)
    new[i] = PAR.sub(lambda m: keep(m.group(0)), r)
out = []; pos = 0
for i, m in enumerate(rows):
    out.append(x[pos:m.start()]); out.append(new.get(i, m.group(0))); pos = m.end()
out.append(x[pos:]); x2 = "".join(out)
minidom.parseString(x2.encode("utf8"))
with zipfile.ZipFile(P + ".tmp", "w", zipfile.ZIP_DEFLATED) as zo:
    for i, d in items:
        if i.filename == "word/document.xml": d = x2.encode("utf8")
        nz = zipfile.ZipInfo(i.filename, i.date_time); nz.compress_type = zipfile.ZIP_DEFLATED; nz.external_attr = i.external_attr
        zo.writestr(nz, d)
os.replace(P + ".tmp", P)
print("ok")
