# Десятый этап: отступ перед блоком подписей «РАБОТОДАТЕЛЬ / EMPLOYER» (строка 163), 04.11.2026
import zipfile, re, os, sys, io
from xml.dom import minidom
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
P = r"D:\Шаблоны\ИТД\Шаблон ИТД ЦАКФ - эталон (проект 2026-11-03).docx"
BEFORE = 240
zi = zipfile.ZipFile(P); items = [(i, zi.read(i.filename)) for i in zi.infolist()]; zi.close()
x = dict((i.filename, d) for i, d in items)["word/document.xml"].decode("utf8")
rows = list(re.finditer(r"<w:tr[ >].*?</w:tr>", x, re.S)); assert len(rows) == 199
r = rows[163].group(0); cells = re.findall(r"<w:tc>.*?</w:tc>", r, re.S); assert len(cells) == 2
nr = r
for c in cells:
    p = re.search(r"<w:p[ >].*?</w:p>", c, re.S).group(0)
    t = re.sub(r"<[^>]+>", "", p).strip(); print(repr(t))
    assert t in ("EMPLOYER", "РАБОТОДАТЕЛЬ")
    p2, n = re.subn(r'(<w:spacing )w:before="\d+"', r'\g<1>w:before="%d"' % BEFORE, p, count=1); assert n == 1
    nr = nr.replace(p, p2, 1)
x2 = x[:rows[163].start()] + nr + x[rows[163].end():]
minidom.parseString(x2.encode("utf8"))
with zipfile.ZipFile(P + ".tmp", "w", zipfile.ZIP_DEFLATED) as zo:
    for i, d in items:
        if i.filename == "word/document.xml": d = x2.encode("utf8")
        n = zipfile.ZipInfo(i.filename, i.date_time); n.compress_type = zipfile.ZIP_DEFLATED; n.external_attr = i.external_attr
        zo.writestr(n, d)
os.replace(P + ".tmp", P)
print("ok")
