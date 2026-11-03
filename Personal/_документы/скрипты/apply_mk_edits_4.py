# Четвертый этап: восстановление точки после буквы/номера в русских ячейках A. и 8. (в них метка и точка - отдельные run)
import zipfile, re, os, sys, io
from xml.dom import minidom
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
P = r"D:\Шаблоны\ИТД\Шаблон ИТД ЦАКФ - эталон (проект 2026-11-03).docx"
OLD = r"D:\Шаблоны\ИТД\_копии\Шаблон ИТД ЦАКФ - эталон - до правок МК 2 (2026-11-03).docx"
ROW = re.compile(r"<w:tr[ >].*?</w:tr>", re.S); CELL = re.compile(r"<w:tc>.*?</w:tc>", re.S); RUN = re.compile(r"<w:r[ >].*?</w:r>", re.S)
old_rows = ROW.findall(zipfile.ZipFile(OLD).read("word/document.xml").decode("utf8"))
zi = zipfile.ZipFile(P); items = [(i, zi.read(i.filename)) for i in zi.infolist()]; zi.close()
x = dict((i.filename, d) for i, d in items)["word/document.xml"].decode("utf8")
rows = list(ROW.finditer(x)); assert len(rows) == 199
new = {}
for i in (1, 49):
    dot = RUN.findall(CELL.findall(old_rows[i])[1])[1]           # старый run с точкой
    assert re.sub(r"<[^>]+>", "", dot) == ".", dot[:80]
    r = rows[i].group(0); c = CELL.findall(r)[1]
    r0 = RUN.findall(c)[0]
    assert "." not in re.sub(r"<[^>]+>", "", RUN.findall(c)[1]).split("\t")[0][:1]
    new[i] = r.replace(c, c.replace(r0, r0 + dot, 1), 1)
out = []; pos = 0
for i, m in enumerate(rows):
    out.append(x[pos:m.start()]); out.append(new.get(i, m.group(0))); pos = m.end()
out.append(x[pos:]); x2 = "".join(out); minidom.parseString(x2.encode("utf8"))
with zipfile.ZipFile(P + ".tmp", "w", zipfile.ZIP_DEFLATED) as zo:
    for i, d in items:
        if i.filename == "word/document.xml": d = x2.encode("utf8")
        n = zipfile.ZipInfo(i.filename, i.date_time); n.compress_type = zipfile.ZIP_DEFLATED; n.external_attr = i.external_attr
        zo.writestr(n, d)
os.replace(P + ".tmp", P); print("ok")
