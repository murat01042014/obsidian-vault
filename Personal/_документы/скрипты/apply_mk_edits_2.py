# Добавка к apply_mk_edits.py: места, разбитые по нескольким run («Индивидуальный» + «договор»; кавычки в п. 2.2)
import zipfile, os, sys
from xml.dom import minidom
P = r"D:\Шаблоны\ИТД\Шаблон ИТД ЦАКФ - эталон (проект 2026-11-03).docx"
zi = zipfile.ZipFile(P); items = [(i, zi.read(i.filename)) for i in zi.infolist()]; zi.close()
def fix(x):
    for a, b in (("Индивидуальный </w:t>", "Трудовой </w:t>"),
                 ('(включительно) (далее "', "(включительно) (далее – «"),
                 ('"), если он не будет прекращен', "»), если он не будет прекращен")):
        assert x.count(a) == 1, a
        x = x.replace(a, b)
    return x
with zipfile.ZipFile(P + ".tmp", "w", zipfile.ZIP_DEFLATED) as zo:
    for i, d in items:
        if i.filename == "word/document.xml":
            t = fix(d.decode("utf8")); minidom.parseString(t.encode("utf8")); d = t.encode("utf8")
        n = zipfile.ZipInfo(i.filename, i.date_time); n.compress_type = zipfile.ZIP_DEFLATED; n.external_attr = i.external_attr
        zo.writestr(n, d)
os.replace(P + ".tmp", P); print("ok")
