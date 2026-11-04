# Четырнадцатый этап (04.11.2026): ФИО ГИД полностью, с отчеством: Қаражан Ерлік Нариманұлы / Karazhan Yerlik Narimanuly
# (порядок «фамилия, имя, отчество» - как у Работника в блоке подписей).
import zipfile, os, sys, io
from xml.dom import minidom
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
P = r"D:\Шаблоны\ИТД\Шаблон ИТД ЦАКФ - эталон (проект 2026-11-03).docx"
zi = zipfile.ZipFile(P); items = [(i, zi.read(i.filename)) for i in zi.infolist()]; zi.close()
x = dict((i.filename, d) for i, d in items)["word/document.xml"].decode("utf8")
SW = [("Yerlik Karazhan", "Karazhan Yerlik Narimanuly"), ("Ерлік Қаражан", "Қаражан Ерлік Нариманұлы")]
for a, b in SW:
    n = x.count(a); print(n, "x", a); assert n == 4
    x = x.replace(a, b)
minidom.parseString(x.encode("utf8"))
with zipfile.ZipFile(P + ".tmp", "w", zipfile.ZIP_DEFLATED) as zo:
    for i, d in items:
        if i.filename == "word/document.xml": d = x.encode("utf8")
        nz = zipfile.ZipInfo(i.filename, i.date_time); nz.compress_type = zipfile.ZIP_DEFLATED; nz.external_attr = i.external_attr
        zo.writestr(nz, d)
os.replace(P + ".tmp", P)
