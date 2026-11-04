# Пятнадцатый этап (04.11.2026): правки Мурата в ТД Серикжановой, перенесённые в эталон:
# 1) таблица сдвинута влево (tblInd -318), правая колонка (RU) шире: 4673 -> 5110;
# 2) в преамбуле жирным остаются только название Фонда и термины «Employer» / «Работодатель» и/или «Фонд»;
# (пустой абзац в блоке подписей Приложения 1 добавляется генератором build_itd_final.py).
import zipfile, re, os, sys, io
from xml.dom import minidom
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
P = r"D:\Шаблоны\ИТД\Шаблон ИТД ЦАКФ - эталон (проект 2026-11-03).docx"
M = r"D:\Кадры\ТД предитоговые\_правки Мурата в ТД Серикжановой\ИТД - Серикжанова А.Т. (правки Мурата 04.11 11-26).docx"
zi = zipfile.ZipFile(P); items = [(i, zi.read(i.filename)) for i in zi.infolist()]; zi.close()
x = dict((i.filename, d) for i, d in items)["word/document.xml"].decode("utf8")
xm = zipfile.ZipFile(M).read("word/document.xml").decode("utf8")
ROW = re.compile(r"<w:tr[ >].*?</w:tr>", re.S)
assert x.count('<w:tcW w:w="4673" w:type="dxa"/>') == 199 and "<w:tblInd" not in x
x = x.replace('<w:tcW w:w="4673" w:type="dxa"/>', '<w:tcW w:w="5110" w:type="dxa"/>')
assert x.count('<w:gridCol w:w="4673"/>') == 1
x = x.replace('<w:gridCol w:w="4673"/>', '<w:gridCol w:w="5110"/>')
a = '<w:tblW w:w="0" w:type="auto"/>'
assert x.count(a) == 1
x = x.replace(a, a + '<w:tblInd w:w="-318" w:type="dxa"/>')
# строка 1 (преамбула) - из файла Мурата (в ней нет данных конкретного работника)
r_e = list(ROW.finditer(x))[1]; r_m = list(ROW.finditer(xm))[1]
assert re.sub(r"<[^>]+>", "", r_e.group(0)) == re.sub(r"<[^>]+>", "", r_m.group(0)), "текст преамбулы отличается"
x = x[:r_e.start()] + r_m.group(0) + x[r_e.end():]
minidom.parseString(x.encode("utf8"))
with zipfile.ZipFile(P + ".tmp", "w", zipfile.ZIP_DEFLATED) as zo:
    for i, d in items:
        if i.filename == "word/document.xml": d = x.encode("utf8")
        nz = zipfile.ZipInfo(i.filename, i.date_time); nz.compress_type = zipfile.ZIP_DEFLATED; nz.external_attr = i.external_attr
        zo.writestr(nz, d)
os.replace(P + ".tmp", P)
print("ok")
