# Девятый этап: основной цвет текста - тёмно-серый 404040 (экономия краски при печати), 04.11.2026
# Цвет задаётся в docDefaults styles.xml; явных цветов в документе нет, поэтому меняется весь текст, включая нумерацию и колонтитул.
import zipfile, os, sys, io
from xml.dom import minidom
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
P = r"D:\Шаблоны\ИТД\Шаблон ИТД ЦАКФ - эталон (проект 2026-11-03).docx"
zi = zipfile.ZipFile(P); items = [(i, zi.read(i.filename)) for i in zi.infolist()]; zi.close()
s = dict((i.filename, d) for i, d in items)["word/styles.xml"].decode("utf8")
a = '<w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="ru-RU"'
assert s.count(a) == 1 and "w:val=\"404040\"" not in s
s = s.replace(a, '<w:color w:val="404040"/>' + a)
minidom.parseString(s.encode("utf8"))
with zipfile.ZipFile(P + ".tmp", "w", zipfile.ZIP_DEFLATED) as zo:
    for i, d in items:
        if i.filename == "word/styles.xml":
            d = s.encode("utf8")
        n = zipfile.ZipInfo(i.filename, i.date_time); n.compress_type = zipfile.ZIP_DEFLATED; n.external_attr = i.external_attr
        zo.writestr(n, d)
os.replace(P + ".tmp", P)
print("ok")
