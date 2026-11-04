# Восьмой этап: в блоке подписей основного договора указаны два адреса Фонда (юридический и фактический), 04.11.2026
import zipfile, os, sys, io
from xml.dom import minidom
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
P = r"D:\Шаблоны\ИТД\Шаблон ИТД ЦАКФ - эталон (проект 2026-11-03).docx"
zi = zipfile.ZipFile(P); items = [(i, zi.read(i.filename)) for i in zi.infolist()]; zi.close()
x = dict((i.filename, d) for i, d in items)["word/document.xml"].decode("utf8")
SW = [("legal address: Astana, Syganak street, 70, office 701, Republic of Kazakhstan:",
       "legal address: Astana, Mangilik El avenue, 11/1, office 704, Republic of Kazakhstan; actual address: Astana, Syganak street, 70, office 701, Republic of Kazakhstan:"),
      ("юридический адрес: Республика Казахстан, г. Астана, ул. Сыганак, д. 70, каб. 701:",
       "юридический адрес: Республика Казахстан, г. Астана, пр. Мангилик Ел, д. 11/1, офис 704; фактический адрес: Республика Казахстан, г. Астана, ул. Сыганак, д. 70, офис 701:")]
for a, b in SW:
    n = x.count(a); print(n, "x", a[:50]); assert n == 1
    x = x.replace(a, b)
minidom.parseString(x.encode("utf8"))
with zipfile.ZipFile(P + ".tmp", "w", zipfile.ZIP_DEFLATED) as zo:
    for i, d in items:
        if i.filename == "word/document.xml":
            d = x.encode("utf8")
        n = zipfile.ZipInfo(i.filename, i.date_time); n.compress_type = zipfile.ZIP_DEFLATED; n.external_attr = i.external_attr
        zo.writestr(n, d)
os.replace(P + ".tmp", P)
