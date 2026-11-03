# Шестой этап: лишние пробелы в заголовках (п. 12 - пробел после номера; п. 2 - пробел в конце), 03.11.2026
import zipfile, re, os, sys, io
from xml.dom import minidom
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
P = r"D:\Шаблоны\ИТД\Шаблон ИТД ЦАКФ - эталон (проект 2026-11-03).docx"
zi = zipfile.ZipFile(P); items = [(i, zi.read(i.filename)) for i in zi.infolist()]; zi.close()
x = dict((i.filename, d) for i, d in items)["word/document.xml"].decode("utf8")
FIX = [(" RIGHTS AND OBLIGATIONS OF THE EMPLOYER", "RIGHTS AND OBLIGATIONS OF THE EMPLOYER"),
       (" ПРАВА И ОБЯЗАННОСТИ РАБОТОДАТЕЛЯ", "ПРАВА И ОБЯЗАННОСТИ РАБОТОДАТЕЛЯ"),
       ("COMMENCEMENT DATE AND TERM ", "COMMENCEMENT DATE AND TERM"),
       ("ДАТА НАЧАЛА И СРОК ДЕЙСТВИЯ ", "ДАТА НАЧАЛА И СРОК ДЕЙСТВИЯ")]
for a, b in FIX:
    pat = re.compile(r"(<w:t(?: [^>]*)?>)" + re.escape(a) + r"(</w:t>)")
    n = len(pat.findall(x)); print(n, "x", repr(a))
    assert n == 1
    x = pat.sub(lambda m: m.group(1) + b + m.group(2), x)
minidom.parseString(x.encode("utf8"))
with zipfile.ZipFile(P + ".tmp", "w", zipfile.ZIP_DEFLATED) as zo:
    for i, d in items:
        if i.filename == "word/document.xml":
            d = x.encode("utf8")
        n = zipfile.ZipInfo(i.filename, i.date_time); n.compress_type = zipfile.ZIP_DEFLATED; n.external_attr = i.external_attr
        zo.writestr(n, d)
os.replace(P + ".tmp", P)
