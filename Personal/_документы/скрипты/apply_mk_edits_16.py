# Шестнадцатый этап (04.11.2026): п. 2.2 - вся фраза о датах жирным (как в правке Мурата в ТД Абенова):
# «Employment commencement date: 2 November 2026, end date: 1 November 2027» / «Дата начала: с 02.11.2026 г., окончания трудового договора: 01.11.2027 г.»
import zipfile, re, os, sys, io
from xml.dom import minidom
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
P = r"D:\Шаблоны\ИТД\Шаблон ИТД ЦАКФ - эталон (проект 2026-11-03).docx"
zi = zipfile.ZipFile(P); items = [(i, zi.read(i.filename)) for i in zi.infolist()]; zi.close()
x = dict((i.filename, d) for i, d in items)["word/document.xml"].decode("utf8")
rows = list(re.finditer(r"<w:tr[ >].*?</w:tr>", x, re.S)); assert len(rows) == 199
r = rows[15].group(0); nr = r; cnt = 0
for lab in ("Employment commencement date: ", "Дата начала: "):
    pat = re.compile(r'<w:r(?: [^>]*)?>(<w:rPr>(?:(?!</w:rPr>).)*</w:rPr>)(?:<w:tab/>)?<w:t(?: [^>]*)?>' + re.escape(lab) + r'</w:t></w:r>', re.S)
    m = pat.search(nr); assert m, lab
    run = m.group(0); rpr = m.group(1)
    if "<w:b/>" not in rpr:
        rpr2 = re.sub(r'(<w:rFonts [^>]*/>)', r'\g<1><w:b/><w:bCs/>', rpr, count=1); assert rpr2 != rpr
        nr = nr.replace(run, run.replace(rpr, rpr2, 1), 1); cnt += 1
print("изменено runs:", cnt)
x2 = x[:rows[15].start()] + nr + x[rows[15].end():]
minidom.parseString(x2.encode("utf8"))
with zipfile.ZipFile(P + ".tmp", "w", zipfile.ZIP_DEFLATED) as zo:
    for i, d in items:
        if i.filename == "word/document.xml": d = x2.encode("utf8")
        nz = zipfile.ZipInfo(i.filename, i.date_time); nz.compress_type = zipfile.ZIP_DEFLATED; nz.external_attr = i.external_attr
        zo.writestr(nz, d)
os.replace(P + ".tmp", P)
