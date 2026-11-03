# Третий этап: «Дополнительно к шаблону ТД» (М.К. Оразбаева, 03.11.2026) - п. 18.4.3, 18.4.4, 18.5, 18.6 + согласование п. 18.1.5
import zipfile, re, os, sys, io
from xml.dom import minidom
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
P = r"D:\Шаблоны\ИТД\Шаблон ИТД ЦАКФ - эталон (проект 2026-11-03).docx"
zi = zipfile.ZipFile(P); items = [(i, zi.read(i.filename)) for i in zi.infolist()]; zi.close()
x = dict((i.filename, d) for i, d in items)["word/document.xml"].decode("utf8")
def esc(s): return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
ROW = re.compile(r"<w:tr[ >].*?</w:tr>", re.S); CELL = re.compile(r"<w:tc>.*?</w:tc>", re.S); PAR = re.compile(r"<w:p[ >].*?</w:p>", re.S)
def set_para(cell, text):
    para = PAR.search(cell).group(0)
    open_tag = re.match(r"<w:p[^>]*>", para).group(0)
    ppr = re.search(r"<w:pPr>.*?</w:pPr>", para, re.S).group(0)
    runs = re.findall(r"<w:r[ >].*?</w:r>", para, re.S)
    r0 = runs[0]; rtab = next((r for r in runs[1:] if "<w:tab/>" in r), runs[1])
    rpr = re.search(r"<w:rPr>.*?</w:rPr>", rtab, re.S).group(0)
    return cell.replace(para, open_tag + ppr + r0 + "<w:r>" + rpr + '<w:tab/><w:t xml:space="preserve">' + esc(text) + "</w:t></w:r></w:p>", 1)
TXT = {
 133: ("violation of this Contract of Employment and the Foundation’s By-laws; termination without notice is permitted only on the grounds set out in clauses 18.4.3 and 18.5;",
       "нарушения условий настоящего Трудового договора и Положения Фонда; увольнение без уведомления допускается только по основаниям, указанным в пунктах 18.4.3 и 18.5;"),
 143: ("the Employee terminates the employment for cause;", "Работник прекращает работу по уважительной причине;"),
 144: ("The Employer terminates the employment for cause on the grounds specified in clause 18.5.", "Работодатель увольняет Работника по основаниям, указанным в пункте 18.5."),
 145: ("An Employer may terminate an Employee’s employment without notice for cause on the following grounds:", "Работодатель может уволить Работника без уведомления по следующим основаниям:"),
 150: ("Termination for cause under clauses 18.4.3 and 18.5 does not include termination under clause 9.4.", "Расторжение по основаниям, указанным в пунктах 18.4.3 и 18.5, не включает расторжение в соответствии с пунктом 9.4."),
}
rows = [m for m in ROW.finditer(x)]; assert len(rows) == 199
new = {}
for i, (en, ru) in TXT.items():
    r = rows[i].group(0); c = CELL.findall(r)
    new[i] = r.replace(c[0], set_para(c[0], en), 1).replace(c[1], set_para(c[1], ru), 1)
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
