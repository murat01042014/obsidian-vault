# Пятый этап: вычитка английского текста эталона (03.11.2026): соответствие русскому тексту и терминам AIFC Employment Regulations
import zipfile, re, os, sys, io
from xml.dom import minidom
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
P = r"D:\Шаблоны\ИТД\Шаблон ИТД ЦАКФ - эталон (проект 2026-11-03).docx"
zi = zipfile.ZipFile(P); items = [(i, zi.read(i.filename)) for i in zi.infolist()]; zi.close()
x = dict((i.filename, d) for i, d in items)["word/document.xml"].decode("utf8")
def esc(s): return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

SWAPS = [
    ("resident of Republic of Kazakhstan", "a citizen of the Republic of Kazakhstan"),     # RU: гражданин РК
    ("with permanent residency at", "residing at"),                                          # RU: проживающий (-ая) по адресу
    ("this employment agreement", "this Contract of Employment"),
    ("Annual leave continues", "Vacation leave continues"),
    ("paid annual leaves (vacation)", "paid vacation leave"),
    ("to provide annual paid leave", "to provide paid vacation leave"),
    ("12 months period", "12-month period"),
    ("shall be paid a sick pay", "shall be paid sick pay"),
    ("Employee is entitled to a payment in lieu of vacation leave earned", "The Employee is entitled to a payment in lieu of vacation leave earned"),
    (" An Employer may require an Employee to take", "An Employer may require an Employee to take"),
    ("accepts the Employee to a position of", "appoints the Employee to the position of"),
    ("In case that bank account indicated by the Employee is not the bank account of Kazakhstan second-tier bank, all relevant expenses for the operations on transfer and cash withdrawal shall be borne by the Employee solely.",
     "If the bank account indicated by the Employee is not an account with a Kazakhstan second-tier bank, all related expenses for transfer and cash withdrawal operations shall be borne solely by the Employee."),
    ("obey the orders of the Employer", "to obey the orders of the Employer"),
    ("comply with the health and safety instructions", "to comply with the health and safety instructions"),
    ("not accept any gift or advantage", "not to accept any gift or advantage"),
    ("in accordance with subclause 18.1.5", "in accordance with clause 18.1.5"),
]
log = {}
def swap_text(m):
    t = m.group(2); o = t
    for a, b in SWAPS:
        if a in t:
            t = t.replace(a, esc(b)); log[a] = log.get(a, 0) + 1
    return m.group(1) + t + m.group(3) if t != o else m.group(0)
x = re.sub(r"(<w:t(?: [^>]*)?>)([^<]*)(</w:t>)", swap_text, x)
for a, b in SWAPS:
    print(f"{log.get(a, 0):>2} x  {a[:70]}")
minidom.parseString(x.encode("utf8"))
with zipfile.ZipFile(P + ".tmp", "w", zipfile.ZIP_DEFLATED) as zo:
    for i, d in items:
        if i.filename == "word/document.xml":
            d = x.encode("utf8")
        n = zipfile.ZipInfo(i.filename, i.date_time); n.compress_type = zipfile.ZIP_DEFLATED; n.external_attr = i.external_attr
        zo.writestr(n, d)
os.replace(P + ".tmp", P)
