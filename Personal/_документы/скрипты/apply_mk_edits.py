# Второй этап правок эталона ИТД: замечания М.К. Оразбаевой к ТД Серикжановой + решения Мурата 03.11.2026
# (принять все, где нет противоречий между пунктами; адрес Фонда - только в конце договора;
# «Индивидуальный договор» -> «Трудовой договор»; английские термины - как в AIFC Employment Regulations).
import zipfile, re, io, sys, os
from xml.dom import minidom
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
P = r"D:\Шаблоны\ИТД\Шаблон ИТД ЦАКФ - эталон (проект 2026-11-03).docx"
zi = zipfile.ZipFile(P); items = [(i, zi.read(i.filename)) for i in zi.infolist()]; zi.close()
x = dict((i.filename, d) for i, d in items)["word/document.xml"].decode("utf8")

def esc(s): return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
def T(p): return re.sub(r"<[^>]+>", "", p)
ROW = re.compile(r"<w:tr[ >].*?</w:tr>", re.S); CELL = re.compile(r"<w:tc>.*?</w:tc>", re.S)
PAR = re.compile(r"<w:p[ >].*?</w:p>", re.S)

# 1) простые замены терминов внутри w:t (без выделенных маркеров)
SWAPS = [
    ("Индивидуальный договор", "Трудовой договор"),
    ("Individual agreement", "Contract of Employment"),
    ("Employment Agreement", "Contract of Employment"),
    ("Payment Period", "Pay Period"),
    ("Annex 1", "Schedule 1"),
    ("the Employment Regulations", "the AIFC Employment Regulations"),
    ("by Employment Regulations", "by the AIFC Employment Regulations"),
    ("5 Рабочих дней", "5 рабочих дней"),
    ('(далее "Срок работы")', "(далее – «Срок работы»)"),
    ("бонусов", "премий"),
    ("базовой заработной платы", "должностного оклада"),
    ("пособие по болезни", "пособие по временной нетрудоспособности"),
    ("the Employee have read", "the Employee has read"),
    ("the Employee have been given", "the Employee has been given"),
    ("Notwithstanding 18.2", "Notwithstanding clause 18.2"),
    ("The Employee’s wages consist", "The Employee’s Wages consist"),
    ("not included in the wages", "not included in the Wages"),
    ("в дополнение к праздничным дням", "в дополнение к Праздничным дням"),
    ("дневную заработную плату", "дневной заработок"),
    ("содержащей Конфиденциальную информацию, Работнику становится известно", "содержащей Конфиденциальную информацию, о которых Работнику становится известно"),
    ("в ходе работы у Работодателя, в результате работы у Работодателя", "в ходе работы у Работодателя или в результате работы у Работодателя"),
    ("in the course of employment with the Employer, as a result of", "in the course of employment with the Employer, or as a result of"),
    ("Если праздничный день выпадает", "Если Праздничный день выпадает"),
    ("продлевается на период праздничного дня", "продлевается на период Праздничного дня"),
    ("The Employee shall work workweek", "The Employee shall work a workweek"),
    ("работать по адресу 5 дней", "работать 5 дней"),
    ("добровольного согласия Работника на работу в его выходной день", "добровольного, информированного предварительного письменного согласия Работника"),
    ("при условии предварительного письменного согласия Работника", "при условии добровольного, информированного предварительного письменного согласия Работника"),
    ("subject to an Employee’s freely given day off, informed prior written consent", "subject to the Employee’s freely given, informed prior written consent"),
    ("the Employee’s freely given day off, informed", "the Employee’s freely given, informed"),
    ("personal account and password details, and", "personal account and password details."),
    ("данные личного аккаунта и пароль, а также", "данные личного аккаунта и пароль."),
    ("увольнения Работника в соответствии с пунктом 9.4;", "трудовые отношения Работника прекращаются в соответствии с пунктом 9.4;"),
    ("по основанию for cause, указанному в пункте 18.5;", "по основанию, указанному в пункте 18.5."),
    ("по иной причине, предусмотренной Правилами по трудовым отношениям МФЦА или настоящим Трудовым договором.", "по иной причине, предусмотренной настоящим Трудовым договором."),
    ("for another reason prescribed by the AIFC Employment Regulations or by this Contract of Employment.", "for another reason prescribed by this Contract of Employment."),
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
    if a not in log:
        print("НЕ применено (разбито по runs или уже заменено):", a[:70])

# 2) замены целых абзацев в ячейках (номер и форматирование сохраняются)
rows = [m for m in ROW.finditer(x)]
assert len(rows) == 199, len(rows)

def set_para(cell, text):
    para = PAR.search(cell).group(0)
    open_tag = re.match(r"<w:p[^>]*>", para).group(0)
    ppr = re.search(r"<w:pPr>.*?</w:pPr>", para, re.S).group(0)
    runs = re.findall(r"<w:r[ >].*?</w:r>", para, re.S)
    r0 = runs[0]
    rtab = next((r for r in runs[1:] if "<w:tab/>" in r), runs[1])
    rpr = re.search(r"<w:rPr>.*?</w:rPr>", rtab, re.S).group(0)
    r1 = "<w:r>" + rpr + '<w:tab/><w:t xml:space="preserve">' + esc(text) + "</w:t></w:r>"
    return cell.replace(para, open_tag + ppr + r0 + r1 + "</w:p>", 1)

def set_par_text(cell, contains, text):
    """Замена текста абзаца без номера (блоки подписей): одна run с форматированием первой run."""
    for p in PAR.findall(cell):
        if contains in T(p):
            open_tag = re.match(r"<w:p[^>]*>", p).group(0)
            ppr = re.search(r"<w:pPr>.*?</w:pPr>", p, re.S)
            ppr = ppr.group(0) if ppr else ""
            r = re.search(r"<w:r[ >].*?</w:r>", p, re.S).group(0)
            rpr = re.search(r"<w:rPr>.*?</w:rPr>", r, re.S)
            rpr = rpr.group(0) if rpr else ""
            return cell.replace(p, open_tag + ppr + "<w:r>" + rpr + '<w:t xml:space="preserve">' + esc(text) + "</w:t></w:r></w:p>", 1)
    raise AssertionError(contains)

TXT = {
 1: ("Central Asia Climate Foundation, a Foundation incorporated on 22 February 2024 in the AIFC, BIN 240240900538 (hereinafter referred to as the “Employer” and/or the “Foundation”), represented by Mr. Yerlik Karazhan, appointed Chief Executive Officer of the Foundation",
     "Центрально-Азиатский Климатический Фонд, Фонд, зарегистрированный 22 февраля 2024 года на территории МФЦА, БИН 240240900538 (далее – «Работодатель» и/или «Фонд»), в лице г-на Ерлік Қаражан, Главного исполнительного директора Фонда"),
 49: ("NATIONAL HOLIDAYS AND DAYS OFF", "ПРАЗДНИЧНЫЕ И ВЫХОДНЫЕ ДНИ"),
 50: ("The Employee is entitled to days off (Saturday and Sunday), as well as to National Holidays and other non-working days announced for the private sector in accordance with the legislation of the Republic of Kazakhstan.",
      "Работнику предоставляются выходные дни (суббота и воскресенье), а также праздничные и иные нерабочие дни, объявленные для частного сектора в соответствии с законодательством Республики Казахстан (далее – «Праздничные дни»)."),
 51: ("A National Holiday that falls on a Business Day is paid at the Employee’s Daily Wage.",
      "Праздничный день, выпадающий на рабочий день, оплачивается в размере дневного заработка Работника."),
 52: ("Work on days off and National Holidays is permitted with the Employee’s written consent or at the Employee’s request on the basis of an act of the Employer. For such work, at the Employee’s choice, the Employee is granted another day off or is paid in the amount established by the labour legislation of the Republic of Kazakhstan.",
      "Работа в выходные и Праздничные дни допускается с письменного согласия Работника или по его просьбе на основании акта Работодателя. За такую работу Работнику по его желанию предоставляется другой день отдыха или производится оплата в размере, установленном трудовым законодательством Республики Казахстан."),
 122: ("The Employer makes contributions to the Social Health Insurance Fund in accordance with the legislation of the Republic of Kazakhstan on compulsory social health insurance.",
       "Работодатель производит отчисления и взносы в Фонд социального медицинского страхования в соответствии с законодательством Республики Казахстан об обязательном социальном медицинском страховании."),
 132: ("expiration of the Employment Term specified in clause 2.2, taking into account clause 2.3;",
       "истечения Срока работы, указанного в пункте 2.2, с учетом пункта 2.3;"),
}
new = {}
for i, (en, ru) in TXT.items():
    r = rows[i].group(0); c = CELL.findall(r)
    new[i] = r.replace(c[0], set_para(c[0], en), 1).replace(c[1], set_para(c[1], ru), 1)

# блоки подписей: адрес Фонда остается только в конце основного договора (строка 163)
ADDR = {
 163: ("Central Asia Climate Foundation, a Foundation incorporated on 22 February 2024 in the AIFC, BIN 240240900538, legal address: Astana, Syganak street, 70, office 701, Republic of Kazakhstan:",
       "Центрально-Азиатский Климатический Фонд, Фонд, зарегистрированный 22 февраля 2024 года на территории МФЦА, БИН 240240900538, юридический адрес: Республика Казахстан, г. Астана, ул. Сыганак, д. 70, каб. 701:"),
 169: ("Central Asia Climate Foundation, a Foundation incorporated on 22 February 2024 in the AIFC, BIN 240240900538:",
       "Центрально-Азиатский Климатический Фонд, Фонд, зарегистрированный 22 февраля 2024 года на территории МФЦА, БИН 240240900538:"),
 198: ("Central Asia Climate Foundation, a Foundation incorporated on 22 February 2024 in the AIFC, BIN 240240900538:",
       "Центрально-Азиатский Климатический Фонд, Фонд, зарегистрированный 22 февраля 2024 года на территории МФЦА, БИН 240240900538:"),
}
for i, (en, ru) in ADDR.items():
    r = rows[i].group(0); c = CELL.findall(r)
    c0 = set_par_text(c[0], "Central Asia Climate Foundation", en)
    c1 = set_par_text(c[1], "Центрально-Азиатский", ru)
    new[i] = r.replace(c[0], c0, 1).replace(c[1], c1, 1)

out = []; pos = 0
for i, m in enumerate(rows):
    out.append(x[pos:m.start()]); out.append(new.get(i, m.group(0))); pos = m.end()
out.append(x[pos:]); x2 = "".join(out)
minidom.parseString(x2.encode("utf8"))
with zipfile.ZipFile(P + ".tmp", "w", zipfile.ZIP_DEFLATED) as zo:
    for i, d in items:
        if i.filename == "word/document.xml":
            d = x2.encode("utf8")
        n = zipfile.ZipInfo(i.filename, i.date_time); n.compress_type = zipfile.ZIP_DEFLATED; n.external_attr = i.external_attr
        zo.writestr(n, d)
os.replace(P + ".tmp", P)
print("применено замен:", sum(log.values()), "| строк", len(ROW.findall(x2)),
      "| абзацев с маркерами", len([p for p in PAR.findall(x2) if "w:highlight" in p]))
