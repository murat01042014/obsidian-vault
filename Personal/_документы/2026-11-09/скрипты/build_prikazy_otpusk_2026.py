# Приказы об очередных отпусках за август - октябрь 2026 по табелям (ТУРВ сотрудники / СОВЕТ), номера с 03-7/17.
# Шаблон: Order №3-7-23 отпуск АК очередной.docx (бланк, стили). Даты приказов и заявлений - жёлтым, их проставляет Мурат.
import zipfile, re, sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
TPL = r"D:\приказы\01_Отпуск_очередной\Order №3-7-23 отпуск АК очередной.docx"
OUTDIR = r"D:\приказы\01_Отпуск_очередной"
def esc(s): return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

zi = zipfile.ZipFile(TPL); items = [(i, zi.read(i.filename)) for i in zi.infolist()]; zi.close()
doc = dict((i.filename, d) for i, d in items)["word/document.xml"].decode("utf8")
head = doc[:doc.find("<w:body>") + 8]
body = doc[doc.find("<w:body>") + 8:doc.find("</w:body>")]
paras = re.findall(r"<w:p[ >].*?</w:p>", body, re.S)
sect = re.search(r"<w:sectPr.*?</w:sectPr>", body, re.S).group(0)
P_ORDER, P_ON, P_BLANK_B = paras[3], paras[4], paras[5]

def r(text, bold=False, blue=False, sz=28, hl=False, tabs=0):
    rpr = ("<w:b/>" if bold else "") + ('<w:color w:val="365F91" w:themeColor="accent1" w:themeShade="BF"/>' if blue else "") + \
          '<w:sz w:val="%d"/><w:szCs w:val="%d"/>' % (sz, sz) + ('<w:highlight w:val="yellow"/>' if hl else "") + '<w:lang w:val="en-US"/>'
    return '<w:r><w:rPr>%s</w:rPr>%s%s</w:r>' % (rpr, "<w:tab/>" * tabs, '<w:t xml:space="preserve">%s</w:t>' % esc(text) if text else "")
def p(runs, jc="both", spacing=True):
    sp = '<w:spacing w:line="259" w:lineRule="auto"/>' if spacing else ""
    return '<w:p><w:pPr>%s<w:jc w:val="%s"/></w:pPr>%s</w:p>' % (sp, jc, "".join(runs))
def blank(spacing=True): return p([], spacing=spacing)

def order_xml(num, date_txt, stmt_date, itd, itd_date, itd_par, who, title, frm, to, n_days, n_words, period, statement, duties):
    n = 1
    hdr = ('<w:p><w:pPr/>' + r(" ", sz=20) + r("No. ", True) + r("03-7/", True) + r(str(num), True) + r("", True, tabs=2) + r("    ", True) +
           r("", True, tabs=4) + r("Date: ", True) + r(date_txt, True, hl=date_txt.startswith("_")) + '</w:p>')
    pre = p([r("In accordance with the provisions of the AIFC Employment Regulations as well as with the terms of the Employment Agreement №%s of %s (paragraph %s), ORDER:" % (itd, itd_date, itd_par), True, True)])
    if duties:
        items = [p([r("%d. I am going on annual paid vacation leave from %s to %s (include), 2026, for a total of %d (%s) working days, for %s." % (n, frm, to, n_days, n_words, period))])]
    else:
        items = [p([r("%d. TO PROVIDE %s, %s, annual paid vacation leave from %s to %s (include), 2026, for a total of %d (%s) working days, for %s." % (n, who, title, frm, to, n_days, n_words, period))])]
    n += 1
    if duties:
        items.append(blank())
        items.append(p([r("%d. For the period of my leave, assign the duties of the Chief Executive Officer of the Central Asia Climate Foundation to the Deputy CEO on administrative issues - member of the Management Council Mr. Murat Temirbekov." % n)]))
        n += 1
        items.append(blank())
        items.append(p([r("%d. To establish for Mr. Murat Temirbekov an allowance of 30 percent of his base salary for the actual period of substitution (%s to %s, 2026), in accordance with Section 7-1 of the HR Policy of the Foundation." % (n, frm, to))]))
        n += 1
    items.append(blank())
    short = who.split(" ", 1)[1]
    if duties:
        items.append(p([r("%d. TO Chief Financial Officer - Chief Accountant: to pay %s for vacation leave and to calculate and pay the allowance to Mr. Murat Temirbekov." % (n, who))]))
    else:
        items.append(p([r("%d. TO Chief Financial Officer - Chief Accountant: to pay %s for vacation leave." % (n, who))]))
    basis = p([r("Basis: Statement of %s dated " % statement), r(stmt_date, hl=stmt_date.startswith("_")), r(".")])
    n_empty = 14 - (10 if duties else 0)
    sig = p([r("CEO of CACF", True, True), r("", True, True, tabs=6), r("     Yerlik KARAZHAN", True, True)])
    read = p([r("I have read the order ________________________ (Surname, name) (personal signature)", sz=20)])
    bd = (hdr + blank(False) + blank(False) + P_ORDER + P_ON + P_BLANK_B + pre + blank() + "".join(items) + blank() + basis +
          blank(False) + blank(False) + sig + blank(False) + blank(False) + "".join(blank() for _ in range(n_empty)) + read + blank() +
          p([r("Date", sz=20)]) + blank(False))
    return head + bd + sect + "</w:body></w:document>"

KAR = ("Mr. Yerlik Karazhan", "Chief Executive Officer of the Central Asia Climate Foundation", "1-1", "February 18, 2025", "6", "Mr. Y. Karazhan")
TEM = ("Mr. Murat Temirbekov", "Deputy Chief Executive Officer on administrative issues of the Central Asia Climate Foundation", "3-3", "February 24, 2025", "6.1", "Mr. M. Temirbekov")
ORA = ("Ms. Maira Orazbayeva", "Chief Financial Officer - Chief Accountant of the Central Asia Climate Foundation", "6-6", "March 3, 2025", "6.12", "Ms. M. Orazbayeva")
KON = ("Mr. Alexander Kononets", "Strategy and R&D Adviser", "10", "February 10, 2026", "6", "Mr. A. Kononets")
PER_KAR1 = "the period of work from February 18, 2025 to February 17, 2026"
PER_KAR2 = "the period of work from February 18, 2026 to February 17, 2027"
PER_KON = "the first year of work - from February 10, 2026 to February 9, 2027"
PER_ORA1 = "the period of work from March 3, 2025 to March 2, 2026"
PER_ORA2 = "the period of work from March 3, 2026 to March 2, 2027"
PER_TEM = "the period of work from February 24, 2026 to February 23, 2027"
# (номер, сотрудник, с, по, дней, прописью, период работы, возложение обязанностей, имя файла)
LIST = [
 (17, KAR, "August 12", "August 19", 6, "six", PER_KAR1, True, "Каражан 12-19.08"),
 (18, KON, "August 24", "September 4", 10, "ten", PER_KON, False, "Кононец 24.08-04.09"),
 (19, KAR, "August 27", "September 2", 5, "five", PER_KAR2, True, "Каражан 27.08-02.09"),
 (20, ORA, "September 1", "September 8", 6, "six", PER_ORA1, False, "Оразбаева 01-08.09"),
 (21, TEM, "September 7", "September 18", 10, "ten", PER_TEM, False, "Темирбеков 07-18.09"),
 (22, ORA, "September 16", "September 29", 10, "ten", PER_ORA2, False, "Оразбаева 16-29.09"),
 (24, KAR, "October 7", "October 14", 6, "six", PER_KAR2, True, "Каражан 07-14.10"),
]
DATES = {17: "August 10, 2026", 18: "August 21, 2026", 19: "August 24, 2026", 20: "August 31, 2026", 21: "September 4, 2026", 22: "September 13, 2026", 24: "October 5, 2026"}
if __name__ == "__main__":
    from xml.dom import minidom
    for num, E, frm, to, nd, nw, per, duties, tag in LIST:
        who, title, itd, itd_date, itd_par, stmt = E
        x = order_xml(num, DATES[num], DATES[num], itd, itd_date, itd_par, who, title, frm, to, nd, nw, per, stmt, duties)
        minidom.parseString(x.encode("utf8"))
        out = os.path.join(OUTDIR, "Order №3-7-%d отпуск %s (проект 2026-11-06).docx" % (num, tag))
        for suf in ("", " (v2)", " (v3)", " (v4)"):
            cand = out.replace(".docx", suf + ".docx")
            if not os.path.exists(cand):
                out = cand; break
            try:
                open(cand, "r+b").close(); out = cand; break
            except PermissionError:
                continue
        if len(sys.argv) > 1 and str(num) not in sys.argv[1:]: continue
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
            for i, d in items:
                if i.filename == "word/document.xml": d = x.encode("utf8")
                elif i.filename == "docProps/app.xml":
                    s = d.decode("utf8")
                    s = re.sub(r"<Company>.*?</Company>", "<Company>CACF</Company>", s) if "<Company>" in s else s.replace("</Properties>", "<Company>CACF</Company></Properties>")
                    d = s.encode("utf8")
                nz = zipfile.ZipInfo(i.filename, i.date_time); nz.compress_type = zipfile.ZIP_DEFLATED; nz.external_attr = i.external_attr
                zo.writestr(nz, d)
        print("готово:", os.path.basename(out))
