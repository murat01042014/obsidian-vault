# Заявления Каражана Е.Н. на очередной отпуск (на имя Верховного Попечителя) по образцу «заявление ЕН на отпуск рус.docx»
import zipfile, re, sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
TPL = r"D:\приказы\11_Заявления_и_согласия_работников\заявление ЕН на отпуск рус.docx"
OUTDIR = r"D:\приказы\11_Заявления_и_согласия_работников"
def esc(s): return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
zi = zipfile.ZipFile(TPL); items = [(i, zi.read(i.filename)) for i in zi.infolist()]; zi.close()
doc = dict((i.filename, d) for i, d in items)["word/document.xml"].decode("utf8")
def set_par(x, idx, text):
    ps = list(re.finditer(r"<w:p[ >].*?</w:p>", x[x.find("<w:body>"):], re.S))
    off = x.find("<w:body>")
    m = ps[idx]; p = m.group(0)
    first = re.search(r"<w:r[ >].*?</w:r>", p, re.S).group(0)
    rpr = re.search(r"<w:rPr>.*?</w:rPr>", first, re.S).group(0)
    ppr = re.search(r"<w:pPr>.*?</w:pPr>", p, re.S)
    ppr = ppr.group(0) if ppr else ""
    popen = re.match(r"<w:p[^>]*>", p).group(0)
    newp = popen + ppr + "<w:r>" + rpr + '<w:t xml:space="preserve">' + esc(text) + "</w:t></w:r></w:p>"
    return x[:off + m.start()] + newp + x[off + m.end():]
P1 = "с 18 февраля 2025 года по 17 февраля 2026 года"
P2 = "с 18 февраля 2026 года по 17 февраля 2027 года"
LIST = [
 ("12-19.08.2026", "6 (шесть) рабочих дней – с 12 по 19 августа (включительно) 2026 года", "«10» августа 2026 г.", P1),
 ("27.08-02.09.2026", "5 (пять) рабочих дней – с 27 августа по 2 сентября (включительно) 2026 года", "«24» августа 2026 г.", P2),
 ("07-14.10.2026", "6 (шесть) рабочих дней – с 7 по 14 октября (включительно) 2026 года", "«5» октября 2026 г.", P2),
]
for tag, days, date, per in LIST:
    x = doc
    x = set_par(x, 11, "В соответствии с пунктом 6.12 Трудового договора №1-1 от 18 февраля 2025 года, прошу Вас предоставить мне оплачиваемый трудовой отпуск в количестве " + days + ", за период " + per + ".")
    x = set_par(x, 15, "Дата: " + date)
    from xml.dom import minidom; minidom.parseString(x.encode("utf8"))
    out = os.path.join(OUTDIR, "заявление ЕН на отпуск %s (проект 2026-11-06).docx" % tag)
    if os.path.exists(out):
        try: open(out, "r+b").close()
        except PermissionError: out = out.replace(".docx", " (v2).docx")
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
