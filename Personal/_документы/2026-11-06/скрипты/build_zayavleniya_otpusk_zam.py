# Заявления Темирбекова М.Е. и Кононца А.В. на очередной отпуск (RU + EN) по их прежним образцам. Даты - по указанию Мурата.
import zipfile, re, sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
D = r"D:\приказы\11_Заявления_и_согласия_работников"
def esc(s): return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
def pars(x):
    off = x.find("<w:body>"); return off, list(re.finditer(r"<w:p[ >].*?</w:p>", x[off:], re.S))
def set_par(x, idx, text):
    off, ps = pars(x); m = ps[idx]; p = m.group(0)
    first = re.search(r"<w:r[ >].*?</w:r>", p, re.S).group(0)
    rpr = re.search(r"<w:rPr>.*?</w:rPr>", first, re.S).group(0)
    ppr = re.search(r"<w:pPr>.*?</w:pPr>", p, re.S); ppr = ppr.group(0) if ppr else ""
    popen = re.match(r"<w:p[^>]*>", p).group(0)
    newp = popen + ppr + "<w:r>" + rpr + '<w:t xml:space="preserve">' + esc(text) + "</w:t></w:r></w:p>"
    return x[:off + m.start()] + newp + x[off + m.end():]
def sub_runs(x, idx, pairs):
    off, ps = pars(x); m = ps[idx]; p = m.group(0)
    for old, new in pairs:
        pat = re.compile(r'(<w:t(?: [^>]*)?>)' + re.escape(old) + r'(</w:t>)')
        assert pat.search(p), (idx, old)
        p = pat.sub(lambda mm: mm.group(1) + esc(new) + mm.group(2), p, count=1)
    return x[:off + m.start()] + p + x[off + m.end():]
def safe(out):
    if os.path.exists(out):
        try:
            open(out, 'r+b').close()
        except PermissionError:
            out = out.replace('.docx', ' (v2).docx')
    return out
def build(tpl, out, edits):
    out = safe(out)
    zi = zipfile.ZipFile(tpl); items = [(i, zi.read(i.filename)) for i in zi.infolist()]; zi.close()
    x = dict((i.filename, d) for i, d in items)["word/document.xml"].decode("utf8")
    x = edits(x)
    from xml.dom import minidom; minidom.parseString(x.encode("utf8"))
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
        for i, d in items:
            if i.filename == "word/document.xml": d = x.encode("utf8")
            elif i.filename == "docProps/app.xml":
                s = d.decode("utf8")
                s = re.sub(r"<Company>.*?</Company>", "<Company>CACF</Company>", s) if "<Company>" in s else s.replace("</Properties>", "<Company>CACF</Company></Properties>")
                d = s.encode("utf8")
            nz = zipfile.ZipFile and zipfile.ZipInfo(i.filename, i.date_time); nz.compress_type = zipfile.ZIP_DEFLATED; nz.external_attr = i.external_attr
            zo.writestr(nz, d)
    print("готово:", os.path.basename(out))

def tem(x):
    x = set_par(x, 10, "На основании пункта 6.1. Трудового договора, прошу Вас предоставить мне трудовой отпуск за период с 24 февраля 2026 года по 23 февраля 2027 года с «7» по «18» сентября (включительно) 2026 года в количестве 10 рабочих дней.")
    x = sub_runs(x, 13, [("30", "4"), ("ию", "сентябр"), ("л", ""), ("я", "я")])
    x = set_par(x, 31, "Pursuant to paragraph 6.1 of the Employment Agreement, I request that you grant me annual leave for the period from February 24, 2026, to February 23, 2027, from September 7 to 18 (include), 2026, for a total of 10 working days.")
    x = sub_runs(x, 34, [("Ju", "Septemb"), ("ly", "er"), ("30", "4")])
    return x
def kon(x):
    x = sub_runs(x, 10, [("24 августа 2026 года по 7 сентября 2026 года включительно", "24 августа 2026 года по 4 сентября 2026 года включительно")])
    x = sub_runs(x, 29, [("August 24 to September 7 (include)", "August 24 to September 4 (include)")])
    return x
build(os.path.join(D, "заявление МЕ на отпуск англ и рус-082026.docx"), os.path.join(D, "заявление МЕ на отпуск 07-18.09.2026 англ и рус (проект 2026-11-06).docx"), tem)
build(os.path.join(D, "Заявление Кононец на отпуск_1.docx"), os.path.join(D, "Заявление Кононец на отпуск 24.08-04.09.2026 англ и рус (проект 2026-11-06).docx"), kon)
