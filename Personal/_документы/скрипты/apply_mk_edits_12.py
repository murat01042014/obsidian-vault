# Двенадцатый этап (04.11.2026): в блоках подписей (строки 163, 169, 198) жирным остаётся только название Фонда
# («Central Asia Climate Foundation» / «Центрально-Азиатский Климатический Фонд»), остальные реквизиты - обычным шрифтом.
import zipfile, re, os, sys, io
from xml.dom import minidom
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
P = r"D:\Шаблоны\ИТД\Шаблон ИТД ЦАКФ - эталон (проект 2026-11-03).docx"
zi = zipfile.ZipFile(P); items = [(i, zi.read(i.filename)) for i in zi.infolist()]; zi.close()
x = dict((i.filename, d) for i, d in items)["word/document.xml"].decode("utf8")
rows = list(re.finditer(r"<w:tr[ >].*?</w:tr>", x, re.S)); assert len(rows) == 199
PAR = re.compile(r"<w:p[ >].*?</w:p>", re.S)
NAMES = ("Central Asia Climate Foundation", "Центрально-Азиатский Климатический Фонд")
def split(p):
    t = re.sub(r"<[^>]+>", "", p)
    nm = next((n for n in NAMES if t.startswith(n)), None)
    if not nm: return p, 0
    m = re.search(r"<w:r>(<w:rPr>.*?</w:rPr>)<w:t xml:space=\"preserve\">([^<]*)</w:t></w:r>", p, re.S)
    assert m and "<w:b/>" not in m.group(1) and m.group(2).startswith(nm)
    rpr = m.group(1); rest = m.group(2)[len(nm):]
    rpr_b = rpr.replace("</w:rFonts>", "</w:rFonts>").replace('w:cs="Arial"/>', 'w:cs="Arial"/><w:b/><w:bCs/>', 1)
    assert "<w:b/>" in rpr_b
    new = ('<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r><w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % (rpr_b, nm, rpr, rest))
    return p.replace(m.group(0), new, 1), 1
new = {}
for i in (163, 169, 198):
    r = rows[i].group(0); cnt = [0]
    def f(m):
        q, c = split(m.group(0)); cnt[0] += c; return q
    new[i] = PAR.sub(f, r); assert cnt[0] == 2, (i, cnt)
out = []; pos = 0
for i, m in enumerate(rows):
    out.append(x[pos:m.start()]); out.append(new.get(i, m.group(0))); pos = m.end()
out.append(x[pos:]); x2 = "".join(out)
minidom.parseString(x2.encode("utf8"))
with zipfile.ZipFile(P + ".tmp", "w", zipfile.ZIP_DEFLATED) as zo:
    for i, d in items:
        if i.filename == "word/document.xml": d = x2.encode("utf8")
        nz = zipfile.ZipInfo(i.filename, i.date_time); nz.compress_type = zipfile.ZIP_DEFLATED; nz.external_attr = i.external_attr
        zo.writestr(nz, d)
os.replace(P + ".tmp", P)
print("ok")
