# Семнадцатый этап (04.11.2026): правка Мурата во всех ТД - абзац «B. Г-н/Г-жа ...» (строка 3, RU-колонка):
# отступ 324/324 вместо 800/800, явная табуляция 800 убрана. Скрипт применим к эталону и к готовым ТД (на месте, без пересборки).
import zipfile, re, os, sys, io
from xml.dom import minidom
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
def patch(P):
    zi = zipfile.ZipFile(P); items = [(i, zi.read(i.filename)) for i in zi.infolist()]; zi.close()
    x = dict((i.filename, d) for i, d in items)["word/document.xml"].decode("utf8")
    rows = list(re.finditer(r"<w:tr[ >].*?</w:tr>", x, re.S)); assert len(rows) == 199
    r = rows[3].group(0); c = re.findall(r"<w:tc>.*?</w:tc>", r, re.S)[1]
    p = re.search(r"<w:p[ >].*?</w:p>", c, re.S).group(0)
    assert re.sub(r"<[^>]+>", "", p).startswith("B."), p[:100]
    if '<w:ind w:left="324" w:hanging="324"/>' in p:
        return "уже"
    p2 = p.replace('<w:ind w:left="800" w:hanging="800"/>', '<w:ind w:left="324" w:hanging="324"/>', 1)
    p2 = re.sub(r'<w:tabs><w:tab w:val="left" w:pos="800"/></w:tabs>', "", p2, count=1)
    assert p2 != p and '<w:ind w:left="324" w:hanging="324"/>' in p2
    x2 = x[:rows[3].start()] + r.replace(p, p2, 1) + x[rows[3].end():]
    minidom.parseString(x2.encode("utf8"))
    tmp = P + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
        for i, d in items:
            if i.filename == "word/document.xml": d = x2.encode("utf8")
            nz = zipfile.ZipInfo(i.filename, i.date_time); nz.compress_type = zipfile.ZIP_DEFLATED; nz.external_attr = i.external_attr
            zo.writestr(nz, d)
    os.replace(tmp, P)
    return "изменён"
if __name__ == "__main__":
    for P in sys.argv[1:]:
        print(os.path.basename(P)[:40], patch(P))
