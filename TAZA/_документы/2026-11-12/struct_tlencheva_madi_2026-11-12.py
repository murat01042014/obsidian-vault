# Структуры ПО ТК: Тленчиева - приход в 2027 (сиреневый), Нуржау М. - главный специалист с 09.11.2026 (зелёный, оклад 1 000 000 там, где оклады показаны)
import zipfile, re, html, sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
D = r"D:\ТАЗА"
FILES = [
    ("Структура ПО Таза Қазақстан – с окладами (2026-11-09).pptx", "Структура ПО Таза Қазақстан – с окладами (2026-11-12).pptx", True),
    ("Структура ПО Таза Қазақстан – без окладов (2026-11-09).pptx", "Структура ПО Таза Қазақстан – без окладов (2026-11-12).pptx", False),
    ("Структура ПО Таза Казахстан - ЦАКФ (2026-11-09).pptx", "Структура ПО Таза Казахстан - ЦАКФ (2026-11-12).pptx", False),
]
T_RE = re.compile(r'(<a:t>)(.*?)(</a:t>)', re.S)
SP_RE = re.compile(r'<p:sp>.*?</p:sp>', re.S)
def esc(s): return html.escape(s, quote=False)
def stext(sp): return html.unescape(" ".join(m.group(2) for m in T_RE.finditer(sp)))
NB = "\xa0"

def replace_in_runs(blk, old, new):
    ms = list(T_RE.finditer(blk)); texts = [html.unescape(m.group(2)) for m in ms]
    full = "".join(texts); idx = full.find(old)
    if idx < 0: return blk, False
    end = idx + len(old); cum = 0; i0 = i1 = None; o0 = o1 = 0
    for i, t in enumerate(texts):
        if i0 is None and idx < cum + len(t): i0, o0 = i, idx - cum
        if i1 is None and end <= cum + len(t): i1, o1 = i, end - cum
        cum += len(t)
        if i0 is not None and i1 is not None: break
    nt = list(texts)
    if i0 == i1: nt[i0] = texts[i0][:o0] + new + texts[i0][o1:]
    else:
        nt[i0] = texts[i0][:o0] + new
        for k in range(i0 + 1, i1): nt[k] = ""
        nt[i1] = texts[i1][o1:]
    out = []; last = 0
    for i, m in enumerate(ms):
        out.append(blk[last:m.start()]); out.append(m.group(1) + esc(nt[i]) + m.group(3)); last = m.end()
    out.append(blk[last:])
    return "".join(out), True

def set_fill(blk, color):
    # заливка самой фигуры (первый solidFill в spPr)
    m = re.search(r"(<p:spPr>.*?<a:solidFill>\s*<a:srgbClr val=\")([0-9A-Fa-f]{6})(\")", blk, re.S)
    assert m, "нет заливки"
    return blk[:m.start(2)] + color + blk[m.end(2):]

def set_y(blk, y):
    return re.sub(r'(<a:off x="\d+" y=")(\d+)(")', lambda m: m.group(1) + str(y) + m.group(3), blk, count=1)
def get_y(blk): return int(re.search(r'<a:off x="\d+" y="(\d+)"', blk).group(1))

for src, dst, with_salary in FILES:
    zin = zipfile.ZipFile(os.path.join(D, src))
    x = zin.read("ppt/slides/slide1.xml").decode("utf8")
    sps = SP_RE.findall(x)
    tl = [s for s in sps if "Тленчиева" in stext(s)]
    nz = [s for s in sps if "Нуржау" in stext(s)]
    bk = [s for s in sps if "Бектемиров" in stext(s)]
    assert len(tl) == 1 and len(nz) == 1 and len(bk) == 1, (len(tl), len(nz), len(bk))
    # --- Тленчиева: 2027 (сиреневый)
    tl2 = set_fill(tl[0], "E6DFF3")
    # --- Нуржау: главный специалист, 2026 (зелёный)
    nz2, ok = replace_in_runs(nz[0], "Ведущий специалист", "Главный специалист"); assert ok
    if with_salary:
        nz2, ok1 = replace_in_runs(nz2, "7", "1"); assert ok1       # первая "7" в рублевой части ("· 7 00 000")
        nz2, ok2 = replace_in_runs(nz2, "00" + NB + "000", NB + "000" + NB + "000")
        if not ok2: nz2, ok2 = replace_in_runs(nz2, "00 000", NB + "000" + NB + "000")
        assert ok2
    nz2 = set_fill(nz2, "E9F3EA")
    # --- порядок: главный специалист Нуржау выше ведущего Бектемирова
    y_nz, y_bk = get_y(nz[0]), get_y(bk[0])
    assert get_y(nz[0]) > get_y(bk[0]), (y_nz, y_bk)
    nz2 = set_y(nz2, y_bk); bk2 = set_y(bk[0], y_nz)
    x2 = x.replace(tl[0], tl2, 1).replace(nz[0], nz2, 1).replace(bk[0], bk2, 1)
    # --- подпись про приход в 2027 (только в версиях ПО)
    legend_done = False
    for s in SP_RE.findall(x2):
        if "в феврале 2027 года приходят ещё 6" in stext(s):
            s2, ok = replace_in_runs(s, "в феврале 2027 года приходят ещё 6", "в январе и феврале 2027 года приходят ещё 6"); assert ok
            x2 = x2.replace(s, s2, 1); legend_done = True; break
    print(dst, "| подпись обновлена:", legend_done, "| y Нуржау/Бектемиров:", y_nz, "->", y_bk)
    outp = os.path.join(D, dst)
    if os.path.exists(outp): print("  уже есть, пропуск"); continue
    with zipfile.ZipFile(outp, "w", zipfile.ZIP_DEFLATED) as zo:
        for it in zin.infolist():
            d = zin.read(it.filename)
            if it.filename == "ppt/slides/slide1.xml": d = x2.encode("utf8")
            elif it.filename == "docProps/app.xml":
                s = d.decode("utf8")
                s = re.sub(r"<Company>.*?</Company>", "<Company>CACF</Company>", s) if "<Company>" in s else s.replace("</Properties>", "<Company>CACF</Company></Properties>")
                d = s.encode("utf8")
            zo.writestr(it, d)
    print("  готово")
