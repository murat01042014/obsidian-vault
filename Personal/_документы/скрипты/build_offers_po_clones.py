# Офферы сотрудникам ПО «Таза Қазақстан» (ноябрь, штат и аутсорс): клоны оффера Серикжановой в редакции Мурата
# (вёрстка, подпись ГИД и список документов не меняются; подставляются номер, адресат, обращение, должность, дата выхода).
# Части в [квадратных скобках] остаются жёлтыми - данных нет в списке/удостоверениях.
import zipfile, re, os, sys, io
from xml.dom import minidom

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
DIR = r"D:\Кадры\Офферы\ПО ТК"
SRC = [f for f in os.listdir(DIR) if f.startswith("Оффер 02-380") and f.endswith(".docx") and not f.startswith("~$")][0]
SRC = os.path.join(DIR, SRC)
DEP_SOC = "Департамента социологических исследований и аналитики"
DEP_PM = "Департамента проектного управления и региональной координации"
DEP_COMM = "Департамента коммуникаций и креативных проектов"

# n - исх. номер; sur/rest - адресат (дательный: фамилия / имя и отчество); hello - обращение; pos - должность (как в списке ПО);
# start - «со 2 ноября»/«с 3 ноября»; outsource - аутсорс (договор вместо трудового); tag - имя файла
P = [
    dict(n=381, sur="Өміртай", rest="Амалбеку Даркенұлы", hello="Уважаемый Амалбек Даркенұлы!", pos="Директор " + DEP_SOC, start="с 3 ноября", tag="Өміртай А. - Директор департамента"),
    dict(n=382, sur="Абенов", rest="Алибеку Даниярулы", hello="Уважаемый Алибек Даниярулы!", pos="Главный специалист " + DEP_SOC, tag="Абенов А.Д. - Главный специалист"),
    dict(n=383, nogreet=True, sur="Нурканат А.", rest="", hello="Уважаемый(ая) [имя, отчество]!", pos="Социолог " + DEP_SOC, outsource=True, tag="Нурканат А. - Социолог (аутсорс)"),
    dict(n=384, start="с декабря", nogreet=True, sur="Тленчиева", rest="Шырын", hello="Уважаемая Шырын [отчество]!", pos="Главный специалист " + DEP_SOC, tag="Тленчиева Ш. - Главный специалист"),
    dict(n=385, sur="Қуандық", rest="Жасұлану Бостандықұлы", hello="Уважаемый Жасұлан Бостандықұлы!", pos="Главный специалист " + DEP_PM, tag="Қуандық Ж.Б. - Главный специалист"),
    dict(n=386, sur="Жанділдаев", rest="Бақдаулету Болатұлы", hello="Уважаемый Бақдаулет Болатұлы!", pos="Главный специалист " + DEP_PM, tag="Жанділдаев Б.Б. - Главный специалист"),
    dict(n=387, sur="Шербаева", rest="Бибинур Байжигит қизи", hello="Уважаемая Бибинур Байжигит қизи!", pos="Главный специалист " + DEP_PM, tag="Шербаева Б.Б. - Главный специалист"),
    dict(n=388, sur="Бектемиров", rest="Амиру Руслановичу", hello="Уважаемый Амир Русланович!", pos="Ведущий специалист " + DEP_PM, tag="Бектемиров А.Р. - Ведущий специалист"),
    dict(n=389, sur="Тастан", rest="Нұрболу Ерболұлы", hello="Уважаемый Нұрбол Ерболұлы!", pos="Директор Департамента цифрового развития", tag="Тастан Н.Е. - Директор департамента"),
    dict(n=390, start="с 9 ноября", nogreet=True, sur="Сағынұлы С.", rest="", hello="Уважаемый [имя, отчество]!", pos="Директор " + DEP_COMM, tag="Сағынұлы С. - Директор департамента"),
    dict(n=391, sur="Талғат", rest="Айман Талғатқызы", hello="Уважаемая Айман Талғатқызы!", pos="Главный специалист " + DEP_COMM, tag="Талғат А.Т. - Главный специалист"),
    dict(n=392, sur="Мухамадиева", rest="Найле Азаматовне", hello="Уважаемая Найля Азаматовна!", pos="Ведущий специалист " + DEP_COMM, tag="Мухамадиева Н.А. - Ведущий специалист"),
    dict(n=393, sur="Жумабай", rest="Бейбарысу Ғалымұлы", hello="Уважаемый Бейбарыс Ғалымұлы!", pos="Главный специалист " + DEP_COMM, tag="Жумабай Б.Ғ. - Главный специалист"),
    dict(n=394, sur="Дутпаева", rest="Алие Кумархановне", hello="Уважаемая Алия Кумархановна!", pos="Ведущий специалист " + DEP_COMM, tag="Дутпаева А.К. - Ведущий специалист"),
]

def esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

def mk_runs(text, rpr):
    out = ""
    for part in [x for x in re.split(r"(\[[^\]]*\])", text) if x != ""]:
        hl = part.startswith("[") and part.endswith("]")
        rp = rpr.replace("</w:rPr>", '<w:highlight w:val="yellow"/></w:rPr>') if (hl and rpr) else (('<w:rPr><w:highlight w:val="yellow"/></w:rPr>') if hl else rpr)
        out += "<w:r>" + rp + '<w:t xml:space="preserve">' + esc(part) + "</w:t></w:r>"
    return out

def edit(x, old, new):
    """Заменяет old на new внутри одного w:t; run перестраивается (new может содержать [маркеры])."""
    for m in re.finditer(r"<w:r\b[^>]*>(?P<rpr><w:rPr>(?:(?!</w:rPr>).)*</w:rPr>)?(?P<mid>(?:(?!</w:r>|<w:t[ >]).)*)<w:t(?: [^>]*)?>(?P<t>[^<]*)</w:t></w:r>", x, re.S):
        t = m.group("t")
        if old in t and "<w:drawing" not in m.group("mid"):
            i = t.index(old)
            new_x = mk_runs(t[:i], m.group("rpr") or "") + mk_runs(new, m.group("rpr") or "") + mk_runs(t[i + len(old):], m.group("rpr") or "")
            return x[:m.start()] + new_x + x[m.end():]
    raise SystemExit("не найдено: " + old)

src = zipfile.ZipFile(SRC)
base = src.read("word/document.xml").decode("utf8")
# правки по просьбе Мурата 03.11.2026: «санированный» -> «сканированный», нумерация справок 1-9 (было 10) вместо 9))
base = base.replace("санированный", "сканированный").replace("санированные", "сканированные").replace("10) справка об отсутствии коррупционных", "9) справка об отсутствии коррупционных")
made = []
for p in P:
    if len(sys.argv) > 1 and str(p["n"]) not in sys.argv[1:]:
        continue
    x = base
    x = edit(x, "02/380", "02/%d" % p["n"])
    x = edit(x, "Серикжанова ", p["sur"] + " ")
    if p["rest"]:
        x = edit(x, "Айгерим Тельжановне", p["rest"])
    else:   # только фамилия и инициалы - в одну строку, второй абзац адресата убираем
        m = [q for q in re.finditer(r"<w:p[ >](?:(?!</w:p>).)*?</w:p>", x, re.S) if "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", q.group(0))) == "Айгерим Тельжановне"]
        assert len(m) == 1
        x = x[:m[0].start()] + x[m[0].end():]
    if p.get("nogreet"):   # нет отчества - обращение «Уважаемый(ая) ...» не пишем (просьба Мурата 03.11.2026)
        m = [q for q in re.finditer(r"<w:p[ >](?:(?!</w:p>).)*?</w:p>", x, re.S) if "Уважаемая Айгерим Тельжановна!" in "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", q.group(0)))]
        assert len(m) == 1
        x = x[:m[0].start()] + x[m[0].end():]
    else:
        x = edit(x, "Уважаемая Айгерим Тельжановна!", p["hello"])
    if p.get("outsource"):   # аутсорс: форма оформления уточняется - пометка после должности
        x = edit(x, "Руководитель Проектного офиса «Таза Қазақстан»»", p["pos"] + "» (возможно оформление на условиях аутсорсинга)")
    else:
        x = edit(x, "Руководитель Проектного офиса «Таза Қазақстан»", p["pos"])
    if p.get("start"):
        x = edit(x, "со 2 ноября", p["start"])
        x = edit(x, "со 2 ноября", p["start"])
    if p.get("outsource"):
        x = edit(x, "трудового договора", "договора")
        x = edit(x, "трудового договора", "договора")
    minidom.parseString(x.encode("utf8"))
    out = os.path.join(DIR, "Оффер 02-%d - %s (проект 2026-11-03).docx" % (p["n"], p["tag"]))
    tmp = out + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
        for it in src.infolist():
            zi = zipfile.ZipInfo(it.filename, it.date_time); zi.compress_type = zipfile.ZIP_DEFLATED   # writestr меняет ZipInfo, поэтому копия
            zo.writestr(zi, x.encode("utf8") if it.filename == "word/document.xml" else src.read(it.filename))
    os.replace(tmp, out)
    made.append(out)
    print(os.path.basename(out))
