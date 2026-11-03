# Предитоговые проекты ИТД: берём последний шаблон ИТД ЦАКФ (с испытательным сроком и т.д.) и подставляем данные
# сотрудника (удостоверение, ШТ 0211, черновик ТД). Жёлтое выделение остаётся только там, где данных пока нет.
import zipfile, re, os, sys, shutil, io
from xml.dom import minidom
from xml.sax.saxutils import escape

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
TPL = r"D:\Шаблоны\ИТД\Шаблон ИТД ЦАКФ - эталон (проект 2026-11-03).docx"
OUTDIR = r"D:\Кадры\ТД предитоговые"
os.makedirs(OUTDIR, exist_ok=True)
NB = "\u00a0"

MONTHS = {"01": "January", "02": "February", "03": "March", "04": "April", "05": "May", "06": "June", "07": "July",
          "08": "August", "09": "September", "10": "October", "11": "November", "12": "December"}

def d_ru(s):  # 02.11.2026 -> 02.11.2026 г.
    return s + " г." if s else None

def d_en(s):  # 02.11.2026 -> 2 November 2026
    if not s:
        return None
    dd, mm, yy = s.split(".")
    return f"{dd if '[' in dd else int(dd)} {MONTHS[mm]} {yy}"

def money_ru(n, words):
    return f"{n:,}".replace(",", NB) + f" ({words})"

def money_en(n, words):
    return f"{n:,}" + f" ({words})"

AUTH_RU = "МВД РК"
AUTH_EN = "the Ministry of Internal Affairs of the Republic of Kazakhstan"

# --- должности по ШТ 0211 (D:\ТАЗА\ШТ 0211.docx): (именительный, родительный, английский)
POS = {
    "po":   ("Руководитель Проектного офиса «Таза Қазақстан»", "руководителя Проектного офиса «Таза Қазақстан»", "Head of the Project Office «Taza Qazaqstan»"),
    "soc":  ("Главный специалист Департамента социологических исследований и аналитики", "главного специалиста Департамента социологических исследований и аналитики", "Chief Specialist of the Department of Sociological Research and Analytics"),
    "pm":   ("Главный специалист Департамента проектного управления и региональной координации", "главного специалиста Департамента проектного управления и региональной координации", "Chief Specialist of the Department of Project Management and Regional Coordination"),
    "dig":  ("Директор Департамента цифрового развития", "директора Департамента цифрового развития", "Director of the Department of Digital Development"),
    "pm_lead": ("Ведущий специалист Департамента проектного управления и региональной координации", "ведущего специалиста Департамента проектного управления и региональной координации", "Leading Specialist of the Department of Project Management and Regional Coordination"),
    "comm_lead": ("Ведущий специалист Департамента коммуникаций и креативных проектов", "ведущего специалиста Департамента коммуникаций и креативных проектов", "Leading Specialist of the Department of Communications and Creative Projects"),
    "soc_dir": ("Директор Департамента социологических исследований и аналитики", "директора Департамента социологических исследований и аналитики", "Director of the Department of Sociological Research and Analytics"),
    "soc_lead": ("Ведущий специалист Департамента социологических исследований и аналитики", "ведущего специалиста Департамента социологических исследований и аналитики", "Leading Specialist of the Department of Sociological Research and Analytics"),
    "pm_dir": ("Директор Департамента проектного управления и региональной координации", "директора Департамента проектного управления и региональной координации", "Director of the Department of Project Management and Regional Coordination"),
    "comm_dir": ("Директор Департамента коммуникаций и креативных проектов", "директора Департамента коммуникаций и креативных проектов", "Director of the Department of Communications and Creative Projects"),
    "comm": ("Главный специалист Департамента коммуникаций и креативных проектов", "главного специалиста Департамента коммуникаций и креативных проектов", "Chief Specialist of the Department of Communications and Creative Projects"),
}
FUNC_PM = ("Главный специалист Департамента проектного управления и региональной координации обеспечивает организационное и аналитическое сопровождение проектов, взаимодействие с региональными представителями и контроль исполнения мероприятий в пределах своей компетенции.",
           "The Chief Specialist of the Department of Project Management and Regional Coordination provides organizational and analytical support of projects, interaction with regional representatives and monitoring of the implementation of activities within the scope of their competence.")
FUNC_PO = ("Осуществляет общее стратегическое и операционное руководство Проектным офисом «Таза Қазақстан», обеспечивает координацию деятельности структурных подразделений, планирование и контроль реализации проектов и инициатив.",
           "Provides overall strategic and operational management of the Project Office «Taza Qazaqstan», ensures coordination of the activities of structural units, planning and control of the implementation of projects and initiatives.")
# подотчётность: сотрудник - руководителю своего департамента, руководитель департамента - Руководителю ПО ТК;
# всем дописано «а также вышестоящему руководству»
UP = (", а также вышестоящему руководству", ", as well as to senior management")
REP_SOC = ("Директору Департамента социологических исследований и аналитики" + UP[0], "the Director of the Department of Sociological Research and Analytics" + UP[1])
REP_PM = ("Директору Департамента проектного управления и региональной координации" + UP[0], "the Director of the Department of Project Management and Regional Coordination" + UP[1])
REP_COMM = ("Директору Департамента коммуникаций и креативных проектов" + UP[0], "the Director of the Department of Communications and Creative Projects" + UP[1])
REP_PO = ("Руководителю Проектного офиса «Таза Қазақстан»" + UP[0], "the Head of the Project Office «Taza Qazaqstan»" + UP[1])
REP_TOP = ("Главному исполнительному директору Фонда", "the Chief Executive Officer of the Foundation")   # Руководитель ПО ТК подчиняется только ГИД

# --- данные сотрудников (удостоверение личности - из PDF; адрес, даты, функции - из черновиков ТД в D:\Кадры\Проекты ТД)
P = [
    dict(file="Серикжанова А.Т.", fem=True, ru="Серикжанова Айгерим Тельжановна", en="Serikzhanova Aigerim Telzhanovna",
         idn="049987124", idd="10.03.2021", addr_ru="г. Астана, ул. Сауран, д. 5д, кв. 44", addr_en="Astana, Sauran St., 5d, apt. 44",
         pos="po", salary=(1400000, "один миллион четыреста тысяч", "one million four hundred thousand"),   # испытательный срок как у всех (Мурат, 03.11.2026)
         start="02.11.2026", end=None, cdate="02.11.2026", rep=REP_TOP, func=FUNC_PO),
    dict(file="Абенов А.Д.", fem=False, ru="Абенов Алибек Даниярулы", en="Abenov Alibek Daniyaruly",
         idn="050903234", idd="26.07.2021", addr_ru="г. Астана, ул. Ч. Айтматова, д. 29а, кв. 17", addr_en="Astana, Ch. Aitmatov St., 29a, apt. 17",
         pos="soc", salary=(1000000, "один миллион", "one million"),
         start="02.11.2026", end=None, cdate="02.11.2026", rep=REP_SOC, func=None),
    dict(file="Қуандық Ж.Б.", fem=False, ru="Қуандық Жасұлан Бостандықұлы", en="Kuandyk Zhassulan Bostandykuly",
         idn="042920885", idd="19.04.2018", addr_ru="г. Астана, ул. Төле би, д. 25, кв. 210", addr_en="Astana, Tole bi St., 25, apt. 210",
         pos="pm", salary=(1000000, "один миллион", "one million"),
         start="02.11.2026", end=None, cdate="02.11.2026", rep=REP_PM, func=FUNC_PM),
    dict(file="Жанділдаев Б.Б.", fem=False, ru="Жанділдаев Бақдаулет Болатұлы", en="Zhandildayev Bakdaulet Bolatuly",
         idn="054092478", idd="07.11.2022", addr_ru="г. Астана, ул. А. Байтұрсынұлы, д. 20, кв. 52", addr_en="Astana, A. Baitursynuly St., 20, apt. 52",
         pos="pm", salary=(1000000, "один миллион", "one million"),
         start="02.11.2026", end=None, cdate="02.11.2026", rep=REP_PM, func=FUNC_PM),
    dict(file="Шербаева Б.Б.", fem=True, ru="Шербаева Бибинур Байжигит қизи", en="Sherbayeva Bibinur Baizhigit kyzy",
         idn="054354788", idd="27.12.2022", addr_ru="г. Астана, ул. Калдаякова, д. 1, кв. 286", addr_en="Astana, Kaldayakov St., 1, apt. 286",
         pos="pm", salary=(1000000, "один миллион", "one million"),
         start="02.11.2026", end=None, cdate="02.11.2026", rep=REP_PM, func=FUNC_PM),
    dict(file="Тастан Н.Е.", fem=False, ru="Тастан Нұрбол Ерболұлы", en="Tastan Nurbol Erboluly",
         idn="053445366", idd="04.08.2022", addr_ru="г. Астана, ул. Кошкарбаева, д. 41, кв. 18", addr_en="Astana, Koshkarbayev St., 41, apt. 18",
         pos="dig", salary=(1300000, "один миллион триста тысяч", "one million three hundred thousand"),
         start="02.11.2026", end=None, cdate="02.11.2026", rep=REP_PO, func=None),
    dict(file="Талғат А.Т.", fem=True, ru="Талғат Айман Талғатқызы", en="Talgat Aiman Talgatkyzy",
         idn="064422834", idd="17.09.2026", addr_ru="г. Астана, ул. Р. Багланова, д. 4, кв. 118", addr_en="Astana, R. Baglanov St., 4, apt. 118",
         pos="comm", salary=(1000000, "один миллион", "one million"),
         start="02.11.2026", end="02.11.2027", cdate="02.11.2026", rep=REP_COMM, func=None),
    # 02.11.2026: удостоверения есть, черновиков ТД не было - адрес и функции пока маркером
    dict(file="Бектемиров А.Р.", fem=False, ru="Бектемиров Амир Русланович", en="Bektemirov Amir Ruslanovich",
         idn="064064248", idd="21.07.2026", addr_ru="г. Астана, ул. Кошкарбаева, д. 54а, кв. 11", addr_en="Astana, Koshkarbayev St., 54a, apt. 11",
         pos="pm_lead", salary=(700000, "семьсот тысяч", "seven hundred thousand"),
         start="02.11.2026", end=None, cdate="02.11.2026", rep=REP_PM, func=None),
    dict(file="Мухамадиева Н.А.", fem=True, ru="Мухамадиева Найля Азаматовна", en="Mukhamadiyeva Nailya Azamatovna",
         idn="052204932", idd="15.02.2022", addr_ru="г. Астана, ул. Кошкарбаева, д. 49, кв. 226", addr_en="Astana, Koshkarbayev St., 49, apt. 226",
         pos="comm_lead", salary=(700000, "семьсот тысяч", "seven hundred thousand"),
         start="02.11.2026", end=None, cdate="02.11.2026", rep=REP_COMM, func=None),
    dict(file="Жумабай Б.Ғ.", fem=False, ru="Жумабай Бейбарыс Ғалымұлы", en="Zhumabay Beibarys Galymuly",
         idn="046760170", idd="22.10.2019", addr_ru="г. Астана, ул. Ш. Айтматова, д. 77/6, кв. 434", addr_en="Astana, Sh. Aitmatov St., 77/6, apt. 434",
         pos="comm", salary=(1000000, "один миллион", "one million"),
         start="02.11.2026", end=None, cdate="02.11.2026", rep=REP_COMM, func=None),
    # 02.11.2026: остальные штатные сотрудники ПО ТК - без удостоверений (ФИО по списку, данные уд.л. и адрес - маркером)
    dict(file="Өміртай А.", fem=False, ru="Өміртай Амалбек Даркенұлы", en="Omirtay Amalbek Darkenuly",
         idn="044781611", idd="25.07.2019", addr_ru="г. Астана, ул. Достык, д. 12/1, кв. 1", addr_en="Astana, Dostyk St., 12/1, apt. 1",
         pos="soc_dir", salary=(1300000, "один миллион триста тысяч", "one million three hundred thousand"),
         start="03.11.2026", end=None, cdate="02.11.2026", rep=REP_PO, func=None),
    dict(file="Тленчиева Ш.", tag=", без уд.л.", fem=True, ru="Тленчиева Шырын [отчество]", en=None, idn=None, idd=None, addr_ru=None, addr_en=None,
         pos="soc", salary=(1000000, "один миллион", "one million"),
         start="[день].12.2026", end=None, cdate="[день].12.2026", rep=REP_SOC, func=None),
    dict(file="Сағынұлы С.", fem=False, ru="Сағынұлы Санжар", en="Sagynuly Sanzhar",
         idn="057345243", idd="24.07.2024", addr_ru="г. Астана, Есильский район, ул. Ақмешіт, д. 17", addr_en="Astana, Yesil District, Akmeshit St., 17",
         pos="comm_dir", salary=(1300000, "один миллион триста тысяч", "one million three hundred thousand"),
         start="09.11.2026", end=None, cdate="02.11.2026", rep=REP_PO, func=None),
    dict(file="Дутпаева А.К.", fem=True, ru="Дутпаева Алия Кумархановна", en="Dutpayeva Aliya Kumarkhanovna",
         idn="044193316", idd="11.04.2018", addr_ru="г. Астана, пр. Бауржан Момышулы, д. 18, кв. 33", addr_en="Astana, Baurzhan Momyshuly Ave., 18, apt. 33",
         pos="comm_lead", salary=(700000, "семьсот тысяч", "seven hundred thousand"),
         start="02.11.2026", end=None, cdate="02.11.2026", rep=REP_COMM, func=None),
]

def end_date(start):
    # срок договора - 1 год от даты выхода на работу, дата окончания включительно (02.11.2026 -> 01.11.2027)
    from datetime import date, timedelta
    if "[" in start:
        return None   # день выхода не определён - дата окончания маркером
    dd, mm, yy = map(int, start.split("."))
    e = date(yy + 1, mm, dd) - timedelta(days=1)
    return e.strftime("%d.%m.%Y")

def make_values(p):
    pos = POS[p["pos"]]
    n, w_ru, w_en = p["salary"]
    en_party = [p["en"], p["idn"], d_en(p["idd"]), AUTH_EN, p["addr_en"]]
    ru_party = [p["ru"], p["idn"], d_ru(p["idd"]), AUTH_RU, p["addr_ru"]]
    V = {}
    V[1] = [p["num"]]; V[3] = [p["num"]]
    V[2] = [d_en(p["cdate"])]; V[4] = [d_ru(p["cdate"])]
    for k in (5, 11, 30): V[k] = en_party
    V[25] = en_party[1:]
    for k in (6,): V[k] = ru_party
    for k in (14, 28, 33): V[k] = ru_party[1:]
    for k in (12, 24, 26, 31): V[k] = [p["en"]]
    for k in (13, 15, 27, 29, 32, 34): V[k] = [p["ru"]]
    V[7] = [pos[2]]; V[8] = [pos[1]]; V[16] = [pos[2]]; V[17] = [pos[0]]
    end = end_date(p["start"])
    V[9] = [d_en(p["start"]), d_en(end)]; V[10] = [d_ru(p["start"]), d_ru(end)]
    V[18] = [pos[2], money_en(n, w_en)]; V[19] = [pos[1], money_ru(n, w_ru)]
    V[20] = [p["rep"][1] if p["rep"] else None]; V[21] = [p["rep"][0] if p["rep"] else None]
    V[22] = [p["func"][1] if p["func"] else None]; V[23] = [p["func"][0] if p["func"] else None]
    return V

PROBATION_OFF = [
    ("A probation period of 3 (three) months is established for the Employee from the employment commencement date specified in clause 2.2. The probation period does not include periods of temporary incapacity for work and leave. The notice of termination during the probation period is determined in accordance with clause 19.2 (not less than 7 (seven) days if the Employee has been continuously employed for 1 month or more but less than 3 months; no notice is required during the first month of employment). Termination of this Employment Agreement during the probation period shall be made in accordance with Section 19.",
     "The Employer does not establish to the Employee a probation period."),
    ("Работнику устанавливается испытательный срок 3 (три) месяца со дня начала работы, указанного в пункте 2.2. В испытательный срок не включаются периоды временной нетрудоспособности и отпуска. Срок уведомления о прекращении трудовых отношений в период испытательного срока определяется пунктом 19.2 (не менее 7 (семи) дней, если Работник непрерывно проработал у Работодателя 1 месяц или более, но менее 3 месяцев; в течение первого месяца работы уведомление не требуется). Прекращение настоящего Трудового договора в период испытательного срока осуществляется в соответствии с разделом 19.",
     "Работодатель не устанавливает Работнику испытательный срок."),
]


# --- должностные функции из черновиков ТД (D:\Кадры\Проекты ТД, от Айгерим): абзацы RU как в черновике, EN - перевод.
# Вставляются абзацами после заголовка п. 4 Приложения 1 (вместо жёлтого маркера).
FP = {
 "Абенов А.Д.": (
  ["Главный специалист Департамента социологических исследований и аналитики - осуществляет аналитическое, информационное и организационно-методическое сопровождение деятельности департамента, обеспечивает сбор, обработку и анализ информации, подготовку аналитических и информационно-справочных материалов по направлениям деятельности департамента.",
   "Должностные функции:",
   "1. Аналитическое сопровождение деятельности департамента: сбор, систематизация, обработка и анализ информации по курируемым направлениям, подготовка аналитических выводов и предложений.",
   "2. Подготовка аналитических и информационно-справочных материалов: разработка аналитических записок, справок, обзоров, докладных записок, презентаций для руководства по вопросам деятельности департамента.",
   "3. Анализ нормативной и методологической базы: изучение и систематизация нормативных правовых актов, документов по направлениям деятельности департамента, подготовка предложений по их совершенствованию.",
   "4. Информационное взаимодействие и координация: подготовка запросов, обобщение информации, поступающей от государственных органов, местных исполнительных органов и иных организаций, участие в подготовке рабочих совещаний и материалов по их итогам."],
  ["The Chief Specialist of the Department of Sociological Research and Analytics provides analytical, informational and organizational and methodological support for the activities of the Department, ensures the collection, processing and analysis of information, and the preparation of analytical and information reference materials in the areas of the Department's activities.",
   "Job functions:",
   "1. Analytical support of the Department's activities: collection, systematization, processing and analysis of information in the supervised areas, preparation of analytical conclusions and proposals.",
   "2. Preparation of analytical and information reference materials: development of analytical notes, reports, reviews, memoranda and presentations for management on matters within the Department's activities.",
   "3. Analysis of the regulatory and methodological framework: study and systematization of regulatory legal acts and documents in the areas of the Department's activities, preparation of proposals for their improvement.",
   "4. Information interaction and coordination: preparation of requests, consolidation of information received from state bodies, local executive bodies and other organizations, participation in preparing working meetings and materials on their outcomes."]),
 "Талғат А.Т.": (
  ["Главный специалист Департамента коммуникаций и креативных проектов обеспечивает информационно-коммуникационное сопровождение проектов и инициатив, участвует в разработке и реализации коммуникационных стратегий, креативных и медиапроектов.",
   "Разработка и реализация коммуникационных стратегий и информационных кампаний.",
   "Информационное сопровождение проектов, инициатив и мероприятий.",
   "Разработка концепций и реализация креативных и медиапроектов.",
   "Подготовка информационных, презентационных и коммуникационных материалов.",
   "Координация производства и распространения медиаконтента.",
   "Взаимодействие со СМИ, блогерами, лидерами общественного мнения и партнёрами.",
   "Участие в организации и проведении информационных и общественных мероприятий.",
   "Мониторинг и анализ эффективности коммуникационных мероприятий."],
  ["The Chief Specialist of the Department of Communications and Creative Projects provides information and communication support for projects and initiatives, participates in the development and implementation of communication strategies, creative and media projects.",
   "Development and implementation of communication strategies and information campaigns.",
   "Information support of projects, initiatives and events.",
   "Development of concepts and implementation of creative and media projects.",
   "Preparation of information, presentation and communication materials.",
   "Coordination of the production and distribution of media content.",
   "Interaction with the media, bloggers, opinion leaders and partners.",
   "Participation in organizing and holding information and public events.",
   "Monitoring and analysis of the effectiveness of communication activities."]),
 "Тастан Н.Е.": (
  ["Формирование концепции платформы: Разработка видения, целевой архитектуры и дорожной карты (Roadmap) развития платформы «Таза Қазақстан».",
   "Контроль разработки (R&D): Координация работы разработки и аналитики.",
   "Интеграция систем: Обеспечение интеграции платформы с существующими компонентами «электронного правительства» (egov, ситуационные центры Akimat, геопорталы).",
   "Информационная безопасность: Обеспечение соответствия платформы единым требованиям информационной безопасности Республики Казахстан."],
  ["Platform concept development: development of the vision, target architecture and roadmap for the development of the «Taza Qazaqstan» platform.",
   "R&D oversight: coordination of the work of development and analytics.",
   "Systems integration: ensuring integration of the platform with existing e-government components (egov, Akimat situation centers, geoportals).",
   "Information security: ensuring compliance of the platform with the unified information security requirements of the Republic of Kazakhstan."]),
 "Мухамадиева Н.А.": (
  ["Описание должности: Обеспечивает разработку визуальных концепций и дизайн-сопровождение коммуникационных, имиджевых и креативных проектов, формирование единого визуального стиля и подготовку графических материалов для цифровых и офлайн-площадок.",
   "Основные должностные функции:",
   "Разработка визуальных концепций, дизайн-макетов и графических материалов.",
   "Создание презентаций, инфографики, баннеров, афиш, постеров и другой визуальной продукции.",
   "Разработка и адаптация визуального контента для социальных сетей, сайтов и цифровых платформ.",
   "Участие в разработке фирменного стиля, айдентики и брендированных материалов проектов.",
   "Подготовка визуальных материалов для мероприятий, рекламных кампаний и информационного сопровождения.",
   "Обеспечение соблюдения единого визуального стиля и требований к брендингу проектов.",
   "Взаимодействие с подрядчиками, типографиями и другими исполнителями по вопросам производства визуальной продукции."],
  ["Position description: Ensures the development of visual concepts and design support for communication, image and creative projects, the formation of a unified visual style and the preparation of graphic materials for digital and offline venues.",
   "Main job functions:",
   "Development of visual concepts, design layouts and graphic materials.",
   "Creation of presentations, infographics, banners, flyers, posters and other visual products.",
   "Development and adaptation of visual content for social media, websites and digital platforms.",
   "Participation in developing corporate style, identity and branded materials for projects.",
   "Preparation of visual materials for events, advertising campaigns and information support.",
   "Ensuring compliance with the unified visual style and branding requirements of projects.",
   "Interaction with contractors, printing houses and other performers on matters of visual product production."]),
}

FP["Өміртай А."] = (
  ["Описание должности: Осуществляет руководство деятельностью департамента, организует и координирует проведение социологических исследований, мониторинг общественных настроений и аналитическую работу. Обеспечивает подготовку экспертных заключений, аналитических материалов и рекомендаций для принятия управленческих решений и реализации проектов.",
   "Основные должностные функции",
   "Организация и координация социологических исследований, опросов и мониторинга общественного мнения.",
   "Анализ общественно-политических процессов, социальных тенденций и региональной специфики.",
   "Подготовка аналитических отчетов, экспертных заключений, прогнозов и рекомендаций.",
   "Разработка методологии исследований, технических заданий и инструментов сбора данных.",
   "Мониторинг эффективности реализуемых проектов и оценка их общественного воздействия.",
   "Координация взаимодействия с научными организациями, экспертным сообществом и исследовательскими центрами.",
   "Руководство работой сотрудников департамента, контроль качества и сроков выполнения задач."],
  ["Position description: Manages the activities of the Department, organizes and coordinates sociological research, monitoring of public sentiment and analytical work. Ensures the preparation of expert opinions, analytical materials and recommendations for management decision-making and project implementation.",
   "Main job functions",
   "Organization and coordination of sociological research, surveys and monitoring of public opinion.",
   "Analysis of socio-political processes, social trends and regional specifics.",
   "Preparation of analytical reports, expert opinions, forecasts and recommendations.",
   "Development of research methodology, terms of reference and data collection tools.",
   "Monitoring of the effectiveness of implemented projects and assessment of their public impact.",
   "Coordination of interaction with scientific organizations, the expert community and research centers.",
   "Management of the Department's staff, control of the quality and timeliness of task completion."])
FP["Жумабай Б.Ғ."] = (
  ["Главный специалист Департамента коммуникаций и креативных проектов обеспечивает информационно-коммуникационное сопровождение проектов и инициатив, участвует в разработке и реализации коммуникационных стратегий, креативных и медиапроектов.",
   "Должностные функции:",
   "Разработка и реализация коммуникационных стратегий и информационных кампаний.",
   "Информационное сопровождение проектов, инициатив и мероприятий.",
   "Разработка концепций и реализация креативных и медиапроектов.",
   "Подготовка информационных, презентационных и коммуникационных материалов.",
   "Координация производства и распространения медиаконтента.",
   "Взаимодействие со СМИ, блогерами, лидерами общественного мнения и партнёрами.",
   "Участие в организации и проведении информационных и общественных мероприятий.",
   "Мониторинг и анализ эффективности коммуникационных мероприятий."],
  ["The Chief Specialist of the Department of Communications and Creative Projects provides information and communication support for projects and initiatives, participates in the development and implementation of communication strategies, creative and media projects.",
   "Job functions:",
   "Development and implementation of communication strategies and information campaigns.",
   "Information support of projects, initiatives and events.",
   "Development of concepts and implementation of creative and media projects.",
   "Preparation of information, presentation and communication materials.",
   "Coordination of the production and distribution of media content.",
   "Interaction with the media, bloggers, opinion leaders and partners.",
   "Participation in organizing and holding information and public events.",
   "Monitoring and analysis of the effectiveness of communication activities."])
FP["Дутпаева А.К."] = (
  ["Описание должности: Обеспечивает взаимодействие с инфлюенсерами, блогерами, лидерами общественного мнения и партнёрскими организациями в целях продвижения проектов, формирования общественной поддержки и расширения информационного охвата.",
   "Основные должностные функции:",
   "Поиск и привлечение блогеров и инфлюенсеров.",
   "Организация рекламных и информационных интеграций.",
   "Координация публикаций и контента в социальных сетях.",
   "Мониторинг охватов и эффективности публикаций.",
   "Ведение базы блогеров, инфлюенсеров."],
  ["Position description: Ensures interaction with influencers, bloggers, opinion leaders and partner organizations in order to promote projects, build public support and expand information reach.",
   "Main job functions:",
   "Search for and engagement of bloggers and influencers.",
   "Organization of advertising and information integrations.",
   "Coordination of publications and content on social media.",
   "Monitoring of reach and effectiveness of publications.",
   "Maintaining a database of bloggers and influencers."])
FP["Бектемиров А.Р."] = (
  ["Обеспечивает организационное сопровождение проектов, координацию проектных задач и контроль реализации мероприятий.",
   "Должностные функции:",
   "Участвует в разработке и реализации проектов.",
   "Координирует выполнение проектных задач и мероприятий.",
   "Осуществляет мониторинг сроков и результатов реализации проектов.",
   "Готовит отчёты, справки и презентационные материалы.",
   "Участвует в организации рабочих встреч и мероприятий.",
   "Выполняет поручения руководства в пределах своей компетенции."],
  ["Provides organizational support for projects, coordination of project tasks and monitoring of the implementation of activities.",
   "Job functions:",
   "Participates in the development and implementation of projects.",
   "Coordinates the performance of project tasks and activities.",
   "Monitors the timing and results of project implementation.",
   "Prepares reports, reference notes and presentation materials.",
   "Participates in organizing working meetings and events.",
   "Performs instructions of management within the scope of their competence."])

# последний пункт функционала у всех (Мурат, 03.11.2026), если по смыслу такого пункта нет
PH = ("Выполняет поручения руководства в пределах своей компетенции.",
      "Performs instructions of management within the scope of their competence.")
for _k, (_ru, _en) in FP.items():
    if any("поручени" in l.lower() for l in _ru):
        continue
    _num = (lambda t: "%d. " % (sum(1 for l in _ru if l[:1].isdigit()) + 1) + t) if _ru[-1][:1].isdigit() else (lambda t: t)
    _ru.append(_num(PH[0])); _en.append(_num(PH[1]))
# у кого функционал был одним абзацем в п. 4 - выносим его отдельным абзацем и добавляем тот же пункт
for _k, _fn in (("Серикжанова А.Т.", FUNC_PO), ("Қуандық Ж.Б.", FUNC_PM), ("Жанділдаев Б.Б.", FUNC_PM), ("Шербаева Б.Б.", FUNC_PM)):
    FP[_k] = ([_fn[0], PH[0]], [_fn[1], PH[1]])
BODY_RPR = '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>'
BODY_P = ('<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="800"/></w:tabs><w:spacing w:after="80" w:line="240" w:lineRule="auto"/>'
          '<w:ind w:left="800" w:hanging="87"/><w:jc w:val="both"/>' + BODY_RPR + '</w:pPr><w:r>' + BODY_RPR + '<w:t xml:space="preserve">{}</w:t></w:r></w:p>')

def func_block(para, lines):
    """Заголовок п. 4 без жёлтого маркера + абзацы функций (как в черновике Айгерим)."""
    head = RUN_RE.sub(lambda m: "" if "w:highlight" in m.group(0) else m.group(0), para)
    return head + "".join(BODY_P.format(escape(t)) for t in lines)

RUN_RE = re.compile(r"<w:r[ >].*?</w:r>", re.S)
HL_RE = re.compile(r"<w:highlight[^>]*/>")
T_RE = re.compile(r"(<w:t(?: [^>]*)?>)([^<]*)(</w:t>)")

def run_text(r):
    return "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", r))

def set_run_text(r, txt, highlight):
    base = r if highlight else HL_RE.sub("", r)
    txt = escape(txt)
    m = T_RE.search(base)
    return base[:m.start()] + '<w:t xml:space="preserve">' + txt + "</w:t>" + base[m.end():]

def emit(run, value):
    """Текст value в копии run; части в [квадратных скобках] остаются выделенными."""
    parts = [x for x in re.split(r"(\[[^\]]*\])", value) if x != ""]
    return "".join(set_run_text(run, x, x.startswith("[") and x.endswith("]")) for x in parts)

def fill_paragraph(para, vals, fem, idx=0):
    runs = RUN_RE.findall(para)
    out = para
    k = 0           # индекс выделенного поля
    first_done = False
    pieces = []
    pos = 0
    for m in RUN_RE.finditer(para):
        r = m.group(0)
        t = run_text(r)
        new = r
        if "w:highlight" in r and t.strip():
            if len(vals) == 1 and para.count("w:highlight") > 0 and re.search(r"прописать подробно|указываются индивидуально", para):
                # поле из нескольких выделенных кусков ("[" + "прописать подробно: " + "...]")
                v = vals[0]
                if v is not None:
                    new = emit(r, v) if not first_done else ""
                    first_done = True
            else:
                v = vals[k] if k < len(vals) else None
                k += 1
                if v is not None:
                    new = emit(r, v)
        else:
            if fem:
                if "Mr./Ms. " in t:
                    new = set_run_text(r, t.replace("Mr./Ms. ", "Ms. "), False)
                elif "Г-н/Г-жа " in t:
                    new = set_run_text(r, t.replace("Г-н/Г-жа ", "Г-жа "), False)
                elif "гражданин Республики" in t:
                    new = set_run_text(r, t.replace("гражданин", "гражданка"), "w:highlight" in r)
                elif "именуемый «" in t and idx == 6:
                    new = set_run_text(r, t.replace("именуемый", "именуемая"), "w:highlight" in r)
            elif fem is False:
                if "Mr./Ms. " in t:
                    new = set_run_text(r, t.replace("Mr./Ms. ", "Mr. "), False)
                elif "Г-н/Г-жа " in t:
                    new = set_run_text(r, t.replace("Г-н/Г-жа ", "Г-н "), False)
        pieces.append(para[pos:m.start()]); pieces.append(new); pos = m.end()
    pieces.append(para[pos:])
    return "".join(pieces)


# --- отступы и подписи по образцу ТД Өміртая, исправленному Муратом (03.11.2026) ---
TABS_EN = '<w:tabs><w:tab w:val="left" w:pos="284"/><w:tab w:val="left" w:pos="993"/></w:tabs>'
TABS_RU = '<w:tabs><w:tab w:val="left" w:pos="290"/><w:tab w:val="left" w:pos="998"/></w:tabs>'
TABS_RU_LIST = '<w:tabs><w:tab w:val="left" w:pos="290"/></w:tabs>'
HEAD_RU = ("Основные должностные функции", "Job functions", "Main job functions")
DESC_START = ("Описание должности", "Position description", "Главный специалист", "The Chief Specialist", "Обеспечивает", "Provides", "Осуществляет")

def _ptext(p):
    return re.sub(r"<[^>]+>", "", p)

def _set_ppr(p, ru, kind, num):
    """kind: 'row' (пп. 1-3), 'plain' (п. 4 и описание), 'item' (нумерованный пункт функций)."""
    m = re.search(r"<w:pPr>(.*?)</w:pPr>", p, re.S)
    inner = m.group(1)
    rpr = re.search(r"<w:rPr>.*?</w:rPr>", inner, re.S)
    rpr = rpr.group(0) if rpr else ""
    spacing = '<w:spacing w:after="%d" w:line="240" w:lineRule="auto"/>' % (0 if (kind == "row" and not ru) else 80)
    jc = '<w:jc w:val="both"/>'
    if kind == "item":
        pre = '<w:pStyle w:val="a5"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="%d"/></w:numPr>' % num
        tabs = TABS_RU_LIST if ru else TABS_EN
        ind = '<w:ind w:left="6" w:hanging="6"/>' if ru else '<w:ind w:left="0" w:firstLine="0"/>'
    else:
        pre = ""
        tabs = TABS_RU if ru else TABS_EN
        ind = '<w:ind w:left="6" w:hanging="6"/>' if ru else ""
    return p[:m.start()] + "<w:pPr>" + pre + tabs + spacing + ind + jc + rpr + "</w:pPr>" + p[m.end():]

def _fix_cell(c, ru, mode, num):
    ps = re.findall(r"<w:p[ >].*?</w:p>", c, re.S)
    out = c
    def _is_item(k):
        t = _ptext(ps[k]).strip()
        return (k >= 1 and not t[:1].isdigit() and not t.endswith(":") and t.rstrip(":") not in HEAD_RU and not t.startswith(DESC_START))
    n_items = sum(1 for k in range(len(ps)) if _is_item(k)) if mode == "func" else 0
    for k, p in enumerate(ps):
        kind = "row" if mode == "row" else "plain"
        if mode == "func" and n_items >= 2 and _is_item(k):    # одиночный пункт списком не оформляем
            kind = "item"
        out = out.replace(p, _set_ppr(p, ru, kind, num), 1)
    return out

def _drop_empty(c, keep):
    """keep(i, ps, texts) -> оставить ли пустой абзац i."""
    ps = re.findall(r"<w:p[ >].*?</w:p>", c, re.S)
    texts = [_ptext(p).strip() for p in ps]
    for i, p in enumerate(ps):
        if texts[i] == "" and not keep(i, texts):
            c = c.replace(p, "", 1)
    return c

def layout(x):
    trs = list(re.finditer(r"<w:tr[ >].*?</w:tr>", x, re.S))
    if len(trs) != 199:
        raise SystemExit("строк в таблице %d, ожидалось 199" % len(trs))
    new = {}
    for i in (165, 166, 167, 168):
        tr = trs[i].group(0)
        cells = re.findall(r"<w:tc>.*?</w:tc>", tr, re.S)
        mode = "row" if i < 168 else "func"
        for ci, c in enumerate(cells):
            tr = tr.replace(c, _fix_cell(c, ci == 1, mode, 9 if ci == 1 else 8), 1)
        if i == 166:
            tr = re.sub(r"(<w:tr[^>]*>)", r'\g<1><w:trPr><w:trHeight w:val="1072"/></w:trPr>', tr, count=1)
        new[i] = tr
    for i in (169, 198):
        tr = trs[i].group(0)
        for c in re.findall(r"<w:tc>.*?</w:tc>", tr, re.S):
            if i == 169:   # оставить один пустой абзац - перед блоком РАБОТНИК
                keep = lambda k, tx: k + 1 < len(tx) and tx[k + 1] in ("EMPLOYEE", "РАБОТНИК")
            else:          # убрать пустые после блока РАБОТНИК
                keep = lambda k, tx: not any(t in ("EMPLOYEE", "РАБОТНИК") for t in tx[:k])
            tr = tr.replace(c, _drop_empty(c, keep), 1)
        new[i] = tr
    out = []; pos = 0
    for i, m in enumerate(trs):
        out.append(x[pos:m.start()]); out.append(new.get(i, m.group(0))); pos = m.end()
    out.append(x[pos:])
    return "".join(out)


# --- уменьшенные отступы в теле договора и в Приложении 2 (по просьбе Мурата, 03.11.2026): меньше страниц при печати ---
TIGHT_AFTER, TIGHT_BEFORE = 20, 40

NUM_SHORT, NUM_LONG = 567, 720      # 10 мм и 12,7 мм (номер вида 2.1.1. шире 10 мм)

def _narrow_indent(pm):
    p = pm.group(0)
    if '<w:ind w:left="800" w:hanging="800"/>' not in p:
        return p
    m = re.search(r"<w:t(?: [^>]*)?>([^<]*)</w:t>", p)
    label = m.group(1).strip() if m else ""
    if not re.fullmatch(r"(?:\d+(?:\.\d+)*\.?|[A-Za-zА-Я]\.)", label):
        return p
    w = NUM_SHORT if len(label) <= 5 else NUM_LONG
    p = p.replace('<w:ind w:left="800" w:hanging="800"/>', '<w:ind w:left="%d" w:hanging="%d"/>' % (w, w))
    return p.replace('<w:tab w:val="left" w:pos="800"/>', '<w:tab w:val="left" w:pos="%d"/>' % w)

def tighten(x):
    trs = list(re.finditer(r"<w:tr[ >].*?</w:tr>", x, re.S))
    def sp(m):
        t = m.group(0)
        t = re.sub(r'w:after="80"', 'w:after="%d"' % TIGHT_AFTER, t)
        t = re.sub(r'w:before="160"', 'w:before="%d"' % TIGHT_BEFORE, t)
        return t
    out = []; pos = 0
    for i, m in enumerate(trs):
        out.append(x[pos:m.start()])
        blk = m.group(0)
        if i <= 163 or 170 <= i <= 197:
            blk = re.sub(r"<w:spacing [^>]*/>", sp, blk)
            blk = re.sub(r"<w:p[ >].*?</w:p>", _narrow_indent, blk, flags=re.S)
        if i in (163, 169, 198):        # подписи сторон: строка не разрывается между листами
            if "<w:trPr>" in blk:
                blk = blk.replace("<w:trPr>", "<w:trPr><w:cantSplit/>", 1)
            else:
                blk = re.sub(r"(<w:tr[ >][^>]*>(?:<w:tblPrEx>.*?</w:tblPrEx>)?)", r"\g<1><w:trPr><w:cantSplit/></w:trPr>", blk, count=1, flags=re.S)
        if i == 162:                    # «В удостоверение чего» остаётся на одном листе с подписями
            def keep(pm):
                p = pm.group(0)
                if "<w:keepNext" in p:
                    return p
                if "<w:pStyle" in p:
                    return re.sub(r"(<w:pStyle [^>]*/>)", r"\g<1><w:keepNext/>", p, count=1)
                return p.replace("<w:pPr>", "<w:pPr><w:keepNext/>", 1)
            blk = re.sub(r"<w:p[ >].*?</w:p>", keep, blk, flags=re.S)
        out.append(blk); pos = m.end()
    out.append(x[pos:])
    return "".join(out)

def numbering_add(n):
    add = "".join('<w:num w:numId="%d"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>' % k for k in (8, 9))
    return n.replace("<w:numIdMacAtCleanup", add + "<w:numIdMacAtCleanup", 1) if "<w:numIdMacAtCleanup" in n else n.replace("</w:numbering>", add + "</w:numbering>")

def build(p):
    z = zipfile.ZipFile(TPL)
    x = z.read("word/document.xml").decode("utf8")
    V = make_values(p)
    counter = [0]
    def repl(m):
        para = m.group(0)
        if "w:highlight" not in para:
            return para
        counter[0] += 1
        if counter[0] in (22, 23) and p["file"] in FP:
            return func_block(para, FP[p["file"]][1 if counter[0] == 22 else 0])
        return fill_paragraph(para, V.get(counter[0], []), p["fem"], counter[0])
    x2 = re.sub(r"<w:p[ >].*?</w:p>", repl, x, flags=re.S)
    x2 = x2.replace("проживающий (-ая)", "проживающая" if p["fem"] else "проживающий")   # формулировка адреса по полу работника
    if p.get("no_probation"):
        # по просьбе Мурата: без испытательного срока (формулировка прежнего шаблона)
        for old, new in PROBATION_OFF:
            if x2.count(old) != 1:
                raise SystemExit("п. 2.5 не найден: " + old[:40])
            x2 = x2.replace(old, new)
    x2 = tighten(layout(x2))
    minidom.parseString(x2.encode("utf8"))
    out = os.path.join(OUTDIR, f"ИТД - {p['file']} (предитоговый 2026-11-03{p.get('tag', '')}).docx")
    tmp = out + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
        for it in z.infolist():
            data = z.read(it.filename)
            if it.filename == "word/document.xml":
                data = x2.encode("utf8")
            elif it.filename == "word/numbering.xml":
                data = numbering_add(data.decode("utf8")).encode("utf8")
            zo.writestr(it, data)
    os.replace(tmp, out)
    left = len(re.findall(r"w:highlight", x2))
    return out, counter[0], left

# порядковые номера индивидуальных договоров: с 11 (Серикжанова), далее по порядку списка P
for _i, _p in enumerate(P):
    _p["num"] = str(11 + _i)

if __name__ == "__main__":
    only = sys.argv[1:]
    for p in P:
        if only and not any(o in p["file"] for o in only):
            continue
        out, n, left = build(p)
        print(f"{p['file']}: paragraphs={n}, highlight left={left}")
