# Сборка эталонного проекта ИТД ЦАКФ по замечаниям М.К. Оразбаевой (файл «примечания МК»), с учетом правил МФЦА
# (AIFC Employment Regulations) и Кадровой политики ЦАКФ. Решения Мурата: 03.11.2026.
# Исходный шаблон не меняется; результат - новый файл рядом. Нумерация: убраны п. 2.4 (нерезидент) и разд. 14 (виза),
# убран п. 11.2.4 (дубль), разделы 15-21 -> 14-20, п. 11.2.5-11.2.14 -> 11.2.4-11.2.13.
import zipfile, re, io, sys, os, shutil
from xml.dom import minidom

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
SRC = r"D:\Шаблоны\ИТД\Шаблон ИТД ЦАКФ (без персональных данных).docx"
OUT = r"D:\Шаблоны\ИТД\Шаблон ИТД ЦАКФ - эталон (проект 2026-11-03).docx"
os.makedirs(r"D:\Шаблоны\ИТД\_копии", exist_ok=True)
BK = r"D:\Шаблоны\ИТД\_копии\Шаблон ИТД ЦАКФ (без персональных данных) - до эталона (2026-11-03).docx"
if not os.path.exists(BK):
    shutil.copy2(SRC, BK)

z = zipfile.ZipFile(SRC)
x = z.read("word/document.xml").decode("utf8")

def esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

def wtext(xml):
    return "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", xml))

ROW_RE = re.compile(r"<w:tr[ >].*?</w:tr>", re.S)
CELL_RE = re.compile(r"<w:tc>.*?</w:tc>", re.S)
rows = [m for m in ROW_RE.finditer(x)]
assert len(rows) == 202

def cells_of(row):
    return CELL_RE.findall(row)

def with_cells(row, new_cells):
    old = cells_of(row)
    assert len(old) == len(new_cells)
    out = row
    for o, n in zip(old, new_cells):
        out = out.replace(o, n, 1)
    return out

def first_t_span(cell):
    m = re.search(r"<w:t(?: [^>]*)?>([^<]*)</w:t>", cell)
    return m

def set_num(cell, num):
    m = first_t_span(cell)
    assert re.match(r"^[\d.()ivx]+$", m.group(1)), m.group(1)
    return cell[:m.start()] + '<w:t xml:space="preserve">' + num + "</w:t>" + cell[m.end():]

def get_num(cell):
    return first_t_span(cell).group(1)

def set_para(cell, text, num=None):
    """Заменяет текст абзаца ячейки (номер и форматирование сохраняются)."""
    para = re.search(r"<w:p[ >].*?</w:p>", cell, re.S).group(0)
    open_tag = re.match(r"<w:p[^>]*>", para).group(0)
    ppr = re.search(r"<w:pPr>.*?</w:pPr>", para, re.S).group(0)
    runs = re.findall(r"<w:r[ >].*?</w:r>", para, re.S)
    r0 = runs[0]
    rtab = next((r for r in runs[1:] if "<w:tab/>" in r), runs[1])
    rpr = re.search(r"<w:rPr>.*?</w:rPr>", rtab, re.S).group(0)
    if num is not None:
        m = re.search(r"<w:t(?: [^>]*)?>([^<]*)</w:t>", r0)
        r0 = r0[:m.start()] + '<w:t xml:space="preserve">' + num + "</w:t>" + r0[m.end():]
    r1 = "<w:r>" + rpr + '<w:tab/><w:t xml:space="preserve">' + esc(text) + "</w:t></w:r>"
    new_para = open_tag + ppr + r0 + r1 + "</w:p>"
    return cell.replace(para, new_para, 1)

def edit_sub(cell, old, new):
    """Замена подстроки внутри одного w:t ячейки (ровно одно вхождение)."""
    hits = [m for m in re.finditer(r"<w:t(?: [^>]*)?>([^<]*)</w:t>", cell) if old in m.group(1)]
    assert len(hits) == 1, (old, len(hits))
    m = hits[0]
    t = m.group(1).replace(old, esc(new), 1)
    return cell[:m.start()] + '<w:t xml:space="preserve">' + t + "</w:t>" + cell[m.end():]

def shift_num(num):
    m = re.match(r"^(\d+)((?:\.\d+)*)(\.?)$", num)
    assert m, num
    first = int(m.group(1))
    assert 15 <= first <= 21, num
    return str(first - 1) + m.group(2) + m.group(3)

def shift_1124(num):
    m = re.match(r"^11\.2\.(\d+)$", num)
    assert m, num
    n = int(m.group(1))
    assert n >= 5
    return "11.2.%d" % (n - 1)

def make_row(template_row, num_en, en, num_ru, ru):
    c = cells_of(template_row)
    return with_cells(template_row, [set_para(c[0], en, num_en), set_para(c[1], ru, num_ru)])

# ------------------------------------------------------------------ тексты (RU / EN)
T = {}

T["2.2tail"] = (
    "unless it is terminated earlier or extended for an additional term in accordance with clause 2.3 or by written agreement of the Parties in accordance with clause 20.2.",
    "если он не будет прекращен досрочно или продлен на дополнительный срок в соответствии с пунктом 2.3 либо по письменному соглашению Сторон в соответствии с пунктом 20.2.")
T["2.3"] = (
    "The Employment Term is extended for 1 (one) year unless either Party notifies the other Party in writing of its refusal to extend not less than 30 days before the Employment Term expires.",
    "Срок действия трудового договора продлевается на 1 (один) год, если ни одна из Сторон не менее чем за 30 дней до его окончания письменно не уведомит другую Сторону об отказе от продления.")
T["2.4"] = (
    "A probation period of 3 (three) months is established for the Employee from the employment commencement date specified in clause 2.2. During the probation period, either Party may terminate this Employment Agreement by giving the other Party 1 (one) week’s written notice, or without notice on the grounds provided for in clauses 18.4.3 and 18.5 of this Employment Agreement.",
    "Работнику устанавливается испытательный срок 3 (три) месяца со дня начала работы, указанного в пункте 2.2. В период испытательного срока любая из Сторон вправе расторгнуть настоящий Трудовой договор, письменно уведомив другую Сторону за 1 (одну) неделю, либо без уведомления по основаниям, предусмотренным пунктами 18.4.3 и 18.5 настоящего Трудового договора.")
T["3h"] = ("REMUNERATION", "ОПЛАТА ТРУДА")
T["3.1"] = (
    "The Employee’s Basic Pay is a fixed monthly amount of remuneration for the performance of job duties where the standard working time has been fully worked. The amount of the Basic Pay is specified in Schedule 1 and does not include bonuses, allowances, supplements, compensation and other payments. Where the standard working time has not been fully worked, the Basic Pay shall be paid pro rata to the time worked.",
    "Должностной оклад Работника – фиксированная ежемесячная сумма оплаты за выполнение должностных обязанностей при полностью отработанной норме рабочего времени. Размер Должностного оклада указан в Приложении 1 и не включает премии, надбавки, доплаты, компенсационные и иные выплаты. При неполностью отработанной норме рабочего времени Должностной оклад выплачивается пропорционально отработанному времени.")
T["3.2"] = (
    "The Employee’s Basic Pay for each calendar month (hereinafter “Payment Period”) must be paid within 5 Business Days after the end of the Payment Period through a direct bank transfer to the personal bank account of the Employee. In case that bank account indicated by the Employee is not the bank account of Kazakhstan second-tier bank, all relevant expenses for the operations on transfer and cash withdrawal shall be borne by the Employee solely.",
    "Должностной оклад Работника за каждый календарный месяц (далее – «Период оплаты») должен быть выплачен в течение 5 Рабочих дней после окончания Периода оплаты путем прямого банковского перевода на личный банковский счёт Работника. В случае, если указанный Работником банковский счёт не является банковским счётом казахстанского банка второго уровня, все соответствующие расходы по операциям перевода и выдачи наличных денежных средств несет исключительно Работник.")
T["3.3"] = (
    "The Employer may, at its discretion, pay the Employee bonuses and other incentive payments in the manner established by the Employer’s internal documents (including the Foundation’s HR Policy). The right to such a payment arises only after the Employer has decided to pay it.",
    "Работодатель вправе по своему усмотрению выплачивать Работнику премии и другие стимулирующие выплаты в порядке, установленном внутренними документами Работодателя (включая Кадровую политику Фонда). Право на такую выплату возникает только после принятия Работодателем решения о ее выплате.")
T["3.4"] = (
    "The Employee’s wages consist of the Basic Pay, supplements and allowances provided for by this Employment Agreement, and bonuses paid in accordance with clause 3.3. Reimbursement of expenses, including business trip expenses, is not included in the wages.",
    "Заработная плата Работника состоит из Должностного оклада, доплат и надбавок, предусмотренных настоящим Трудовым договором, и премий, выплаченных в соответствии с пунктом 3.3. Возмещение расходов, в том числе командировочных, в состав заработной платы не входит.")
T["6.4"] = (
    "An Employee entitled to vacation leave may take it in periods of not less than 3 Business Days.",
    "Работник, имеющий право на отпуск, может использовать его периодами не менее 3 рабочих дней.")
T["6.11"] = (
    "During the first year of employment the Employee may take vacation leave only within the days accrued as at the date the leave begins, less the days already taken in that year. Granting vacation leave in excess of accrued days (in advance) is not permitted.",
    "В течение первого года работы Работник вправе использовать отпуск только в пределах дней, начисленных на дату начала отпуска, за вычетом уже использованных в этом году. Предоставление отпуска сверх начисленных дней (авансом) не допускается.")
T["6.12"] = (
    "For the purposes of clause 6.11, during the first year of employment 1/12 of the annual vacation entitlement under clause 6.1 is credited to the Employee on the first day of each month of employment (the first crediting being on the start date). Months of employment are counted from the start date.",
    "Для целей пункта 6.11 в течение первого года работы Работнику начисляется 1/12 годового отпуска, предусмотренного пунктом 6.1, в первый день каждого месяца работы (первое начисление – в дату начала работы). Месяцы работы считаются от даты начала работы.")
T["6.13"] = (
    "Vacation pay must be paid no later than 5 Business Days before the day the vacation leave starts.",
    "Оплата отпуска выплачивается не позднее чем за 5 рабочих дней до начала отпуска.")
T["7h"] = ("UNPAID AND ACADEMIC LEAVE", "НЕОПЛАЧИВАЕМЫЙ И АКАДЕМИЧЕСКИЙ ОТПУСК")
T["7.1"] = (
    "At the Employee’s written request and with the Employer’s consent, the Employee may be granted unpaid leave for family-related circumstances, for study purposes (academic leave) and for other valid reasons.",
    "По письменному заявлению Работника и с согласия Работодателя Работнику может быть предоставлен неоплачиваемый отпуск по семейным обстоятельствам, в связи с обучением (академический отпуск) и по иным уважительным причинам.")
T["8.1ru"] = "Работник имеет право на отдых в праздничные дни, выпадающие на рабочий день."
T["8.2ru"] = "Работник имеет право на оплату Дневной заработной платы за праздничный день, предоставленный в качестве отдыха, согласно пункту 8.1."
T["8.3"] = (
    "Leave on a National Holiday may be replaced by a day in lieu, a payment in lieu or a pro rata amount if the Employer and the Employee so agree in writing.",
    "Отдых в праздничный день может быть заменен другим днем отдыха, денежной выплатой или пропорциональной суммой по письменному соглашению Работодателя и Работника.")
T["10.1ru"] = "Работник имеет право на отпуск по материнству или отцовству (по беременности и родам, по уходу за ребенком) в соответствии с действующим законодательством."
T["10.2ru"] = "Ежегодный отпуск продолжает начисляться Работнику в течение минимального периода отпуска по материнству или отцовству и может быть использован отдельно."
T["10.3"] = (
    "During the minimum period of paternity or maternity leave of an Employee under clause 10.1, the Employer must pay the Employee the Daily Wage in the amount and manner provided for by the AIFC Employment Regulations (sections 38 and 39).",
    "В течение минимального периода отпуска по материнству или отцовству Работника в соответствии с пунктом 10.1 Работодатель выплачивает Работнику дневную заработную плату в размере и порядке, предусмотренных Правилами по трудовым отношениям МФЦА (статьи 38 и 39).")
T["10.4ru"] = "Работник не может получить компенсацию вместо отпуска по материнству или отцовству."
T["10.5ru"] = "Если праздничный день выпадает на рабочий день во время минимального периода отпуска по материнству или отцовству, минимальный период такого отпуска продлевается на период праздничного дня."
T["11.1.3"] = (
    "to the timely and complete payment of the Basic Pay and other payments provided for by this Employment Agreement, as well as bonuses where the Employer has decided to pay them;",
    "на своевременную и полную выплату Должностного оклада и других выплат, предусмотренных настоящим Трудовым договором, а также премий – в случае принятия Работодателем решения об их выплате;")
T["11.2.6"] = (  # бывш. 11.2.7 (обходной лист)
    "to submit the termination checklist to the Employer 1 day prior to the termination of this Employment Agreement or, in case of immediate termination (clauses 9.4 and 18.5), on the last working day;",
    "предоставить Работодателю обходной лист за 1 день до расторжения настоящего Трудового договора, а при немедленном прекращении трудовых отношений (пункты 9.4 и 18.5) – в последний рабочий день;")
T["11.2.10"] = (  # бывш. 11.2.11
    "to take reasonable care of any of the Employer’s property that is in the Employee’s possession, under the Employee’s control or used by the Employee or to which the Employee has access, and to treat with care the property of fellow employees;",
    "проявлять разумную заботу о любом имуществе Работодателя, которое находится во владении Работника, под контролем Работника, используется Работником или к которому Работник имеет доступ, а также бережно относиться к имуществу коллег по работе;")
T["11.2.12"] = (  # бывш. 11.2.13 (оргтехника)
    "to use the personal office equipment and devices issued by the Employer under clause 12.2.9 solely for the performance of job duties, not to transfer them to third parties, not to install software not authorised by the Employer, to promptly notify the Employer of any loss, damage or malfunction, and to return them in working condition (allowing for normal wear and tear) under a handover and acceptance certificate upon termination of this Employment Agreement or at the Employer’s earlier request, with the return recorded in the termination checklist;",
    "использовать персональную оргтехнику и оборудование, выданные Работодателем в соответствии с пунктом 12.2.9, исключительно для выполнения должностных обязанностей, не передавать их третьим лицам, не устанавливать программное обеспечение, не санкционированное Работодателем, незамедлительно уведомлять Работодателя об утрате, повреждении или неисправности, а также вернуть их в исправном состоянии (с учетом нормального износа) по акту приема-передачи при прекращении настоящего Трудового договора или по более раннему требованию Работодателя с отметкой о возврате в обходном листе;")
T["11.2.13"] = (  # бывш. 11.2.14 (ответственность за имущество)
    "to be liable for the Employer’s property (including office equipment, technical devices, documents, materials and other property) that is in the Employee’s possession or use, and in case of its loss or damage through the Employee’s fault (other than normal wear and tear), including as a result of non-compliance with clause 11.2.12, to compensate the Employer for the actual damage, taking into account the residual value of the property. Deductions from the Employee’s pay for this purpose are made only in the cases permitted by section 20 of the AIFC Employment Regulations, including with the Employee’s prior written consent.",
    "нести ответственность за имущество Работодателя (включая оргтехнику, технические устройства, документы, материалы и иное имущество), находящееся во владении или пользовании Работника, и в случае его утраты или порчи по вине Работника (за исключением нормального износа), в том числе вследствие нарушения пункта 11.2.12, возместить Работодателю фактический ущерб с учетом остаточной стоимости имущества. Удержания из заработной платы Работника для этой цели производятся только в случаях, допускаемых статьей 20 Правил по трудовым отношениям МФЦА (AIFC Employment Regulations), в том числе с предварительного письменного согласия Работника.")
T["12.2.4"] = (
    "to pay the Employee in full and on time the Basic Pay and other payments provided for by this Employment Agreement, as well as bonuses where the Employer has decided to pay them;",
    "полностью и своевременно выплачивать Работнику Должностной оклад и другие выплаты, предусмотренные настоящим Трудовым договором, а также премии – в случае принятия Работодателем решения об их выплате;")
T["12.2.8ru"] = "обрабатывать Персональные данные/чувствительные Персональные данные в соответствии с Положением о защите данных МФЦА;"
T["12.2.9"] = (
    "when personal office equipment or devices (computer/laptop, peripherals, mobile phone and other) are issued to the Employee, to hand them over for the term of this Employment Agreement under a handover and acceptance certificate and to ensure their technical maintenance.",
    "при выдаче Работнику персональной оргтехники или оборудования (компьютер/ноутбук, периферийные устройства, мобильный телефон и др.) передавать их на период срока действия трудового договора по акту приема-передачи и обеспечивать их техническое обслуживание.")
T["14.1ru"] = "Работодатель оставляет за собой право по своему усмотрению принимать решение о выплате компенсации Работнику в размере, который определит Работодатель, в случае расторжения настоящего Трудового договора в связи с роспуском или ликвидацией Фонда."
T["15.1"] = (
    "The Employer must enroll the Employee in the Kazakhstan pension scheme and make mandatory pension contributions for the Employee in accordance with the legislation of the Republic of Kazakhstan.",
    "Работодатель обязан зачислить Работника в казахстанскую пенсионную систему и производить обязательные пенсионные взносы за Работника в соответствии с законодательством Республики Казахстан.")
T["17.1"] = (
    "By signing this Employment Agreement, the Employee agrees with the terms of the Non-Disclosure Obligation set out in Schedule 2.",
    "Подписывая настоящий Трудовой договор, Работник соглашается с условиями Обязательства о неразглашении, изложенными в Приложении 2.")
T["17.3"] = (
    "Work products created by the Employee in the course of performing job duties belong to the Employer on the terms of paragraph 3 of Schedule 2.",
    "Результаты работы, созданные Работником при исполнении должностных обязанностей, принадлежат Работодателю на условиях пункта 3 Приложения 2.")
T["18.1.5"] = (
    "violation of this Employment Agreement and the Foundation’s By-laws; termination without notice is permitted only on the grounds set out in clause 18.5;",
    "нарушения условий настоящего Трудового договора и Положения Фонда; увольнение без уведомления допускается только по основаниям пункта 18.5;")
T["18.4.1"] = (
    "the Employee’s employment is terminated during the probationary period under clause 2.4;",
    "трудовые отношения Работника прекращаются в период испытательного срока в соответствии с пунктом 2.4;")
T["18.4.3ru"] = "Работник прекращает работу по основанию for cause;"
T["18.4.4ru"] = "Работодатель увольняет Работника по основанию for cause, указанному в пункте 18.5;"
T["18.5ru"] = "Работодатель может уволить Работника без уведомления по основанию for cause:"
T["18.5.1"] = (
    "because the Employee has committed a crime, if the commission of the crime has a material and detrimental impact on the Employer; or",
    "совершение Работником преступления, если оно оказало существенное и пагубное влияние на Работодателя; или")
T["18.5.2ru"] = "совершение Работником существенного нарушения положения Трудового договора; или"
T["18.5.3ru"] = "Работник существенно нарушил Правила по трудовым отношениям МФЦА (AIFC Employment Regulations), и это нарушение оказало существенное и пагубное влияние на Работодателя; или"
T["18.5.4ru"] = "по иной причине, предусмотренной Правилами по трудовым отношениям МФЦА или настоящим Трудовым договором."
T["18.6ru"] = "Расторжение по основанию for cause не включает расторжение в соответствии с пунктом 9.4."
T["19.1ru"] = "Настоящий Трудовой договор регулируется действующим законодательством МФЦА."
T["19.2"] = (
    "The following matters are governed by the Employer’s internal rules, which the Employer may change at its discretion by giving written notice to the Employee: working hours and rest; vacation leave; business trips; information security and use of IT resources; discipline and grievance procedures; the procedure and conditions for payment of bonuses. Such changes do not alter the terms of this Employment Agreement; the Employment Agreement may be amended only in accordance with clause 20.2.",
    "Следующие вопросы регулируются внутренними правилами Работодателя, которые он вправе изменять по своему усмотрению, письменно уведомив Работника: режим работы и отдыха; предоставление отпусков; командировки; информационная безопасность и использование ИТ-ресурсов; дисциплина и порядок рассмотрения жалоб; порядок и условия выплаты премий. Такие изменения не изменяют условий настоящего Трудового договора; его изменение осуществляется в соответствии с пунктом 20.2.")
T["20.1add"] = (
    " The Employee consents to the transfer of Personal Data to state authorities, banks, insurance and other organisations engaged by the Employer, including outside the AIFC, solely for the purposes specified in this clause.",
    " Работник дает согласие на передачу Персональных данных государственным органам, банкам, страховым и иным организациям, привлекаемым Работодателем, в том числе за пределы МФЦА, исключительно для целей, указанных в настоящем пункте.")
T["20.4"] = (
    "This Employment Agreement is made in English and Russian, one original copy for each Party. In case of discrepancies, the English text shall prevail.",
    "Настоящий Трудовой договор составлен на английском и русском языках, по одному оригинальному экземпляру для каждой Стороны. В случае расхождений преимущественную силу имеет текст на английском языке.")

# ------------------------------------------------------------------ обработка строк
new_rows = {}

def R(i):
    return rows[i].group(0)

def both(i, fn_en=None, fn_ru=None):
    c = cells_of(R(i))
    en, ru = c[0], c[1]
    if fn_en: en = fn_en(en)
    if fn_ru: ru = fn_ru(ru)
    return with_cells(R(i), [en, ru])

deleted = {17, 47, 48, 84, 117, 118, 119, 120}   # 2.4; 7.1.1; 7.2; 11.2.4; разд. 14 (4 строки)

for i in range(len(rows)):
    if i in deleted:
        new_rows[i] = []
        continue
    row = R(i)
    c = cells_of(row)
    out = [row]

    # ---- раздел 2
    if i == 15:
        row = both(i, lambda e: edit_sub(e, "unless it is terminated earlier or extended for an additional term under Foundation’s By-Laws.", T["2.2tail"][0]),
                      lambda r: edit_sub(r, "если он не будет прекращен досрочно или продлен на дополнительный срок в соответствии с Положением Фонда.", T["2.2tail"][1]))
        out = [row]
    elif i == 16:
        out = [with_cells(row, [set_para(c[0], T["2.3"][0]), set_para(c[1], T["2.3"][1])])]
    elif i == 18:
        out = [with_cells(row, [set_para(c[0], T["2.4"][0], "2.4."), set_para(c[1], T["2.4"][1], "2.4.")])]
    # ---- сквозные правки терминов/знаков
    elif i == 4:
        out = [both(i, None, lambda r: edit_sub(r, "а вместе - ", "а вместе – "))]
    elif i == 24:
        out = [both(i, lambda e: edit_sub(e, "retain the position, Grade and the rate of Basic Pay", "retain the position and the rate of Basic Pay"), None)]
    elif i == 27:
        out = [both(i, None, lambda r: edit_sub(r, "стандартный рабочий день - 8 часов", "стандартный рабочий день – 8 часов"))]
    elif i == 68:
        out = [both(i, None, lambda r: edit_sub(r, "уставными документами Фонда", "Положением Фонда"))]
    # ---- раздел 3
    elif i == 19:
        out = [with_cells(row, [set_para(c[0], T["3h"][0]), set_para(c[1], T["3h"][1])])]
    elif i == 20:
        out = [with_cells(row, [set_para(c[0], T["3.1"][0]), set_para(c[1], T["3.1"][1])])]
    elif i == 21:
        out = [with_cells(row, [set_para(c[0], T["3.2"][0]), set_para(c[1], T["3.2"][1])]),
               make_row(row, "3.3.", T["3.3"][0], "3.3.", T["3.3"][1]),
               make_row(row, "3.4.", T["3.4"][0], "3.4.", T["3.4"][1])]
    # ---- раздел 6
    elif i == 31:
        out = [both(i, None, lambda r: edit_sub(edit_sub(r, "20 рабочих дня", "20 рабочих дней"), "национальным и государственным праздникам", "праздничным дням"))]
    elif i == 34:
        out = [with_cells(row, [set_para(c[0], T["6.4"][0]), set_para(c[1], T["6.4"][1])])]
    elif i == 40:
        out = [both(i, lambda e: edit_sub(e, "at least 7 days prior written notice", "at least 5 Business Days’ prior written notice"),
                       lambda r: edit_sub(edit_sub(r, "не менее чем за 7 дней", "не менее чем за 5 рабочих дней"), "в соответствии с пунктом 6.9", "в соответствии с пунктом 6.9."))]
    elif i == 41:
        out = [both(i, lambda e: edit_sub(e, "at least 7 days prior written notice", "at least 15 Business Days’ prior written notice"),
                       lambda r: edit_sub(r, "не менее чем за 7 дней", "не менее чем за 15 рабочих дней"))]
    elif i == 43:
        out = [with_cells(row, [set_para(c[0], T["6.11"][0]), set_para(c[1], T["6.11"][1])])]
    elif i == 44:
        out = [with_cells(row, [set_para(c[0], T["6.12"][0]), set_para(c[1], T["6.12"][1])]),
               make_row(row, "6.13.", T["6.13"][0], "6.13.", T["6.13"][1])]
    # ---- раздел 7
    elif i == 45:
        out = [with_cells(row, [set_para(c[0], T["7h"][0]), set_para(c[1], T["7h"][1])])]
    elif i == 46:
        out = [with_cells(row, [set_para(c[0], T["7.1"][0]), set_para(c[1], T["7.1"][1])])]
    # ---- раздел 8
    elif i == 49:
        out = [both(i, None, lambda r: edit_sub(r, "ОТПУСК В ДНИ НАЦИОНАЛЬНЫХ И ГОСУДАРСТВЕННЫХ ПРАЗДНИКОВ", "ПРАЗДНИЧНЫЕ ДНИ"))]
    elif i == 50:
        out = [with_cells(row, [c[0], set_para(c[1], T["8.1ru"])])]
    elif i == 51:
        out = [with_cells(row, [c[0], set_para(c[1], T["8.2ru"])]),
               make_row(row, "8.3.", T["8.3"][0], "8.3.", T["8.3"][1])]
    # ---- раздел 10
    elif i == 60:
        out = [with_cells(row, [c[0], set_para(c[1], "ОТПУСК ПО МАТЕРИНСТВУ И ОТЦОВСТВУ")])]
    elif i == 61:
        out = [with_cells(row, [c[0], set_para(c[1], T["10.1ru"])])]
    elif i == 62:
        out = [with_cells(row, [c[0], set_para(c[1], T["10.2ru"])])]
    elif i == 63:
        out = [with_cells(row, [set_para(c[0], T["10.3"][0]), set_para(c[1], T["10.3"][1])])]
    elif i == 64:
        out = [with_cells(row, [c[0], set_para(c[1], T["10.4ru"])])]
    elif i == 65:
        out = [with_cells(row, [c[0], set_para(c[1], T["10.5ru"])])]
    # ---- раздел 11
    elif i == 70:
        out = [with_cells(row, [set_para(c[0], T["11.1.3"][0]), set_para(c[1], T["11.1.3"][1])])]
    elif i == 79:
        out = [both(i, lambda e: edit_sub(e, "Employer’s representative;", "Employer’s representative."),
                       lambda r: edit_sub(r, "представителя Работодателя;", "представителя Работодателя."))]
    elif 85 <= i <= 97:
        if i in (90, 91, 92):
            out = [row]
        else:
            nn = shift_1124(get_num(c[0]))
            if i == 85:
                row = both(i, lambda e: edit_sub(e, "section 18", "section 17"), lambda r: edit_sub(r, "разделе 18", "разделе 17"))
                c = cells_of(row)
                out = [with_cells(row, [set_num(c[0], nn), set_num(c[1], nn)])]
            elif i == 87:
                out = [with_cells(row, [set_para(c[0], T["11.2.6"][0], nn), set_para(c[1], T["11.2.6"][1], nn)])]
            elif i == 94:
                out = [with_cells(row, [set_para(c[0], T["11.2.10"][0], nn), set_para(c[1], T["11.2.10"][1], nn)])]
            elif i == 96:
                out = [with_cells(row, [set_para(c[0], T["11.2.12"][0], nn), set_para(c[1], T["11.2.12"][1], nn)])]
            elif i == 97:
                out = [with_cells(row, [set_para(c[0], T["11.2.13"][0], nn), set_para(c[1], T["11.2.13"][1], nn)])]
            else:
                out = [with_cells(row, [set_num(c[0], nn), set_num(c[1], nn)])]
    # ---- раздел 12
    elif i == 100:
        out = [both(i, None, lambda r: edit_sub(r, "уставными документами Фонда", "Положением Фонда"))]
    elif i == 109:
        out = [with_cells(row, [set_para(c[0], T["12.2.4"][0]), set_para(c[1], T["12.2.4"][1])])]
    elif i == 113:
        out = [with_cells(row, [c[0], set_para(c[1], T["12.2.8ru"])])]
    elif i == 114:
        out = [with_cells(row, [set_para(c[0], T["12.2.9"][0]), set_para(c[1], T["12.2.9"][1])])]
    # ---- разделы 15-21 -> 14-20 (строки 121..164)
    elif 121 <= i <= 164:
        ne, nr = shift_num(get_num(c[0])), shift_num(get_num(c[1]))
        if i == 122:    # 15.1 -> 14.1
            out = [with_cells(row, [set_num(c[0], ne), set_para(c[1], T["14.1ru"], nr)])]
        elif i == 124:  # 16.1 -> 15.1
            out = [with_cells(row, [set_para(c[0], T["15.1"][0], ne), set_para(c[1], T["15.1"][1], nr)])]
        elif i == 128:  # 18.1 -> 17.1
            out = [with_cells(row, [set_para(c[0], T["17.1"][0], ne), set_para(c[1], T["17.1"][1], nr)])]
        elif i == 129:  # 18.2 -> 17.2 (+ новый 17.3)
            cc = cells_of(both(i, lambda e: edit_sub(e, "subclause 19.1.5", "subclause 18.1.5"), lambda r: edit_sub(r, "подпунктом 19.1.5", "подпунктом 18.1.5")))
            out = [with_cells(row, [set_num(cc[0], ne), set_num(cc[1], nr)]),
                   make_row(row, "17.3.", T["17.3"][0], "17.3.", T["17.3"][1])]
        elif i == 136:  # 19.1.5 -> 18.1.5
            out = [with_cells(row, [set_para(c[0], T["18.1.5"][0], ne), set_para(c[1], T["18.1.5"][1], nr)])]
        elif i == 137:  # 19.1.6
            cc = cells_of(both(i, lambda e: edit_sub(e, "clause 19.4", "clause 18.4"), lambda r: edit_sub(r, "п. 19.4", "п. 18.4")))
            out = [with_cells(row, [set_num(cc[0], ne), set_num(cc[1], nr)])]
        elif i == 138:  # 19.2
            cc = cells_of(both(i, lambda e: edit_sub(e, "clause 19.4", "clause 18.4"), lambda r: edit_sub(r, "пункта 19.4", "пункта 18.4")))
            out = [with_cells(row, [set_num(cc[0], ne), set_num(cc[1], nr)])]
        elif i in (139, 140, 141):  # 19.2.1-3: номер без точки, тире «–»
            nr2 = nr.rstrip(".")
            r_txt = lambda r: edit_sub(r, "—", "–")
            cc = cells_of(both(i, None, r_txt))
            out = [with_cells(row, [set_num(cc[0], ne), set_num(cc[1], nr2)])]
        elif i == 142:  # 19.3
            cc = cells_of(both(i, lambda e: edit_sub(e, "Notwithstanding 19.2", "Notwithstanding 18.2"), lambda r: edit_sub(r, "Несмотря на пункт 19.2", "Несмотря на пункт 18.2")))
            out = [with_cells(row, [set_num(cc[0], ne), set_num(cc[1], nr)])]
        elif i == 143:  # 19.4
            cc = cells_of(both(i, lambda e: edit_sub(e, "Clause 19.2", "Clause 18.2"), lambda r: edit_sub(r, "Пункт 19.2", "Пункт 18.2")))
            out = [with_cells(row, [set_num(cc[0], ne), set_num(cc[1], nr)])]
        elif i == 144:  # 19.4.1
            out = [with_cells(row, [set_para(c[0], T["18.4.1"][0], ne), set_para(c[1], T["18.4.1"][1], nr)])]
        elif i == 146:  # 19.4.3
            out = [with_cells(row, [set_num(c[0], ne), set_para(c[1], T["18.4.3ru"], nr)])]
        elif i == 147:  # 19.4.4
            cc = cells_of(both(i, lambda e: edit_sub(e, "clause 19.5", "clause 18.5"), None))
            out = [with_cells(row, [set_num(cc[0], ne), set_para(c[1], T["18.4.4ru"], nr)])]
        elif i == 148:  # 19.5
            out = [with_cells(row, [set_num(c[0], ne), set_para(c[1], T["18.5ru"], nr)])]
        elif i == 149:
            out = [with_cells(row, [set_para(c[0], T["18.5.1"][0], ne), set_para(c[1], T["18.5.1"][1], nr)])]
        elif i == 150:
            out = [with_cells(row, [set_num(c[0], ne), set_para(c[1], T["18.5.2ru"], nr)])]
        elif i == 151:
            out = [with_cells(row, [set_num(c[0], ne), set_para(c[1], T["18.5.3ru"], nr)])]
        elif i == 152:
            out = [with_cells(row, [set_num(c[0], ne), set_para(c[1], T["18.5.4ru"], nr)])]
        elif i == 153:  # 19.6
            out = [with_cells(row, [set_num(c[0], ne), set_para(c[1], T["18.6ru"], nr)])]
        elif i == 156:  # 20.1
            out = [with_cells(row, [set_num(c[0], ne), set_para(c[1], T["19.1ru"], nr)])]
        elif i == 157:  # 20.2
            out = [with_cells(row, [set_para(c[0], T["19.2"][0], ne), set_para(c[1], T["19.2"][1], nr)])]
        elif i == 161:  # 21.1 -> 20.1 (+ получатели данных)
            en_t = wtext(c[0]).split(get_num(c[0]), 1)[1]
            ru_t = wtext(c[1]).split(get_num(c[1]), 1)[1]
            out = [with_cells(row, [set_para(c[0], en_t.rstrip() + T["20.1add"][0], ne), set_para(c[1], ru_t.rstrip() + T["20.1add"][1], nr)])]
        elif i == 164:  # 21.4 -> 20.4
            out = [with_cells(row, [set_para(c[0], T["20.4"][0], ne), set_para(c[1], T["20.4"][1], nr)])]
        else:
            out = [with_cells(row, [set_num(c[0], ne), set_num(c[1], nr)])]
    # ---- приложения
    elif i == 167:
        out = [both(i, None, lambda r: edit_sub(r, "ОСНОВНАЯ ЗАРАБОТНАЯ ПЛАТА", "ДОЛЖНОСТНОЙ ОКЛАД"))]
    elif i == 169:
        out = [both(i, lambda e: edit_sub(e, " before taxes and other obligatory deductions.", " per month before taxes and other mandatory deductions."),
                       lambda r: edit_sub(r, " до вычета налогов и других обязательных отчислений.", " в месяц до вычета налогов и других обязательных удержаний."))]
    elif i == 173:
        out = [both(i, None, lambda r: edit_sub(r, "ОБЯЗАТЕЛЬСТВО НЕРАСКРЫТИЯ ИНФОРМАЦИИ", "ОБЯЗАТЕЛЬСТВО О НЕРАЗГЛАШЕНИИ ИНФОРМАЦИИ"))]
    elif i == 177:
        out = [both(i, None, lambda r: edit_sub(r, "в связи с неисполнением или исполнением", "в связи с исполнением или принудительным исполнением"))]
    elif i == 185:
        out = [both(i, None, lambda r: edit_sub(r, "является единственным владельцем все такие рабочие продукты.", "является единственным владельцем всех таких рабочих продуктов."))]
    elif i == 193:
        out = [both(i, None, lambda r: edit_sub(edit_sub(r, "удалена с рабочего места", "вынесена с рабочего места"), "загружена. на", "загружена на"))]
    elif i == 197:
        out = [both(i, None, lambda r: edit_sub(r, "настоящим Соглашением", "настоящим Трудовым договором"))]
    elif i == 198:
        out = [both(i, None, lambda r: edit_sub(edit_sub(r, "уличенный в нарушении или нарушившем обязательство", "нарушивший обязательство"), "настоящим Соглашением", "настоящим Трудовым договором"))]
    elif i == 199:
        out = [both(i, None, lambda r: edit_sub(edit_sub(r, "по настоящему Соглашению", "по настоящему Трудовому договору"), "предмета настоящего Соглашения", "предмета настоящего Трудового договора"))]
    new_rows[i] = out

# сборка документа
parts, last = [], 0
for i, m in enumerate(rows):
    parts.append(x[last:m.start()])
    parts.append("".join(new_rows[i]))
    last = m.end()
parts.append(x[last:])
x2 = "".join(parts)

# нумерация страниц (нижний колонтитул)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
FOOT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr ' + W + '><w:p><w:pPr><w:jc w:val="center"/></w:pPr>'
        '<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
        '<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r>'
        '<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr><w:t>1</w:t></w:r>'
        '<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>')
assert x2.count("<w:sectPr") == 1
assert "footerReference" not in x2
m = re.search(r"<w:sectPr[^>]*>", x2)
x2 = x2[:m.end()] + '<w:footerReference w:type="default" r:id="rIdFooterEt1"/>' + x2[m.end():]
assert 'xmlns:r=' in x2[:2000]
minidom.parseString(x2.encode("utf8"))

ct = z.read("[Content_Types].xml").decode("utf8")
ct = ct.replace("</Types>", '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>')
rels = z.read("word/_rels/document.xml.rels").decode("utf8")
rels = rels.replace("</Relationships>", '<Relationship Id="rIdFooterEt1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>')

tmp = OUT + ".tmp"
with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
    for it in z.infolist():
        data = z.read(it.filename)
        if it.filename == "word/document.xml": data = x2.encode("utf8")
        elif it.filename == "[Content_Types].xml": data = ct.encode("utf8")
        elif it.filename == "word/_rels/document.xml.rels": data = rels.encode("utf8")
        zi = zipfile.ZipInfo(it.filename, it.date_time); zi.compress_type = zipfile.ZIP_DEFLATED
        zo.writestr(zi, data)
    zi = zipfile.ZipInfo("word/footer1.xml"); zi.compress_type = zipfile.ZIP_DEFLATED
    zo.writestr(zi, FOOT.encode("utf8"))
os.replace(tmp, OUT)
print("собран:", OUT)
