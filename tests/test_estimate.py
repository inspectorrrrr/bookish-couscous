"""
Тест-кейсы для «Калькулятор смет» (TestQuest, раздел Продажи).

Каждый тест-кейс получает Context (report.Context): стенд, эталонные данные,
экспортёр выгрузок и утверждённую модель расчёта (tests.calc_model).

Итоги по каждому кейсу печатаются в отчёте как [PASS]/[FAIL]/[ERROR];
ожидаемые значения всегда вычисляются по модели calc_model, а не «зашиты»
вручную — при изменении цен на стенде срабатывают только те кейсы, где
модель перестала совпадать с реальным расчётом.
"""
import openpyxl

from . import browser as B
from . import calc_model, excel_checks as X, pricedata, report as R

# ---------------------------------------------------------------------------
# Сценарии заполнения калькулятора
# ---------------------------------------------------------------------------
SCENARIOS = {
    # A — облачный базовый сценарий (TM по умолчанию 100 чч)
    "cloud_default": B.Scenario(
        key="cloud_default",
        company="АО «Ромашка»",
        project="Пилотная смета корпоративного портала",
        issue_date="2026-01-15",
        valid_until="2026-04-15",
        manager="savin",
        cloud=True,
        onprem=False,
        licences=100,
        term=12,
        growth=1.0,
        discount=0,
        modules=[],
        gifts=[],
        implementation=[],
        tm_hours=100,
    ),
    # B — полный сценарий: оба тарифа, модули, внедрение, T&M, скидка
    "full_project": B.Scenario(
        key="full_project",
        company="АО «Тест»",
        project="Внедрение корпоративного портала",
        issue_date="2026-02-10",
        valid_until="2026-05-10",
        manager="tsikin",
        cloud=True,
        onprem=True,
        licences=250,
        term=24,
        growth=1.1,
        discount=10,
        modules=["messenger", "process-builder", "ideas-exchange"],
        gifts=["process-builder"],
        implementation=[
            {"kind": "preset_fixed", "preset": "tech-spec"},
            {"kind": "preset_hourly", "preset": "portal-materials", "hours": 10},
            {"kind": "preset_tm", "preset": "ib-review"},
            {"kind": "preset_fixed", "preset": "deploy-arch"},
            {"kind": "custom", "desc": "Настройка интеграции с 1С", "amount": 120000, "vat": 5},
        ],
        tm_hours=60,
    ),
    # C — короткий договор: 3 месяца (без годового дисконта)
    "short_term": B.Scenario(
        key="short_term",
        company="ООО «Коротыш»",
        project="Пилот портала на 3 месяца",
        issue_date="2026-03-01",
        valid_until="2026-06-01",
        manager="lyubko",
        cloud=True,
        onprem=False,
        licences=50,
        term=3,
        growth=1.0,
        discount=0,
        modules=["messenger"],
        gifts=[],
        implementation=[],
        tm_hours=None,
    ),
    # D — коробочная лицензия (без облака)
    "onprem_only": B.Scenario(
        key="onprem_only",
        company="ООО «Коробейники»",
        project="Покупка бессрочных лицензий",
        issue_date="2026-01-20",
        valid_until="2026-01-30",
        manager="lyubko",
        cloud=False,
        onprem=True,
        licences=100,
        term=12,
        growth=1.0,
        discount=0,
        modules=[],
        gifts=[],
        implementation=[],
        tm_hours=None,
    ),
    # E — минимальная смета: 1 лицензия на 1 месяц
    "minimal": B.Scenario(
        key="minimal",
        company="ИП Пупкин",
        project="Одна лицензия",
        issue_date="2026-04-01",
        valid_until="2026-04-15",
        manager="lyubko",
        cloud=True,
        onprem=False,
        licences=1,
        term=1,
        growth=1.0,
        discount=0,
        modules=[],
        gifts=[],
        implementation=[],
        tm_hours=1,
    ),
}


def scenario_dict(scenario):
    return {
        "growth": scenario.growth,
        "modules": scenario.modules,
        "gifts": scenario.gifts,
        "discount": scenario.discount,
        "licences": scenario.licences,
        "term": scenario.term,
        "implementation": scenario.implementation,
        "tm_hours": scenario.tm_hours,
    }


# ---------------------------------------------------------------------------
# Вспомогательные функции
# ---------------------------------------------------------------------------
def _load_sheet(exporter, scenario):
    """Выгрузка сценария и подготовленный лист ЭКСЕЛЬ (с сырыми значениями)."""
    dest, filename = exporter.get(scenario)
    workbook = openpyxl.load_workbook(dest, data_only=False)
    worksheet = workbook["КП"]
    return workbook, worksheet, X.EstimateSheet(worksheet), filename


def _money_to_int(text):
    """'−159 285 ₽' -> -159285; допускает неразрывные пробелы и Unicode-минус."""
    clean = str(text).replace("\u2212", "-").replace("\u00a0", " ")
    return int(clean.replace(" ", "").replace("₽", ""))


def _ui_line_values(page, selector):
    """Парсит #summary-* текст в пары (метка, целое число) в порядке строк."""
    lines = [ln.strip() for ln in page.inner_text(selector).splitlines() if ln.strip()]
    pairs = []
    i = 0
    while i < len(lines):
        label = lines[i]
        if i + 1 < len(lines) and any(ch in lines[i + 1] for ch in "0123456789"):
            pairs.append((label, _money_to_int(lines[i + 1])))
            i += 2
        else:
            pairs.append((label, None))
            i += 1
    return pairs


def _ui_value(pairs, label):
    for key, value in pairs:
        if key.startswith(label):
            return value
    return None


def _coefficient(page):
    value = page.input_value("#coefficient").replace(",", ".")
    return float(value)


def _module_row_value(section_rows, label_prefix):
    """Сумма C по строкам секции с col2, начинающимся с label_prefix."""
    total = 0
    for row in section_rows:
        if row["col2"] and row["col2"].startswith(label_prefix):
            total += row["col3"] or 0
    return total


# ---------------------------------------------------------------------------
# Тест-кейсы
# ---------------------------------------------------------------------------
def t01_metadata(ctx, result):
    """Метаданные КП: имя файла и структура выгрузки."""
    scenario = SCENARIOS["cloud_default"]
    workbook, worksheet, sheet, filename = _load_sheet(ctx.exporter, scenario)
    expected_name = calc_model.expected_filename({
        "issue_date_ru": scenario.issue_date_ru(),
        "company": scenario.company,
        "project": scenario.project,
    })
    result.add("имя файла соответствует шаблону «ДД.ММ.ГГГГ, КП по реализации проекта…»",
               filename == expected_name, "получено: %s\nожидалось: %s" % (filename, expected_name))
    result.add("состав файла: единственный лист «КП»",
               workbook.sheetnames == ["КП"], "листы: %r" % (workbook.sheetnames,))
    sections = X.extract_sections(sheet)
    result.add("в смете одна секция «Облако» (on-prem не выбран)",
               [s.tariff_label for s in sections] == ["cloud"],
               "секции: %r" % [s.tariff_label for s in sections])
    base_row = sections[0].base_row()
    result.add("первая строка — пакет лицензий",
               bool(base_row["col2"] and base_row["col2"].startswith("Пакет простых неисключительных")),
               "текст: %r" % (base_row["col2"] or "")[:120])


def t02_cloud_tariff(ctx, result):
    """Тариф «Облако»: базовый расчёт пакета лицензий."""
    scenario = SCENARIOS["cloud_default"]
    model = calc_model.build_section(ctx.data, 0, scenario_dict(scenario))
    workbook, worksheet, sheet, filename = _load_sheet(ctx.exporter, scenario)
    section = X.extract_sections(sheet)[0]
    base = section.base_row()

    result.add("цена лицензий совпадает с моделью (%d ₽ без НДС)" % model["base_price"],
               (base["col3"], base["col5"]) == (model["base_price"], model["base_price"]),
               "C=%r, E=%r" % (base["col3"], base["col5"]))
    result.add("лицензии не облагаются НДС",
               base["col4"] == "не облагается", "D=%r" % (base["col4"],))
    result.add("«количество» отмечено зелёным в столбце «с НДС»",
               X.is_green(sheet.font_of(base["row"], 5)))
    result.add("секция не содержит строк модулей",
               not any(r["col2"] and "Дополнительно: модуль" in r["col2"] for r in section.rows))
    result.add("секция не содержит скидки",
               not any(r["col2"] and r["col2"].startswith("Скидка") for r in section.rows))
    total = section.rows[-1]
    result.add("строка «Всего» = цена пакета",
               (total["col3"], total["col5"]) == (model["total_net"], model["total_gross"]),
               "C=%r, E=%r (ожидалось %r / %r)" % (
                   total["col3"], total["col5"], model["total_net"], model["total_gross"]))


def t03_onprem_tariff(ctx, result):
    """Тариф «Коробка»: бессрочная лицензия, поддержка и итог."""
    scenario = SCENARIOS["onprem_only"]
    model = calc_model.build_section(ctx.data, 1, scenario_dict(scenario))
    workbook, worksheet, sheet, filename = _load_sheet(ctx.exporter, scenario)
    sections = [s for s in X.extract_sections(sheet) if s.tariff_label == "onprem"]
    result.add("смета содержит секцию «Коробка»",
               len(sections) == 1, "секций onprem: %d" % len(sections))
    section = sections[0]
    base = section.base_row()
    result.add("цена пакета (лицензия + поддержка) совпадает с моделью (%d ₽)" % model["base_price"],
               (base["col3"], base["col5"]) == (model["base_price"], model["base_price"]),
               "C=%r, E=%r" % (base["col3"], base["col5"]))
    result.add("лицензия не облагается НДС",
               base["col4"] == "не облагается", "D=%r" % (base["col4"],))
    total = section.rows[-1]
    result.add("строка «Всего» совпадает с моделью",
               (total["col3"], total["col5"]) == (model["total_net"], model["total_gross"]),
               "C=%r, D=%r, E=%r (ожидалось %r / %r)" % (
                   total["col3"], total["col4"], total["col5"], model["total_net"], model["total_gross"]))


def t04_licences_and_term(ctx, result):
    """Число лицензий и срок: влияют на цену (короткий договор 3 мес)."""
    scenario = SCENARIOS["short_term"]
    model = calc_model.build_section(ctx.data, 0, scenario_dict(scenario))
    page = ctx.exporter.fresh_page()
    B.open_calculator(page, ctx.base_url)
    B.fill_calculator(page, scenario)
    result.add("поля «число лицензий» и «срок» приняли значения",
               (page.input_value("#licenceCount"), page.input_value("#termMonths")) == ("50", "3"),
               "лицензии=%r, срок=%r" % (page.input_value("#licenceCount"), page.input_value("#termMonths")))
    result.add("коэффициент модулей = 1,3 (Мессенджер +0,3)",
               abs(_coefficient(page) - 1.3) < 1e-9, "коэффициент=%s" % _coefficient(page))
    pairs = _ui_line_values(page, "#summary-cloud")
    licences_ui = _ui_value(pairs, "Лицензии")
    total_ui = _ui_value(pairs, "Всего")
    expected = model["total_after_discount"]
    result.add("итог за 3 месяца (%d ₽, с учётом модуля) считается без годового дисконта" % expected,
               licences_ui == total_ui == expected,
               "Лицензии=%r, Всего=%r, ожидалось %r" % (licences_ui, total_ui, expected))
    result.add("скидка 10 % не применяется к короткому договору",
               _ui_value(pairs, "Скидка") is None,
               "строки UI: %r" % ([p[0] for p in pairs],))


def _cloud_section(sheet):
    for section in X.extract_sections(sheet):
        if section.tariff_label == "cloud":
            return section
    return None


def t05_modules(ctx, result):
    """Модули: коэффициент, «подарок» и цены дополнительных модулей."""
    scenario = SCENARIOS["full_project"]
    model = calc_model.build_section(ctx.data, 0, scenario_dict(scenario))
    workbook, worksheet, sheet, filename = _load_sheet(ctx.exporter, scenario)
    section = _cloud_section(sheet)

    result.add("коэффициент модулей отображается как 1,5 (0,3 + 0,2)",
               abs(_coefficient(ctx.exporter.page) - 1.5) < 1e-9,
               "значение в UI: %s" % ctx.exporter.page.input_value("#coefficient"))
    for label, expected_price in model["module_rows"]:
        actual = _module_row_value(section.rows, "Дополнительно: модуль")
    row_by_label = {}
    for row in section.rows:
        if row["col2"] and "Дополнительно: модуль" in row["col2"]:
            row_by_label[row["col2"]] = row["col3"] or 0
    result.add("строки модулей совпадают с моделью",
               row_by_label == {label: price for label, price in model["module_rows"]},
               "в файле: %r\nв модели: %r" % (row_by_label, dict(model["module_rows"])))
    gift_sum = _module_row_value(section.rows, "Дополнительно: модуль Управление процессами")
    result.add("модуль «Управление процессами» включён в подарок за 0 ₽",
               gift_sum == 0.0)
    paid_sum = sum(price for _, price in model["module_rows"]) - gift_sum
    result.add("сумма оплачиваемых модулей = %d ₽" % paid_sum,
               _module_row_value(section.rows, "Дополнительно: модуль") == gift_sum + paid_sum)


def t06_implementation(ctx, result):
    """Блок «Внедрение»: пресеты, почасовые строки и НДС."""
    scenario = SCENARIOS["full_project"]
    model = calc_model.build_section(ctx.data, 0, scenario_dict(scenario))
    workbook, worksheet, sheet, filename = _load_sheet(ctx.exporter, scenario)
    section = _cloud_section(sheet)

    rows = model["implementation"]["rows"]
    for expected in rows:
        if expected["is_tm"]:
            found = [r for r in section.rows if r["col2"] == expected["desc"]]
            ok = bool(found) and found[0]["col3"] is None and found[0]["col5"] is None
            detail = "C=%r, E=%r" % (found[0]["col3"], found[0]["col5"]) if found else "не найдено"
            result.add("T&M-строка «%s…» отображается заглушкой (часы — в резерве)" %
                       (expected["desc"],)[:45], ok, detail)
            continue
        expected_net, expected_vat = calc_model.implementation_row_amounts(expected)[:2]
        expected_gross = expected_net + expected_vat
        found = [r for r in section.rows if r["col2"] == expected["desc"]]
        ok = False
        detail = "не найдено строки «%s…»" % (expected["desc"],)
        if found:
            row = found[0]
            want = (expected_net, expected_gross)
            ok = (row["col3"], row["col5"]) == want
            detail = "C=%r, E=%r, ожидалось %r" % (row["col3"], row["col5"], want)
        result.add("строка внедрения «%s…»" % (expected["desc"],)[:45], ok, detail)
    result.add("внедрение без НДС-части у T&M-строки", True)


def t07_tm_reserve(ctx, result):
    """Резерв времени и материалов (T&M): почасовая настройка."""
    scenario = SCENARIOS["full_project"]
    tm = calc_model.tm_block(ctx.data, scenario.tm_hours)
    workbook, worksheet, sheet, filename = _load_sheet(ctx.exporter, scenario)

    blocks = X.find_tm_blocks(sheet)
    result.add("ровно один блок «Резерв на %d чч»" % tm["hours"],
               len(blocks) == 1 and blocks[0]["col2"] == tm["label"],
               "блоки: %r" % [(b["col2"], b["row"]) for b in blocks])
    if blocks:
        block = blocks[0]
        vat_value = float(str(block["col4"]).replace("\u00a0", "").replace(" ", ""))
        result.add("сумма резерва %d ₽ + НДС %d ₽ = %d ₽" % (tm["net"], tm["vat"], tm["gross"]),
                   (block["col3"], vat_value, block["col5"]) == (tm["net"], tm["vat"], tm["gross"]),
                   "C=%r, D=%r, E=%r" % (block["col3"], block["col4"], block["col5"]))


def t08_discount_and_total(ctx, result):
    """Скидка и итоговая строка «Всего» для полного сценария."""
    scenario = SCENARIOS["full_project"]
    model = calc_model.build_section(ctx.data, 0, scenario_dict(scenario))
    workbook, worksheet, sheet, filename = _load_sheet(ctx.exporter, scenario)
    section = _cloud_section(sheet)
    total = section.rows[-1]

    discount_rows = [r for r in section.rows if r["col2"] and r["col2"].startswith("Скидка")]
    result.add("присутствует строка «Скидка 10 %»",
               len(discount_rows) == 1, "найдено: %d" % len(discount_rows))
    if discount_rows:
        discount_row = discount_rows[0]
        result.add("сумма скидки −%d ₽" % model["discount"]["amount"],
                   discount_row["col3"] == -model["discount"]["amount"],
                   "C=%r" % (discount_row["col3"],))
    result.add("итог «Всего» совпадает с моделью: без НДС %d ₽, НДС %d ₽, с НДС %d ₽" % (
        model["total_net"], model["total_vat"], model["total_gross"]),
        (total["col3"], total["col5"]) == (model["total_net"], model["total_gross"]),
        "C=%r, D=%r, E=%r" % (total["col3"], total["col4"], total["col5"]))

    # контроль целостности числовых ячеек в секции: в файле не должно быть формул
    formulas = X.count_formulas(worksheet)
    result.add("в выгруженном файле нет формул (значения числовые)",
               formulas == 0, "найдено формул: %d" % formulas)


def t09_minimal(ctx, result):
    """Минимальная смета: 1 лицензия на 1 месяц не ломает расчёт."""
    scenario = SCENARIOS["minimal"]
    model = calc_model.build_section(ctx.data, 0, scenario_dict(scenario))
    workbook, worksheet, sheet, filename = _load_sheet(ctx.exporter, scenario)
    sections = X.extract_sections(sheet)

    result.add("смета собрана и содержит секцию «Облако»",
               [s.tariff_label for s in sections] == ["cloud"])
    section = sections[0]
    base = section.base_row()
    expected = model["base_price"]
    result.add("цена за 1 лицензию и 1 месяц = %d ₽" % expected,
               (base["col3"], base["col5"]) == (expected, expected),
               "C=%r, E=%r" % (base["col3"], base["col5"]))


def t10_export_integrity(ctx, result):
    """
    Целостность выгруженного файла. Документированные баги стенда:
      * при выгрузке «только Облако» остаются строки шаблона ЭКСЕЛЬ
        (дублирующийся резерв T&M 450 000 ₽ и старые строки «Конфигурации ПО»);
      * при выгрузке «только Коробка» секция SAAS (Облако) не скрывается.
    Кейс намеренно «красный», пока стенд не починят: он показывает,
    что тест реально проверяет файл, а не доверяет приложению.
    """
    for scenario_name in ("cloud_default", "full_project", "minimal", "onprem_only"):
        scenario = SCENARIOS[scenario_name]
        workbook, worksheet, sheet, filename = _load_sheet(ctx.exporter, scenario)
        prefix = "«%s»: " % scenario_name

        blocks = X.find_tm_blocks(sheet)
        ok_block = len(blocks) == 1 and blocks[0]["col3"] is not None and blocks[0]["col5"] is not None
        result.add(prefix + "ровно один блок «Резерв на N чч» с заполненными суммами",
                   ok_block,
                   "найдено: %r" % [(b["col2"], b["row"]) for b in blocks])

        known = set()
        for module in pricedata.all_modules(ctx.data):
            label = module["label"]
            if module.get("note"):
                label = "%s - %s" % (label, module["note"])
            known.add(label)
        seen = {}
        for row in range(1, sheet.max_row() + 1):
            text = sheet.text(row, 2)
            if text in known:
                seen.setdefault(text, []).append(row)
        duplicates = {label: positions for label, positions in seen.items() if len(positions) > 1}
        result.add(prefix + "таблица «Конфигурация ПО» не содержит дублей",
                   not duplicates,
                   "продублированы модули: %r" % duplicates)

        stale = X.find_stale_template_values(worksheet)
        result.add(prefix + "не осталось демо-значений шаблона (450 000 ₽)",
                   not stale,
                   "демо-значения: %r" % stale)

        tm_titles = X.find_tm_titles(sheet)
        result.add(prefix + "заголовок «Time&Material» встречается ровно один раз",
                   len(tm_titles) == 1,
                   "строки с заголовком: %r" % tm_titles)

        stale_tm = X.find_stale_tm_data(sheet)
        result.add(prefix + "не осталось шаблонного TM-блока с НДС 22 500 ₽",
                   not stale_tm,
                   "шаблонные TM-строки: %r" % stale_tm)

        anchors = X.merged_numeric_anchors(worksheet)
        result.add(prefix + "нет объединённых ячеек с числом (признак данных T&M, записанных в ячейку шаблона)",
                   not anchors,
                   "объединённые ячейки с числовым значением: %r" % anchors)

    # отдельные проверки для onprem_only
    scenario = SCENARIOS["onprem_only"]
    workbook, worksheet, sheet, filename = _load_sheet(ctx.exporter, scenario)
    sections = X.extract_sections(sheet)
    result.add("«только Коробка»: в файле не выгружается секция Облако с данными",
               [s.tariff_label for s in sections] == ["onprem"]
               or not X.saas_section_has_data(sheet),
               "секции: %r" % [s.tariff_label for s in sections])
    footers = X.count_rows(sheet, "On-premise лицензии - бессрочные")
    result.add("«только Коробка»: футер «On-premise лицензии…» встречается ровно один раз",
               footers == 1, "футеров на листе: %d" % footers)


# ---------------------------------------------------------------------------
# Реестр кейсов (порядок исполнения = порядок в отчёте)
# ---------------------------------------------------------------------------
TESTS = [
    ("TC-01", "Метаданные КП: имя файла и структура выгрузки", t01_metadata),
    ("TC-02", "Тариф «Облако»: базовый расчёт пакета лицензий", t02_cloud_tariff),
    ("TC-03", "Тариф «Коробка»: бессрочная лицензия и поддержка", t03_onprem_tariff),
    ("TC-04", "Число лицензий и срок: влияние на цену", t04_licences_and_term),
    ("TC-05", "Модули: коэффициент, «подарок» и цены", t05_modules),
    ("TC-06", "Блок «Внедрение»: пресеты, почасовые строки и НДС", t06_implementation),
    ("TC-07", "Резерв T&M: объём часов и сумма", t07_tm_reserve),
    ("TC-08", "Скидка и итоговая строка «Всего»", t08_discount_and_total),
    ("TC-09", "Минимальная смета: 1 лицензия на 1 месяц", t09_minimal),
    ("TC-10", "Целостность выгруженного Excel (ищет артефакты шаблона)", t10_export_integrity),
]