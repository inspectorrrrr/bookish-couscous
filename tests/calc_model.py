"""
Эталонная модель расчёта сметы — точный порт логики js/calc.js + js/app.js.

Используется для вычисления «правильных» значений, которые затем сверяются
с выгруженным из калькулятора Excel (и с итогами на странице).
"""
import math

from . import pricedata


def js_round(x):
    """Math.round() в JS: округление половины вверх (по абсолютной величине для положительных)."""
    return math.floor(x + 0.5)


def clamp(value, low, high):
    return max(low, min(high, value))


# ---------------------------------------------------------------------------
# Расчёт пакета лицензий (PCalc.calculate + fillPrices из js/calc.js)
# ---------------------------------------------------------------------------
def price_base(data, variant):
    for base_row in data.get("PRICE_DATA", {}).get("baseValues", []):
        if base_row.get("variant") == variant:
            return base_row.get("baseCloudPrice", 0.0)
    return 0.0


def year_support_fee(data, variant):
    for base_row in data.get("PRICE_DATA", {}).get("baseValues", []):
        if base_row.get("variant") == variant:
            return base_row.get("baseYearSupportFee", 0.0)
    return 0.0


def term_discount_factor(data, term):
    """getTermDiscount из js/calc.js: (1 - скидка), максимальная применимая скидка."""
    discounts = [d for d in data.get("PRICE_DATA", {}).get("termDiscounts", [])
                 if d.get("minTerm", 0) <= term]
    if not discounts:
        return 1.0
    best = max(discounts, key=lambda d: d.get("minTerm", 0))
    return 1 - best.get("discount", 0.0) / 100.0


def licence_rows(licence_count):
    """Разбивка лицензий на строки по 100 шт (как в js/calc.js)."""
    licence_tail = licence_count % 100
    row_count = (licence_count - licence_tail) // 100
    rows = [{"licenceCount": 100} for _ in range(row_count)]
    rows.append({"licenceCount": licence_tail})
    return rows


def fill_prices(data, rows, coef, variant):
    """fillPrices из js/calc.js: цены со скидками-делителями за объём."""
    divisors = data.get("PRICE_DATA", {}).get("priceDivisors", [])
    base_price = price_base(data, variant) * coef
    cur = 0
    for idx, row in enumerate(rows):
        cur += 100
        if idx < len(divisors):
            divisor = divisors[idx].get("Divisor", 1.0)
        else:
            divisor = cur / 100.0
        row["divisorCoef"] = 1 / math.sqrt(divisor)
        row["licencePrice"] = base_price * row["divisorCoef"]
    return rows


def package_total(data, tariff_index, licence_count, term_months, coef):
    """
    runPackageCalculation -> calculate() из js/calc.js.
    coef здесь уже включает к-т модулей И к-т роста г/г.
    """
    if tariff_index == 0 and (term_months == 0 or licence_count == 0):
        return 0.0

    calc_term = term_months if tariff_index == 0 else 24
    discount = term_discount_factor(data, calc_term)

    rows = licence_rows(licence_count)
    fill_prices(data, rows, coef, 0)  # ProductVariant.OnlyPryaniks

    month_sum = sum(r["licencePrice"] * r["licenceCount"] for r in rows)
    total = month_sum * calc_term * discount

    if tariff_index == 1:  # Quest / on-prem
        support_years = math.ceil(term_months / 12)
        support_fee = year_support_fee(data, 0)
        magister_base = total / (1 + support_fee / 100.0)
        support_coef = 1.0
        if support_years > 1:
            support_coef = 1 + (support_fee / 100.0) * (support_years - 1)
        total = magister_base * support_coef

    return total


# ---------------------------------------------------------------------------
# Модули и коэффициент
# ---------------------------------------------------------------------------
def module_coef_add(module):
    if module.get("mandatory"):
        return 0.0
    if module.get("coefAdd") is not None:
        return float(module.get("coefAdd"))
    group = module.get("group") or {}
    return float(group.get("coefAdd", 0.0))


def module_coef_from_selection(data, selected_ids, gift_ids):
    adds = 0.0
    for module in pricedata.all_modules(data):
        if module["id"] in selected_ids and module["id"] not in gift_ids:
            adds += module_coef_add(module)
    return round(adds * 100) / 100.0  # roundCoef


# ---------------------------------------------------------------------------
# Внедрение
# ---------------------------------------------------------------------------
def implementation_row_amounts(item):
    if item.get("is_tm"):
        return (0, 0, 0)  # net, vat, gross
    amount_net = js_round(float(item.get("amount_net", 0) or 0))
    vat_rate = float(item.get("vat_rate", 5))
    vat_amount = js_round(amount_net * vat_rate / 100.0)
    return (amount_net, vat_amount, amount_net + vat_amount)


def _preset_order(data):
    order = []
    for group in data.get("IMPLEMENTATION_PRESET_GROUPS", []):
        for item in group.get("items", []):
            order.append(item.get("id"))
    return order


def resolve_implementation(data, items, tariff_index):
    """
    Приводит конкретные строки внедрения к единым записям и сортирует их
    так же, как DOM приложения: пресеты — в порядке каталога, произвольные
    строки — в конце (sortImplementationRowsInDom из js/app.js).
    Для облачного тарифа строки с пометкой onPremOnly исключаются.
    """
    preset_order = _preset_order(data)
    rate = float(data.get("TM_HOURLY_RATE", 5850))
    rows = []
    for item in items:
        kind = item.get("kind")
        preset_id = item.get("preset")
        preset = pricedata.get_preset(data, preset_id) if preset_id else None
        if kind == "preset_fixed":
            rows.append({
                "desc": preset["label"],
                "amount_net": preset.get("defaultPrice", 0),
                "vat_rate": 5,
                "is_tm": False,
                "group_title": preset.get("group_title", ""),
                "onprem_only": bool(preset.get("onPremOnly")),
                "preset_order": preset_order.index(preset_id) if preset_id in preset_order else 999,
            })
        elif kind == "preset_hourly":
            hours = int(item.get("hours", 0))
            amount = hours * rate
            desc = preset["label"].replace("{hours}", str(hours))
            rows.append({
                "desc": desc,
                "amount_net": amount,
                "vat_rate": 5,
                "is_tm": False,
                "group_title": preset.get("group_title", ""),
                "onprem_only": bool(preset.get("onPremOnly")),
                "preset_order": preset_order.index(preset_id) if preset_id in preset_order else 999,
            })
        elif kind == "preset_tm":
            rows.append({
                "desc": preset["label"],
                "amount_net": 0,
                "vat_rate": 0,
                "is_tm": True,
                "group_title": preset.get("group_title", ""),
                "onprem_only": bool(preset.get("onPremOnly")),
                "preset_order": preset_order.index(preset_id) if preset_id in preset_order else 999,
            })
        elif kind == "custom":
            rows.append({
                "desc": item.get("desc", ""),
                "amount_net": float(item.get("amount", 0) or 0),
                "vat_rate": float(item.get("vat", 5) or 5),
                "is_tm": False,
                "group_title": "",
                "onprem_only": False,
                "preset_order": 999,
            })
    rows.sort(key=lambda r: r["preset_order"])
    if tariff_index == 0:
        rows = [r for r in rows if not r["onprem_only"]]
    return rows


# ---------------------------------------------------------------------------
# Построение «эталонного» описания таблицы для одного тарифа
# ---------------------------------------------------------------------------
def build_section(data, tariff_index, scenario):
    """
    Возвращает словарь с ожидаемыми числами для секции тарифа:
      base_price, base_title, module_rows[(label, price)], discount,
      total_net, total_vat, total_gross, impl_totals...
    """
    growth = float(scenario.get("growth", 1.0))
    selected_ids = set(scenario.get("modules", []))
    gift_ids = set(scenario.get("gifts", []))
    module_coef = 1 + module_coef_from_selection(data, selected_ids, gift_ids)
    discount_pct = clamp(float(scenario.get("discount", 0.0)), 0.0, 100.0)

    licences = int(scenario.get("licences", 0))
    term = int(scenario.get("term", 0))

    base_package = package_total(data, tariff_index, licences, term, 1.0 * growth)
    total_package = package_total(data, tariff_index, licences, term, module_coef * growth)
    total_after_discount = total_package * (1 - discount_pct / 100.0)
    discount_amount = total_package - total_after_discount

    module_rows = []
    # Порядок строк модулей в смете такой же, как в DOM приложения:
    # по группам MODULE_GROUPS (getSelectedOptionalModules из js/app.js).
    for module in pricedata.all_modules(data):
        module_id = module["id"]
        if module_id not in selected_ids:
            continue
        is_gift = module_id in gift_ids
        price = 0 if is_gift else base_package * module_coef_add(module)
        label = "Дополнительно: модуль %s%s" % (
            pricedata.export_label(module), " (в подарок)" if is_gift else "")
        module_rows.append((label, js_round(price)))

    implementation = resolve_implementation(data, scenario.get("implementation", []), tariff_index)

    impl_net = 0
    impl_vat = 0
    impl_gross = 0
    for item in implementation:
        net, vat, gross = implementation_row_amounts(item)
        if not item.get("is_tm"):
            impl_net += net
            impl_vat += vat
            impl_gross += gross

    base_title_cloud = (
        "Пакет простых неисключительных неименных неконкурентных лицензий на ПО «Бублики», "
        "дающий право на регистрацию %d учетных записей пользователей сроком на %d месяцев. "
        "Базовая конфигурация (см. список модулей ниже)" % (licences, term)
    )
    base_title_onprem = (
        "Пакет простых неисключительных неименных неконкурентных бессрочных лицензий "
        "на ПО «Бублики», дающий право на регистрацию %d учетных записей пользователей. "
        "Базовая конфигурация (см. список модулей ниже)" % licences
    )

    section = {
        "tariff_index": tariff_index,
        "base_price": js_round(base_package),
        "base_title": base_title_cloud if tariff_index == 0 else base_title_onprem,
        "module_rows": module_rows,
        "implementation": dict(
            rows=implementation,
            net=impl_net,
            vat=impl_vat,
            gross=impl_gross,
        ),
        "discount": {
            "percent": discount_pct,
            "amount": js_round(discount_amount),
        },
        # Точная семантика js/exportEstimate.js: round(totalAfterDiscount + impl_*)
        "total_net": js_round(total_after_discount + impl_net),
        "total_vat": impl_vat,
        "total_gross": js_round(total_after_discount + impl_gross),
        "total_after_discount": js_round(total_after_discount),
        "module_coef": module_coef,
    }
    return section


def tm_block(data, tm_hours):
    rate = float(data.get("TM_HOURLY_RATE", 5850))
    hours = tm_hours if tm_hours is not None else data.get("TM_DEFAULT_HOURS", 100)
    net = js_round(hours * rate)
    vat = js_round(net * 0.05)
    gross = js_round(net + vat)
    return {
        "hours": hours,
        "net": net,
        "vat": vat,
        "gross": gross,
        "label": "Резерв на %d чч" % hours,
    }


def module_config_table(data, selected_ids, gift_ids):
    """buildModuleConfigurationForExport из js/app.js: статус каждого модуля."""
    rows = []
    index = 0
    for module in pricedata.all_modules(data):
        group = module.get("group") or {}
        if module.get("mandatory") or group.get("id") == "base":
            status = "базовая конфигурация"
        else:
            if module["id"] in selected_ids:
                status = "в подарок" if module["id"] in gift_ids else "оценено дополнительно"
            else:
                status = "не входит"
        index += 1
        label = module.get("label", "")
        if module.get("note"):
            label = "%s - %s" % (label, module["note"])
        rows.append({"index": index, "label": label, "status": status})
    return rows


def expected_filename(scenario):
    date_part = scenario.get("issue_date_ru", "")
    company = _sanitize(scenario.get("company", ""))
    project = _sanitize(scenario.get("project", ""))
    project_part = (project + " ") if project else ""
    return (
        "%s, КП по реализации проекта %sна платформе Бублики для компании %s, ООО Технологии защиты.xlsx"
        % (date_part, project_part, company)
    )


def _sanitize(name):
    for ch in '\\/:*?"<>|':
        name = name.replace(ch, "_")
    return name.strip()


def date_ru(date_value):
    parts = (date_value or "").split("-")
    if len(parts) != 3:
        return date_value or ""
    return "%s.%s.%s" % (parts[2], parts[1], parts[0])