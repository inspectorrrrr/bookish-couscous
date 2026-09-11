"""
Парсинг эталонных данных калькулятора (js/data.js со стенда).

Данные цен и тарифов берутся НЕ из «зашитых» в тесты констант, а со стенда —
JS-объект `window.PryanikyData` разбирается лёгким синтаксическим парсером,
поэтому тесты остаются валидными при изменении цен/модулей на стенде.

Если страницу/файл data.js не удалось загрузить, используется встроенный
фолбэк с теми же значениями (помечается в логе).
"""
import json
import os
import re

# ---------------------------------------------------------------------------
# Фолбэк-константы на случай, если стенд недоступен или data.js не парсится.
# ---------------------------------------------------------------------------
FALLBACK = {
    "PRICE_DATA": {
        "baseValues": [
            {"variant": 0, "baseCloudPrice": 208.78, "baseYearSupportFee": 30.0},
            {"variant": 1, "baseCloudPrice": 218.4, "baseYearSupportFee": 30.0},
            {"variant": 2, "baseCloudPrice": 62.4, "baseYearSupportFee": 30.0},
        ],
        "priceDivisors": [
            {"PriceTag": 100, "Divisor": 1.0},
            {"PriceTag": 200, "Divisor": 1.0},
            {"PriceTag": 300, "Divisor": 1.5},
            {"PriceTag": 400, "Divisor": 1.5},
            {"PriceTag": 500, "Divisor": 1.5},
            {"PriceTag": 600, "Divisor": 2.0},
            {"PriceTag": 700, "Divisor": 3.0},
            {"PriceTag": 800, "Divisor": 4.0},
            {"PriceTag": 900, "Divisor": 5.0},
        ],
        "termDiscounts": [
            {"minTerm": 6, "discount": 10.0},
            {"minTerm": 12, "discount": 20.0},
        ],
    },
    "TariffType": {"Test": 0, "Quest": 1},
    "ProductVariant": {"OnlyPryaniks": 0, "PryaniksIdeas": 1, "OnlyIdeas": 2},
    "TARIFF_LABELS": ["TestQuest Облако", "TestQuest Коробка"],
    "TM_DEFAULT_HOURS": 100,
    "TM_HOURLY_RATE": 5850,
    "KP_MANAGERS": [
        {"id": "lyubko", "name": "Любко Евгения", "phone": "89265826505", "email": "es@pryaniky.com"},
        {"id": "tsikin", "name": "Цыкин Алексей", "phone": "89919781154", "email": "ats@pryaniky.ru"},
        {"id": "savin", "name": "Савин Марк", "phone": "89264844907", "email": "mas@pryaniky.ru"},
    ],
    "MODULE_GROUPS": [
        {"id": "base", "title": "Базовая конфигурация", "coefAdd": 0,
         "modules": [
             {"id": "base-services", "label": "Базовые сервисы", "mandatory": True},
             {"id": "corp-network", "label": "Корпоративная социальная сеть", "mandatory": True},
             {"id": "knowledge-base", "label": "База знаний", "mandatory": True},
             {"id": "personal-account", "label": "Личный кабинет", "mandatory": True},
         ]},
        {"id": "extended", "title": "Ключевые модули", "coefAdd": 0.2,
         "modules": [
             {"id": "process-builder", "label": "Конструктор процессов", "mandatory": False},
             {"id": "corp-university", "label": "Корпоративный университет", "mandatory": False},
             {"id": "survey-builder", "label": "Конструктор опросников", "mandatory": False},
             {"id": "assessment-360", "label": "Оценка 360", "mandatory": False},
             {"id": "ideas-exchange", "label": "Биржа идей", "mandatory": False},
             {"id": "gamification-builder", "label": "Конструктор геймификации", "mandatory": False},
             {"id": "messenger", "label": "Мессенджер", "mandatory": False, "coefAdd": 0.3},
         ]},
        {"id": "other", "title": "Дополнительные модули", "coefAdd": 0.1,
         "modules": [
             {"id": "vacation-calendar", "label": "Календарь отпусков", "mandatory": False},
             {"id": "room-booking", "label": "Бронь переговорных", "mandatory": False},
             {"id": "ceo-office", "label": "Приемная руководителя", "mandatory": False},
             {"id": "mailing-builder", "label": "Конструктор рассылок", "mandatory": False},
             {"id": "vacancies", "label": "Вакансии", "mandatory": False},
             {"id": "tasks", "label": "Задачи", "mandatory": False},
             {"id": "skills", "label": "Навыки", "mandatory": False},
             {"id": "mood-color", "label": "Цвет настроения", "mandatory": False},
             {"id": "postcards", "label": "Открытки", "mandatory": False},
             {"id": "secret-santa", "label": "Тайный Санта", "mandatory": False},
         ]},
        {"id": "planned", "title": "Планируются к выходу", "coefAdd": 0.2,
         "modules": [
             {"id": "adaptation", "label": "Адаптация", "note": "планируется к выходу в Q3/2026", "mandatory": False},
             {"id": "idr", "label": "ИПР", "note": "планируется к выходу в Q3/2026", "mandatory": False},
         ]},
    ],
    "IMPLEMENTATION_PRESET_GROUPS": [
        {"id": "design", "title": "ПРОЕКТИРОВАНИЕ",
         "items": [
             {"id": "tech-spec", "label": "Подготовка технического задания", "pricing": "fixed", "defaultPrice": 117000},
             {"id": "deploy-arch", "label": "Подготовка частной архитектуры развертывания", "pricing": "fixed", "defaultPrice": 58500, "onPremOnly": True},
             {"id": "acceptance-tests", "label": "Подготовка приемочных тестов", "pricing": "fixed", "defaultPrice": 58500},
             {"id": "ib-review", "label": "Прохождение проверок со стороны ИБ", "pricing": "tm"},
         ]},
        {"id": "development", "title": "РАЗРАБОТКА",
         "items": [
             {"id": "test-env", "label": "Развертывание тестовой среды в контуре заказчика", "pricing": "fixed", "defaultPrice": 58500, "onPremOnly": True},
             {"id": "user-sync-single", "label": "Настройка синхронизации пользователей с одним источником данных заказчика", "pricing": "fixed", "defaultPrice": 187200},
             {"id": "user-sync-dual", "label": "Настройка схемы обмен данными о пользователях с комплексом систем (двумя источниками)", "pricing": "fixed", "defaultPrice": 327600},
             {"id": "portal-materials", "label": "Настройка портала по материалам заказчика в объеме {hours} чч: главная страница, навигация, брендирование, наполнение контентом, ролевая модель, уведомления и пр.", "pricing": "hourly"},
             {"id": "portal-tz", "label": "Настройка портала в соответствии с ТЗ", "pricing": "free"},
             {"id": "portal-training", "label": "Обучение рабочей группы заказчика управлению настройками портала в формате онлайн-сессий общим объемом до {hours} чч", "pricing": "hourly"},
             {"id": "acceptance-run", "label": "Выполнение приемочных испытаний в тестовой среде заказчика", "pricing": "fixed", "defaultPrice": 117000},
             {"id": "documentation", "label": "Подготовка документации", "pricing": "fixed", "defaultPrice": 58500},
         ]},
        {"id": "pilot", "title": "ПИЛОТНАЯ ЭКСПЛУАТАЦИЯ",
         "items": [
             {"id": "prod-migration", "label": "Перенос настроенного решения в продуктивный контур заказчика, подключение боевых интеграций", "pricing": "fixed", "defaultPrice": 58500, "onPremOnly": True},
             {"id": "pilot-feedback", "label": "Внесение изменений в конфигурацию портала на основе обратной связи пилотной группы", "pricing": "tm"},
         ]},
        {"id": "production", "title": "ПРОМЫШЛЕННАЯ ЭКСПЛУАТАЦИЯ",
         "items": [
             {"id": "cleanup-test-content", "label": "Зачистка тестового контента на продуктивных серверах", "pricing": "fixed", "defaultPrice": 23400},
         ]},
    ],
}

# Экспортные метки модулей (переносятся из js/exportEstimate.js в приложении).
MODULE_EXPORT_LABELS = {
    "ideas-exchange": "Биржа идей",
    "survey-builder": "Конструктор опросов",
    "assessment-360": "Оценка 360",
    "process-builder": "Управление процессами",
}

# ---------------------------------------------------------------------------
# Лёгкий парсер JS-объектного литерала (без функций/выражений).
# ---------------------------------------------------------------------------
_TOKEN_RE = re.compile(r"""
    (?P<ws>\s+)
  | (?P<comment>//[^\n]*|/\*[^*]*\*+(?:[^/*][^*]*\*+)*/)
  | (?P<str>'(?:[^'\\]|\\.)*'|"(?:[^"\\]|\\.)*")
  | (?P<num>-?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?)
  | (?P<kw>true|false|null)
  | (?P<punct>[{}[\],:])
  | (?P<ident>[A-Za-z_$][A-Za-z0-9_$]*)
""", re.VERBOSE)


def _tokenize(source):
    tokens = []
    pos = 0
    while pos < len(source):
        m = _TOKEN_RE.match(source, pos)
        if not m:
            skip = False
            # Позволяем остаточный мусор вида ';' и завершающий комментарий
            ch = source[pos]
            if ch == ';':
                pos += 1
                skip = True
            if not skip:
                raise ValueError("parser: неожиданный символ %r на %d" % (ch, pos))
            continue
        kind = m.lastgroup
        if kind in ("ws", "comment"):
            pass
        elif kind == "str":
            raw = m.group("str")
            tokens.append(("str", _unescape_js_string(raw)))
        elif kind == "num":
            tokens.append(("num", float(m.group("num"))))
        elif kind == "kw":
            tokens.append(("kw", m.group("kw")))
        elif kind == "punct":
            tokens.append(("punct", m.group("punct")))
        elif kind == "ident":
            tokens.append(("ident", m.group("ident")))
        pos = m.end()
    return tokens


def _unescape_js_string(raw):
    body = raw[1:-1]
    out = []
    i = 0
    while i < len(body):
        ch = body[i]
        if ch == "\\" and i + 1 < len(body):
            nxt = body[i + 1]
            mapping = {"n": "\n", "t": "\t", "r": "\r", '"': '"', "'": "'", "\\": "\\"}
            out.append(mapping.get(nxt, nxt))
            i += 2
        else:
            out.append(ch)
            i += 1
    return "".join(out)


class _Parser(object):
    def __init__(self, tokens):
        self.toks = tokens
        self.i = 0

    def peek(self):
        return self.toks[self.i] if self.i < len(self.toks) else (None, None)

    def advance(self):
        tok = self.peek()
        self.i += 1
        return tok

    def parse(self):
        kind, value = self.peek()
        if kind == "punct" and value == "{":
            return self._object()
        if kind == "punct" and value == "[":
            return self._array()
        raise ValueError("ожидался объект/массив, найдено %r" % (value,))

    def _object(self):
        self.advance()  # {
        obj = {}
        while True:
            kind, value = self.peek()
            if kind == "punct" and value == "}":
                self.advance()
                return obj
            if kind in ("ident", "str"):
                key = value
                self.advance()
                self._expect(":")
                obj[key] = self._value()
            else:
                # пропускаем trailing comma
                self.advance()

    def _array(self):
        self.advance()  # [
        arr = []
        while True:
            kind, value = self.peek()
            if kind == "punct" and value == "]":
                self.advance()
                return arr
            if kind == "punct" and value == ",":
                self.advance()
                continue
            arr.append(self._value())

    def _value(self):
        kind, value = self.peek()
        if kind == "punct" and value == "{":
            return self._object()
        if kind == "punct" and value == "[":
            return self._array()
        kind, value = self.advance()
        if kind == "str":
            return value
        if kind == "num":
            return value
        if kind == "kw":
            return {"true": True, "false": False, "null": None}[value]
        raise ValueError("неожиданное значение %r" % (value,))

    def _expect(self, punct):
        kind, value = self.advance()
        if kind != "punct" or value != punct:
            raise ValueError("ожидалось %r" % punct)


def parse_data_js(text):
    """Вытаскивает `window.PryanikyData = {...}` из data.js и строит словарь."""
    match = re.search(r"window\.PryanikyData\s*=\s*(\{[\s\S]*?\})\s*;", text)
    if not match:
        raise ValueError("в data.js не найден объект PryanikyData")
    tokens = _tokenize(match.group(1))
    return _Parser(tokens).parse()


def load_price_data(base_url, fetch, log):
    """
    Загружает и разбирает data.js со стенда.

    Параметры:
        base_url - корень стенда
        fetch    - callable(url) -> текст ответа (обычно urllib.request)
        log      - callable(str) для записи в лог
    Возвращает (data, источник) где источник: 'stand' | 'fallback'.
    """
    try:
        text = fetch(base_url.rstrip("/") + "/js/data.js")
    except Exception as exc:  # noqa: BLE001
        log("  [warn] data.js не загружен (%s); используются встроенные эталонные значения" % exc)
        return FALLBACK, "fallback"
    try:
        parsed = parse_data_js(text)
    except Exception as exc:  # noqa: BLE001
        log("  [warn] data.js не распознан (%s); используются встроенные эталонные значения" % exc)
        return FALLBACK, "fallback"
    log("  [ok] эталонные данные взяты со стенда (json)")
    return parsed, "stand"


def module_by_id(data, module_id):
    for group in data.get("MODULE_GROUPS", []):
        for module in group.get("modules", []):
            if module.get("id") == module_id:
                return module
    return None


def all_modules(data):
    result = []
    for group in data.get("MODULE_GROUPS", []):
        for module in group.get("modules", []):
            item = dict(module)
            item["group"] = group
            result.append(item)
    return result


def export_label(module):
    return MODULE_EXPORT_LABELS.get(module.get("id"), module.get("label"))


def get_preset(data, preset_id):
    for group in data.get("IMPLEMENTATION_PRESET_GROUPS", []):
        for item in group.get("items", []):
            if item.get("id") == preset_id:
                entry = dict(item)
                entry["group_title"] = group.get("title", "")
                return entry
    return None