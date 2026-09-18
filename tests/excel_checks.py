"""
Парсинг и проверка выгруженного из калькулятора файла .xlsx.

Файл генерируется приложением по шаблону КП, поэтому координаты ячеек
плавают (вставка строк). Все проверки работают по маркерам-текстам.
"""
import re

import openpyxl

GREEN = "00B050"


def cell_text(value):
    """Нормализует значение ячейки (в т.ч. rich text) в строку."""
    if value is None:
        return ""
    if isinstance(value, str):
        return value
    if hasattr(value, "text"):  # openpyxl RichText
        return value.text
    if isinstance(value, (int, float)):
        return str(value)
    return str(value)


def cell_number(value):
    """Значение ячейки как число, или None, если это не число."""
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    return None


def is_green(font):
    try:
        color = font and font.color
        if color is None or color.rgb is None:
            return False
        return str(color.rgb).upper().endswith(GREEN)
    except Exception:  # noqa: BLE001
        return False


def load_workbook(path):
    return openpyxl.load_workbook(path, data_only=False)


def sheet(workbook):
    return workbook.worksheets[0]


class EstimateSheet(object):
    """Обёртка над листом с удобным доступом к значениям поп-строкам."""

    def __init__(self, worksheet):
        self.ws = worksheet
        self.rows = []
        for row in worksheet.iter_rows():
            self.rows.append([c.value for c in row])

    def max_row(self):
        return len(self.rows)

    def text(self, row, col):
        if row < 1 or col < 1 or row > len(self.rows):
            return ""
        row_data = self.rows[row - 1]
        if col > len(row_data):
            return ""
        return cell_text(row_data[col - 1])

    def num(self, row, col):
        if row < 1 or col < 1 or row > len(self.rows):
            return None
        row_data = self.rows[row - 1]
        if col > len(row_data):
            return None
        return cell_number(row_data[col - 1])

    def font_of(self, row, col):
        cell = self.ws.cell(row=row, column=col)
        return cell.font

    # --- поиск строк по маркерам -----------------------------------------
    def find_rows(self, text, col=2):
        """Все строки, где текст в колонке col начинается с text."""
        out = []
        for r, row in enumerate(self.rows, start=1):
            if cell_text(row[col - 1] if col <= len(row) else None).startswith(text):
                out.append(r)
        return out

    def find_row(self, text, col=1):
        """Первая строка, где текст в колонке col начинается с text."""
        found = self.find_rows(text, col)
        return found[0] if found else None

    def all_text_rows(self, col=2):
        return {r: cell_text(self.text(r, col)) for r in range(1, self.max_row() + 1)
                if self.text(r, col)}


# ---------------------------------------------------------------------------
# Выделение секций сметы
# ---------------------------------------------------------------------------
CLOUD_MARKER = "Расчет для проекта с размещением ПО в Yandex.Cloud"
ONPREM_MARKER = "Расчет для проекта с размещением ПО в контуре"
VAT_NOT_APPLICABLE = "не облагается"
TM_TITLE = "Time&Material"
MODULE_CFG_HEADER = "Модуль"
TOTAL_TEXT = "Всего"
RESERVE_PREFIX = "Резерв на "


class Section(object):
    """Таблица одного тарифа (облако/коробка): строки до строки «Всего»."""

    def __init__(self, tariff_label, rows, total_row):
        self.tariff_label = tariff_label
        self.rows = rows          # list of dicts: row, col1..col5 texts/numbers
        self.total_row = total_row

    def base_row(self):
        for r in self.rows:
            if cell_text(r["col2"]).startswith("Пакет простых"):
                return r
        return None

    def module_rows(self):
        return [r for r in self.rows if cell_text(r["col2"]).startswith("Дополнительно: модуль ")]

    def implementation_rows(self):
        out = []
        for r in self.rows:
            text = cell_text(r["col2"])
            if not text:
                continue
            if text.startswith("Пакет простых") or text.startswith("Дополнительно: модуль "):
                continue
            if text == "Скидка" or text.startswith("Скидка "):
                continue
            if text == TOTAL_TEXT:
                continue
            out.append(r)
        return out

    def discount_row(self):
        for r in self.rows:
            if cell_text(r["col2"]).startswith("Скидка "):
                return r
        return None

    def total(self):
        return self.rows[-1] if self.rows else None


def extract_sections(sheet_):
    """
    Ищет секции по маркерам и возвращает список Section.
    Секция заканчивается строкой, где текст в B == 'Всего'.
    """
    sections = []
    starts = []
    for r in range(1, sheet_.max_row() + 1):
        t = sheet_.text(r, 1)
        if t.startswith(CLOUD_MARKER):
            starts.append((r, 0))
        elif t.startswith(ONPREM_MARKER):
            starts.append((r, 1))

    for idx, (start_row, tariff) in enumerate(starts):
        end = starts[idx + 1][0] if idx + 1 < len(starts) else sheet_.max_row() + 1
        cur = start_row + 2  # после заголовка «Nпп/Позиция/…»
        rows = []
        total_row = None
        while cur < end:
            text = sheet_.text(cur, 2)
            if text == TOTAL_TEXT:
                total_row = cur
                rows.append({
                    "row": cur,
                    "col1": sheet_.text(cur, 1),
                    "col2": text,
                    "col3": sheet_.num(cur, 3),
                    "col4": sheet_.text(cur, 4),
                    "col5": sheet_.num(cur, 5),
                })
                break
            if sheet_.text(cur, 2) or sheet_.text(cur, 3):
                rows.append({
                    "row": cur,
                    "col1": sheet_.text(cur, 1),
                    "col2": sheet_.text(cur, 2),
                    "col3": sheet_.num(cur, 3),
                    "col4": sheet_.text(cur, 4),
                    "col5": sheet_.num(cur, 5),
                    "col3_font_green": is_green(sheet_.font_of(cur, 3)),
                })
            cur += 1
        sections.append(
            Section("cloud" if tariff == 0 else "onprem", rows, total_row)
        )
    return sections


def find_module_config(sheet_):
    """
    Таблица «Конфигурация ПО»: возвращает список {index, label, status} и кол-во строк.
    Ищется по заголовку в B == 'Модуль' и D, начинающимся с 'входит в конфигурацию'.
    """
    start = None
    for r in range(1, sheet_.max_row() + 1):
        if sheet_.text(r, 2) == MODULE_CFG_HEADER:
            start = r
            break
    if start is None:
        return None, []
    rows = []
    r = start + 1
    while r <= sheet_.max_row():
        label = sheet_.text(r, 2)
        status = sheet_.text(r, 4)
        if not label and not status:
            break
        if sheet_.text(r, 1).startswith("Nпп"):
            r += 1
            continue
        rows.append({
            "row": r,
            "index": sheet_.text(r, 1),
            "label": label,
            "status": status,
        })
        r += 1
    return start, rows


def find_tm_blocks(sheet_):
    """Все строки с «Резервом на N чч» по всему листу."""
    out = []
    for r in range(1, sheet_.max_row() + 1):
        text = sheet_.text(r, 2)
        if text.startswith(RESERVE_PREFIX):
            out.append({
                "row": r,
                "col2": text,
                "col3": sheet_.num(r, 3),
                "col4": sheet_.text(r, 4),
                "col5": sheet_.num(r, 5),
            })
    return out


def count_formulas(ws):
    return sum(1 for row in ws.iter_rows() for cell in row if cell.data_type == "f")


def find_stale_template_values(ws):
    """
    Ищет «наследников» шаблонного демо-значения 100 чч × 4500 ₽ = 450 000 ₽,
    а также развёрнутые формулы (dtype 'f') — признаки невычищенных строк шаблона.
    """
    issues = []
    for row in ws.iter_rows():
        for cell in row:
            if cell.data_type == "f":
                issues.append("найден незаменённый формульный ячейка %s = %s" % (cell.coordinate, cell.value))
    for r in range(1, ws.max_row + 1):
        for col in (3, 5):
            v = ws.cell(row=r, column=col).value
            if isinstance(v, (int, float)) and abs(v - 450000) < 0.001:
                issues.append("строке %s осталось демо-значение шаблона %s" % (r, 450000))
    return issues


def find_tm_titles(sheet_):
    """Все строки с заголовком «Time&Material» в колонке B или в объединённом заголовке A."""
    out = []
    for r in range(1, sheet_.max_row() + 1):
        if sheet_.text(r, 2) == TM_TITLE or sheet_.text(r, 1).startswith(TM_TITLE):
            out.append(r)
    return out


def is_cell_merged(ws, row, col):
    """Проверяет, является ли ячейка частью объединённого диапазона."""
    for merged_range in ws.merged_cells.ranges:
        if (merged_range.min_row <= row <= merged_range.max_row
                and merged_range.min_col <= col <= merged_range.max_col):
            return True
    return False


def find_stale_tm_data(sheet_):
    """
    Ищет строки шаблонного TM-блока: «Резерв на N чч» с НДС 22 500 ₽
    (шаблонный расчёт 100 чч × 4500 ₽ × 5% = 22 500 ₽). Приложение пишет
    своё значение НДС, поэтому 22500 — признак невычищенной строки шаблона.
    """
    issues = []
    for r in range(1, sheet_.max_row() + 1):
        text = sheet_.text(r, 2)
        if not text.startswith(RESERVE_PREFIX):
            continue
        d_val = sheet_.num(r, 4)
        if d_val is not None and abs(d_val - 22500) < 0.01:
            issues.append("строка %d: шаблонный TM-блок «%s» с НДС %s" % (r, text, d_val))
    return issues


def merged_numeric_anchors(ws):
    """
    Объединённые диапазоны, у которых в «главной» ячейке лежит число.
    Так выглядят данные, попавшие в объединённую ячейку шаблона:
    значение пишется не в текст, а как индекс (число) — например, строка
    «Резерв», записанная приложением в шёблонный футер A35:E35.
    """
    out = []
    for merged_range in sorted(ws.merged_cells.ranges, key=lambda r: (r.min_row, r.min_col)):
        if merged_range.min_col != 1 or merged_range.max_col <= merged_range.min_col:
            continue
        anchor = ws.cell(row=merged_range.min_row, column=merged_range.min_col).value
        if isinstance(anchor, (int, float)) and not isinstance(anchor, bool):
            out.append((merged_range.coord, anchor))
    return out


def count_rows(sheet_, prefix, col=1):
    """Число строк, где текст в колонке col начинается с prefix."""
    count = 0
    for r in range(1, sheet_.max_row() + 1):
        if sheet_.text(r, col).startswith(prefix):
            count += 1
    return count


def saas_section_has_data(sheet_):
    """Проверяет, есть ли в секции SAAS (Yandex.Cloud) заполненные строки данных."""
    for section in extract_sections(sheet_):
        if section.tariff_label == "cloud":
            data_rows = [r for r in section.rows
                         if r["col2"] and r["col2"] != TOTAL_TEXT
                         and not r["col2"].startswith("Пакет простых")
                         and not r["col2"].startswith("Дополнительно: модуль")
                         and not r["col2"].startswith("Скидка")]
            return len(data_rows) > 0
    return False