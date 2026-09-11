"""
Автоматизация калькулятора смет TestQuest через Playwright.

Единственное место, где выполняются действия в браузере: открыть стенд,
перейти в «Продажи → Калькулятор смет», заполнить форму по сценарию и
выгрузить смету в Excel (скачать файл, который потом проверяется).
"""
import os
import re
import time

import playwright
from playwright.sync_api import sync_playwright

from . import calc_model

SELECT_CALCULATOR = '[data-page="calculator"]'
EXPORT_BTN = "#export-btn"

DEFAULT_MANAGER = "lyubko"


class ScenarioError(RuntimeError):
    pass


class Scenario(object):
    """Описание сценария заполнения калькулятора."""

    def __init__(self, key, **kwargs):
        self.key = key
        self.company = kwargs.get("company", "АО «Тест»")
        self.project = kwargs.get("project", "Внедрение корпоративного портала")
        self.issue_date = kwargs.get("issue_date", "2026-01-15")
        self.valid_until = kwargs.get("valid_until", "2026-04-15")
        self.manager = kwargs.get("manager", DEFAULT_MANAGER)
        self.cloud = kwargs.get("cloud", True)
        self.onprem = kwargs.get("onprem", False)
        self.licences = kwargs.get("licences", 100)
        self.term = kwargs.get("term", 12)
        self.growth = kwargs.get("growth", 1.0)
        self.discount = kwargs.get("discount", 0.0)
        self.modules = kwargs.get("modules", [])          # ids
        self.gifts = kwargs.get("gifts", [])              # ids
        self.implementation = kwargs.get("implementation", [])  # concrete rows
        self.tm_hours = kwargs.get("tm_hours", None)      # None -> оставить по умолчанию
        self.note = kwargs.get("note", "")

    def issue_date_ru(self):
        return calc_model.date_ru(self.issue_date)

    def valid_until_ru(self):
        return calc_model.date_ru(self.valid_until)

    def manager_contact(self, data):
        for manager in data.get("KP_MANAGERS", []):
            if manager.get("id") == self.manager:
                return "%s, %s, %s" % (manager["name"], manager["phone"], manager["email"])
        return ""


def open_calculator(page, base_url, timeout_ms=90000):
    """Открывает стенд и переходит в раздел «Продажи → Калькулятор смет»."""
    page.goto(base_url, wait_until="domcontentloaded", timeout=timeout_ms)
    page.wait_for_selector(SELECT_CALCULATOR, timeout=30000)
    page.click(SELECT_CALCULATOR)
    page.wait_for_selector("#companyName", timeout=30000)
    # ждём, пока приложение инициализирует форму расчёта
    page.wait_for_function("document.getElementById('export-btn') !== null", timeout=30000)
    time.sleep(0.3)


def _toggle_tariff(page, selector, wanted):
    checked = page.is_checked(selector)
    if checked != wanted:
        # чекбокс скрыт CSS (opacity: 0), эмулируем клик пользователя по label
        page.click("label.checkbox-option:has(%s) .checkbox-text" % selector)
    time.sleep(0.1)


def _set_numeric(page, selector, value):
    if value is None:
        return
    page.fill(selector, str(value))


def _switch_tab(page, tab_name):
    page.click('.config-tab[data-tab="%s"]' % tab_name)
    time.sleep(0.15)


def _apply_tariffs(page, scenario):
    if not scenario.cloud and not scenario.onprem:
        raise ScenarioError("не задан ни один тариф")
    _toggle_tariff(page, "#tariff-cloud", scenario.cloud)
    _toggle_tariff(page, "#tariff-onprem", scenario.onprem)


def _apply_meta(page, scenario):
    page.fill("#companyName", scenario.company)
    page.fill("#kpIssueDate", scenario.issue_date)
    page.fill("#kpValidUntil", scenario.valid_until)
    page.fill("#projectName", scenario.project)
    if scenario.manager:
        options = page.eval_on_selector_all(
            "#kpManager option", "els => els.map(o => o.value)")
        if scenario.manager in options:
            page.select_option("#kpManager", scenario.manager)
        else:
            page.select_option("#kpManager", index=0)


def _apply_modules(page, scenario):
    for module_id in scenario.modules:
        page.check("#module-" + module_id)
    for module_id in scenario.gifts:
        if module_id not in scenario.modules:
            page.check("#module-" + module_id)
        page.click('.module-item[data-module-id="%s"] .module-gift-btn' % module_id)


def _apply_implementation(page, scenario):
    if not scenario.implementation:
        return
    _switch_tab(page, "implementation")
    for item in scenario.implementation:
        kind = item.get("kind")
        if kind == "preset_fixed":
            page.check('.implementation-preset-checkbox[data-preset-id="%s"]' % item["preset"])
        elif kind == "preset_hourly":
            page.check('.implementation-preset-checkbox[data-preset-id="%s"]' % item["preset"])
            page.fill(
                '.implementation-row[data-preset-id="%s"] .implementation-hours' % item["preset"],
                str(item["hours"]),
            )
        elif kind == "preset_tm":
            page.check('.implementation-preset-checkbox[data-preset-id="%s"]' % item["preset"])
        elif kind == "custom":
            page.click("#implementation-add-row")
            page.fill(".implementation-row:last-child .implementation-desc", item["desc"])
            page.fill(".implementation-row:last-child .implementation-amount", str(item["amount"]))
            if item.get("vat") is not None:
                page.fill(".implementation-row:last-child .implementation-vat", str(item["vat"]))
        else:
            raise ScenarioError("неизвестный вид строки внедрения: %r" % (kind,))
    _switch_tab(page, "modules")


def _apply_tm(page, scenario):
    if scenario.tm_hours is None:
        return
    _switch_tab(page, "tm")
    page.fill("#tmHours", str(scenario.tm_hours))
    _switch_tab(page, "modules")


def fill_calculator(page, scenario):
    _apply_meta(page, scenario)
    _apply_tariffs(page, scenario)
    _set_numeric(page, "#licenceCount", scenario.licences)
    _set_numeric(page, "#termMonths", scenario.term)
    _set_numeric(page, "#coefficientGrowth", scenario.growth)
    _set_numeric(page, "#discountPercent", scenario.discount)
    _apply_modules(page, scenario)
    _apply_implementation(page, scenario)
    _apply_tm(page, scenario)


def export_estimate(page, scenario, dest_dir):
    """Выгрузка сметы по сценарию. Возвращает путь к скачанному .xlsx."""
    fill_calculator(page, scenario)
    try:
        with page.expect_download(timeout=30000) as download_info:
            page.click(EXPORT_BTN)
        download = download_info.value
        if not download.suggested_filename.endswith(".xlsx"):
            raise ScenarioError("скачался не xlsx: %s" % download.suggested_filename)
        os.makedirs(dest_dir, exist_ok=True)
        safe_key = reuse_safe(scenario.key)
        dest = os.path.join(dest_dir, "%s.xlsx" % safe_key)
        download.save_as(dest)
        return dest, download.suggested_filename
    except playwright.sync_api.Error as exc:
        hint = page.eval_on_selector("#status-hint", "(el) => el.textContent || ''")
        raise ScenarioError("не удалось выгрузить смету: %s; подсказка на странице: %s" % (exc, hint))


def reuse_safe(key):
    return re.sub(r"[^A-Za-z0-9_-]+", "_", key)


class EstimateExporter(object):
    """
    Управляет браузером и кэширует выгрузки по ключу сценария.
    Гарантирует, что каждый сценарий выгружается ровно один раз.
    """

    def __init__(self, base_url, playwright_obj, browser, dest_dir, log):
        self.base_url = base_url
        self.playwright = playwright_obj
        self.browser = browser
        self.dest_dir = dest_dir
        self.log = log
        self.cache = {}
        self._page = None

    @property
    def page(self):
        if self._page is None or self._page.is_closed():
            context = self.browser.new_context(accept_downloads=True)
            self._page = context.new_page()
        return self._page

    def fresh_page(self):
        """Новый чистый контекст (сбрасывает состояние формы)."""
        if self._page is not None and not self._page.is_closed():
            try:
                self._page.context.close()
            except Exception:  # noqa: BLE001
                pass
        self._page = None
        return self.page

    def get(self, scenario):
        """Возвращает путь к выгруженному .xlsx для сценария (с кэшем)."""
        if scenario.key in self.cache:
            return self.cache[scenario.key]
        page = self.fresh_page()
        open_calculator(page, self.base_url)
        self.log("    сценарий «%s»: заполнение и выгрузка…" % scenario.key)
        dest, filename = export_estimate(page, scenario, self.dest_dir)
        self.log("    сценарий «%s»: файл «%s» сохранён" % (scenario.key, filename))
        self.cache[scenario.key] = (dest, filename)
        return self.cache[scenario.key]

    def close(self):
        try:
            self._page.context.close()
        except Exception:  # noqa: BLE001
            pass
        try:
            self.browser.close()
        except Exception:  # noqa: BLE001
            pass