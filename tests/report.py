"""
Мини-каркас тестового прогона: обработка тест-кейсов, агрегация и вердикт.

Формат вывода рассчитан на автоматическую разборку песочницей:
по каждому блоку печатается [PASS]/[FAIL], в конце — строка RESULT: PASSED|FAILED.
"""
import sys
import time
import traceback


class Check(object):
    def __init__(self, label, ok, detail=""):
        self.label = label
        self.ok = bool(ok)
        self.detail = detail or ""


class TestResult(object):
    def __init__(self, case_id, title):
        self.case_id = case_id
        self.title = title
        self.checks = []
        self.error = None

    def add(self, label, ok, detail=""):
        self.checks.append(Check(label, ok, detail))

    def passed(self):
        if self.error:
            return False
        return bool(self.checks) and all(c.ok for c in self.checks)

    @property
    def status(self):
        if self.error:
            return "ERROR"
        return "PASS" if self.passed() else "FAIL"


class Context(object):
    """Объект, передаваемый каждому тест-кейсу: стенд, данные, экспортёр, лог."""

    def __init__(self, base_url, data, exporter, artifacts_dir, log):
        self.base_url = base_url
        self.data = data
        self.exporter = exporter
        self.artifacts_dir = artifacts_dir
        self.log = log


def cerr(value, text=""):
    try:
        sys.stderr.write(text + value + "\n")
    except Exception:  # noqa: BLE001
        pass


def run_case(test_func, case_id, title, ctx):
    result = TestResult(case_id, title)
    started = time.time()
    try:
        test_func(ctx, result)
    except Exception as exc:  # noqa: BLE001
        result.error = "%s: %s" % (exc.__class__.__name__, exc)
        result.checks.append(Check("выполнение тест-кейса", False, traceback.format_exc()))
    result.elapsed = round(time.time() - started, 2)
    return result


def format_checks(result, indent="    "):
    lines = []
    for check in result.checks:
        mark = "OK  " if check.ok else "FAIL"
        lines.append("%s[%s] %s" % (indent, mark, check.label))
        if check.detail:
            for detail_line in str(check.detail).splitlines():
                lines.append("%s      %s" % (indent, detail_line))
    return lines


def print_report(base_url, results):
    lines = []
    lines.append("")
    lines.append("=" * 74)
    lines.append("TestQuest — Калькулятор смет")
    lines.append("Стенд: %s" % base_url)
    lines.append("Прогон: %s" % time.strftime("%Y-%m-%d %H:%M:%S"))
    lines.append("=" * 74)
    for result in results:
        lines.append("")
        lines.append("[%s] %s — %s (%s с)" % (
            result.status, result.case_id, result.title, getattr(result, "elapsed", "?")))
        if result.error:
            lines.append("    ERROR: %s" % result.error.splitlines()[0])
        lines.extend(format_checks(result))
    lines.append("")
    lines.append("-" * 74)
    passed_count = sum(1 for r in results if r.status == "PASS")
    failed_count = sum(1 for r in results if r.status != "PASS")
    verdict = "PASSED" if failed_count == 0 else "FAILED"
    lines.append("Блоков: %d, PASS: %d, FAIL/ERROR: %d" % (
        len(results), passed_count, failed_count))
    lines.append("RESULT: %s" % verdict)
    lines.append("")
    return "\n".join(lines), verdict, failed_count