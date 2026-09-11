"""
Командный интерфейс: запуск всех тест-кейсов и печать отчёта.

Использование (из корня проекта):
    python -m tests.runner [--base-url https://testquest.pryaniky.com]
    BASE_URL=... ./run-tests.sh        # или run-tests.bat в Windows

Код возврата: 0 — все кейсы PASS, 1 — есть FAIL/ERROR.
"""
import argparse
import os
import subprocess
import sys

from tests import browser as B
from tests import pricedata, report as R
from tests.test_estimate import TESTS

DEFAULT_BASE_URL = "https://testquest.pryaniky.com"


def _log(message):
    sys.stderr.write(message + "\n")
    sys.stderr.flush()


def _ensure_browser(playwright_obj, log):
    """Запускает Chromium; при отсутствии бинарника устанавливает headless-shell."""
    attempts = []

    def launch(args):
        try:
            return playwright_obj.chromium.launch(**args)
        except Exception as exc:  # noqa: BLE001
            attempts.append(str(exc).splitlines()[0])
            return None

    browser = launch({"headless": True, "args": ["--no-sandbox", "--disable-dev-shm-usage"]})
    if browser:
        return browser

    log("chromium не запустился, ставлю headless-shell…")
    subprocess.run(
        [sys.executable, "-m", "playwright", "install", "chromium", "--only-shell"],
        check=False,
    )
    browser = launch({"headless": True, "args": ["--no-sandbox", "--disable-dev-shm-usage"]})
    if browser:
        return browser

    browser = launch({"channel": "chrome", "headless": True,
                      "args": ["--no-sandbox", "--disable-dev-shm-usage"]})
    if browser:
        return browser

    raise RuntimeError(
        "Не удалось запустить браузер. Пробовали: %s\n"
        "Установите chromium вручную: python -m playwright install chromium --only-shell" % attempts)


def main(argv=None):
    parser = argparse.ArgumentParser(description="Тесты «Калькулятор смет» TestQuest")
    parser.add_argument("--base-url", default=os.environ.get("BASE_URL") or DEFAULT_BASE_URL)
    parser.add_argument("--artifacts-dir", default=os.path.abspath("artifacts"))
    args = parser.parse_args(argv)

    data, source = pricedata.load_price_data(
        args.base_url,
        lambda url: __import__("urllib.request").request.urlopen(url, timeout=30).read().decode("utf-8"),
        _log,
    )
    _log("Эталонные данные: %s" % source)

    with B.sync_playwright() as playwright_obj:
        browser = _ensure_browser(playwright_obj, _log)
        exporter = B.EstimateExporter(args.base_url, playwright_obj, browser,
                                      args.artifacts_dir, _log)
        try:
            ctx = R.Context(args.base_url, data, exporter, args.artifacts_dir, _log)
            results = []
            for case_id, title, func in TESTS:
                result = R.run_case(func, case_id, title, ctx)
                results.append(result)
        finally:
            exporter.close()

    text, verdict, failed_count = R.print_report(args.base_url, results)
    sys.stdout.write(text + "\n")
    sys.stdout.flush()
    _log("Итог: %d кейсов, отчёт выше" % len(results))
    return 1 if failed_count else 0


if __name__ == "__main__":
    sys.exit(main())