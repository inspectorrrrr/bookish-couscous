#!/usr/bin/env sh
#
# Тесты «Калькулятор смет» (TestQuest, стенд testquest.pryaniky.com).
# Готово «из коробки»: создаёт виртуальное окружение, ставит зависимости,
# запускает все кейсы и возвращает код выхода 0/1.
#
#   BASE_URL=https://testquest.pryaniky.com ./run-tests.sh
#
set -u

ROOT="$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)"
cd "$ROOT"

BASE_URL="${BASE_URL:-https://testquest.pryaniky.com}"

echo "== TestQuest: Калькулятор смет =="
echo "== Стенд: $BASE_URL"

if command -v python3 >/dev/null 2>&1; then
    PY=python3
else
    PY=python
fi

if [ ! -d .venv ]; then
    echo "== создаю виртуальное окружение .venv"
    "$PY" -m venv .venv
fi

# shellcheck disable=SC1091
. .venv/bin/activate

if ! python -c "import playwright, openpyxl" >/dev/null 2>&1; then
    echo "== устанавливаю зависимости (tests/requirements.txt)"
    python -m pip install --quiet -r tests/requirements.txt
fi

echo "== запускаю тесты (может занять 2–3 минуты)"
RUNNER_STATUS=0
python -m tests.runner --base-url "$BASE_URL"
RUNNER_STATUS=$?

echo ""
echo "EXIT_CODE: $RUNNER_STATUS"
exit $RUNNER_STATUS