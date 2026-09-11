@echo off
REM Тесты «Калькулятор смет» (TestQuest, стенд testquest.pryaniky.com).
REM Готово «из коробки»: создаёт виртуальное окружение, ставит зависимости,
REM запускает все кейсы и возвращает код выхода 0/1.
REM
REM   set BASE_URL=https://testquest.pryaniky.com
REM   run-tests.bat

setlocal
cd /d "%~dp0"

if "%BASE_URL%"=="" set "BASE_URL=https://testquest.pryaniky.com"

echo == TestQuest: Калькулятор смет ==
echo == Стенд: %BASE_URL%

if not exist ".venv" (
    echo == создаю виртуальное окружение .venv
    python -m venv .venv
)

call .venv\Scripts\activate.bat

python -c "import playwright, openpyxl" >nul 2>&1
if errorlevel 1 (
    echo == устанавливаю зависимости ^(tests/requirements.txt^)
    python -m pip install --quiet -r tests\requirements.txt
)

echo == запускаю тесты (может занять 2-3 минуты)

python -m tests.runner --base-url "%BASE_URL%"
set RUNNER_STATUS=%errorlevel%

echo.
echo EXIT_CODE: %RUNNER_STATUS%
exit /b %RUNNER_STATUS%