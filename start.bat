@echo off
echo ========================================
echo  Report Automate - Швидкий старт
echo ========================================
echo.

REM Перевірка Node.js
where node >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo [ПОМИЛКА] Node.js не встановлено!
    echo.
    echo Завантажте з https://nodejs.org/
    pause
    exit /b 1
)

echo [OK] Node.js встановлено
node --version
echo.

REM Перевірка npm
where npm >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo [ПОМИЛКА] npm не встановлено!
    pause
    exit /b 1
)

echo [OK] npm встановлено
npm --version
echo.

REM Перевірка node_modules
if not exist "node_modules\" (
    echo [INFO] Встановлення залежностей...
    call npm install
    if %ERRORLEVEL% NEQ 0 (
        echo [ПОМИЛКА] Не вдалось встановити залежності
        pause
        exit /b 1
    )
    echo.
)

REM Перевірка .env
if not exist ".env" (
    echo [УВАГА] Файл .env не знайдено!
    echo.
    if exist ".env.example" (
        echo Створюю .env з .env.example...
        copy .env.example .env >nul
        echo [OK] .env створено
        echo.
        echo ВАЖЛИВО: Відредагуйте файл .env з вашими параметрами!
        echo Натисніть будь-яку клавішу для відкриття .env...
        pause >nul
        notepad .env
    ) else (
        echo [ПОМИЛКА] .env.example не знайдено!
        pause
        exit /b 1
    )
)

REM Створення папки output
if not exist "output\" (
    mkdir output
    echo [OK] Створено папку output/
)

echo.
echo ========================================
echo  Запуск програми...
echo ========================================
echo.

REM Запуск програми
call npm start

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ПОМИЛКА] Програма завершилась з помилкою
    pause
    exit /b 1
)

pause