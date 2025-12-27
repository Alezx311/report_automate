#!/bin/bash

echo "========================================"
echo " Report Automate - Швидкий старт"
echo "========================================"
echo ""

# Перевірка Node.js
if ! command -v node &> /dev/null; then
    echo "[ПОМИЛКА] Node.js не встановлено!"
    echo ""
    echo "Завантажте з https://nodejs.org/"
    exit 1
fi

echo "[OK] Node.js встановлено"
node --version
echo ""

# Перевірка npm
if ! command -v npm &> /dev/null; then
    echo "[ПОМИЛКА] npm не встановлено!"
    exit 1
fi

echo "[OK] npm встановлено"
npm --version
echo ""

# Перевірка node_modules
if [ ! -d "node_modules" ]; then
    echo "[INFO] Встановлення залежностей..."
    npm install
    if [ $? -ne 0 ]; then
        echo "[ПОМИЛКА] Не вдалось встановити залежності"
        exit 1
    fi
    echo ""
fi

# Перевірка .env
if [ ! -f ".env" ]; then
    echo "[УВАГА] Файл .env не знайдено!"
    echo ""
    if [ -f ".env.example" ]; then
        echo "Створюю .env з .env.example..."
        cp .env.example .env
        echo "[OK] .env створено"
        echo ""
        echo "ВАЖЛИВО: Відредагуйте файл .env з вашими параметрами!"
        echo "Відкрийте .env у текстовому редакторі"
        
        # Спроба відкрити редактор
        if command -v nano &> /dev/null; then
            read -p "Відкрити в nano? (y/n) " -n 1 -r
            echo ""
            if [[ $REPLY =~ ^[Yy]$ ]]; then
                nano .env
            fi
        fi
    else
        echo "[ПОМИЛКА] .env.example не знайдено!"
        exit 1
    fi
fi

# Створення папки output
if [ ! -d "output" ]; then
    mkdir output
    echo "[OK] Створено папку output/"
fi

echo ""
echo "========================================"
echo " Запуск програми..."
echo "========================================"
echo ""

# Запуск програми
npm start

if [ $? -ne 0 ]; then
    echo ""
    echo "[ПОМИЛКА] Програма завершилась з помилкою"
    exit 1
fi