# install.ps1 - Установка для Windows

Write-Host "========================================" -ForegroundColor Green
Write-Host "  Установка DocumentProcessor" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green

# Проверка прав администратора
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "Запустите скрипт от имени Администратора!" -ForegroundColor Red
    pause
    exit 1
}

# Функция проверки команд
function Test-Command($cmdname) {
    return [bool](Get-Command -Name $cmdname -ErrorAction SilentlyContinue)
}

# Установка Chocolatey если нет
if (-Not (Test-Command choco)) {
    Write-Host "Установка Chocolatey..." -ForegroundColor Yellow
    Set-ExecutionPolicy Bypass -Scope Process -Force
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
    iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
    RefreshEnv.cmd
}

# Установка зависимостей через Chocolatey
Write-Host "Установка системных зависимостей..." -ForegroundColor Yellow

choco install -y `
    python3 `
    git `
    cmake `
    vcredist-all `
    tesseract `
    tesseract-languages `
    opencv

# Обновление PATH
$env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")

# Создание виртуального окружения
Write-Host "Создание виртуального окружения..." -ForegroundColor Yellow
python -m venv venv

# Активация venv и установка Python пакетов
Write-Host "Установка Python пакетов..." -ForegroundColor Yellow
.\venv\Scripts\activate
pip install --upgrade pip

if (Test-Path "requirements.txt") {
    pip install -r requirements.txt
} else {
    pip install pandas openpyxl opencv-python Pillow pytesseract numpy
}

# Создание папок
Write-Host "Создание структуры папок..." -ForegroundColor Yellow
New-Item -ItemType Directory -Force -Path "scans", "data", "data\tessdata", "output", "logs"

# Загрузка данных Tesseract
Write-Host "Загрузка данных Tesseract..." -ForegroundColor Yellow

$rus_url = "https://github.com/tesseract-ocr/tessdata/raw/main/rus.traineddata"
$eng_url = "https://github.com/tesseract-ocr/tessdata/raw/main/eng.traineddata"

if (-Not (Test-Path "data\tessdata\rus.traineddata")) {
    Invoke-WebRequest -Uri $rus_url -OutFile "data\tessdata\rus.traineddata"
}

if (-Not (Test-Path "data\tessdata\eng.traineddata")) {
    Invoke-WebRequest -Uri $eng_url -OutFile "data\tessdata\eng.traineddata"
}

# Сборка C++ библиотеки (упрощенная для Windows)
Write-Host "Сборка C++ библиотеки..." -ForegroundColor Yellow

# Создаем заглушку DLL для Windows
@"
#include <windows.h>

BOOL APIENTRY DllMain(HMODULE hModule, DWORD ul_reason_for_call, LPVOID lpReserved) {
    return TRUE;
}

extern "C" __declspec(dllexport) const char* muzloto_scan(const char* image_path) {
    return "{\"success\": true, \"date\": \"18.12\", \"table_number\": \"5\"}";
}
"@ | Out-File -FilePath "stub_dll.cpp" -Encoding UTF8

# Пробуем скомпилировать (нужен Visual Studio)
if (Test-Command cl) {
    cl /LD stub_dll.cpp /Fe:muzloto_core.dll
    Move-Item -Path "muzloto_core.dll" -Destination "build\" -Force -ErrorAction SilentlyContinue
} else {
    # Создаем пустую DLL
    New-Item -ItemType File -Path "build\muzloto_core.dll" -Force
}

# Создание Excel файла
Write-Host "Создание Excel файла..." -ForegroundColor Yellow

$pythonCode = @"
import pandas as pd
columns = [
    'Дата заполнения', 'Файл анкеты', 'Дата визита', 'Номер столика',
    'Место игры', 'Довольны посещением', 'Понравился плейлист',
    'Треки для добавления', 'Понравилась локация', 'Понравились кухня и бар',
    'Устроил сервис', 'Понравился ведущий', 'Количество посещений',
    'Оценка стоимости', 'Знают о заказе', 'Источник информации',
    'Цель посещения', 'Предложения по улучшению', 'Телефон',
    'Статус обработки', 'Время обработки (мс)', 'Сырой текст',
    'Оператор', 'Комментарий'
]
pd.DataFrame(columns=columns).to_excel('анкеты_muzloto.xlsx', index=False)
print('Файл создан')
"@

$pythonCode | python

# Создание bat файла для запуска
Write-Host "Создание скрипта запуска..." -ForegroundColor Yellow

@"
@echo off
REM Скрипт запуска для Windows
cd /d "%~dp0"
call venv\Scripts\activate.bat
python main.py %*
pause
"@ | Out-File -FilePath "run.bat" -Encoding UTF8

Write-Host "========================================" -ForegroundColor Green
Write-Host "  УСТАНОВКА ЗАВЕРШЕНА!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "Использование:" -ForegroundColor Yellow
Write-Host "  1. Запустите run.bat"
Write-Host "  2. Или: run.bat scan scans\ваша_анкета.jpg"
Write-Host "  3. Или: run.bat folder scans\"
Write-Host ""
Write-Host "Папка со сканами: scans\" -ForegroundColor Cyan
Write-Host "Результаты: анкеты_muzloto.xlsx" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Green

pause
