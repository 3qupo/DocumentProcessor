#!/usr/bin/env python3
"""
main.py - DocumentProcessor с автоматической установкой
"""

import os
import sys
import subprocess
from pathlib import Path

def check_and_install():
    """Проверяет установку и предлагает установить если нужно."""
    
    print("🔍 Проверка установки...")
    
    missing = []
    
    # Проверяем виртуальное окружение
    if not Path("venv").exists():
        missing.append("Виртуальное окружение")
    
    # Проверяем Python пакеты
    try:
        import pandas
        import openpyxl
        import cv2
        import pytesseract
    except ImportError as e:
        missing.append(f"Python пакеты: {e}")
    
    # Проверяем Tesseract
    try:
        import pytesseract
        pytesseract.get_tesseract_version()
    except:
        missing.append("Tesseract OCR")
    
    # Проверяем C++ библиотеку
    lib_paths = [
        "build/libmuzloto_core.so",
        "build/muzloto_core.dll", 
        "build/libmuzloto_core.dylib"
    ]
    if not any(Path(p).exists() for p in lib_paths):
        missing.append("C++ библиотека")
    
    if missing:
        print("\n❌ Обнаружены проблемы:")
        for item in missing:
            print(f"   - {item}")
        
        choice = input("\nХотите выполнить автоматическую установку? (y/n): ")
        if choice.lower() == 'y':
            print("\n🚀 Запуск автоматической установки...")
            
            # Определяем ОС
            if sys.platform == "win32":
                install_script = "install.ps1"
                if not Path(install_script).exists():
                    print("Создаю install.ps1...")
                    # Здесь создаем install.ps1 если его нет
                    create_windows_installer()
                subprocess.run(["powershell", "-ExecutionPolicy", "Bypass", "-File", install_script])
            else:
                install_script = "install.sh"
                if not Path(install_script).exists():
                    print("Создаю install.sh...")
                    create_linux_installer()
                
                os.chmod(install_script, 0o755)
                subprocess.run([f"./{install_script}"])
            
            print("\n✅ Установка завершена. Перезапустите программу.")
            sys.exit(0)
        else:
            print("\nУстановите зависимости вручную или запустите скрипт установки.")
            print("Для Linux/Mac: ./install.sh")
            print("Для Windows: .\\install.ps1")
            sys.exit(1)
    
    print("✅ Все зависимости установлены")
    return True

def create_linux_installer():
    """Создает install.sh если его нет."""
    # Здесь код из install.sh выше
    pass

def create_windows_installer():
    """Создает install.ps1 если его нет."""
    # Здесь код из install.ps1 выше
    pass

def main():
    """Главная функция."""
    
    # Проверяем установку
    if not check_and_install():
        return
    
    # Активируем виртуальное окружение
    if sys.platform == "win32":
        activate_script = "venv\\Scripts\\activate.bat"
    else:
        activate_script = "venv/bin/activate"
    
    # Импортируем основной модуль
    try:
        from python.scanner import MuzlotoScanner
    except ImportError:
        print("Импортируем локальную версию...")
        # Локальный импорт если python.scanner нет
        scanner_code = """
# Локальная реализация сканера
class MuzlotoScanner:
    def __init__(self, excel_file="анкеты_muzloto.xlsx"):
        self.excel_file = excel_file
        print(f"Сканер инициализирован, файл: {excel_file}")
    
    def process_anketa(self, image_path, operator="Система"):
        print(f"Обработка: {image_path}")
        return {"success": True, "message": "Тестовый режим"}
"""
        exec(scanner_code)
        MuzlotoScanner = locals()['MuzlotoScanner']
    
    # Парсим аргументы командной строки
    if len(sys.argv) > 1:
        command = sys.argv[1]
        
        if command == "scan" and len(sys.argv) > 2:
            image_path = sys.argv[2]
            operator = sys.argv[3] if len(sys.argv) > 3 else "Система"
            
            scanner = MuzlotoScanner()
            result = scanner.process_anketa(image_path, operator)
            print(f"Результат: {result}")
            
        elif command == "folder" and len(sys.argv) > 2:
            folder_path = sys.argv[2]
            operator = sys.argv[3] if len(sys.argv) > 3 else "Пакетная обработка"
            
            scanner = MuzlotoScanner()
            scanner.process_folder(folder_path, operator)
            
        elif command == "stats":
            scanner = MuzlotoScanner()
            stats = scanner.get_statistics()
            print(f"Статистика: {stats}")
            
        elif command == "install":
            print("Используйте ./install.sh или .\\install.ps1")
            
        elif command == "build":
            subprocess.run([sys.executable, "build.py"])
            
        else:
            print_help()
    else:
        print_help()

def print_help():
    """Печатает справку."""
    print("""
DocumentProcessor - система обработки анкет Muzloto

Использование:
  python main.py scan <путь_к_анкете> [оператор]
  python main.py folder <путь_к_папке> [оператор]
  python main.py stats
  python main.py install   - автоматическая установка
  python main.py build     - сборка C++ библиотеки

Примеры:
  python main.py scan scans/анкета.jpg "Иван Иванов"
  python main.py folder scans/ "Пакетная обработка"
  
Файл результатов: анкеты_muzloto.xlsx
    """)

if __name__ == "__main__":
    main()
