#!/usr/bin/env python3
"""
build.py - автоматическая сборка C++ библиотеки
"""

import os
import sys
import subprocess
import platform
import shutil
from pathlib import Path

def print_colored(text, color):
    """Вывод цветного текста."""
    colors = {
        'red': '\033[91m',
        'green': '\033[92m',
        'yellow': '\033[93m',
        'blue': '\033[94m',
        'end': '\033[0m'
    }
    print(f"{colors.get(color, '')}{text}{colors['end']}")

def check_dependencies():
    """Проверка необходимых зависимостей."""
    print_colored("Проверка зависимостей...", "blue")
    
    deps = {
        'cmake': 'CMake',
        'g++': 'GCC',
        'cl': 'MSVC (Windows)'
    }
    
    missing = []
    for cmd, name in deps.items():
        try:
            subprocess.run([cmd, '--version'], capture_output=True, check=False)
            print_colored(f"  ✓ {name}", "green")
        except:
            print_colored(f"  ✗ {name} не найден", "yellow")
            missing.append(name)
    
    return len(missing) == 0

def create_simple_library():
    """Создание простой C++ библиотеки-заглушки."""
    print_colored("Создание простой библиотеки...", "blue")
    
    source_code = '''
#include <cstring>

#ifdef _WIN32
    #define EXPORT __declspec(dllexport)
#else
    #define EXPORT __attribute__((visibility("default")))
#endif

extern "C" {
    EXPORT const char* muzloto_scan_image(const char* image_path) {
        static const char* result = 
            "{"
            "\\"success\\": true,"
            "\\"date\\": \\"18.12\\","
            "\\"table_number\\": \\"5\\","
            "\\"location\\": \\"Борщина куца\\","
            "\\"satisfaction\\": \\"Да\\","
            "\\"playlist_liked\\": \\"Да\\","
            "\\"tracks_to_add\\": \\"Рок, Поп\\","
            "\\"location_liked\\": \\"Да\\","
            "\\"kitchen_liked\\": \\"Да\\","
            "\\"service_ok\\": \\"Да\\","
            "\\"host_work\\": \\"Да\\","
            "\\"visits_count\\": \\"3\\","
            "\\"ticket_price\\": \\"достойно и дорого\\","
            "\\"know_booking\\": \\"Да\\","
            "\\"source_info\\": \\"Друзья\\","
            "\\"purpose\\": \\"Развлечение\\","
            "\\"improvements\\": \\"Больше музыки\\","
            "\\"phone_number\\": \\"+79991234567\\""
            "}";
        return result;
    }
    
    EXPORT void free_string(const char* str) {
        // Заглушка
    }
}
'''
    
    # Определяем имя файла в зависимости от ОС
    system = platform.system()
    if system == "Windows":
        lib_name = "muzloto_core.dll"
        source_file = "stub_dll.cpp"
        compile_cmd = ["cl", "/LD", source_file, f"/Fe:{lib_name}"]
    elif system == "Darwin":
        lib_name = "libmuzloto_core.dylib"
        source_file = "stub_lib.cpp"
        compile_cmd = ["g++", "-shared", "-fPIC", "-o", lib_name, source_file]
    else:  # Linux
        lib_name = "libmuzloto_core.so"
        source_file = "stub_lib.cpp"
        compile_cmd = ["g++", "-shared", "-fPIC", "-o", lib_name, source_file]
    
    # Создаем исходный файл
    with open(source_file, "w") as f:
        f.write(source_code)
    
    # Пробуем скомпилировать
    try:
        print_colored(f"Компиляция {lib_name}...", "blue")
        result = subprocess.run(compile_cmd, capture_output=True, text=True)
        
        if result.returncode == 0:
            # Перемещаем в build/
            build_dir = Path("build")
            build_dir.mkdir(exist_ok=True)
            shutil.move(lib_name, build_dir / lib_name)
            print_colored(f"✓ Библиотека создана: build/{lib_name}", "green")
            
            # Удаляем временные файлы
            os.remove(source_file)
            if system == "Windows":
                for ext in [".obj", ".exp", ".lib"]:
                    temp_file = lib_name.replace(".dll", ext)
                    if os.path.exists(temp_file):
                        os.remove(temp_file)
            
            return True
        else:
            print_colored(f"✗ Ошибка компиляции: {result.stderr}", "red")
            return False
            
    except Exception as e:
        print_colored(f"✗ Ошибка: {e}", "red")
        return False

def build_with_cmake():
    """Сборка с помощью CMake."""
    print_colored("Сборка с CMake...", "blue")
    
    build_dir = Path("build")
    build_dir.mkdir(exist_ok=True)
    
    try:
        # Конфигурация
        print_colored("  Конфигурация CMake...", "blue")
        config_cmd = ["cmake", "..", "-DCMAKE_BUILD_TYPE=Release"]
        result = subprocess.run(config_cmd, cwd=build_dir, capture_output=True, text=True)
        
        if result.returncode != 0:
            print_colored(f"  ✗ Ошибка конфигурации: {result.stderr}", "yellow")
            return False
        
        # Сборка
        print_colored("  Компиляция...", "blue")
        
        # Определяем количество ядер
        import multiprocessing
        cores = multiprocessing.cpu_count()
        
        build_cmd = ["cmake", "--build", ".", "--config", "Release", "-j", str(cores)]
        result = subprocess.run(build_cmd, cwd=build_dir, capture_output=True, text=True)
        
        if result.returncode == 0:
            print_colored("  ✓ Сборка успешна", "green")
            
            # Проверяем что создалась библиотека
            lib_files = list(build_dir.glob("*muzloto_core*"))
            if lib_files:
                print_colored(f"  ✓ Библиотека: {lib_files[0].name}", "green")
                return True
            else:
                print_colored("  ✗ Библиотека не найдена после сборки", "yellow")
                return False
        else:
            print_colored(f"  ✗ Ошибка сборки: {result.stderr}", "yellow")
            return False
            
    except Exception as e:
        print_colored(f"  ✗ Исключение: {e}", "yellow")
        return False

def main():
    """Главная функция сборки."""
    print_colored("=" * 50, "blue")
    print_colored("  АВТОМАТИЧЕСКАЯ СБОРКА C++ БИБЛИОТЕКИ", "blue")
    print_colored("=" * 50, "blue")
    
    # Проверяем зависимости
    if not check_dependencies():
        print_colored("\nНекоторые зависимости отсутствуют.", "yellow")
        choice = input("Продолжить с созданием простой библиотеки? (y/n): ")
        if choice.lower() != 'y':
            return 1
    
    # Пробуем собрать через CMake
    print_colored("\nПопытка сборки через CMake...", "blue")
    if build_with_cmake():
        print_colored("\n✅ Сборка завершена успешно!", "green")
        return 0
    
    # Если CMake не сработал, создаем простую библиотеку
    print_colored("\nСоздание простой библиотеки...", "blue")
    if create_simple_library():
        print_colored("\n✅ Простая библиотека создана!", "green")
        print_colored("   Приложение будет работать в тестовом режиме.", "yellow")
        return 0
    else:
        print_colored("\n❌ Не удалось создать библиотеку", "red")
        return 1

if __name__ == "__main__":
    sys.exit(main())
