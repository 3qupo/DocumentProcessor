#!/usr/bin/env python3
"""
Muzloto Анкета Сканер
Автоматическое распознавание анкет Muzloto и сохранение в один Excel файл.

Использование:
    python main.py scan <путь_к_изображению> [оператор]
    python main.py folder <путь_к_папке> [оператор]
    python main.py stats
"""

import sys
import argparse
from pathlib import Path
from datetime import datetime
from python.scanner import MuzlotoScanner

def main():
    parser = argparse.ArgumentParser(
        description="Сканер анкет Muzloto - распознавание и сохранение в Excel",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Примеры:
  %(prog)s scan анкета.jpg "Иван Иванов"
  %(prog)s folder ./сканы_анкет/ "Оператор 1"
  %(prog)s stats
  
Файл результатов: анкеты_muzloto.xlsx
        """
    )
    
    subparsers = parser.add_subparsers(dest='command', help='Команда')
    
    # Команда scan
    scan_parser = subparsers.add_parser('scan', help='Сканировать одну анкету')
    scan_parser.add_argument('image_path', help='Путь к изображению анкеты')
    scan_parser.add_argument('operator', nargs='?', default='Авто', 
                           help='Имя оператора (по умолчанию: Авто)')
    
    # Команда folder
    folder_parser = subparsers.add_parser('folder', help='Обработать папку с анкетами')
    folder_parser.add_argument('folder_path', help='Путь к папке с анкетами')
    folder_parser.add_argument('operator', nargs='?', default='Пакетная обработка',
                             help='Имя оператора (по умолчанию: Пакетная обработка)')
    
    # Команда stats
    stats_parser = subparsers.add_parser('stats', help='Показать статистику')
    
    args = parser.parse_args()
    
    # Создаем сканер
    try:
        scanner = MuzlotoScanner(
            excel_file="анкеты_muzloto.xlsx",
            tessdata_path="./data/tessdata"  # Путь к данным Tesseract
        )
    except Exception as e:
        print(f"❌ Ошибка инициализации сканера: {e}")
        print("Убедитесь, что:")
        print("  1. C++ библиотека скомпилирована")
        print("  2. Установлен Tesseract OCR")
        print("  3. Данные Tesseract (rus.traineddata) в папке data/tessdata/")
        return 1
    
    if args.command == 'scan':
        # Обработка одной анкеты
        result = scanner.process_anketa(
            image_path=args.image_path,
            operator=args.operator
        )
        
        if result["success"]:
            print(f"\n✅ Анкета успешно обработана!")
            print(f"   Сохранено в строку: {result['row_number']}")
            print(f"   Файл: {result['excel_file']}")
        else:
            print(f"\n❌ Ошибка: {result['message']}")
            
    elif args.command == 'folder':
        # Обработка папки
        result = scanner.process_folder(
            folder_path=args.folder_path,
            operator=args.operator
        )
        
    elif args.command == 'stats':
        # Статистика
        stats = scanner.get_statistics()
        
        print("\n📊 СТАТИСТИКА ОБРАБОТКИ АНКЕТ")
        print("=" * 50)
        print(f"Файл с анкетами: {stats.get('excel_file', '—')}")
        print(f"Всего записей в Excel: {stats.get('total_records', 0)}")
        print(f"Успешно обработанных: {stats.get('successful_records', 0)}")
        print(f"Уникальных дат: {stats.get('unique_dates', 0)}")
        
        proc_stats = stats.get('processing_stats', {})
        print(f"\nТекущая сессия:")
        print(f"  Всего обработано: {proc_stats.get('total', 0)}")
        print(f"  Успешно: {proc_stats.get('success', 0)}")
        print(f"  С ошибками: {proc_stats.get('failed', 0)}")
        
        if proc_stats.get('last_file'):
            print(f"  Последний файл: {proc_stats.get('last_file')}")
        
        print("=" * 50)
        
    else:
        parser.print_help()
    
    return 0

if __name__ == "__main__":
    sys.exit(main())