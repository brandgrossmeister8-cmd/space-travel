#!/bin/bash
# Скрипт экспорта отчёта: MD → DOCX с форматированием
#
# Использование:
#   ./export_report.sh [имя_файла.md]
#
# Если файл не указан, используется ФИНАЛЬНЫЙ_ОТЧЁТ_ЗАЩИТА.md

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"

# Определить входной файл
if [ -n "$1" ]; then
    INPUT_FILE="$1"
else
    INPUT_FILE="$PROJECT_DIR/ФИНАЛЬНЫЙ_ОТЧЁТ_ЗАЩИТА.md"
fi

# Проверить существование файла
if [ ! -f "$INPUT_FILE" ]; then
    echo "Ошибка: файл не найден: $INPUT_FILE"
    exit 1
fi

# Определить выходной файл
OUTPUT_FILE="${INPUT_FILE%.md}.docx"

echo "=========================================="
echo "Экспорт отчёта Space Travel"
echo "=========================================="
echo ""
echo "Исходный файл: $INPUT_FILE"
echo "Выходной файл: $OUTPUT_FILE"
echo ""

# Шаг 1: Конвертация MD → DOCX
echo "1. Конвертация MD → DOCX..."
pandoc "$INPUT_FILE" -o "$OUTPUT_FILE"
if [ $? -ne 0 ]; then
    echo "Ошибка при конвертации!"
    exit 1
fi
echo "   Готово"

# Шаг 2: Форматирование DOCX
echo "2. Форматирование документа..."
python3 "$SCRIPT_DIR/format_docx.py" "$OUTPUT_FILE"
if [ $? -ne 0 ]; then
    echo "Ошибка при форматировании!"
    exit 1
fi

# Шаг 3: Открыть документ
echo ""
echo "3. Открываю документ..."
open "$OUTPUT_FILE"

echo ""
echo "=========================================="
echo "Готово!"
echo "=========================================="
