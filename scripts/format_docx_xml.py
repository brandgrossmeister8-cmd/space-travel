#!/usr/bin/env python3
"""
Форматирование DOCX через прямую работу с XML.
Не требует python-docx - работает с ZIP/XML напрямую.

Форматирование:
- Таблицы: границы, голубая шапка (#B8CCE4), авто-ширина
- Шрифт: Tahoma 9pt
- Заголовки: синие (#003399)
"""

import sys
import zipfile
import shutil
import re
from pathlib import Path
from xml.etree import ElementTree as ET

# Namespaces
NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}

# Регистрация namespaces
for prefix, uri in NS.items():
    ET.register_namespace(prefix, uri)

# Дополнительные namespaces из DOCX
EXTRA_NS = {
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
    'w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
    'wpc': 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
    'wpg': 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
    'wpi': 'http://schemas.microsoft.com/office/word/2010/wordprocessingInk',
    'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
}
for prefix, uri in EXTRA_NS.items():
    ET.register_namespace(prefix, uri)


def qn(tag):
    """Создать qualified name с namespace."""
    if ':' in tag:
        prefix, local = tag.split(':')
        return '{%s}%s' % (NS.get(prefix, EXTRA_NS.get(prefix, '')), local)
    return tag


def create_border_element(tag, color='000000', size='4', space='0', val='single'):
    """Создать элемент границы."""
    el = ET.Element(qn(f'w:{tag}'))
    el.set(qn('w:val'), val)
    el.set(qn('w:sz'), size)
    el.set(qn('w:space'), space)
    el.set(qn('w:color'), color)
    return el


def set_table_borders(tbl):
    """Установить границы таблицы."""
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = ET.SubElement(tbl, qn('w:tblPr'))
        tbl.insert(0, tblPr)

    # Удалить старые границы
    old_borders = tblPr.find(qn('w:tblBorders'))
    if old_borders is not None:
        tblPr.remove(old_borders)

    # Создать новые границы
    tblBorders = ET.SubElement(tblPr, qn('w:tblBorders'))
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        tblBorders.append(create_border_element(border_name))


def set_cell_shading(tc, color='B8CCE4'):
    """Установить цвет фона ячейки."""
    tcPr = tc.find(qn('w:tcPr'))
    if tcPr is None:
        tcPr = ET.SubElement(tc, qn('w:tcPr'))
        tc.insert(0, tcPr)

    # Удалить старый shading
    old_shd = tcPr.find(qn('w:shd'))
    if old_shd is not None:
        tcPr.remove(old_shd)

    # Создать новый
    shd = ET.SubElement(tcPr, qn('w:shd'))
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color)


def set_run_font(rPr, font_name='Tahoma', font_size='18', bold=False, color=None):
    """Установить шрифт для run properties."""
    # Шрифт
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = ET.SubElement(rPr, qn('w:rFonts'))
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)
    rFonts.set(qn('w:cs'), font_name)

    # Размер (в half-points, 18 = 9pt)
    sz = rPr.find(qn('w:sz'))
    if sz is None:
        sz = ET.SubElement(rPr, qn('w:sz'))
    sz.set(qn('w:val'), font_size)

    szCs = rPr.find(qn('w:szCs'))
    if szCs is None:
        szCs = ET.SubElement(rPr, qn('w:szCs'))
    szCs.set(qn('w:val'), font_size)

    # Bold
    b = rPr.find(qn('w:b'))
    if bold:
        if b is None:
            ET.SubElement(rPr, qn('w:b'))
    else:
        if b is not None:
            rPr.remove(b)

    # Color
    if color:
        c = rPr.find(qn('w:color'))
        if c is None:
            c = ET.SubElement(rPr, qn('w:color'))
        c.set(qn('w:val'), color)


def format_table(tbl):
    """Форматировать таблицу."""
    # Установить границы
    set_table_borders(tbl)

    # Установить авто-ширину
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is not None:
        tblW = tblPr.find(qn('w:tblW'))
        if tblW is None:
            tblW = ET.SubElement(tblPr, qn('w:tblW'))
        tblW.set(qn('w:type'), 'auto')
        tblW.set(qn('w:w'), '0')

    # Форматировать строки
    rows = tbl.findall(qn('w:tr'))
    for row_idx, tr in enumerate(rows):
        cells = tr.findall(qn('w:tc'))
        for tc in cells:
            # Голубая шапка для первой строки
            if row_idx == 0:
                set_cell_shading(tc, 'B8CCE4')

            # Форматировать текст в ячейке
            for p in tc.findall(qn('w:p')):
                for r in p.findall(qn('w:r')):
                    rPr = r.find(qn('w:rPr'))
                    if rPr is None:
                        rPr = ET.SubElement(r, qn('w:rPr'))
                        r.insert(0, rPr)

                    # Шрифт и bold для шапки
                    set_run_font(rPr, 'Tahoma', '18', bold=(row_idx == 0))


def is_heading_paragraph(p):
    """Проверить, является ли параграф заголовком."""
    pPr = p.find(qn('w:pPr'))
    if pPr is not None:
        pStyle = pPr.find(qn('w:pStyle'))
        if pStyle is not None:
            style_val = pStyle.get(qn('w:val'), '').lower()
            if 'heading' in style_val or 'заголовок' in style_val:
                return True

    # Проверить текст на #
    for r in p.findall(qn('w:r')):
        for t in r.findall(qn('w:t')):
            if t.text and t.text.strip().startswith('#'):
                return True

    return False


def format_paragraph(p):
    """Форматировать параграф."""
    is_heading = is_heading_paragraph(p)

    for r in p.findall(qn('w:r')):
        rPr = r.find(qn('w:rPr'))
        if rPr is None:
            rPr = ET.SubElement(r, qn('w:rPr'))
            r.insert(0, rPr)

        if is_heading:
            set_run_font(rPr, 'Tahoma', '32', bold=True, color='003399')  # 16pt = 32
        else:
            set_run_font(rPr, 'Tahoma', '18')  # 9pt = 18


def process_document(input_path, output_path=None):
    """Обработать документ."""
    input_path = Path(input_path)
    if output_path is None:
        output_path = input_path
    else:
        output_path = Path(output_path)

    # Создать временную директорию
    temp_dir = Path('/tmp/docx_format_temp')
    if temp_dir.exists():
        shutil.rmtree(temp_dir)
    temp_dir.mkdir()

    try:
        # Распаковать DOCX
        print(f"Распаковка: {input_path}")
        with zipfile.ZipFile(input_path, 'r') as zf:
            zf.extractall(temp_dir)

        # Обработать document.xml
        doc_path = temp_dir / 'word' / 'document.xml'
        if not doc_path.exists():
            print("Ошибка: document.xml не найден")
            return

        print("Обработка document.xml...")

        # Парсить с сохранением всех атрибутов
        tree = ET.parse(doc_path)
        root = tree.getroot()

        # Найти body
        body = root.find(qn('w:body'))
        if body is None:
            print("Ошибка: body не найден")
            return

        # Форматировать таблицы
        tables = body.findall('.//' + qn('w:tbl'))
        print(f"Найдено таблиц: {len(tables)}")
        for tbl in tables:
            format_table(tbl)

        # Форматировать параграфы (вне таблиц)
        paragraphs = body.findall(qn('w:p'))
        print(f"Найдено параграфов: {len(paragraphs)}")
        for p in paragraphs:
            format_paragraph(p)

        # Сохранить document.xml
        tree.write(doc_path, encoding='UTF-8', xml_declaration=True)

        # Запаковать обратно
        print(f"Сохранение: {output_path}")
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for file_path in temp_dir.rglob('*'):
                if file_path.is_file():
                    arc_name = file_path.relative_to(temp_dir)
                    zf.write(file_path, arc_name)

        print("Готово!")

    finally:
        # Очистить временные файлы
        if temp_dir.exists():
            shutil.rmtree(temp_dir)


def main():
    if len(sys.argv) < 2:
        print("Использование: python format_docx_xml.py <файл.docx> [выходной.docx]")
        sys.exit(1)

    input_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else None

    if not Path(input_path).exists():
        print(f"Ошибка: файл не найден: {input_path}")
        sys.exit(1)

    process_document(input_path, output_path)


if __name__ == "__main__":
    main()
