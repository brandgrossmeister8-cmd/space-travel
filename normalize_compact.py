#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Компактная нормализация данных анкет L1, L2, L3.
Использует короткие коды для департаментов, критичности, приоритетов и т.д.
"""

import json
import re
from pathlib import Path
from typing import Dict, List, Any, Optional

BASE_DIR = Path(__file__).parent

# ==================== СПРАВОЧНИКИ ====================

DEPARTMENTS = {
    "HR": "Департамент HR",
    "MKT": "Департамент маркетинга и рекламы",
    "FIN": "Департамент финансов",
    "IT": "Департамент информационных технологий",
    "OPS": "Операционный департамент",
    "REV": "Департамент по управлению операционными доходами",
    "AGN": "Департамент развития взаимодействия с агентствами",
    "INB": "Департамент въездного туризма, MICE и корпоративного обслуживания",
    "SRV": "Управление клиентского сервиса",
    "UNK": "Не указан"
}

CRITICALITY = {
    "H": "Высокая",
    "M": "Средняя",
    "L": "Низкая"
}

AUTOMATION_PRIORITY = {
    "1": "Да, в первую очередь (критично)",
    "2": "Да, хотелось бы (важно)",
    "3": "Можно, но некритично (желательно)",
    "0": "Не нужно"
}

REGULATION_STATUS = {
    "exists": "Есть утверждённый регламент",
    "draft": "Есть наброски/блок-схема",
    "none": "Регламента нет"
}

TASK_TYPE = {
    "R": "Рутинная задача",
    "P": "Проблемная задача"
}

ROUTINE_LEVEL = {
    "VH": "Очень высокий",
    "H": "Высокий",
    "M": "Средний",
    "L": "Низкий"
}

FREQUENCY = {
    "D": "Ежедневно",
    "W": "Еженедельно",
    "M": "Ежемесячно",
    "Q": "Ежеквартально",
    "R": "По запросу"
}

# ==================== ЭТАЛОННЫЙ МАППИНГ ОТДЕЛ → ДЕПАРТАМЕНТ ====================
# Источник: «Структура департаментов и отделов_новые и старые названия_новый.xlsx»
# Приоритет: если название отдела (division) совпадает — используем этот маппинг,
# а НЕ то, что сотрудник написал в поле «департамент» (часто ошибочно).

DIVISION_TO_DEPARTMENT = {
    # --- AGN: Департамент развития взаимодействия с агентствами ---
    "Управление по развитию региональных офисов": "AGN",
    "Офис г. Екатеринбург": "AGN",
    "Офис г. Казань": "AGN",
    "Офис г. Краснодар": "AGN",
    "Офис г. Новосибирск": "AGN",
    "Офис г. Пятигорск": "AGN",
    "Офис г. Самара": "AGN",
    "Офис г. Санкт-Петербург": "AGN",
    "Офис г. Хабаровск": "AGN",
    "Отдел по работе с ключевыми клиентами": "AGN",
    "Группа по привлечению агентств": "AGN",
    # --- REV: Департамент по управлению операционными доходами ---
    "Отдел продукта_группа стран Ближний Восток": "REV",
    "Отдел продукта_группа стран Экзотика": "REV",
    "Отдел продукта_группа стран FIT": "REV",
    "Отдел продукта_группа стран_FIT": "REV",
    "Отдел продукта_группа стран ЮВА": "REV",
    "Группа тарификации": "REV",
    "Чартерная группа": "REV",
    "Группа стран FIT": "REV",
    "Группа стран «Ближний Восток»": "REV",
    "Группа стран «Экзотика»": "REV",
    "Группа стран «ЮВА»": "REV",
    # --- INB: Департамент въездного туризма, MICE и корп. обслуживания ---
    "Отдел MICE": "INB",
    "Отдел группового бронирования": "INB",
    "Отдел бизнес-тревел": "INB",
    "Отдел въездного туризма": "INB",
    # --- OPS: Операционный департамент ---
    "ВИП-отдел": "OPS",
    "ВИП отдел": "OPS",
    "Группа визового сопровождения": "OPS",
    "Транспортный отдел": "OPS",
    "Отдел бронирования": "OPS",
    "Отдел продаж": "OPS",
    # --- SRV: Управление клиентского сервиса ---
    "Отдел клиентской поддержки": "SRV",
    "Отдел круглосуточной работы с клиентами": "SRV",
    # --- MKT: Департамент маркетинга и рекламы ---
    "Отдел рекламы": "MKT",
    "Отдел контент-маркетинга": "MKT",
    "Отдел маркетинга": "MKT",
    "Маркетинг": "MKT",
    # --- FIN: Департамент финансов ---
    "Бухгалтерия группы компаний_Бухгалтерия": "FIN",
    "Бухгалтерия группы компаний": "FIN",
    "Финансовый отдел": "FIN",
    "Финансовый отдел_Банковская группа": "FIN",
    "Финансовый отдел_Группа по расчетам с иностранными поставщиками": "FIN",
    "Банковская группа": "FIN",
    "Группа по расчетам с иностранными поставщиками": "FIN",
    "Бухгалтерия": "FIN",
}

# Нормализация названий отделов (варианты написания → каноническое название)
DIVISION_NAME_NORMALIZE = {
    "ВИП отдел": "ВИП-отдел",
    "Вип-отдел": "ВИП-отдел",
    "Вип отдел": "ВИП-отдел",
    "отдел рекламы": "Отдел рекламы",
    "отдел бронирования": "Отдел бронирования",
    "офис г. Санкт-Петербург": "Офис г. Санкт-Петербург",
    "офис г. Екатеринбург": "Офис г. Екатеринбург",
    "Офис в г.Санкт-Петербурге": "Офис г. Санкт-Петербург",
    "Офис г.Санкт-Петербурге": "Офис г. Санкт-Петербург",
    "Офис г.Казань": "Офис г. Казань",
    "Офис г.Краснодар": "Офис г. Краснодар",
    "Представительсто в Новосибирске": "Офис г. Новосибирск",
    "представительство в г. Пятигорск": "Офис г. Пятигорск",
    "Отдел круглосуочной работы с клиентами": "Отдел круглосуточной работы с клиентами",
    "Группа стран «Ближний Восток»": "Отдел продукта_группа стран Ближний Восток",
    "Группа стран «Экзотика»": "Отдел продукта_группа стран Экзотика",
    "Группа стран «ЮВА»": "Отдел продукта_группа стран ЮВА",
    "Группа стран FIT": "Отдел продукта_группа стран FIT",
    "не имеется": "[Нет отдела]",
    "нет отделов": "[Нет отдела]",
}


# ==================== ФУНКЦИИ МАППИНГА ====================

def normalize_division_name(div: str) -> str:
    """Нормализует название отдела по справочнику."""
    if not div:
        return ""
    div = div.strip()
    # Точное совпадение
    if div in DIVISION_NAME_NORMALIZE:
        return DIVISION_NAME_NORMALIZE[div]
    # Совпадение без учёта регистра
    for key, val in DIVISION_NAME_NORMALIZE.items():
        if div.lower() == key.lower():
            return val
    return div


def detect_department_by_division(div: str) -> Optional[str]:
    """Определяет код департамента по названию отдела (эталонный маппинг).
    Возвращает код или None, если отдел не найден в справочнике."""
    if not div:
        return None
    div = div.strip()
    # Точное совпадение
    if div in DIVISION_TO_DEPARTMENT:
        return DIVISION_TO_DEPARTMENT[div]
    # Без учёта регистра
    div_lower = div.lower()
    for key, code in DIVISION_TO_DEPARTMENT.items():
        if div_lower == key.lower():
            return code
    # Нечёткое: если нормализованное имя есть в справочнике
    norm = normalize_division_name(div)
    if norm in DIVISION_TO_DEPARTMENT:
        return DIVISION_TO_DEPARTMENT[norm]
    return None


def detect_department_code(text: str, division: str = "") -> str:
    """Определяет код департамента.
    Приоритет: 1) по названию отдела (эталон), 2) по тексту департамента (regex)."""
    # Сначала пробуем по эталонному маппингу отдела
    if division:
        code = detect_department_by_division(division)
        if code:
            return code

    if not text:
        return "UNK"

    text_lower = text.lower()

    # Порядок важен - более специфичные паттерны первыми
    patterns = [
        # REV - Департамент по управлению операционными доходами (продукт, тарификация, группы стран)
        (r'операционн\w*\s+доход|revenue|тарификац|продукт|группа\s+стран|ближн\w+\s+восток|экзотик|юва|fit', 'REV'),
        # HR
        (r'hr|эйчар|human\s*resource', 'HR'),
        # MKT - маркетинг и реклама
        (r'маркетинг|реклам|контент', 'MKT'),
        # FIN - финансы
        (r'финанс|бухгалтер', 'FIN'),
        # IT
        (r'\bit\b|информац|технолог|айти', 'IT'),
        # INB - въездной туризм, MICE
        (r'въезд|mice|корпоратив|бизнес.?тревел', 'INB'),
        # AGN - работа с агентствами, регионы
        (r'агентств|развит\w+\s+взаимо|региональн|ключев\w+\s+клиент', 'AGN'),
        # SRV - клиентский сервис
        (r'сервис|клиентск|поддержк', 'SRV'),
        # OPS - операционный департамент (продажи, бронирование, VIP)
        (r'операционн|продаж|брониров|вип|vip|чартер|транспорт|виз', 'OPS'),
    ]

    for pattern, code in patterns:
        if re.search(pattern, text_lower):
            return code

    return "UNK"


def map_criticality(text: str) -> str:
    """Маппит критичность в код."""
    if not text:
        return "M"
    text_lower = text.lower()
    if 'высок' in text_lower or 'high' in text_lower:
        return "H"
    elif 'низк' in text_lower or 'low' in text_lower:
        return "L"
    return "M"


def map_automation_priority(text: str) -> str:
    """Маппит приоритет автоматизации в код."""
    if not text:
        return "0"
    text_lower = text.lower()

    if 'уже есть' in text_lower or 'уже автоматизир' in text_lower:
        return "0"
    if 'нет' in text_lower and 'да' not in text_lower:
        return "0"
    if 'первую очередь' in text_lower or 'приоритет 1' in text_lower or 'критично' in text_lower:
        return "1"
    if 'хотелось бы' in text_lower or 'приоритет 2' in text_lower or 'важно' in text_lower:
        return "2"
    if 'можно' in text_lower or 'приоритет 3' in text_lower or 'некритично' in text_lower:
        return "3"
    if 'да' in text_lower:
        return "2"
    return "0"


def map_regulation_status(text: str) -> str:
    """Маппит статус регламента в код."""
    if not text:
        return "none"
    text_lower = text.lower()

    if 'нет' in text_lower or 'отсутств' in text_lower:
        return "none"
    if 'наброс' in text_lower or 'блок-схем' in text_lower or 'схем' in text_lower:
        return "draft"
    if 'есть' in text_lower or 'утвержд' in text_lower:
        return "exists"
    return "none"


def map_task_type(text: str) -> str:
    """Маппит тип задачи в код."""
    if not text:
        return "R"
    text_lower = text.lower()
    if 'проблем' in text_lower:
        return "P"
    return "R"


def map_routine_level(text: str) -> str:
    """Маппит уровень рутинности в код."""
    if not text:
        return "M"
    text_lower = text.lower()

    if 'очень высок' in text_lower:
        return "VH"
    if 'высок' in text_lower:
        return "H"
    if 'низк' in text_lower:
        return "L"
    return "M"


def map_frequency(text: str) -> str:
    """Маппит частоту в код."""
    if not text:
        return "R"
    text_lower = text.lower()

    if 'ежедневн' in text_lower or 'каждый день' in text_lower:
        return "D"
    if 'еженедельн' in text_lower or 'каждую неделю' in text_lower:
        return "W"
    if 'ежемесячн' in text_lower or 'каждый месяц' in text_lower:
        return "M"
    if 'ежеквартальн' in text_lower or 'квартал' in text_lower:
        return "Q"
    return "R"


def shorten_fio(fio: str) -> str:
    """Сокращает ФИО до Фамилия И.О."""
    if not fio:
        return ""
    parts = fio.strip().split()
    if len(parts) == 1:
        return parts[0]
    if len(parts) == 2:
        return f"{parts[0]} {parts[1][0]}."
    if len(parts) >= 3:
        return f"{parts[0]} {parts[1][0]}.{parts[2][0]}."
    return fio


def normalize_text(text: str) -> str:
    """Очищает текст."""
    if not text:
        return ""
    text = str(text).strip()
    text = re.sub(r'\s+', ' ', text)
    return text


def load_json_files(folder: str, exclude_patterns: List[str] = None) -> List[Dict]:
    """Загружает JSON файлы из папки."""
    folder_path = BASE_DIR / folder
    exclude_patterns = exclude_patterns or []
    files = []

    for f in sorted(folder_path.glob("*.json")):
        skip = any(p in f.name for p in exclude_patterns)
        if skip:
            continue
        try:
            with open(f, 'r', encoding='utf-8') as file:
                data = json.load(file)
                data['_filename'] = f.name
                files.append(data)
        except Exception as e:
            print(f"Ошибка загрузки {f}: {e}")
    return files


# ==================== ПАРСЕРЫ ПО УРОВНЯМ ====================

def parse_l1(data: Dict, index: int) -> Dict:
    """Парсит L1 (департамент) в компактный формат."""
    respondent = data.get('respondent', {})
    dept_text = respondent.get('department', '') or data.get('source_file', '')
    dep_code = detect_department_code(dept_text)

    result = {
        "id": f"{dep_code}-L1-{index:03d}",
        "fio": shorten_fio(respondent.get('fio', '')),
        "lvl": 1,
        "dep": dep_code,
        "num": respondent.get('division_size', ''),
        "goals": [],
        "processes": []
    }

    # Цели
    for goal in data.get('goals', []):
        goal_data = {
            "id": goal.get('row', 0),
            "text": normalize_text(goal.get('text', ''))
        }

        # Задачи для этой цели
        tasks_for_goal = []
        for task in data.get('tasks', []):
            if task.get('goal_row') == goal.get('row'):
                tasks_for_goal.append(normalize_text(task.get('text', '')))

        if tasks_for_goal:
            goal_data["tasks"] = tasks_for_goal

        # Процессы для этой цели (через задачу)
        processes_for_goal = []
        for proc in data.get('processes', []):
            if proc.get('task_row') == goal.get('row'):
                processes_for_goal.append(proc.get('row'))

        if processes_for_goal:
            goal_data["processes"] = processes_for_goal

        result["goals"].append(goal_data)

    # Процессы
    for proc in data.get('processes', []):
        auto_text = normalize_text(proc.get('needs_automation', ''))
        auto_priority = map_automation_priority(auto_text)

        proc_data = {
            "n": normalize_text(proc.get('name', '')),
            "cr": map_criticality(proc.get('criticality', '')),
            "staff": normalize_text(proc.get('employees_involved', '')),
            "depts": normalize_text(proc.get('departments_involved', '')),
            "reg": map_regulation_status(proc.get('has_regulation', '')),
        }

        # Проблемы
        pb = normalize_text(proc.get('problems', ''))
        if pb and pb.lower() not in ['нет', '-', 'нет проблем']:
            proc_data["pb"] = pb

        cs = normalize_text(proc.get('causes', ''))
        if cs and cs.lower() not in ['нет', '-']:
            proc_data["cs"] = cs

        # Автоматизация
        proc_data["auto"] = {
            "need": auto_priority != "0",
            "priority": auto_priority,
        }

        if 'уже' in auto_text.lower() or 'автоматизиров' in auto_text.lower():
            proc_data["auto"]["done"] = True

        reason = normalize_text(proc.get('why', ''))
        if reason and reason != '-':
            proc_data["auto"]["reason"] = reason

        effect = normalize_text(proc.get('expected_effect', ''))
        if effect and effect != '-':
            proc_data["auto"]["effect"] = effect

        result["processes"].append(proc_data)

    return result


def parse_l2(data: Dict, index: int) -> Dict:
    """Парсит L2 (отдел) в компактный формат."""
    respondent = data.get('respondent', {})

    # Определяем департамент: приоритет — по названию отдела (эталон)
    div_raw = normalize_text(respondent.get('division', ''))
    div_normalized = normalize_division_name(div_raw)
    dept_text = respondent.get('department', '') or div_raw or data.get('source_file', '')
    dep_code = detect_department_code(dept_text, division=div_raw)

    result = {
        "id": f"{dep_code}-L2-{index:03d}",
        "fio": shorten_fio(respondent.get('fio', '')),
        "lvl": 2,
        "dep": dep_code,
        "div": div_normalized,
        "pos": normalize_text(respondent.get('position', '')),
        "num": respondent.get('division_size', ''),
        "g": [],
        "tasks": [],
        "items": []
    }

    # Цели отдела
    for goal in data.get('division_goals', []):
        if goal:
            result["g"].append(normalize_text(goal))

    # Задачи и процессы
    for task in data.get('division_tasks', []):
        task_name = normalize_text(task.get('task_name', ''))
        if task_name:
            result["tasks"].append(task_name)

        # Процессы внутри задачи
        for proc in task.get('task_processes', []):
            # Системы
            systems = []
            sys_raw = proc.get('process_systems', '')
            if isinstance(sys_raw, list):
                systems = [s for s in sys_raw if s]
            elif sys_raw:
                systems = [sys_raw]

            # Базы данных
            databases = []
            db_raw = proc.get('process_databases', [])
            if isinstance(db_raw, list):
                databases = [d for d in db_raw if d]
            elif db_raw:
                databases = [db_raw]

            proc_data = {
                "n": normalize_text(proc.get('process_name', '')),
                "desc": normalize_text(proc.get('process_description', '')),
                "tr": normalize_text(proc.get('process_trigger', '')),
                "rs": normalize_text(proc.get('process_result', '')),
            }

            if systems:
                proc_data["sy"] = systems
            if databases:
                proc_data["db"] = databases

            proc_data["stages_count"] = normalize_text(proc.get('process_steps_count', ''))

            # Собираем все шаги в компактном виде
            stages_desc = []
            delays = []
            errors = []
            depts_involved = []
            dept_roles = []
            auto_needed = []

            for step in proc.get('process_steps', []):
                desc = normalize_text(step.get('step_description', ''))
                if desc:
                    stages_desc.append(desc)

                delay = normalize_text(step.get('step_delays', ''))
                if delay and delay.lower() not in ['нет', '']:
                    delays.append(delay)

                err = normalize_text(step.get('step_errors', ''))
                if err and err.lower() not in ['нет', '']:
                    errors.append(err)

                dept = normalize_text(step.get('step_involved_departments', ''))
                if dept:
                    depts_involved.append(dept)

                role = normalize_text(step.get('step_department_roles', ''))
                if role:
                    dept_roles.append(role)

                auto = normalize_text(step.get('step_needs_automation', ''))
                if auto and 'да' in auto.lower():
                    auto_needed.append(auto)

            if stages_desc:
                proc_data["stages_desc"] = stages_desc
            if delays:
                proc_data["delays"] = delays
            if errors:
                proc_data["errors"] = errors
            if depts_involved:
                proc_data["depts_involved"] = list(set(depts_involved))
            if dept_roles:
                proc_data["dept_roles"] = dept_roles
            if auto_needed:
                proc_data["auto_needed"] = True
                proc_data["auto_reason"] = auto_needed

            result["items"].append(proc_data)

    return result


def parse_l3(data: Dict, index: int) -> Dict:
    """Парсит L3 (сотрудник) в компактный формат."""
    respondent = data.get('respondent', {})

    # Определяем департамент: приоритет — по названию отдела (эталон)
    div_raw = normalize_text(respondent.get('division', ''))
    div_normalized = normalize_division_name(div_raw)
    dept_text = respondent.get('department', '') or div_raw or data.get('source_file', '')
    dep_code = detect_department_code(dept_text, division=div_raw)

    result = {
        "id": f"{dep_code}-L3-{index:03d}",
        "fio": shorten_fio(respondent.get('fio', '')),
        "lvl": 3,
        "dep": dep_code,
        "div": div_normalized,
        "pos": normalize_text(respondent.get('position', '')),
        "items": []
    }

    for task in data.get('tasks', []):
        # Пропускаем служебные записи
        task_name = task.get('task_name', '')
        if not task_name or '#NAME?' in str(task_name):
            continue

        task_data = {
            "n": normalize_text(task_name),
            "tp": map_task_type(task.get('task_type', '')),
            "rt": map_routine_level(task.get('routine_level', '')),
            "fr": map_frequency(task.get('regularity', '')),
            "tm": task.get('time_minutes', 0),
            "ch": normalize_text(task.get('task_character', '')),
        }

        # Проблемы
        pb = normalize_text(task.get('problem', ''))
        if pb and pb.lower() not in ['проблем нет', 'нет', '-']:
            task_data["pb"] = pb

        cs = normalize_text(task.get('problem_cause', ''))
        if cs and cs.lower() not in ['нет', '-']:
            task_data["cs"] = cs

        # Приоритет автоматизации (из priority)
        priority = str(task.get('priority', '3'))
        if priority in ['1', '2', '3']:
            task_data["auto_priority"] = priority

        # Взаимодействия
        interactions = task.get('interactions', [])
        if interactions:
            ints = []
            for inter in interactions:
                if isinstance(inter, dict):
                    dept = inter.get('department', '')
                    desc = inter.get('description', '')
                    if dept and dept != 'нет отделов':
                        ints.append({
                            "dept": detect_department_code(dept),
                            "dept_name": normalize_text(dept),
                            "desc": normalize_text(desc)
                        })
            if ints:
                task_data["interactions"] = ints

        # Ресурсы
        internal = task.get('internal_resources', [])
        external = task.get('external_resources', [])

        if internal:
            if isinstance(internal, list):
                task_data["res_int"] = [normalize_text(r) for r in internal if r]
            else:
                task_data["res_int"] = [normalize_text(internal)]

        if external:
            if isinstance(external, list):
                task_data["res_ext"] = [normalize_text(r) for r in external if r]
            else:
                task_data["res_ext"] = [normalize_text(external)]

        result["items"].append(task_data)

    return result


# ==================== MAIN ====================

def main():
    print("Загрузка данных L1...")
    l1_raw = load_json_files('L1_json')
    print(f"  Загружено {len(l1_raw)} анкет")

    print("Загрузка данных L2...")
    l2_raw = load_json_files('L2_json', exclude_patterns=['ID_Дата'])
    print(f"  Загружено {len(l2_raw)} анкет")

    print("Загрузка данных L3...")
    l3_raw = load_json_files('L3_json', exclude_patterns=[' 2.json', '_backup.json'])
    print(f"  Загружено {len(l3_raw)} анкет")

    print("\nПарсинг L1...")
    l1_parsed = [parse_l1(d, i+1) for i, d in enumerate(l1_raw)]

    print("Парсинг L2...")
    l2_parsed = [parse_l2(d, i+1) for i, d in enumerate(l2_raw)]

    print("Парсинг L3...")
    l3_parsed = [parse_l3(d, i+1) for i, d in enumerate(l3_raw)]

    # Статистика
    total_processes_l1 = sum(len(item.get('processes', [])) for item in l1_parsed)
    total_items_l2 = sum(len(item.get('items', [])) for item in l2_parsed)
    total_items_l3 = sum(len(item.get('items', [])) for item in l3_parsed)

    # Формируем итоговую структуру
    normalized_data = {
        "meta": {
            "version": "2.0-compact",
            "counts": {
                "l1": len(l1_parsed),
                "l2": len(l2_parsed),
                "l3": len(l3_parsed),
                "processes_l1": total_processes_l1,
                "items_l2": total_items_l2,
                "items_l3": total_items_l3
            }
        },
        "dictionaries": {
            "departments": DEPARTMENTS,
            "criticality": CRITICALITY,
            "automation_priority": AUTOMATION_PRIORITY,
            "regulation_status": REGULATION_STATUS,
            "task_type": TASK_TYPE,
            "routine_level": ROUTINE_LEVEL,
            "frequency": FREQUENCY
        },
        "fields_l1": {
            "id": "Уникальный идентификатор (DEP-L1-NNN)",
            "fio": "ФИО (сокращённое)",
            "lvl": "Уровень (1=руководитель департамента)",
            "dep": "Код департамента",
            "num": "Численность департамента",
            "goals": "Массив целей департамента",
            "goals[].id": "Номер цели",
            "goals[].text": "Текст цели",
            "goals[].tasks": "Задачи для достижения цели",
            "goals[].processes": "Процессы, относящиеся к цели",
            "processes[].n": "Название процесса",
            "processes[].cr": "Критичность (H/M/L)",
            "processes[].staff": "Количество сотрудников в процессе",
            "processes[].depts": "Отделы-участники",
            "processes[].reg": "Статус регламента (exists/draft/none)",
            "processes[].pb": "Описание проблем в процессе",
            "processes[].cs": "Причина проблем",
            "processes[].auto.need": "Нужна ли автоматизация (true/false)",
            "processes[].auto.priority": "Приоритет автоматизации (1/2/3/0)",
            "processes[].auto.done": "Уже автоматизировано (true если да)",
            "processes[].auto.reason": "Обоснование необходимости автоматизации",
            "processes[].auto.effect": "Ожидаемый эффект от автоматизации"
        },
        "fields_l2": {
            "id": "Уникальный идентификатор (DEP-L2-NNN)",
            "fio": "ФИО (сокращённое)",
            "lvl": "Уровень (2=руководитель отдела)",
            "dep": "Код департамента",
            "div": "Название отдела",
            "pos": "Должность",
            "num": "Численность отдела",
            "g": "Цели отдела",
            "tasks": "Задачи отдела (массив разделенных задач)",
            "items": "Процессы отдела",
            "items[].n": "Название процесса",
            "items[].desc": "Описание процесса",
            "items[].tr": "Триггер процесса",
            "items[].rs": "Результат процесса",
            "items[].sy": "Используемые системы",
            "items[].db": "Используемые базы данных и базы знаний",
            "items[].stages_count": "Количество этапов в процессе",
            "items[].stages_desc": "Описание всех этапов процесса",
            "items[].delays": "Задержки на этапах",
            "items[].errors": "Ошибки на этапах",
            "items[].depts_involved": "Какие отделы вовлечены в этапы",
            "items[].dept_roles": "Как каждый отдел участвует в этапах",
            "items[].auto_needed": "Нужна ли автоматизация",
            "items[].auto_reason": "Обоснование необходимости автоматизации"
        },
        "fields_l3": {
            "id": "Уникальный идентификатор (DEP-L3-NNN)",
            "fio": "ФИО (сокращённое)",
            "lvl": "Уровень (3=сотрудник)",
            "dep": "Код департамента",
            "div": "Название отдела",
            "pos": "Должность",
            "items": "Задачи сотрудника",
            "items[].n": "Название задачи",
            "items[].tp": "Тип задачи (R/P)",
            "items[].rt": "Уровень рутинности (VH/H/M/L)",
            "items[].fr": "Частота (D/W/M/Q/R)",
            "items[].tm": "Время выполнения в минутах",
            "items[].ch": "Характер задачи",
            "items[].pb": "Проблемы",
            "items[].cs": "Причины проблем",
            "items[].auto_priority": "Приоритет автоматизации (1/2/3)",
            "items[].interactions": "Взаимодействия с другими отделами",
            "items[].res_int": "Внутренние ресурсы",
            "items[].res_ext": "Внешние ресурсы"
        },
        "data": {
            "l1": l1_parsed,
            "l2": l2_parsed,
            "l3": l3_parsed
        }
    }

    # Сохраняем результат
    output_path = BASE_DIR / 'normalized_compact.json'
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(normalized_data, f, ensure_ascii=False, indent=2)

    print(f"\n✓ Данные сохранены в {output_path}")
    print(f"\nСтатистика:")
    print(f"  L1 (департаменты): {len(l1_parsed)}")
    print(f"  L2 (отделы): {len(l2_parsed)}")
    print(f"  L3 (сотрудники): {len(l3_parsed)}")
    print(f"  Процессов L1: {total_processes_l1}")
    print(f"  Процессов L2: {total_items_l2}")
    print(f"  Задач L3: {total_items_l3}")

    # Распределение по департаментам
    print(f"\nРаспределение по департаментам:")
    dept_counts = {}
    for item in l1_parsed + l2_parsed + l3_parsed:
        dep = item.get('dep', 'UNK')
        dept_counts[dep] = dept_counts.get(dep, 0) + 1

    for dep, count in sorted(dept_counts.items(), key=lambda x: -x[1]):
        print(f"  {dep}: {count} ({DEPARTMENTS.get(dep, 'Unknown')})")


if __name__ == '__main__':
    main()
