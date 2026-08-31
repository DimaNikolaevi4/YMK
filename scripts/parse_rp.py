#!/usr/bin/env python3
"""
Скрипт для парсинга рабочих программ (РП) и извлечения данных о МДК.

Использование:
    python3 scripts/parse_rp.py RP/РП_ПМ_03_КИП.txt

Вывод: структурированные данные по каждому МДК в формате JSON.
"""

import re
import sys
import json


def extract_mdks(text):
    """Извлечь перечень МДК и формы аттестации."""
    mdks = []
    # Ищем блок с таблицей контроля знаний
    block = text
    # Ищем строки с МДК и формой аттестации
    pattern = r'МДК\s+(\d+\.\d+)\s+(.+?)(?:\s+\|\s*(Дифференцированный зачет|Экзамен|Зачет))'
    for m in re.finditer(pattern, text, re.IGNORECASE):
        mdks.append({
            'code': m.group(1).strip(),
            'name': m.group(2).strip(),
            'exam_form': m.group(3).strip()
        })
    # Альтернативный поиск - построчный
    if not mdks:
        lines = text.split('\n')
        current_mdk = None
        for line in lines:
            if 'МДК' in line and re.search(r'МДК\s+\d+\.\d+', line):
                m = re.search(r'(МДК\s+\d+\.\d+)\s+(.+?)(?:\s+\|\s*(Дифференцированный зачет|Экзамен|Зачет)|$)', line)
                if m:
                    current_mdk = {
                        'code': m.group(1).strip().replace('МДК ', ''),
                        'name': m.group(2).strip(),
                        'exam_form': (m.group(3) or '').strip()
                    }
                    if current_mdk['name'] and current_mdk['code']:
                        mdks.append(current_mdk)
    return mdks


def extract_thematic_plan(text):
    """Извлечь данные из тематического плана (раздел 3.1)."""
    data = {}
    # Ищем блок тематического плана
    lines = text.split('\n')
    in_plan = False
    current_mdk = None

    for i, line in enumerate(lines):
        s = line.strip()
        if 'Тематический план' in s or '3.1' in s:
            in_plan = True
        if in_plan and 'Контроль и оцен' in s:
            break
        if in_plan:
            # Ищем строки МДК с часами
            mdk_match = re.search(r'МДК\s+(\d+\.\d+)', s)
            if mdk_match and 'Тема' not in s:
                code = mdk_match.group(1)
                if code not in data:
                    data[code] = {'topics': []}
                current_mdk = code
    return data


def extract_topics(text, mdk_code):
    """Извлечь темы для конкретного МДК."""
    topics = []
    lines = text.split('\n')
    in_mdk = False

    for i, line in enumerate(lines):
        s = line.strip()
        # Определяем, что мы в нужном МДК
        if f'МДК {mdk_code}' in s or f'МДК\s+{mdk_code}' in s:
            in_mdk = True
            continue
        # Если встретили другой МДК - выходим
        if in_mdk and re.search(r'МДК\s+\d+\.\d+', s):
            if mdk_code not in s:
                break
        # Ищем темы
        if in_mdk:
            topic_match = re.match(r'Тема\s+(\d+\.\d+)\s+(.+?)(?:\s*\|\s*(\d+)\s*\|)', s)
            if topic_match:
                topics.append({
                    'number': topic_match.group(1),
                    'name': topic_match.group(2).strip(),
                    'total_hours': int(topic_match.group(3)) if topic_match.group(3) else 0
                })
    return topics


def extract_literature(text):
    """Извлечь списки литературы."""
    result = {'basic': [], 'additional': []}
    lines = text.split('\n')
    section = None

    for line in lines:
        s = line.strip()
        if 'Основные источники' in s:
            section = 'basic'
            continue
        if 'Дополнительные источники' in s:
            section = 'additional'
            continue
        if 'Средства массовой информации' in s or '4.3' in s:
            section = None
            continue
        if section and s and re.match(r'^\d+\.', s):
            entry = re.sub(r'^\d+\.\s*', '', s).strip()
            if entry:
                result[section].append(entry)
    return result


def extract_equipment(text):
    """Извлечь перечень оборудования."""
    equipment = []
    lines = text.split('\n')
    in_section = False

    for line in lines:
        s = line.strip()
        if '4.1' in s and 'материал' in s.lower():
            in_section = True
            continue
        if in_section and '4.2' in s:
            break
        if in_section and s.startswith('-'):
            item = s.lstrip('- ').strip()
            if item:
                equipment.append(item)
    return equipment


def main():
    if len(sys.argv) < 2:
        print('Использование: python3 parse_rp.py <путь_к_файлу_РП.txt>')
        sys.exit(1)

    filepath = sys.argv[1]
    with open(filepath, 'r', encoding='utf-8', errors='replace') as f:
        text = f.read()

    # Извлекаем все данные
    mdks = extract_mdks(text)
    literature = extract_literature(text)
    equipment = extract_equipment(text)

    result = {
        'mdks': mdks,
        'literature': literature,
        'equipment': equipment
    }

    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
