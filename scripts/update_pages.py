#!/usr/bin/env python3
"""Обновляет страницы и коды книг в generate_ktp_05_02.py
по оценочным данным из ОИ2 (Нестеренко) и ОИ3 (Сидорова).

Поддерживает оба формата:
  многострочный:  "equipment": MAIN_BOOK_CODE,\n  "task": "Стр. X-Y",
  компактный:    "equipment": MAIN_BOOK_CODE, "task": "Стр. X-Y", "note":
"""

import re

SCRIPT_PATH = "/home/z/my-project/YMK/scripts/generate_ktp_05_02.py"

# Маппинг: старые страницы (ОИ1) -> (новый код книги, новые страницы)
PAGE_MAP = {
    "Стр. 5-12":   ("BOOK_OI3", "Стр. 5-89"),      # 2.1 Безопасность труда
    "Стр. 13-22":  ("BOOK_OI2", "Стр. 79-93"),      # 2.2 ч.1
    "Стр. 23-30":  ("BOOK_OI2", "Стр. 94-108"),     # 2.2 ч.2
    "Стр. 31-40":  ("BOOK_OI2", "Стр. 109-123"),    # 2.2 ч.3
    "Стр. 41-50":  ("BOOK_OI2", "Стр. 124-138"),    # 2.2 ч.4
    "Стр. 51-60":  ("BOOK_OI2", "Стр. 139-153"),    # 2.2 ч.5
    "Стр. 61-70":  ("BOOK_OI2", "Стр. 154-167"),    # 2.2 ч.6
    "Стр. 71-78":  ("BOOK_OI2", "Стр. 228-241"),    # 2.3 ч.1
    "Стр. 79-88":  ("BOOK_OI2", "Стр. 242-255"),    # 2.3 ч.2
    "Стр. 89-100": ("BOOK_OI2", "Стр. 256-269"),    # 2.3 ч.3
    "Стр. 101-112":("BOOK_OI2", "Стр. 270-283"),    # 2.3 ч.4
    "Стр. 113-120":("BOOK_OI2", "Стр. 124-167"),    # 2.4 Заземление
    "Стр. 121-130":("BOOK_OI3", "Стр. 203-218"),    # 2.5 ч.1
    "Стр. 131-140":("BOOK_OI3", "Стр. 219-234"),    # 2.5 ч.2
    "Стр. 141-150":("BOOK_OI3", "Стр. 263-273"),    # 2.5 ч.3
    "Стр. 151-160":("BOOK_OI3", "Стр. 274-284"),    # 2.5 ч.4
    "Стр. 161-170":("BOOK_OI3", "Стр. 175-181"),    # 2.6 ч.1
    "Стр. 171-180":("BOOK_OI3", "Стр. 182-188"),    # 2.6 ч.2
    "Стр. 181-190":("BOOK_OI3", "Стр. 189-194"),    # 2.6 ч.3
    "Стр. 191-200":("BOOK_OI3", "Стр. 285-296"),    # 2.7 ч.1
    "Стр. 201-210":("BOOK_OI3", "Стр. 297-308"),    # 2.7 ч.2
    "Стр. 211-220":("BOOK_OI3", "Стр. 309-320"),    # 2.7 ч.3
    "Стр. 221-230":("BOOK_OI3", "Стр. 203-218"),    # 2.8 ч.1
    "Стр. 231-240":("BOOK_OI3", "Стр. 219-234"),    # 2.8 ч.2
    "Стр. 241-250":("BOOK_OI3", "Стр. 263-284"),    # 2.8 ч.3
}

with open(SCRIPT_PATH, 'r', encoding='utf-8') as f:
    content = f.read()

replaced = 0
not_found = 0

# Простой подход: для каждой старой страницы заменить
# "equipment": MAIN_BOOK_CODE, "task": "Стр. X-Y"
# на
# "equipment": NEW_BOOK, "task": "Стр. N-M"
for old_pages, (new_book, new_pages) in PAGE_MAP.items():
    # Компактный формат (однострочный): 
    # "equipment": MAIN_BOOK_CODE, "task": "Стр. X-Y", "note":
    old_compact = f'"equipment": MAIN_BOOK_CODE, "task": "{old_pages}"'
    new_compact = f'"equipment": {new_book}, "task": "{new_pages}"'
    
    # Многострочный формат:
    # "equipment": MAIN_BOOK_CODE,\n             "task": "Стр. X-Y",
    old_multi = f'"equipment": MAIN_BOOK_CODE,\n             "task": "{old_pages}"'
    new_multi = f'"equipment": {new_book},\n             "task": "{new_pages}"'
    
    if old_compact in content:
        content = content.replace(old_compact, new_compact)
        replaced += 1
        print(f"  OK (compact): {old_pages} -> {new_book} {new_pages}")
    elif old_multi in content:
        content = content.replace(old_multi, new_multi)
        replaced += 1
        print(f"  OK (multi):   {old_pages} -> {new_book} {new_pages}")
    else:
        not_found += 1
        print(f"  НЕ НАЙДЕНО: {old_pages}")

print(f"\nЗаменено: {replaced}, не найдено: {not_found}")

# Проверка: нет ли оставшихся MAIN_BOOK_CODE
remaining = content.count('MAIN_BOOK_CODE')
if remaining > 0:
    print(f"ВНИМАНИЕ: осталось {remaining} ссылок на MAIN_BOOK_CODE")
else:
    print("OK: ссылок на MAIN_BOOK_CODE не осталось")

with open(SCRIPT_PATH, 'w', encoding='utf-8') as f:
    f.write(content)

print("\nФайл сохранён.")