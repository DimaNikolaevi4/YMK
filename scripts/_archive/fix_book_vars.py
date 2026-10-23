#!/usr/bin/env python3
"""Fix: change Сидорова topics (2.1, 2.5-2.8) from BOOK_OI1 back to BOOK_OI2."""
path = "/home/z/my-project/YMK/scripts/generate_ktp_05_02.py"

with open(path, 'r', encoding='utf-8') as f:
    content = f.read()

# Page ranges for Сидорова (ОИ2) topics: 2.1, 2.5, 2.6, 2.7, 2.8
SIDOROVA_PAGES = [
    "Стр. 5-89",      # 2.1
    "Стр. 203-218",   # 2.5 ч.1 / 2.8 ч.1
    "Стр. 219-234",   # 2.5 ч.2 / 2.8 ч.2
    "Стр. 263-273",   # 2.5 ч.3
    "Стр. 274-284",   # 2.5 ч.4
    "Стр. 175-181",   # 2.6 ч.1
    "Стр. 182-188",   # 2.6 ч.2
    "Стр. 189-194",   # 2.6 ч.3
    "Стр. 285-296",   # 2.7 ч.1
    "Стр. 297-308",   # 2.7 ч.2
    "Стр. 309-320",   # 2.7 ч.3
    "Стр. 263-284",   # 2.8 ч.3
]

fixed = 0
for pages in SIDOROVA_PAGES:
    old = f'BOOK_OI1, "task": "{pages}"'
    new = f'BOOK_OI2, "task": "{pages}"'
    if old in content:
        content = content.replace(old, new)
        fixed += 1
        print(f"  {pages} -> BOOK_OI2")
    else:
        print(f"  НЕ НАЙДЕНО: {pages}")

with open(path, 'w', encoding='utf-8') as f:
    f.write(content)

print(f"\nИсправлено: {fixed}")

# Verify
with open(path, 'r', encoding='utf-8') as f:
    v = f.read()
print(f"BOOK_OI1 count: {v.count('BOOK_OI1')}")
print(f"BOOK_OI2 count: {v.count('BOOK_OI2')}")