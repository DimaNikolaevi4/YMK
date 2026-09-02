#!/usr/bin/env python3
"""Rename BOOK_OI3->BOOK_OI2, BOOK_OI2->BOOK_OI1 in generate_ktp_05_02.py."""
path = "/home/z/my-project/YMK/scripts/generate_ktp_05_02.py"

with open(path, 'r', encoding='utf-8') as f:
    content = f.read()

# Step 1: BOOK_OI3 -> BOOK_OI2 (Сидорова)
content = content.replace('BOOK_OI3', 'BOOK_OI2')
print("BOOK_OI3 -> BOOK_OI2")

# Step 2: BOOK_OI2 -> BOOK_OI1 (Нестеренко)
content = content.replace('BOOK_OI2', 'BOOK_OI1')
print("BOOK_OI2 -> BOOK_OI1")

# Fix: both vars became BOOK_OI1, restore BOOK_OI2 = "ОИ2"
content = content.replace('BOOK_OI1 = "ОИ2"', 'BOOK_OI2 = "ОИ2"')
print("Restored BOOK_OI2 = ОИ2")

# Fix validation: (BOOK_OI1, BOOK_OI1) -> (BOOK_OI1, BOOK_OI2)
content = content.replace('eq in (BOOK_OI1, BOOK_OI1)', 'eq in (BOOK_OI1, BOOK_OI2)')
print("Fixed validation")

# Fix comment
content = content.replace('BOOK_OI1 (темы 2.2-2.4 монтаж) или BOOK_OI1', 'BOOK_OI1 (темы 2.2-2.4 монтаж) или BOOK_OI2')
print("Fixed comment")

with open(path, 'w', encoding='utf-8') as f:
    f.write(content)

# Verify
with open(path, 'r', encoding='utf-8') as f:
    v = f.read()
print(f"Осталось BOOK_OI3: {v.count('BOOK_OI3')}")
for i, line in enumerate(v.split('\n'), 1):
    if 'BOOK_OI' in line and ('=' in line or 'equipment' in line or 'ok_eq' in line):
        print(f"  L{i}: {line.strip()}")