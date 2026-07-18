#!/usr/bin/env python3
import sqlite3
import json
from pathlib import Path

DB_PATH = Path(__file__).resolve().parents[1] / "data" / "lotto_history.db"

draws = [
    ('2026-06-13', [5, 13, 17, 20, 31, 32]),
    ('2026-06-11', [7, 14, 18, 31, 34, 39]),
    ('2026-06-09', [2, 4, 24, 30, 39, 41]),
    ('2026-06-06', [7, 8, 17, 32, 36, 47]),
    ('2026-06-04', [13, 17, 20, 41, 43, 46]),
    ('2026-06-02', [8, 20, 28, 31, 36, 39]),
    ('2026-05-30', [1, 6, 10, 32, 35, 42]),
    ('2026-05-28', [6, 7, 15, 22, 42, 43]),
    ('2026-05-26', [3, 4, 5, 16, 23, 46]),
    ('2026-05-23', [2, 9, 14, 17, 30, 46]),
    ('2026-05-21', [11, 21, 24, 28, 42, 44]),
    ('2026-05-19', [17, 18, 20, 38, 42, 49]),
    ('2026-05-16', [3, 11, 12, 20, 23, 49]),
    ('2026-05-14', [9, 15, 17, 24, 42, 44]),
]

conn = sqlite3.connect(str(DB_PATH))
cursor = conn.cursor()

for date, numbers in draws:
    numbers_json = json.dumps(numbers)
    cursor.execute("INSERT INTO draws (draw_date, numbers) VALUES (?, ?)", (date, numbers_json))

conn.commit()
conn.close()

print(f"✓ Dodano {len(draws)} losowań do bazy")
