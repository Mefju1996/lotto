#!/usr/bin/env python3
"""Test SQLite integration without GUI."""

import sqlite3
import json
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "lotto_history.db"

print(f"Testing SQLite database at: {DB_PATH}")
print(f"Database exists: {DB_PATH.exists()}")

if DB_PATH.exists():
    try:
        conn = sqlite3.connect(str(DB_PATH))
        cursor = conn.execute("SELECT COUNT(*) FROM draws")
        count = cursor.fetchone()[0]
        print(f"✓ Database connected")
        print(f"✓ Total records: {count}")
        
        # Test query for a sample combination
        cursor = conn.execute("SELECT * FROM draws LIMIT 5")
        print(f"\nSample records:")
        for row in cursor:
            print(f"  ID: {row[0]}, Numbers: {row[2]}")
        
        # Test if combination exists
        test_numbers = [7, 8, 16, 34, 42, 7351]
        test_json = json.dumps(sorted(test_numbers))
        cursor = conn.execute("SELECT COUNT(*) FROM draws WHERE numbers = ?", (test_json,))
        exists = cursor.fetchone()[0] > 0
        print(f"\n✓ Test combination {test_numbers}: {'EXISTS' if exists else 'NOT FOUND'}")
        
        conn.close()
        print("\n✓ All tests passed!")
    except Exception as e:
        print(f"✗ Error: {e}")
else:
    print("✗ Database file not found!")
