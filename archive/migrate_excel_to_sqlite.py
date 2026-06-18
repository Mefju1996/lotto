"""
Migrate lottery draw history from Excel to SQLite.
"""
import sqlite3
import json
from pathlib import Path
import pandas as pd
from datetime import datetime


def create_database(db_path):
    """Create SQLite database with draw history table."""
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS draws (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        draw_date TEXT,
        numbers TEXT,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
    """)
    
    conn.commit()
    conn.close()
    print(f"Database created: {db_path}")


def migrate_from_excel(excel_path, sheet_name, db_path):
    """
    Migrate draw history from Excel to SQLite.
    
    Excel columns expected: A-F (6 columns with numbers)
    """
    # Read Excel
    try:
        df = pd.read_excel(excel_path, sheet_name=sheet_name, engine="openpyxl", usecols="A:F")
    except Exception as e:
        print(f"Error reading Excel: {e}")
        return False
    
    if df.empty:
        print("Excel sheet is empty")
        return False
    
    # Connect to database
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    # Migrate rows
    migrated_count = 0
    skipped_count = 0
    
    for idx, row in df.iterrows():
        try:
            # Extract and validate numbers
            numbers = []
            for val in row:
                if pd.notna(val):
                    numbers.append(int(val))
            
            if len(numbers) == 6:
                numbers_json = json.dumps(sorted(numbers))
                draw_date = datetime.now().isoformat()
                
                cursor.execute(
                    "INSERT INTO draws (draw_date, numbers) VALUES (?, ?)",
                    (draw_date, numbers_json)
                )
                migrated_count += 1
            else:
                skipped_count += 1
        except Exception as e:
            print(f"Row {idx} error: {e}")
            skipped_count += 1
    
    conn.commit()
    conn.close()
    
    print(f"Migration complete: {migrated_count} rows migrated, {skipped_count} skipped")
    return True


def check_database(db_path):
    """Display summary of imported data."""
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    cursor.execute("SELECT COUNT(*) FROM draws")
    count = cursor.fetchone()[0]
    
    cursor.execute("SELECT * FROM draws LIMIT 5")
    rows = cursor.fetchall()
    
    conn.close()
    
    print(f"\nDatabase summary:")
    print(f"Total draws: {count}")
    print(f"Sample rows:")
    for row in rows:
        print(f"  ID: {row[0]}, Date: {row[1]}, Numbers: {row[2]}")


if __name__ == "__main__":
    base_dir = Path(__file__).resolve().parent
    excel_path = base_dir / "wyniki_lotto.xlsx"
    db_path = base_dir / "lotto_history.db"
    sheet_name = "Arkusz1"
    
    if not excel_path.exists():
        print(f"Error: Excel file not found: {excel_path}")
        exit(1)
    
    print(f"Migrating from: {excel_path}")
    print(f"To database: {db_path}")
    
    # Create database and migrate
    create_database(db_path)
    success = migrate_from_excel(excel_path, sheet_name, db_path)
    
    if success:
        check_database(db_path)
        print("\n✓ Migration successful!")
    else:
        print("\n✗ Migration failed")
