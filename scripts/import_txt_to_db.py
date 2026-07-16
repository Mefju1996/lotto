#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
import_txt_to_db.py

Importuje wyniki Lotto z plików TXT (data/raw/lotto/wyniki*.txt)
do bazy SQLite (data/lotto_history.db) kompatybilnej z lotto_generator_tkinter.py.

Format wejściowy TXT (obsugiwane warianty):
  6 14 21 32 38 45                        <- sama 6 liczb (np. wynik_all.txt)
  14-07-2026  7 9 20 31 38 43             <- data + 6 liczb (wyniki_scraped_YYYY.txt)

Struktura tabeli draws (zgodna z tkinter):
  id          INTEGER PRIMARY KEY AUTOINCREMENT
  draw_date   TEXT    (YYYY-MM-DD lub pusty)
  numbers     TEXT    (JSON array, np. [7,9,20,31,38,43])

Użycie:
    python scripts/import_txt_to_db.py
    python scripts/import_txt_to_db.py --dry-run   # tylko podgląd, bez zapisu
    python scripts/import_txt_to_db.py --file data/raw/lotto/wyniki_scraped_2026.txt
    python scripts/import_txt_to_db.py --clear      # czyść DB przed importem

Wymagania: tylko biblioteka standardowa Python (re, json, sqlite3, pathlib)
"""

import argparse
import json
import re
import sqlite3
import sys
from datetime import datetime
from pathlib import Path

# ============================================================
# ŚCIEŻKI
# ============================================================
ROOT_DIR = Path(__file__).resolve().parents[1]
RAW_DIR  = ROOT_DIR / "data" / "raw" / "lotto"
DB_PATH  = ROOT_DIR / "data" / "lotto_history.db"

# Wzorce do rozpoznawania linii
# Wariant A: "14-07-2026  7 9 20 31 38 43"  (wyniki_scraped_*.txt)
PAT_DATE_NUMS = re.compile(
    r"(\d{2}[-./]\d{2}[-./]\d{4})\s+(\d{1,2})\s+(\d{1,2})\s+(\d{1,2})\s+(\d{1,2})\s+(\d{1,2})\s+(\d{1,2})"
)
# Wariant B: "7 9 20 31 38 43"  (wynik_all.txt, wyniki_all.txt itp.)
PAT_NUMS_ONLY = re.compile(
    r"^\s*(\d{1,2})\s+(\d{1,2})\s+(\d{1,2})\s+(\d{1,2})\s+(\d{1,2})\s+(\d{1,2})\s*$"
)


# ============================================================
# PARSOWANIE
# ============================================================

def parse_date(raw: str) -> str:
    """Konwertuje różne formaty daty na YYYY-MM-DD lub zwraca pusty string."""
    raw = raw.strip()
    for fmt in ("%d-%m-%Y", "%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(raw, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return ""


def parse_line(line: str) -> tuple[str, list[int]] | None:
    """
    Zwraca (data_str, [n1..n6]) lub None jeśli linia nie pasuje.
    data_str może być pustym stringiem jeśli linia zawiera tylko liczby.
    """
    line = line.strip()
    if not line or line.startswith("#"):
        return None

    # Wariant A: data + liczby
    m = PAT_DATE_NUMS.search(line)
    if m:
        date_str = parse_date(m.group(1))
        nums = [int(m.group(i)) for i in range(2, 8)]
        return date_str, nums

    # Wariant B: tylko liczby
    m = PAT_NUMS_ONLY.match(line)
    if m:
        nums = [int(m.group(i)) for i in range(1, 7)]
        if all(1 <= n <= 49 for n in nums) and len(set(nums)) == 6:
            return "", nums

    return None


def parse_file(path: Path) -> list[tuple[str, list[int]]]:
    """Parsuje jeden plik TXT, zwraca listę (data, liczby)."""
    results = []
    skipped = 0
    with open(path, encoding="utf-8", errors="replace") as f:
        for lineno, line in enumerate(f, 1):
            parsed = parse_line(line)
            if parsed is None:
                if line.strip():
                    skipped += 1
            else:
                results.append(parsed)
    if skipped:
        print(f"     ⚠ Pominięto {skipped} niepasujących linii w {path.name}")
    return results


# ============================================================
# BAZA DANYCH
# ============================================================

def init_db(db_path: Path) -> sqlite3.Connection:
    """Tworzy/otwiera SQLite i upewnia się, że tabela draws istnieje."""
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(db_path))
    conn.execute("""
        CREATE TABLE IF NOT EXISTS draws (
            id        INTEGER PRIMARY KEY AUTOINCREMENT,
            draw_date TEXT,
            numbers   TEXT NOT NULL
        )
    """)
    conn.commit()
    return conn


def get_existing_numbers(conn: sqlite3.Connection) -> set[str]:
    """Zwraca zbiór JSON-owych stringów już zapisanych losowań."""
    rows = conn.execute("SELECT numbers FROM draws").fetchall()
    return {r[0] for r in rows}


def insert_draws(
    conn: sqlite3.Connection,
    draws: list[tuple[str, list[int]]],
    existing: set[str],
    dry_run: bool = False,
) -> tuple[int, int]:
    """
    Wstawia nowe losowania do DB.
    Zwraca (dodane, pominięte_duplikaty).
    """
    added = 0
    skipped = 0
    rows_to_insert = []

    for date_str, nums in draws:
        nums_sorted = sorted(nums)
        nums_json = json.dumps(nums_sorted)
        if nums_json in existing:
            skipped += 1
            continue
        existing.add(nums_json)   # zapobiegaj duplikatom wewnątrz sesji
        rows_to_insert.append((date_str, nums_json))
        added += 1

    if not dry_run and rows_to_insert:
        conn.executemany(
            "INSERT INTO draws (draw_date, numbers) VALUES (?, ?)",
            rows_to_insert,
        )
        conn.commit()

    return added, skipped


# ============================================================
# MAIN
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description="Import wyników Lotto z plików TXT do lotto_history.db"
    )
    parser.add_argument(
        "--file", type=Path, default=None,
        help="Konkretny plik TXT do importu (domyślnie: wszystkie wyniki*.txt w data/raw/lotto/)",
    )
    parser.add_argument(
        "--db", type=Path, default=DB_PATH,
        help=f"Plik SQLite (domyślnie: {DB_PATH})",
    )
    parser.add_argument(
        "--dry-run", action="store_true",
        help="Tylko podgląd — nie zapisuje do bazy",
    )
    parser.add_argument(
        "--clear", action="store_true",
        help="Czyść tabelę draws przed importem (UWAGA: usuwa wszystkie dane!)",
    )
    args = parser.parse_args()

    print("=" * 60)
    print("IMPORT TXT → SQLite (lotto_history.db)")
    if args.dry_run:
        print("  [DRY-RUN — tylko podgląd, bez zapisu]")
    print("=" * 60)

    # Zbierz pliki do przetworzenia
    if args.file:
        if not args.file.exists():
            print(f"✗ Plik nie istnieje: {args.file}", file=sys.stderr)
            sys.exit(1)
        txt_files = [args.file]
    else:
        txt_files = sorted(RAW_DIR.glob("wyniki*.txt"))
        # Dodaj też wynik_all.txt jeśli istnieje
        extra = [RAW_DIR / "wynik_all.txt", RAW_DIR / "wyniki_all.txt"]
        for f in extra:
            if f.exists() and f not in txt_files:
                txt_files.insert(0, f)

    if not txt_files:
        print(f"✗ Nie znaleziono plików TXT w: {RAW_DIR}", file=sys.stderr)
        print("  Upewnij się, że uruchomiłeś najpierw scraper_megalotto.py")
        sys.exit(1)

    print(f"\nPliki źródłowe ({len(txt_files)})")
    for f in txt_files:
        print(f"  · {f.name}")

    # Otwieraj/twórz bazę
    print(f"\nBaza danych: {args.db}")
    conn = init_db(args.db)

    if args.clear and not args.dry_run:
        conn.execute("DELETE FROM draws")
        conn.commit()
        print("  ✓ Tabela draws wyczyszczona")

    existing = get_existing_numbers(conn)
    print(f"  ✓ Istniejących rekordów w DB: {len(existing)}")

    # Import
    total_added = 0
    total_skipped = 0

    print("\nParsowanie i import...")
    for txt_path in txt_files:
        draws = parse_file(txt_path)
        added, skipped = insert_draws(conn, draws, existing, dry_run=args.dry_run)
        total_added += added
        total_skipped += skipped
        status = "(dry-run)" if args.dry_run else "✓"
        print(f"  {status} {txt_path.name}: {len(draws)} linii → +{added} nowych, {skipped} dup.")

    conn.close()

    print("\n" + "=" * 60)
    print(f"WYNIK: +{total_added} nowych losowań, {total_skipped} duplikatów pominięto")
    if not args.dry_run:
        final_conn = sqlite3.connect(str(args.db))
        count = final_conn.execute("SELECT COUNT(*) FROM draws").fetchone()[0]
        final_conn.close()
        print(f"Stan DB po imporcie: {count} losowań łącznie")
        print(f"\nMożesz teraz otworzyć lotto_generator_tkinter.py")
        print(f"i kliknąć 'Plik historii' → wskazać: {args.db}")
    print("=" * 60)


if __name__ == "__main__":
    main()
