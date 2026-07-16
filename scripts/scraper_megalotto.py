#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Scraper wyników Lotto z https://megalotto.pl/wyniki/lotto
Zsynchronizowany z:
  - lotto_generator_tkinter.py  → data/lotto_history.db  (tabela draws)
  - generate_lotto_stats_final.py → data/wyniki_lotto.xlsx / Arkusz1

Użycie:
    python scripts/scraper_megalotto.py --update-xlsx
    python scripts/scraper_megalotto.py --year 2025 --update-xlsx
    python scripts/scraper_megalotto.py --all-years --update-xlsx

Wymagania:
    pip install requests beautifulsoup4 openpyxl pandas
"""

import argparse
import json
import sqlite3
import time
from datetime import datetime
from pathlib import Path

import pandas as pd
import requests
from bs4 import BeautifulSoup

# ============================================================
# ŚCIEŻKI
# ============================================================
ROOT_DIR  = Path(__file__).resolve().parents[1]
DATA_DIR  = ROOT_DIR / "data"
RAW_DIR   = DATA_DIR / "raw" / "lotto"
XLSX_FILE = DATA_DIR / "wyniki_lotto.xlsx"
DB_FILE   = DATA_DIR / "lotto_history.db"   # używane przez lotto_generator_tkinter.py
SHEET     = "Arkusz1"

# Kolumny wymagane przez generate_lotto_stats_final.py
COLUMNS = ["data", "nr_losowania", "pierwsza", "druga", "trzecia", "czwarta", "piąta", "szósta"]

BASE_URL  = "https://megalotto.pl/wyniki/lotto"
HEADERS   = {"User-Agent": "Mozilla/5.0 (compatible; LottoScraper/1.0)"}
DELAY_SEC = 1.5


# ============================================================
# PARSOWANIE STRONY
# ============================================================
def _parse_page(html: str) -> list[dict]:
    soup = BeautifulSoup(html, "html.parser")
    results = []

    nr_tags   = soup.find_all("li", class_="nr_in_list")
    date_tags = soup.find_all("li", class_="date_in_list")
    num_tags  = soup.find_all("li", class_="numbers_in_list")

    if len(num_tags) % 6 != 0:
        return results

    n_draws = min(len(nr_tags), len(date_tags), len(num_tags) // 6)

    for i in range(n_draws):
        try:
            nr_losowania = int(nr_tags[i].get_text(strip=True))
        except (ValueError, IndexError):
            continue

        raw_date = date_tags[i].get_text(strip=True)
        draw_date = None
        for fmt in ("%d-%m-%Y", "%Y-%m-%d", "%d.%m.%Y"):
            try:
                draw_date = datetime.strptime(raw_date, fmt).date()
                break
            except ValueError:
                continue

        numbers = []
        for j in range(6):
            try:
                numbers.append(int(num_tags[i * 6 + j].get_text(strip=True)))
            except (ValueError, IndexError):
                break

        if len(numbers) == 6 and draw_date:
            results.append({
                "data":         draw_date,
                "nr_losowania": nr_losowania,
                "pierwsza":     numbers[0],
                "druga":        numbers[1],
                "trzecia":      numbers[2],
                "czwarta":      numbers[3],
                "piąta":        numbers[4],
                "szósta":       numbers[5],
            })

    return results


# ============================================================
# SCRAPING JEDNEGO ROKU
# ============================================================
def scrape_year(year: int) -> list[dict]:
    url = f"{BASE_URL}?year={year}"
    print(f"  → GET {url}")
    try:
        resp = requests.get(url, headers=HEADERS, timeout=20)
        resp.raise_for_status()
    except requests.RequestException as exc:
        print(f"  ✗ Błąd pobierania {url}: {exc}")
        return []

    time.sleep(DELAY_SEC)
    records = _parse_page(resp.text)
    print(f"     Sparsowano {len(records)} losowań z roku {year}")
    return records


# ============================================================
# ZAPIS RAW TXT
# ============================================================
def save_raw_txt(records: list[dict], year: int) -> None:
    RAW_DIR.mkdir(parents=True, exist_ok=True)
    out_file = RAW_DIR / f"wyniki_scraped_{year}.txt"
    lines = []
    for r in sorted(records, key=lambda x: x["nr_losowania"], reverse=True):
        nums = " ".join(str(r[c]) for c in ["pierwsza", "druga", "trzecia", "czwarta", "piąta", "szósta"])
        lines.append(f"{r['nr_losowania']}\t{r['data']}\t{nums}")
    out_file.write_text("\n".join(lines), encoding="utf-8")
    print(f"  ✓ Zapisano raw: {out_file}")


# ============================================================
# AKTUALIZACJA SQLITE — wymagana przez lotto_generator_tkinter.py
# Schemat: draws(id INTEGER PK, draw_date TEXT, numbers TEXT JSON)
# ============================================================
def update_db(records: list[dict]) -> int:
    """
    Wstawia nowe rekordy do lotto_history.db.
    Klucz deduplikacji: draw_date (jeden losowanie = jedna data).
    Zwraca liczbę nowo wstawionych wierszy.
    """
    if not records:
        return 0

    DATA_DIR.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(str(DB_FILE))
    cur = con.cursor()

    # Utwórz tabelę jeśli nie istnieje (zgodnie ze schematem w tkinter)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS draws (
            id        INTEGER PRIMARY KEY AUTOINCREMENT,
            draw_date TEXT    UNIQUE,
            numbers   TEXT
        )
    """)
    con.commit()

    inserted = 0
    for r in records:
        draw_date_str = str(r["data"])          # 'YYYY-MM-DD'
        numbers = sorted([
            r["pierwsza"], r["druga"], r["trzecia"],
            r["czwarta"],  r["piąta"], r["szósta"]
        ])
        numbers_json = json.dumps(numbers)

        try:
            cur.execute(
                "INSERT OR IGNORE INTO draws (draw_date, numbers) VALUES (?, ?)",
                (draw_date_str, numbers_json)
            )
            if cur.rowcount:
                inserted += 1
        except sqlite3.Error as exc:
            print(f"  ⚠ DB błąd dla {draw_date_str}: {exc}")

    con.commit()
    con.close()
    print(f"  ✓ Baza {DB_FILE.name}: +{inserted} nowych rekordów")
    return inserted


# ============================================================
# AKTUALIZACJA XLSX — wymagana przez generate_lotto_stats_final.py
# ============================================================
def update_xlsx(records: list[dict]) -> None:
    """
    Aktualizuje Arkusz1 w wyniki_lotto.xlsx.
    Kolumny: data | nr_losowania | pierwsza | druga | trzecia | czwarta | piąta | szósta
    Sortowanie: malejące wg daty (df.head(50) = ostatnie 50 losowań).
    """
    if not records:
        print("  ℹ Brak rekordów do zapisania w XLSX.")
        return

    new_df = pd.DataFrame(records, columns=COLUMNS)
    new_df["data"] = pd.to_datetime(new_df["data"], errors="coerce").dt.date

    if XLSX_FILE.exists():
        try:
            existing_df = pd.read_excel(XLSX_FILE, sheet_name=SHEET, engine="openpyxl")
            if "data" in existing_df.columns:
                existing_df["data"] = pd.to_datetime(
                    existing_df["data"], errors="coerce"
                ).dt.date
        except Exception as exc:
            print(f"  ⚠ Nie mogę wczytać istniejącego XLSX ({exc}), tworzę nowy.")
            existing_df = pd.DataFrame(columns=COLUMNS)
    else:
        existing_df = pd.DataFrame(columns=COLUMNS)

    combined = pd.concat([existing_df, new_df], ignore_index=True)
    combined = combined.drop_duplicates(subset=["nr_losowania"], keep="last")
    combined = combined.sort_values("data", ascending=False).reset_index(drop=True)

    for col in ["nr_losowania", "pierwsza", "druga", "trzecia", "czwarta", "piąta", "szósta"]:
        if col in combined.columns:
            combined[col] = combined[col].astype("Int64")

    DATA_DIR.mkdir(parents=True, exist_ok=True)

    if XLSX_FILE.exists():
        with pd.ExcelWriter(
            XLSX_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace"
        ) as writer:
            combined.to_excel(writer, sheet_name=SHEET, index=False)
    else:
        with pd.ExcelWriter(XLSX_FILE, engine="openpyxl") as writer:
            combined.to_excel(writer, sheet_name=SHEET, index=False)

    print(f"  ✓ XLSX zaktualizowany: {XLSX_FILE}")
    print(f"     Łącznie rekordów w '{SHEET}': {len(combined)}")
    if not combined.empty:
        row0 = combined.iloc[0]
        print(
            f"     Najnowsze: {row0['data']}  "
            f"nr {row0['nr_losowania']}: "
            f"{row0[['pierwsza','druga','trzecia','czwarta','piąta','szósta']].tolist()}"
        )


# ============================================================
# GŁÓWNA LOGIKA
# ============================================================
def main():
    parser = argparse.ArgumentParser(
        description="Scraper wyników Lotto z megalotto.pl"
    )
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--year",      type=int,
                       help="Pobierz konkretny rok (np. 2025)")
    group.add_argument("--all-years", action="store_true",
                       help="Pobierz całą historię (od 1957 do dziś)")
    parser.add_argument("--update-xlsx", action="store_true",
                        help="Zaktualizuj data/wyniki_lotto.xlsx")
    parser.add_argument("--no-db", action="store_true",
                        help="Pomiń zapis do lotto_history.db")
    args = parser.parse_args()

    current_year = datetime.now().year

    if args.all_years:
        years = range(1957, current_year + 1)
    elif args.year:
        years = [args.year]
    else:
        years = [current_year]

    all_records: list[dict] = []

    print("=" * 60)
    print("SCRAPER MEGALOTTO.PL — START")
    print("=" * 60)

    for year in years:
        print(f"\nRok {year}:")
        records = scrape_year(year)
        if records:
            save_raw_txt(records, year)
            all_records.extend(records)

    print(f"\nŁącznie pobrano: {len(all_records)} losowań")

    if all_records:
        # Zawsze aktualizuj DB (używaną przez Tkinter)
        if not args.no_db:
            print("\nAktualizacja SQLite (lotto_history.db)...")
            update_db(all_records)

        # Opcjonalnie aktualizuj XLSX (dla statystyk)
        if args.update_xlsx:
            print("\nAktualizacja XLSX (wyniki_lotto.xlsx)...")
            update_xlsx(all_records)

    print("\n✅ Gotowe!")
    if all_records and not args.update_xlsx:
        print("   Podpowiedź: dodaj --update-xlsx żeby zaktualizować wyniki_lotto.xlsx")


if __name__ == "__main__":
    main()
