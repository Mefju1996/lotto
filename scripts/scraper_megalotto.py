#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Scraper wyników Lotto z https://megalotto.pl/wyniki/lotto
Zsynchronizowany z generate_lotto_stats_final.py

Generuje:
  - data/raw/lotto/wyniki_scraped_YYYY.txt  (surowe wyniki)
  - data/wyniki_lotto.xlsx / Arkusz1         (gotowe dla generate_lotto_stats_final.py)

Kolumny xlsx: data | nr_losowania | pierwsza | druga | trzecia | czwarta | piąta | szósta

Użycie:
    # Bieżący rok + aktualizacja xlsx
    python scripts/scraper_megalotto.py --update-xlsx

    # Konkretny rok
    python scripts/scraper_megalotto.py --year 2025 --update-xlsx

    # Cała historia (od 1957)
    python scripts/scraper_megalotto.py --all-years --update-xlsx

Wymagania:
    pip install requests beautifulsoup4 openpyxl pandas
"""

import argparse
import time
from datetime import datetime
from pathlib import Path

import pandas as pd
import requests
from bs4 import BeautifulSoup

# ============================================================
# ŚCIEŻKI — względem katalogu głównego repo
# ============================================================
ROOT_DIR = Path(__file__).resolve().parents[1]
DATA_DIR  = ROOT_DIR / "data"
RAW_DIR   = DATA_DIR / "raw" / "lotto"
XLSX_FILE = DATA_DIR / "wyniki_lotto.xlsx"
SHEET     = "Arkusz1"

# Kolumny wymagane przez generate_lotto_stats_final.py
COLUMNS = ["data", "nr_losowania", "pierwsza", "druga", "trzecia", "czwarta", "piąta", "szósta"]

BASE_URL  = "https://megalotto.pl/wyniki/lotto"
HEADERS   = {"User-Agent": "Mozilla/5.0 (compatible; LottoScraper/1.0)"}
DELAY_SEC = 1.5   # grzeczne opóźnienie między requestami


# ============================================================
# PARSOWANIE JEDNEJ STRONY
# ============================================================
def _parse_page(html: str) -> list[dict]:
    """
    Parsuje HTML strony wyników megalotto.pl i zwraca listę słowników
    z kluczami zgodnymi z kolumnami XLSX.
    """
    soup = BeautifulSoup(html, "html.parser")
    results = []

    # Każde losowanie jest opakowane w <ul class="lista_wynikow"> lub podobny kontener;
    # konkretne elementy to li z klasami nr_in_list, date_in_list, numbers_in_list
    nr_tags   = soup.find_all("li", class_="nr_in_list")
    date_tags = soup.find_all("li", class_="date_in_list")
    num_tags  = soup.find_all("li", class_="numbers_in_list")

    # Liczby idą po 6 na losowanie
    if len(num_tags) % 6 != 0:
        return results

    n_draws = min(len(nr_tags), len(date_tags), len(num_tags) // 6)

    for i in range(n_draws):
        try:
            nr_losowania = int(nr_tags[i].get_text(strip=True))
        except (ValueError, IndexError):
            continue

        raw_date = date_tags[i].get_text(strip=True)
        try:
            draw_date = datetime.strptime(raw_date, "%d-%m-%Y").date()
        except ValueError:
            try:
                draw_date = datetime.strptime(raw_date, "%Y-%m-%d").date()
            except ValueError:
                draw_date = None

        numbers = []
        for j in range(6):
            try:
                numbers.append(int(num_tags[i * 6 + j].get_text(strip=True)))
            except (ValueError, IndexError):
                break

        if len(numbers) == 6:
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
    """
    Pobiera wszystkie wyniki dla danego roku.
    megalotto.pl obsługuje parametr ?year=YYYY w URL.
    """
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
# ZAPIS DO PLIKU TXT (raw)
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
# AKTUALIZACJA XLSX
# ============================================================
def update_xlsx(new_records: list[dict]) -> None:
    """
    Dodaje nowe rekordy do XLSX, deduplikując po nr_losowania.
    Arkusz jest posortowany malejąco wg daty (najnowsze u góry),
    co jest wymagane przez generate_lotto_stats_final.py (df.head(50) = ostatnie 50).
    """
    if not new_records:
        print("  ℹ Brak nowych rekordów do zapisania.")
        return

    new_df = pd.DataFrame(new_records, columns=COLUMNS)

    if XLSX_FILE.exists():
        try:
            existing_df = pd.read_excel(XLSX_FILE, sheet_name=SHEET, engine="openpyxl")
            # Normalizuj datę istniejących danych
            if "data" in existing_df.columns:
                existing_df["data"] = pd.to_datetime(existing_df["data"], errors="coerce").dt.date
        except Exception as exc:
            print(f"  ⚠ Nie można wczytać istniejącego XLSX ({exc}), tworzę nowy.")
            existing_df = pd.DataFrame(columns=COLUMNS)
    else:
        existing_df = pd.DataFrame(columns=COLUMNS)

    # Normalizuj datę nowych danych
    new_df["data"] = pd.to_datetime(new_df["data"], errors="coerce").dt.date

    combined = pd.concat([existing_df, new_df], ignore_index=True)
    combined = combined.drop_duplicates(subset=["nr_losowania"], keep="last")

    # Sortuj malejąco wg daty — generate_lotto_stats_final używa df.head(50) jako "ostatnie 50"
    combined = combined.sort_values("data", ascending=False).reset_index(drop=True)

    # Upewnij się, że liczby są int (nie float) — wymagane przez stats_final
    for col in ["nr_losowania", "pierwsza", "druga", "trzecia", "czwarta", "piąta", "szósta"]:
        if col in combined.columns:
            combined[col] = combined[col].astype("Int64")

    DATA_DIR.mkdir(parents=True, exist_ok=True)

    # Zachowaj pozostałe arkusze jeśli plik istnieje
    if XLSX_FILE.exists():
        with pd.ExcelWriter(XLSX_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            combined.to_excel(writer, sheet_name=SHEET, index=False)
    else:
        with pd.ExcelWriter(XLSX_FILE, engine="openpyxl") as writer:
            combined.to_excel(writer, sheet_name=SHEET, index=False)

    print(f"  ✓ Zaktualizowano XLSX: {XLSX_FILE}")
    print(f"     Łącznie rekordów w arkuszu '{SHEET}': {len(combined)}")
    if not combined.empty:
        print(f"     Najnowsze: {combined.iloc[0]['data']}  "
              f"nr {combined.iloc[0]['nr_losowania']}: "
              f"{combined.iloc[0][['pierwsza','druga','trzecia','czwarta','piąta','szósta']].tolist()}")


# ============================================================
# GŁÓWNA LOGIKA
# ============================================================
def main():
    parser = argparse.ArgumentParser(
        description="Scraper wyników Lotto z megalotto.pl"
    )
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--year",      type=int, help="Pobierz konkretny rok (np. 2025)")
    group.add_argument("--all-years", action="store_true",
                       help="Pobierz całą historię (od 1957 do dziś)")
    parser.add_argument("--update-xlsx", action="store_true",
                        help="Zaktualizuj data/wyniki_lotto.xlsx po pobraniu")
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

    if args.update_xlsx and all_records:
        print("\nAktualizacja XLSX...")
        update_xlsx(all_records)

    print("\n✅ Gotowe!")
    if all_records and not args.update_xlsx:
        print("   Podpowiedź: uruchom z --update-xlsx żeby zaktualizować wyniki_lotto.xlsx")


if __name__ == "__main__":
    main()
