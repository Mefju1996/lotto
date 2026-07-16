#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Scrapper wyników Lotto z megalotto.pl
Źródło: https://megalotto.pl/wyniki/lotto

Pobranie wyników za dany rok (lub bieżącego) i zapisanie do:
  - data/raw/lotto/wyniki_scraped_YYYY.txt  (format: DD-MM-YYYY  N1 N2 N3 N4 N5 N6)
  - data/wyniki_lotto.xlsx                  (kompatybilny z generate_lotto_stats_final.py)

Wymagania:
    pip install requests beautifulsoup4 openpyxl pandas

Użycie:
    python scripts/scraper_megalotto.py            # pobiera bieżący rok
    python scripts/scraper_megalotto.py --year 2025
    python scripts/scraper_megalotto.py --year 2025 --update-xlsx  # dołącza do xlsx
    python scripts/scraper_megalotto.py --all-years --update-xlsx  # pełna historia
"""

import argparse
import time
import re
import sys
from datetime import datetime
from pathlib import Path

import requests
from bs4 import BeautifulSoup
import pandas as pd

# ============================================================
# KONFIGURACJA
# ============================================================
BASE_URL   = "https://megalotto.pl/wyniki/lotto/losowania-z-roku-{year}"
MAIN_URL   = "https://megalotto.pl/wyniki/lotto"
HEADERS    = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0 Safari/537.36"
    ),
    "Accept-Language": "pl-PL,pl;q=0.9",
}
DELAY_SEC  = 1.5   # przerwa między żądaniami (uprzejme scrapowanie)

ROOT_DIR   = Path(__file__).resolve().parents[1]
RAW_DIR    = ROOT_DIR / "data" / "raw" / "lotto"
XLSX_FILE  = ROOT_DIR / "data" / "wyniki_lotto.xlsx"

# Kolumny zgodne z generate_lotto_stats_final.py
COLS       = ["data", "nr_losowania", "pierwsza", "druga", "trzecia",
              "czwarta", "piąta", "szósta"]


# ============================================================
# FUNKCJE POMOCNICZE
# ============================================================

def fetch_page(url: str) -> BeautifulSoup | None:
    """Pobiera stronę i zwraca obiekt BeautifulSoup lub None przy błędzie."""
    try:
        resp = requests.get(url, headers=HEADERS, timeout=15)
        resp.raise_for_status()
        resp.encoding = "utf-8"
        return BeautifulSoup(resp.text, "html.parser")
    except requests.RequestException as e:
        print(f"   ✗ Błąd pobierania {url}: {e}", file=sys.stderr)
        return None


def parse_drawings(soup: BeautifulSoup) -> list[dict]:
    """
    Parsuje losowania ze strony megalotto.
    Struktura HTML:
      <div class="lista_ostatnich_losowan">
        <ul>
          <li class="nr_in_list">7378. </li>
          <li class="date_in_list">14-07-2026</li>
          <li class="numbers_in_list">7 </li>
          ... (6 liczb)
        </ul>
      ...
    """
    results = []
    container = soup.find("div", class_="lista_ostatnich_losowan")
    if not container:
        return results

    for ul in container.find_all("ul"):
        items = ul.find_all("li")

        nr_tag    = ul.find("li", class_="nr_in_list")
        date_tag  = ul.find("li", class_="date_in_list")
        num_tags  = ul.find_all("li", class_="numbers_in_list")

        if not (nr_tag and date_tag and len(num_tags) == 6):
            continue

        nr_text = re.sub(r"[^\d]", "", nr_tag.get_text())
        date_text = date_tag.get_text(strip=True)  # DD-MM-YYYY

        try:
            nr = int(nr_text) if nr_text else None
            date_obj = datetime.strptime(date_text, "%d-%m-%Y").date()
            numbers = [int(n.get_text(strip=True)) for n in num_tags]
        except (ValueError, AttributeError):
            continue

        results.append({
            "data":        date_obj,
            "nr_losowania": nr,
            "pierwsza":    numbers[0],
            "druga":       numbers[1],
            "trzecia":     numbers[2],
            "czwarta":     numbers[3],
            "piąta":       numbers[4],
            "szósta":      numbers[5],
        })

    return results


def get_available_years(soup: BeautifulSoup) -> list[int]:
    """Odczytuje listę dostępnych lat z linków na stronie głównej."""
    years = []
    for a in soup.find_all("a", href=re.compile(r"losowania-z-roku-(\d{4})")):
        m = re.search(r"(\d{4})", a["href"])
        if m:
            years.append(int(m.group(1)))
    return sorted(set(years))


def scrape_year(year: int) -> list[dict]:
    """Scrapuje wyniki dla jednego roku."""
    url = BASE_URL.format(year=year)
    print(f"   → Pobieram rok {year}: {url}")
    soup = fetch_page(url)
    if soup is None:
        return []
    drawings = parse_drawings(soup)
    print(f"     ✓ Znaleziono {len(drawings)} losowań")
    time.sleep(DELAY_SEC)
    return drawings


def save_txt(drawings: list[dict], year: int) -> Path:
    """Zapisuje wyniki do pliku TXT (format kompatybilny z projektem)."""
    RAW_DIR.mkdir(parents=True, exist_ok=True)
    out_path = RAW_DIR / f"wyniki_scraped_{year}.txt"
    with open(out_path, "w", encoding="utf-8") as f:
        for d in sorted(drawings, key=lambda x: x["data"]):
            nums = " ".join(str(d[k]) for k in
                            ["pierwsza", "druga", "trzecia", "czwarta", "piąta", "szósta"])
            f.write(f"{d['data'].strftime('%d-%m-%Y')}  {nums}\n")
    print(f"   ✓ Zapisano TXT: {out_path}")
    return out_path


def save_or_update_xlsx(drawings: list[dict]) -> None:
    """
    Dołącza nowe losowania do data/wyniki_lotto.xlsx.
    Jeśli plik nie istnieje, tworzy go od zera.
    Duplikaty (wg nr_losowania lub daty) są pomijane.
    """
    XLSX_FILE.parent.mkdir(parents=True, exist_ok=True)
    new_df = pd.DataFrame(drawings, columns=COLS)
    new_df["data"] = pd.to_datetime(new_df["data"])

    if XLSX_FILE.exists():
        try:
            existing = pd.read_excel(XLSX_FILE, engine="openpyxl", sheet_name="Arkusz1")
            existing["data"] = pd.to_datetime(existing["data"])
            combined = pd.concat([new_df, existing], ignore_index=True)
        except Exception as e:
            print(f"   ⚠ Nie udało się wczytać istniejącego XLSX: {e}")
            combined = new_df
    else:
        combined = new_df

    # Usuń duplikaty - priorytet dla nowych rekordów (head)
    if "nr_losowania" in combined.columns:
        combined = combined.drop_duplicates(subset="nr_losowania", keep="first")
    else:
        combined = combined.drop_duplicates(subset="data", keep="first")

    combined = combined.sort_values("data", ascending=False).reset_index(drop=True)

    with pd.ExcelWriter(XLSX_FILE, engine="openpyxl") as writer:
        combined.to_excel(writer, sheet_name="Arkusz1", index=False)

    print(f"   ✓ Zapisano XLSX: {XLSX_FILE} ({len(combined)} losowań łącznie)")


# ============================================================
# MAIN
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description="Scrapper wyników Lotto z megalotto.pl"
    )
    parser.add_argument(
        "--year", type=int, default=datetime.now().year,
        help="Rok do pobrania (domyślnie: bieżący)"
    )
    parser.add_argument(
        "--all-years", action="store_true",
        help="Pobierz wszystkie dostępne lata (ostrzeżenie: dużo requestów)"
    )
    parser.add_argument(
        "--update-xlsx", action="store_true",
        help="Dołącz wyniki do data/wyniki_lotto.xlsx"
    )
    args = parser.parse_args()

    print("=" * 60)
    print("SCRAPPER MEGALOTTO.PL - Wyniki Lotto")
    print(f"Data uruchomienia: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)

    all_drawings = []

    if args.all_years:
        print("\n1. Pobieranie listy dostępnych lat...")
        soup = fetch_page(MAIN_URL)
        if soup is None:
            print("   ✗ Nie udało się pobrać strony głównej.", file=sys.stderr)
            sys.exit(1)
        years = get_available_years(soup)
        print(f"   ✓ Znaleziono lata: {years[0]}–{years[-1]} ({len(years)} lat)")

        print("\n2. Scrapowanie wyników...")
        for year in years:
            drawings = scrape_year(year)
            all_drawings.extend(drawings)
            save_txt(drawings, year)
    else:
        year = args.year
        print(f"\n1. Scrapowanie wyników dla roku {year}...")
        # Bieżący rok może być na stronie głównej
        url = BASE_URL.format(year=year) if year != datetime.now().year else MAIN_URL
        soup = fetch_page(url)
        if soup is None:
            sys.exit(1)
        drawings = parse_drawings(soup)
        if not drawings and year == datetime.now().year:
            # fallback na stronę roczną
            drawings = scrape_year(year)
        else:
            print(f"   ✓ Znaleziono {len(drawings)} losowań (rok {year})")
        all_drawings = drawings
        save_txt(drawings, year)

    if not all_drawings:
        print("\n✗ Brak danych do zapisania.")
        sys.exit(1)

    if args.update_xlsx:
        print("\n3. Aktualizacja pliku XLSX...")
        save_or_update_xlsx(all_drawings)
    else:
        print(
            "\n💡 Wskazówka: dodaj --update-xlsx żeby zaktualizować "
            "data/wyniki_lotto.xlsx (potrzebne dla generate_lotto_stats_final.py)"
        )

    print("\n" + "=" * 60)
    print(f"✓ Gotowe! Pobrano {len(all_drawings)} losowań.")
    print("=" * 60)


if __name__ == "__main__":
    main()
