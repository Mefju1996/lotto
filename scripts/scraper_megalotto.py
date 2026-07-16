#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Scraper wyników Lotto z https://megalotto.pl/wyniki/lotto
Fallback: lotto.pl (oficjalna strona) gdy megalotto zmieni strukturę HTML.

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
import logging
import sqlite3
import time
from datetime import datetime
from pathlib import Path

import pandas as pd
import requests
from bs4 import BeautifulSoup

# ============================================================
# LOGGING
# ============================================================
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)

# ============================================================
# ŚCIEŻKI
# ============================================================
ROOT_DIR  = Path(__file__).resolve().parents[1]
DATA_DIR  = ROOT_DIR / "data"
RAW_DIR   = DATA_DIR / "raw" / "lotto"
XLSX_FILE = DATA_DIR / "wyniki_lotto.xlsx"
DB_FILE   = DATA_DIR / "lotto_history.db"
SHEET     = "Arkusz1"

# Kolumny wymagane przez generate_lotto_stats_final.py
COLUMNS = ["data", "nr_losowania", "pierwsza", "druga", "trzecia", "czwarta", "piąta", "szósta"]

BASE_URL  = "https://megalotto.pl/wyniki/lotto"
HEADERS   = {"User-Agent": "Mozilla/5.0 (compatible; LottoScraper/1.0)"}
DELAY_SEC = 1.5


# ============================================================
# PARSOWANIE — MEGALOTTO (pierwotna metoda)
# ============================================================
def _parse_megalotto(html: str, year: int) -> list[dict]:
    """Parsuje HTML z megalotto.pl. Zwraca listę słowników (może być pusta)."""
    soup = BeautifulSoup(html, "html.parser")
    results = []

    nr_tags   = soup.find_all("li", class_="nr_in_list")
    date_tags = soup.find_all("li", class_="date_in_list")
    num_tags  = soup.find_all("li", class_="numbers_in_list")

    logger.debug(
        "Megalotto – tagi: nr=%d, date=%d, num=%d",
        len(nr_tags), len(date_tags), len(num_tags),
    )

    if not nr_tags or not date_tags or not num_tags:
        logger.warning(
            "Megalotto: nie znaleziono tagów dla roku %d "
            "(nr=%d, date=%d, num=%d) – być może zmieniono strukturę HTML.",
            year, len(nr_tags), len(date_tags), len(num_tags),
        )
        return results

    if len(num_tags) % 6 != 0:
        logger.warning(
            "Megalotto: liczba tagów liczbowych (%d) nie jest wielokrotnością 6 dla roku %d.",
            len(num_tags), year,
        )
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
# PARSOWANIE — LOTTO.PL (fallback)
# ============================================================
def _parse_lotto_pl(html: str, year: int) -> list[dict]:
    """
    Fallback: parsuje HTML z lotto.pl/wyniki-i-wygrane/lotto.
    Strona używa innej struktury (tabela lub data-atrybuty).
    Zwraca listę słowników lub [] gdy struktura nieznana.
    """
    soup = BeautifulSoup(html, "html.parser")
    results = []

    # lotto.pl: wiersze tabeli z wynikami
    rows = soup.select("table.results-table tr[data-draw-date]")
    if not rows:
        # Alternatywna struktura: divs z klasą draw-result
        rows = soup.select("div.draw-result")

    if not rows:
        logger.warning("Lotto.pl fallback: nie rozpoznano struktury HTML dla roku %d.", year)
        return results

    for row in rows:
        try:
            raw_date = row.get("data-draw-date") or row.select_one(".draw-date").get_text(strip=True)
            draw_date = None
            for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d.%m.%Y"):
                try:
                    draw_date = datetime.strptime(raw_date, fmt).date()
                    break
                except ValueError:
                    continue
            if not draw_date:
                continue

            num_els = row.select(".ball, .number, td.draw-number")
            numbers = [int(el.get_text(strip=True)) for el in num_els if el.get_text(strip=True).isdigit()]
            if len(numbers) < 6:
                continue
            numbers = numbers[:6]

            nr_el = row.select_one(".draw-id, td.draw-number-id")
            nr_losowania = int(nr_el.get_text(strip=True)) if nr_el else 0

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
        except Exception as exc:
            logger.debug("Lotto.pl: pominięto wiersz (%s)", exc)
            continue

    return results


# ============================================================
# SCRAPING JEDNEGO ROKU (z fallbackiem)
# ============================================================
def scrape_year(year: int) -> list[dict]:
    """
    Pobiera wyniki dla danego roku.
    Kolejność prób:
      1. megalotto.pl  (pierwotna)
      2. lotto.pl      (fallback gdy megalotto zwróci 0 rekordów)
    """
    # --- Próba 1: megalotto.pl ---
    url_megalotto = f"{BASE_URL}?year={year}"
    logger.info("→ GET %s", url_megalotto)
    records: list[dict] = []

    try:
        resp = requests.get(url_megalotto, headers=HEADERS, timeout=20)
        resp.raise_for_status()
        records = _parse_megalotto(resp.text, year)
        logger.info("Megalotto: sparsowano %d losowań z roku %d", len(records), year)
    except requests.RequestException as exc:
        logger.error("Błąd pobierania megalotto (%s): %s", url_megalotto, exc)

    time.sleep(DELAY_SEC)

    if records:
        return records

    # --- Próba 2: lotto.pl (fallback) ---
    url_lotto = f"https://www.lotto.pl/wyniki-i-wygrane/lotto?date={year}"
    logger.warning("Megalotto zwróciło 0 rekordów — próba fallbacku: %s", url_lotto)
    try:
        resp2 = requests.get(url_lotto, headers=HEADERS, timeout=20)
        resp2.raise_for_status()
        records = _parse_lotto_pl(resp2.text, year)
        logger.info("Lotto.pl fallback: sparsowano %d losowań z roku %d", len(records), year)
    except requests.RequestException as exc:
        logger.error("Błąd pobierania lotto.pl (%s): %s", url_lotto, exc)

    time.sleep(DELAY_SEC)

    if not records:
        logger.error(
            "Brak wyników dla roku %d — obie źródła zawiodły. "
            "Sprawdź połączenie sieciowe lub strukturę HTML źródłowych stron.",
            year,
        )

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
    logger.info("Zapisano raw: %s", out_file)


# ============================================================
# AKTUALIZACJA SQLITE
# ============================================================
def update_db(records: list[dict]) -> int:
    if not records:
        return 0

    DATA_DIR.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(str(DB_FILE))
    cur = con.cursor()

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
        draw_date_str = str(r["data"])
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
            logger.warning("DB błąd dla %s: %s", draw_date_str, exc)

    con.commit()
    con.close()
    logger.info("Baza %s: +%d nowych rekordów", DB_FILE.name, inserted)
    return inserted


# ============================================================
# AKTUALIZACJA XLSX
# ============================================================
def update_xlsx(records: list[dict]) -> None:
    if not records:
        logger.info("Brak rekordów do zapisania w XLSX.")
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
            logger.warning("Nie mogę wczytać istniejącego XLSX (%s), tworzę nowy.", exc)
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

    logger.info("XLSX zaktualizowany: %s", XLSX_FILE)
    logger.info("Łącznie rekordów w '%s': %d", SHEET, len(combined))
    if not combined.empty:
        row0 = combined.iloc[0]
        logger.info(
            "Najnowsze: %s  nr %s: %s",
            row0["data"], row0["nr_losowania"],
            row0[["pierwsza", "druga", "trzecia", "czwarta", "piąta", "szósta"]].tolist(),
        )


# ============================================================
# GŁÓWNA LOGIKA
# ============================================================
def main():
    parser = argparse.ArgumentParser(
        description="Scraper wyników Lotto z megalotto.pl (fallback: lotto.pl)"
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
    parser.add_argument("--verbose", "-v", action="store_true",
                        help="Pokaż szczegółowe logi debugowania")
    args = parser.parse_args()

    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)

    current_year = datetime.now().year

    if args.all_years:
        years = range(1957, current_year + 1)
    elif args.year:
        years = [args.year]
    else:
        years = [current_year]

    all_records: list[dict] = []

    logger.info("=" * 50)
    logger.info("SCRAPER MEGALOTTO.PL — START (fallback: lotto.pl)")
    logger.info("=" * 50)

    for year in years:
        logger.info("Rok %d:", year)
        records = scrape_year(year)
        if records:
            save_raw_txt(records, year)
            all_records.extend(records)

    logger.info("Łącznie pobrano: %d losowań", len(all_records))

    if all_records:
        if not args.no_db:
            update_db(all_records)
        if args.update_xlsx:
            update_xlsx(all_records)
    else:
        logger.error(
            "Nie pobrano żadnych danych. "
            "Sprawdź połączenie sieciowe lub strukturę HTML źródłowych stron."
        )

    logger.info("KONIEC")


if __name__ == "__main__":
    main()
