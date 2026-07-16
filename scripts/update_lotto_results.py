#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Skrypt aktualizacji wyników Lotto — wywoływany przez przycisk '⟳ Aktualizuj'
w lotto_generator_tkinter.py.

Pobiera bieżący rok z megalotto.pl i zapisuje dane do:
  - data/lotto_history.db  (tabela draws — sprawdzanie historii w Tkinter)
  - data/wyniki_lotto.xlsx (Arkusz1    — generowanie statystyk)

Nie wymaga żadnych argumentów.
"""

from pathlib import Path
import sys

# Dodaj katalog scripts do ścieżki importu
sys.path.insert(0, str(Path(__file__).parent))

from scraper_megalotto import scrape_year, update_db, update_xlsx
from datetime import datetime

def main():
    current_year = datetime.now().year
    print(f"Aktualizacja wyników Lotto ({current_year})...")

    records = scrape_year(current_year)
    if records:
        update_db(records)
        update_xlsx(records)
        print(f"✅ Zaktualizowano {len(records)} losowań.")
    else:
        print("⚠ Brak nowych wyników do zaktualizowania.")

if __name__ == "__main__":
    main()
