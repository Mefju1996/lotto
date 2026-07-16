#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generator statystyk Lotto - WERSJA POPRAWIONA + ZSYNCHRONIZOWANA ZE SCRAPEREM
Analizuje bazę historycznych wyników i generuje raport XLSX z wykresami PNG

Użycie:
    python scripts/generate_lotto_stats_final.py

Wymagania:
    - data/wyniki_lotto.xlsx (wypełniany przez scripts/scraper_megalotto.py)
    - pandas, matplotlib, openpyxl, numpy
"""

import pandas as pd
import matplotlib.pyplot as plt
from collections import Counter
from itertools import combinations
import numpy as np
import os
from datetime import datetime, date, timedelta
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT_DIR / 'data'

print("="*70)
print("GENERATOR STATYSTYK LOTTO - START")
print("="*70)

# ============================================================
# WCZYTANIE DANYCH
# ============================================================
print("\n1. Wczytywanie danych z pliku wyniki_lotto.xlsx...")
file = DATA_DIR / 'wyniki_lotto.xlsx'
df = pd.read_excel(file, engine='openpyxl', sheet_name='Arkusz1')

# FIX #1: Normalizuj kolumnę 'data' do datetime (scraper zapisuje jako date/datetime)
if 'data' in df.columns:
    df['data'] = pd.to_datetime(df['data'], errors='coerce')

# Upewnij się, że dane posortowane od najnowszego (scraper zapisuje malejąco, ale dla pewności)
if 'data' in df.columns and df['data'].notna().any():
    df = df.sort_values('data', ascending=False).reset_index(drop=True)

print(f"   ✓ Wczytano {len(df)} losowań")

liczby_cols = ['pierwsza', 'druga', 'trzecia', 'czwarta', 'piąta', 'szósta']
wszystkie_liczby = df[liczby_cols].values.flatten()

print(f"   ✓ Najnowsze losowanie: {df.iloc[0][liczby_cols].tolist()}")
print(f"   ✓ Najstarsze losowanie w bazie: {df.iloc[-1][liczby_cols].tolist()}")

# ============================================================
# KONFIGURACJA KATALOGU WYJŚCIOWEGO
# ============================================================
timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
input_dir = ROOT_DIR / 'analysis' / f'statystyki_{timestamp}'
os.makedirs(input_dir, exist_ok=True)
output_dir = str(input_dir)
print(f"\n   ✓ Katalog wyjściowy: {output_dir}")

def out(filename):
    """Zwraca pełną ścieżkę do pliku w katalogu wyjściowym."""
    return os.path.join(output_dir, filename)

# ============================================================
# OBLICZENIE KOLUMN POMOCNICZYCH (WAŻNE: PRZED df_recent!)
# ============================================================
def count_parzyste(row):
    return sum(1 for x in row if x % 2 == 0)

def count_niskie(row):
    return sum(1 for x in row if x <= 24)

df['spread'] = df['szósta'] - df['pierwsza']
df['parzyste'] = df[liczby_cols].apply(count_parzyste, axis=1)
df['nieparzyste'] = 6 - df['parzyste']
df['niskie'] = df[liczby_cols].apply(count_niskie, axis=1)
df['wysokie'] = 6 - df['niskie']

# FIX #2: Kolumna 'Suma' musi być obliczona TUTAJ, przed df_recent i przed sekcją sumy
df['Suma'] = df[liczby_cols].sum(axis=1)

df_recent = df.head(50).copy()
wszystkie_liczby_recent = df_recent[liczby_cols].values.flatten()

# ============================================================
# ZAKŁADKA 1: CZĘSTOTLIWOŚĆ LICZB (CAŁA HISTORIA)
# ============================================================
print("\n2. Generowanie statystyk częstotliwości...")
counter = Counter(wszystkie_liczby)
freq_df = pd.DataFrame([
    {'Liczba': int(num), 'Wystąpienia': count, 'Procent': round(count/len(df)*100, 2)}
    for num, count in sorted(counter.items())
])

plt.figure(figsize=(14, 6))
plt.bar(freq_df['Liczba'], freq_df['Wystąpienia'], color='steelblue', alpha=0.7)
plt.xlabel('Liczba', fontsize=12)
plt.ylabel('Liczba wystąpień', fontsize=12)
plt.title(f'Częstotliwość wylosowania liczb 1-49 ({len(df)} losowań)', fontsize=14, fontweight='bold')
plt.xticks(range(1, 50))
plt.grid(axis='y', alpha=0.3)
plt.tight_layout()
plt.savefig(out('wykres_czestotliwosc.png'), dpi=150, bbox_inches='tight')
plt.close()
print("   ✓ Wygenerowano wykres_czestotliwosc.png")

# ============================================================
# ZAKŁADKA 2 & 3: HOT & COLD NUMBERS
# ============================================================
top_20 = freq_df.nlargest(20, 'Wystąpienia').reset_index(drop=True)
top_20.index += 1
top_20.index.name = 'Ranking'

bottom_20 = freq_df.nsmallest(20, 'Wystąpienia').reset_index(drop=True)
bottom_20.index += 1
bottom_20.index.name = 'Ranking'

fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))

top_10 = freq_df.nlargest(10, 'Wystąpienia')
ax1.barh(top_10['Liczba'].astype(str), top_10['Wystąpienia'], color='crimson', alpha=0.7)
ax1.set_xlabel('Wystąpienia', fontsize=11)
ax1.set_ylabel('Liczba', fontsize=11)
ax1.set_title('TOP 10 Hot Numbers', fontsize=12, fontweight='bold')
ax1.invert_yaxis()
ax1.grid(axis='x', alpha=0.3)

bottom_10 = freq_df.nsmallest(10, 'Wystąpienia')
ax2.barh(bottom_10['Liczba'].astype(str), bottom_10['Wystąpienia'], color='dodgerblue', alpha=0.7)
ax2.set_xlabel('Wystąpienia', fontsize=11)
ax2.set_ylabel('Liczba', fontsize=11)
ax2.set_title('TOP 10 Cold Numbers', fontsize=12, fontweight='bold')
ax2.invert_yaxis()
ax2.grid(axis='x', alpha=0.3)

plt.tight_layout()
plt.savefig(out('wykres_hot_cold.png'), dpi=150, bbox_inches='tight')
plt.close()
print("   ✓ Wygenerowano wykres_hot_cold.png")

# ============================================================
# ZAKŁADKA 4: ROZKŁAD PARZYSTOŚĆ/NIEPARZYSTOŚĆ
# ============================================================
print("\n3. Generowanie rozkładów parzystość/nieparzystość...")

rozklad_parz = df.groupby('parzyste').size().reset_index(name='Liczba_losowan')
rozklad_parz['Rozkład'] = rozklad_parz['parzyste'].astype(str) + 'P-' + (6 - rozklad_parz['parzyste']).astype(str) + 'N'
rozklad_parz['Procent'] = round(rozklad_parz['Liczba_losowan'] / len(df) * 100, 2)

plt.figure(figsize=(10, 6))
plt.bar(rozklad_parz['Rozkład'], rozklad_parz['Liczba_losowan'], color='mediumseagreen', alpha=0.7)
plt.xlabel('Rozkład (P=Parzyste, N=Nieparzyste)', fontsize=12)
plt.ylabel('Liczba losowań', fontsize=12)
plt.title('Rozkład parzystości w losowaniach', fontsize=14, fontweight='bold')
for i, row in rozklad_parz.iterrows():
    plt.text(i, row['Liczba_losowan'] + 50, f"{row['Procent']}%", ha='center', fontsize=10)
plt.grid(axis='y', alpha=0.3)
plt.tight_layout()
plt.savefig(out('wykres_parzyste.png'), dpi=150, bbox_inches='tight')
plt.close()
print("   ✓ Wygenerowano wykres_parzyste.png")

# ============================================================
# ZAKŁADKA 5: ROZKŁAD NISKIE/WYSOKIE (1-24 vs 25-49)
# ============================================================
print("\n4. Generowanie rozkładów niskie/wysokie...")

rozklad_nh = df.groupby('niskie').size().reset_index(name='Liczba_losowan')
rozklad_nh['Rozkład'] = rozklad_nh['niskie'].astype(str) + 'L-' + (6 - rozklad_nh['niskie']).astype(str) + 'H'
rozklad_nh['Procent'] = round(rozklad_nh['Liczba_losowan'] / len(df) * 100, 2)

plt.figure(figsize=(10, 6))
plt.bar(rozklad_nh['Rozkład'], rozklad_nh['Liczba_losowan'], color='darkorange', alpha=0.7)
plt.xlabel('Rozkład (L=Niskie 1-24, H=Wysokie 25-49)', fontsize=12)
plt.ylabel('Liczba losowań', fontsize=12)
plt.title('Rozkład niskie/wysokie w losowaniach', fontsize=14, fontweight='bold')
for i, row in rozklad_nh.iterrows():
    plt.text(i, row['Liczba_losowan'] + 50, f"{row['Procent']}%", ha='center', fontsize=10)
plt.grid(axis='y', alpha=0.3)
plt.tight_layout()
plt.savefig(out('wykres_niskie_wysokie.png'), dpi=150, bbox_inches='tight')
plt.close()
print("   ✓ Wygenerowano wykres_niskie_wysokie.png")

# ============================================================
# ZAKŁADKA 6: NAJCZĘSTSZE PARY LICZB
# ============================================================
print("\n5. Generowanie statystyk par liczb...")
pary = []
for _, row in df[liczby_cols].iterrows():
    nums = sorted(row.values)
    for para in combinations(nums, 2):
        pary.append(f"{int(para[0])}-{int(para[1])}")

counter_pary = Counter(pary)
pary_df = pd.DataFrame([
    {'Para': para, 'Wystąpienia': count}
    for para, count in counter_pary.most_common(50)
])
pary_df.index += 1
pary_df.index.name = 'Ranking'

top_15_pary = pary_df.head(15)
plt.figure(figsize=(12, 7))
plt.barh(top_15_pary['Para'], top_15_pary['Wystąpienia'], color='purple', alpha=0.6)
plt.xlabel('Wystąpienia', fontsize=12)
plt.ylabel('Para liczb', fontsize=12)
plt.title('TOP 15 najczęstszych par liczb', fontsize=12, fontweight='bold')
plt.gca().invert_yaxis()
plt.grid(axis='x', alpha=0.3)
plt.tight_layout()
plt.savefig(out('wykres_pary.png'), dpi=150, bbox_inches='tight')
plt.close()
print("   ✓ Wygenerowano wykres_pary.png")

# ============================================================
# ZAKŁADKA 7 & 8: SPREAD (rozpiętość)
# ============================================================
print("\n6. Generowanie statystyk spreadu...")
spread_stats = df['spread'].describe()

spread_grouped = df.groupby('spread').size().reset_index(name='Liczba_losowan')
spread_grouped['Procent'] = round(spread_grouped['Liczba_losowan'] / len(df) * 100, 2)

plt.figure(figsize=(12, 6))
plt.hist(df['spread'], bins=40, color='teal', alpha=0.7, edgecolor='black')
plt.xlabel('Spread (różnica między max a min)', fontsize=12)
plt.ylabel('Liczba losowań', fontsize=12)
plt.title(f'Rozkład spreadu (średnia: {spread_stats["mean"]:.1f})', fontsize=14, fontweight='bold')
plt.axvline(spread_stats['mean'], color='red', linestyle='--', linewidth=2, label=f'Średnia: {spread_stats["mean"]:.1f}')
plt.legend()
plt.grid(axis='y', alpha=0.3)
plt.tight_layout()
plt.savefig(out('wykres_spread.png'), dpi=150, bbox_inches='tight')
plt.close()
print("   ✓ Wygenerowano wykres_spread.png")

spread_summary = pd.DataFrame({
    'Statystyka': ['Min', 'Max', 'Średnia', 'Mediana', '25% kwantyl', '75% kwantyl'],
    'Wartość': [
        spread_stats['min'],
        spread_stats['max'],
        round(spread_stats['mean'], 2),
        spread_stats['50%'],
        spread_stats['25%'],
        spread_stats['75%']
    ]
})

spread_detail = spread_grouped.sort_values('Liczba_losowan', ascending=False).head(30).reset_index(drop=True)
spread_detail.index += 1
spread_detail.index.name = 'Ranking'

# ============================================================
# ZAKŁADKA 9 & 10: SUMA LICZB
# ============================================================
print("\n7. Generowanie statystyk sum...")
# FIX #2: df['Suma'] już obliczona wcześniej — używamy bezpośrednio
suma_grouped = df.groupby('Suma').size().reset_index(name='Liczba_losowan')
suma_grouped['Procent'] = round(suma_grouped['Liczba_losowan'] / len(df) * 100, 2)

plt.figure(figsize=(14, 6))
plt.hist(df['Suma'], bins=50, color='navy', alpha=0.6, edgecolor='black')
plt.xlabel('Suma 6 liczb', fontsize=12)
plt.ylabel('Liczba losowań', fontsize=12)
plt.title(f'Rozkład sum (średnia: {df["Suma"].mean():.1f})', fontsize=14, fontweight='bold')
plt.axvline(df['Suma'].mean(), color='red', linestyle='--', linewidth=2, label=f'Średnia: {df["Suma"].mean():.1f}')
plt.axvline(130, color='green', linestyle=':', linewidth=2, label='Zakres algorytmu: 130-205')
plt.axvline(205, color='green', linestyle=':', linewidth=2)
plt.legend()
plt.grid(axis='y', alpha=0.3)
plt.tight_layout()
plt.savefig(out('wykres_suma.png'), dpi=150, bbox_inches='tight')
plt.close()
print("   ✓ Wygenerowano wykres_suma.png")

suma_top = suma_grouped.sort_values('Liczba_losowan', ascending=False).head(50).reset_index(drop=True)
suma_top.index += 1
suma_top.index.name = 'Ranking'

# ============================================================
# ZAKŁADKA 10: WALIDACJA REGUŁ ALGORYTMU
# ============================================================
print("\n8. Walidacja reguł algorytmu...")
def check_rules(row):
    nums = row[liczby_cols].values
    suma = row['Suma']

    rule1 = 130 <= suma <= 205

    diffs = [abs(nums[i] - nums[i-1]) for i in range(1, 6)]
    rule2 = all(d <= 17 for d in diffs)

    suma_diff = sum(diffs)
    rule3 = 27 <= suma_diff <= 47

    return pd.Series({
        'Regula_suma': rule1,
        'Regula_max_diff': rule2,
        'Regula_suma_diff': rule3,
        'Wszystkie_reguly': rule1 and rule2 and rule3
    })

rules_df = df.apply(check_rules, axis=1)
validation_summary = pd.DataFrame({
    'Reguła': ['Suma 130-205', 'Max różnica ≤17', 'Suma różnic 27-47', 'WSZYSTKIE spełnione'],
    'Losowań_spełnia': [
        rules_df['Regula_suma'].sum(),
        rules_df['Regula_max_diff'].sum(),
        rules_df['Regula_suma_diff'].sum(),
        rules_df['Wszystkie_reguly'].sum()
    ],
    'Procent': [
        round(rules_df['Regula_suma'].sum() / len(df) * 100, 2),
        round(rules_df['Regula_max_diff'].sum() / len(df) * 100, 2),
        round(rules_df['Regula_suma_diff'].sum() / len(df) * 100, 2),
        round(rules_df['Wszystkie_reguly'].sum() / len(df) * 100, 2)
    ]
})

plt.figure(figsize=(10, 6))
colors = ['green' if p > 50 else 'orange' if p > 30 else 'red' for p in validation_summary['Procent']]
bars = plt.bar(validation_summary['Reguła'], validation_summary['Procent'], color=colors, alpha=0.7)
plt.ylabel('Procent losowań spełniających regułę', fontsize=12)
plt.title('Walidacja reguł algorytmu na rzeczywistych danych', fontsize=14, fontweight='bold')
plt.ylim(0, 105)
for i, (bar, row) in enumerate(zip(bars, validation_summary.itertuples())):
    plt.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 2,
             f"{row.Procent}%\n({row.Losowań_spełnia})", ha='center', fontsize=10, fontweight='bold')
plt.xticks(rotation=15, ha='right')
plt.grid(axis='y', alpha=0.3)
plt.tight_layout()
plt.savefig(out('wykres_walidacja.png'), dpi=150, bbox_inches='tight')
plt.close()
print("   ✓ Wygenerowano wykres_walidacja.png")

# ============================================================
# ZAKŁADKA 11: OSTATNIE 50 LOSOWAŃ - HOT NUMBERS
# ============================================================
print("\n9. Generowanie statystyk dla ostatnich 50 losowań...")

counter_recent = Counter(wszystkie_liczby_recent)
freq_recent_df = pd.DataFrame([
    {
        'Liczba': int(num),
        'Wystąpienia_ostatnie_50': counter_recent.get(num, 0),
        'Wystąpienia_ogółem': counter[num],
        'Oczekiwane_w_50': round(counter[num] / len(df) * 50, 2),
        'Różnica': round(counter_recent.get(num, 0) - (counter[num] / len(df) * 50), 2)
    }
    for num in range(1, 50)
])
freq_recent_df = freq_recent_df.sort_values('Wystąpienia_ostatnie_50', ascending=False).reset_index(drop=True)
freq_recent_df.index += 1
freq_recent_df.index.name = 'Ranking'

hot_recent = freq_recent_df.head(20)

plt.figure(figsize=(12, 7))
colors_recent = ['darkred' if r > 5 else 'crimson' if r > 2 else 'orange' if r > 0 else 'gray' for r in hot_recent['Różnica']]
plt.barh(hot_recent['Liczba'].astype(str), hot_recent['Wystąpienia_ostatnie_50'], color=colors_recent, alpha=0.7)
plt.xlabel('Wystąpienia w ostatnich 50 losowaniach', fontsize=12)
plt.ylabel('Liczba', fontsize=12)
plt.title('TOP 20 Hot Numbers - ostatnie 50 losowań', fontsize=12, fontweight='bold')
plt.gca().invert_yaxis()
plt.grid(axis='x', alpha=0.3)
plt.tight_layout()
plt.savefig(out('wykres_hot_recent50.png'), dpi=150, bbox_inches='tight')
plt.close()
print("   ✓ Wygenerowano wykres_hot_recent50.png")

# ============================================================
# ZAKŁADKA 12: STATYSTYKI PORÓWNAWCZE (OSTATNIE 50 vs CAŁA HISTORIA)
# ============================================================
print("\n10. Generowanie statystyk porównawczych...")

try:
    najczestszy_parz_recent = df_recent.groupby(['parzyste', 'nieparzyste']).size().idxmax()
    najczestszy_parz_all = df.groupby(['parzyste', 'nieparzyste']).size().idxmax()
    najczestszy_nh_recent = df_recent.groupby(['niskie', 'wysokie']).size().idxmax()
    najczestszy_nh_all = df.groupby(['niskie', 'wysokie']).size().idxmax()
except Exception:
    najczestszy_parz_recent = 'N/A'
    najczestszy_parz_all = 'N/A'
    najczestszy_nh_recent = 'N/A'
    najczestszy_nh_all = 'N/A'

recent_stats = pd.DataFrame({
    'Statystyka': [
        'Średnia suma',
        'Średni spread',
        'Najczęstszy rozkład P/N',
        'Najczęstszy rozkład L/H'
    ],
    'Ostatnie_50': [
        round(df_recent['Suma'].mean(), 2),
        round(df_recent['spread'].mean(), 2),
        f"{int(df_recent.groupby(['parzyste','nieparzyste']).size().idxmax()[0])}P-{int(df_recent.groupby(['parzyste','nieparzyste']).size().idxmax()[1])}N",
        f"{int(df_recent.groupby(['niskie','wysokie']).size().idxmax()[0])}L-{int(df_recent.groupby(['niskie','wysokie']).size().idxmax()[1])}H"
    ],
    'Cała_historia': [
        round(df['Suma'].mean(), 2),
        round(df['spread'].mean(), 2),
        f"{int(df.groupby(['parzyste','nieparzyste']).size().idxmax()[0])}P-{int(df.groupby(['parzyste','nieparzyste']).size().idxmax()[1])}N",
        f"{int(df.groupby(['niskie','wysokie']).size().idxmax()[0])}L-{int(df.groupby(['niskie','wysokie']).size().idxmax()[1])}H"
    ]
})

# ============================================================
# ZAKŁADKA 13: COLD STREAKS
# ============================================================
print("\n11. Generowanie statystyk cold streaks...")

def find_last_occurrence(liczba):
    for idx, row in df[liczby_cols].iterrows():
        if liczba in row.values:
            return idx
    return len(df)

cold_streaks = []
for num in range(1, 50):
    last_idx = find_last_occurrence(num)
    cold_streaks.append({
        'Liczba': num,
        'Losowań_temu': last_idx,
        'Status': 'HOT 🔥' if last_idx < 5 else 'WARM ☀️' if last_idx < 15 else 'MEDIUM 🌤️' if last_idx < 30 else 'COLD ❄️'
    })

cold_streaks_df = pd.DataFrame(cold_streaks).sort_values('Losowań_temu', ascending=False).reset_index(drop=True)
cold_streaks_df.index += 1
cold_streaks_df.index.name = 'Ranking'

top_cold_streaks = cold_streaks_df.head(10)

plt.figure(figsize=(10, 7))
plt.barh(top_cold_streaks['Liczba'].astype(str), top_cold_streaks['Losowań_temu'], color='dodgerblue', alpha=0.7)
plt.xlabel('Liczba losowań od ostatniego wystąpienia', fontsize=12)
plt.ylabel('Liczba', fontsize=12)
plt.title('TOP 10 Cold Streaks - najdłużej nie wylosowane liczby', fontsize=12, fontweight='bold')
plt.gca().invert_yaxis()
plt.grid(axis='x', alpha=0.3)
plt.tight_layout()
plt.savefig(out('wykres_cold_streaks.png'), dpi=150, bbox_inches='tight')
plt.close()
print("   ✓ Wygenerowano wykres_cold_streaks.png")

# ============================================================
# ZAKŁADKA 14: NAJCZĘSTSZE TRÓJKI LICZB
# ============================================================
print("\n12b. Generowanie statystyk trójek liczb...")
trojki = []
for _, row in df[liczby_cols].iterrows():
    nums = sorted(row.values)
    for trojka in combinations(nums, 3):
        trojki.append(f"{int(trojka[0])}-{int(trojka[1])}-{int(trojka[2])}")

counter_trojki = Counter(trojki)
trojki_df = pd.DataFrame([
    {'Trójka': t, 'Wystąpienia': count}
    for t, count in counter_trojki.most_common(50)
])
trojki_df.index += 1
trojki_df.index.name = 'Ranking'

top_15_trojki = trojki_df.head(15)
plt.figure(figsize=(12, 7))
plt.barh(top_15_trojki['Trójka'], top_15_trojki['Wystąpienia'], color='darkviolet', alpha=0.6)
plt.xlabel('Wystąpienia', fontsize=12)
plt.ylabel('Trójka liczb', fontsize=12)
plt.title('TOP 15 najczęstszych trójek liczb', fontsize=12, fontweight='bold')
plt.gca().invert_yaxis()
plt.grid(axis='x', alpha=0.3)
plt.tight_layout()
plt.savefig(out('wykres_trojki.png'), dpi=150, bbox_inches='tight')
plt.close()
print("   ✓ Wygenerowano wykres_trojki.png")

# ============================================================
# ZAKŁADKA 15: ANTYPARY (liczby najrzadziej razem)
# ============================================================
print("\n15. Generowanie antypary...")
wszystkie_mozliwe_pary = [
    f"{a}-{b}" for a in range(1, 50) for b in range(a+1, 50)
]
antypary_df = pd.DataFrame([
    {'Para': para, 'Wystąpienia': counter_pary.get(para, 0)}
    for para in wszystkie_mozliwe_pary
]).sort_values('Wystąpienia').reset_index(drop=True)
antypary_df.index += 1
antypary_df.index.name = 'Ranking'

top_15_antypary = antypary_df.head(15)
plt.figure(figsize=(12, 7))
plt.barh(top_15_antypary['Para'], top_15_antypary['Wystąpienia'], color='gray', alpha=0.6)
plt.xlabel('Wystąpienia', fontsize=12)
plt.ylabel('Para liczb', fontsize=12)
plt.title('TOP 15 Antypary - liczby najrzadziej wypadające razem', fontsize=12, fontweight='bold')
plt.gca().invert_yaxis()
plt.grid(axis='x', alpha=0.3)
plt.tight_layout()
plt.savefig(out('wykres_antypary.png'), dpi=150, bbox_inches='tight')
plt.close()
print("   ✓ Wygenerowano wykres_antypary.png")

# ============================================================
# ZAKŁADKA 16: CZĘSTOTLIWOŚĆ KROCZĄCA (rolling frequency)
# ============================================================
print("\n16. Generowanie częstotliwości kroczącej...")
# Wybieramy 9 liczb: top 3 hot, top 3 cold, 3 środkowe
top3_hot = freq_df.nlargest(3, 'Wystąpienia')['Liczba'].tolist()
top3_cold = freq_df.nsmallest(3, 'Wystąpienia')['Liczba'].tolist()
srednie = freq_df.sort_values('Wystąpienia').iloc[len(freq_df)//2 - 1 : len(freq_df)//2 + 2]['Liczba'].tolist()
tracked_numbers = top3_hot + top3_cold + srednie

window = 100
rolling_data = {}
for num in tracked_numbers:
    presence = df[liczby_cols].apply(lambda row: int(num in row.values), axis=1)
    rolling_data[str(num)] = presence[::-1].rolling(window=window).sum().values[::-1]

rolling_df_plot = pd.DataFrame(rolling_data)

plt.figure(figsize=(14, 7))
for col in rolling_df_plot.columns:
    valid = rolling_df_plot[col].dropna()
    x = range(len(rolling_df_plot) - len(valid), len(rolling_df_plot))
    plt.plot(x, valid.values, label=f'Liczba {col}', linewidth=1.5)
plt.xlabel('Indeks losowania (od najnowszego)', fontsize=12)
plt.ylabel(f'Wystąpienia w oknie {window} losowań', fontsize=12)
plt.title(f'Częstotliwość krocząca (okno={window}) - wybrane liczby', fontsize=12, fontweight='bold')
plt.legend(ncol=3, fontsize=9)
plt.grid(alpha=0.3)
plt.tight_layout()
plt.savefig(out('wykres_rolling.png'), dpi=150, bbox_inches='tight')
plt.close()

rolling_export = pd.DataFrame({'Indeks_losowania': range(len(df))})
for num in tracked_numbers:
    presence = df[liczby_cols].apply(lambda row: int(num in row.values), axis=1)
    rolling_export[f'Liczba_{num}_roll{window}'] = presence[::-1].rolling(window=window).sum().values[::-1]
print("   ✓ Wygenerowano wykres_rolling.png")

# ============================================================
# ZAKŁADKA 17: HEATMAPA WG ROKU
# ============================================================
print("\n17. Generowanie heatmapy wg roku...")

if 'data' in df.columns and df['data'].notna().sum() > 0:
    df['rok'] = df['data'].dt.year
    print("   ✓ Używam rzeczywistych dat z kolumny 'data'")
else:
    base_date = date.today()
    df['data_syntetyczna'] = [base_date - timedelta(days=2 * i) for i in range(len(df))]
    df['rok'] = df['data_syntetyczna'].apply(lambda d: d.year)
    print("   ℹ Brak kolumny 'data' — używam dat syntetycznych (co 2 dni wstecz od dziś)")

heatmap_data = {}
for num in range(1, 50):
    presence = df[liczby_cols].apply(lambda row: int(num in row.values), axis=1)
    heatmap_data[num] = df.groupby('rok').apply(lambda g: presence[g.index].sum())

heatmap_df = pd.DataFrame(heatmap_data).T
heatmap_df.index.name = 'Liczba'
heatmap_df = heatmap_df[sorted(heatmap_df.columns)]

title_suffix = '' if ('data' in df.columns and df['data'].notna().sum() > 0) else '\n(daty syntetyczne: co 2 dni wstecz od dziś)'
fig, ax = plt.subplots(figsize=(max(12, len(heatmap_df.columns)), 14))
im = ax.imshow(heatmap_df.values, aspect='auto', cmap='YlOrRd')
ax.set_xticks(range(len(heatmap_df.columns)))
ax.set_xticklabels(heatmap_df.columns, rotation=45, ha='right', fontsize=9)
ax.set_yticks(range(len(heatmap_df.index)))
ax.set_yticklabels(heatmap_df.index, fontsize=8)
ax.set_xlabel('Rok', fontsize=12)
ax.set_ylabel('Liczba', fontsize=12)
ax.set_title(f'Heatmapa wystąpień liczb wg roku{title_suffix}', fontsize=13, fontweight='bold')
plt.colorbar(im, ax=ax, label='Liczba wystąpień w roku')
plt.tight_layout()
plt.savefig(out('wykres_heatmapa_rok.png'), dpi=150, bbox_inches='tight')
plt.close()
print("   ✓ Wygenerowano wykres_heatmapa_rok.png")

# ============================================================
# ZAKŁADKA 18: STATYSTYKI POZYCYJNE
# FIX: użyj tick_labels= zamiast labels= (matplotlib 3.9+)
# ============================================================
print("\n18. Generowanie statystyk pozycyjnych...")
pozycje_stats = pd.DataFrame({
    f'Pozycja_{i+1}': df[liczby_cols[i]].describe()
    for i in range(6)
}).T
pozycje_stats.index.name = 'Pozycja'

srednie_pozycji = df[liczby_cols].mean()

pozycja_najczesciej = []
for num in range(1, 50):
    counts = [
        (df[col] == num).sum()
        for col in liczby_cols
    ]
    najczesciej = liczby_cols[counts.index(max(counts))]
    pozycja_najczesciej.append({
        'Liczba': num,
        'Najczęstsza_pozycja': najczesciej,
        'Wystąpień_na_tej_pozycji': max(counts),
        **{f'Poz_{i+1}': counts[i] for i in range(6)}
    })

pozycja_df = pd.DataFrame(pozycja_najczesciej)

fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))

ax1.bar(range(1, 7), srednie_pozycji.values, color='steelblue', alpha=0.7)
ax1.set_xlabel('Pozycja (1=najmniejsza, 6=największa)', fontsize=11)
ax1.set_ylabel('Średnia wartość liczby', fontsize=11)
ax1.set_title('Średnia wartość liczby na każdej pozycji', fontsize=11, fontweight='bold')
ax1.set_xticks(range(1, 7))
for i, v in enumerate(srednie_pozycji.values):
    ax1.text(i+1, v + 0.3, f'{v:.1f}', ha='center', fontsize=10)
ax1.grid(axis='y', alpha=0.3)

# Kompatybilność matplotlib: tick_labels= (>=3.9) lub labels= (<3.9)
import matplotlib
_mpl_ver = tuple(int(x) for x in matplotlib.__version__.split('.')[:2])
_boxplot_kwargs = dict(
    patch_artist=True,
    boxprops=dict(facecolor='lightblue', alpha=0.7)
)
if _mpl_ver >= (3, 9):
    _boxplot_kwargs['tick_labels'] = [f'P{i+1}' for i in range(6)]
else:
    _boxplot_kwargs['labels'] = [f'P{i+1}' for i in range(6)]

ax2.boxplot([df[col].values for col in liczby_cols], **_boxplot_kwargs)
ax2.set_xlabel('Pozycja', fontsize=11)
ax2.set_ylabel('Wartość liczby', fontsize=11)
ax2.set_title('Rozkład wartości na każdej pozycji (boxplot)', fontsize=11, fontweight='bold')
ax2.grid(axis='y', alpha=0.3)

plt.tight_layout()
plt.savefig(out('wykres_pozycje.png'), dpi=150, bbox_inches='tight')
plt.close()
print("   ✓ Wygenerowano wykres_pozycje.png")

# ============================================================
# ZAKŁADKA 19: ANALIZA RÓŻNIC R1-R5
# ============================================================
print("\n19. Generowanie analizy różnic R1-R5...")
r_cols_present = [c for c in ['R1', 'R2', 'R3', 'R4', 'R5'] if c in df.columns]

if r_cols_present:
    r_stats = df[r_cols_present].describe().T
    r_stats.index.name = 'Różnica'

    fig, axes = plt.subplots(1, len(r_cols_present), figsize=(14, 5), sharey=False)
    if len(r_cols_present) == 1:
        axes = [axes]
    colors_r = ['#e74c3c', '#e67e22', '#2ecc71', '#3498db', '#9b59b6']
    for i, (col, ax) in enumerate(zip(r_cols_present, axes)):
        ax.hist(df[col].dropna(), bins=20, color=colors_r[i], alpha=0.7, edgecolor='black')
        ax.set_title(f'{col}\nśr={df[col].mean():.1f}', fontsize=10, fontweight='bold')
        ax.set_xlabel('Wartość różnicy', fontsize=9)
        ax.set_ylabel('Liczba losowań', fontsize=9)
        ax.grid(axis='y', alpha=0.3)
    plt.suptitle('Rozkład różnic między kolejnymi liczbami (R1=P2-P1, itd.)', fontsize=12, fontweight='bold')
    plt.tight_layout()
    plt.savefig(out('wykres_roznice_r1r5.png'), dpi=150, bbox_inches='tight')
    plt.close()
    print("   ✓ Wygenerowano wykres_roznice_r1r5.png")
    has_r_cols = True
else:
    # Oblicz różnice na bieżąco jeśli nie ma kolumn R1-R5
    print("   ℹ Brak kolumn R1-R5 w pliku - obliczam na bieżąco...")
    for i, col in enumerate(['R1', 'R2', 'R3', 'R4', 'R5']):
        df[col] = df[liczby_cols[i+1]] - df[liczby_cols[i]]
    r_cols_present = ['R1', 'R2', 'R3', 'R4', 'R5']
    r_stats = df[r_cols_present].describe().T
    r_stats.index.name = 'Różnica'

    fig, axes = plt.subplots(1, 5, figsize=(14, 5), sharey=False)
    colors_r = ['#e74c3c', '#e67e22', '#2ecc71', '#3498db', '#9b59b6']
    for i, (col, ax) in enumerate(zip(r_cols_present, axes)):
        ax.hist(df[col].dropna(), bins=20, color=colors_r[i], alpha=0.7, edgecolor='black')
        ax.set_title(f'{col}\nśr={df[col].mean():.1f}', fontsize=10, fontweight='bold')
        ax.set_xlabel('Wartość różnicy', fontsize=9)
        ax.set_ylabel('Liczba losowań', fontsize=9)
        ax.grid(axis='y', alpha=0.3)
    plt.suptitle('Rozkład różnic między kolejnymi liczbami (R1=P2-P1, itd.)', fontsize=12, fontweight='bold')
    plt.tight_layout()
    plt.savefig(out('wykres_roznice_r1r5.png'), dpi=150, bbox_inches='tight')
    plt.close()
    print("   ✓ Wygenerowano wykres_roznice_r1r5.png")
    has_r_cols = True

# ============================================================
# ZAKŁADKA 20: PRAWIE DUPLIKATY — szybka wersja z numpy
# ============================================================
print("\n20. Szukanie prawie duplikatów (≥5 wspólnych liczb)...")

draws_array = df[liczby_cols].values  # shape (N, 6)
prawie_duplikaty = []

masks = np.zeros(len(draws_array), dtype=np.int64)
for i, row in enumerate(draws_array):
    for num in row:
        masks[i] |= (1 << int(num))

window_dup = min(1000, len(masks))
for i in range(window_dup):
    for j in range(i + 1, window_dup):
        common = bin(masks[i] & masks[j]).count('1')
        if common >= 5:
            prawie_duplikaty.append({
                'Losowanie_A': i,
                'Losowanie_B': j,
                'Wspólnych_liczb': common,
                'Zestaw_A': str(sorted(draws_array[i].tolist())),
                'Zestaw_B': str(sorted(draws_array[j].tolist()))
            })

if prawie_duplikaty:
    prawie_dup_df = pd.DataFrame(prawie_duplikaty).sort_values(
        'Wspólnych_liczb', ascending=False
    ).reset_index(drop=True)
    prawie_dup_df.index += 1
    prawie_dup_df.index.name = 'Ranking'
else:
    prawie_dup_df = pd.DataFrame({
        'Info': ['Brak losowań z ≥5 wspólnymi liczbami w ostatnich 1000']
    })

print(f"   ✓ Znaleziono {len(prawie_duplikaty)} par z ≥5 wspólnymi liczbami")

sample_size = min(300, len(draws_array))
wspolne_counts = Counter()
for i in range(sample_size):
    for j in range(i + 1, sample_size):
        common = bin(masks[i] & masks[j]).count('1')
        wspolne_counts[common] += 1

plt.figure(figsize=(10, 5))
x_vals = sorted(wspolne_counts.keys())
y_vals = [wspolne_counts[x] for x in x_vals]
plt.bar(x_vals, y_vals, color='coral', alpha=0.7, edgecolor='black')
plt.xlabel('Liczba wspólnych liczb między parą losowań', fontsize=12)
plt.ylabel('Liczba par losowań', fontsize=12)
plt.title(f'Rozkład podobieństwa losowań (próbka {sample_size} losowań)', fontsize=12, fontweight='bold')
plt.xticks(x_vals)
plt.grid(axis='y', alpha=0.3)
plt.tight_layout()
plt.savefig(out('wykres_podobienstwo.png'), dpi=150, bbox_inches='tight')
plt.close()
print("   ✓ Wygenerowano wykres_podobienstwo.png")


# ============================================================
# ZAKŁADKA 21: WALIDACJA REGUŁ - SLIDING WINDOW
# ============================================================
print("\n21. Generowanie sliding window walidacji reguł...")
window_val = 100
results_sliding = []

df_rev = df.iloc[::-1].reset_index(drop=True)  # od najstarszego
rules_rev = rules_df.iloc[::-1].reset_index(drop=True)

for start in range(0, len(df_rev) - window_val, window_val // 2):
    end = start + window_val
    chunk = rules_rev.iloc[start:end]
    results_sliding.append({
        'Okno_start': start,
        'Okno_end': end,
        'Regula_suma_%': round(chunk['Regula_suma'].mean() * 100, 1),
        'Regula_max_diff_%': round(chunk['Regula_max_diff'].mean() * 100, 1),
        'Regula_suma_diff_%': round(chunk['Regula_suma_diff'].mean() * 100, 1),
        'Wszystkie_%': round(chunk['Wszystkie_reguly'].mean() * 100, 1)
    })

sliding_df = pd.DataFrame(results_sliding)

plt.figure(figsize=(14, 6))
x = range(len(sliding_df))
plt.plot(x, sliding_df['Regula_suma_%'], label='Suma 130-205', linewidth=2, marker='o', markersize=3)
plt.plot(x, sliding_df['Regula_max_diff_%'], label='Max różnica ≤17', linewidth=2, marker='s', markersize=3)
plt.plot(x, sliding_df['Regula_suma_diff_%'], label='Suma różnic 27-47', linewidth=2, marker='^', markersize=3)
plt.plot(x, sliding_df['Wszystkie_%'], label='WSZYSTKIE', linewidth=2.5, marker='D', markersize=3, linestyle='--', color='black')
plt.xlabel(f'Okno czasowe (co {window_val//2} losowań, od najstarszego)', fontsize=11)
plt.ylabel('% losowań spełniających regułę', fontsize=11)
plt.title(f'Stabilność reguł algorytmu w czasie (okno={window_val})', fontsize=12, fontweight='bold')
plt.legend(fontsize=10)
plt.ylim(0, 105)
plt.grid(alpha=0.3)
plt.tight_layout()
plt.savefig(out('wykres_sliding_window.png'), dpi=150, bbox_inches='tight')
plt.close()
print("   ✓ Wygenerowano wykres_sliding_window.png")

# ============================================================
# ZAKŁADKA 22: GAP ANALYSIS — odstępy między wystąpieniami
# ============================================================
print("\n22. Generowanie gap analysis...")

gap_stats = []
gap_detail = []

for num in range(1, 50):
    indices = np.where(
        df[liczby_cols].apply(lambda row: num in row.values, axis=1).values
    )[0].tolist()
    if len(indices) < 2:
        continue

    gaps = [indices[i+1] - indices[i] for i in range(len(indices) - 1)]
    aktualny_gap = indices[0]

    gap_stats.append({
        'Liczba': num,
        'Śr_odstęp': round(np.mean(gaps), 2),
        'Min_odstęp': int(np.min(gaps)),
        'Max_odstęp': int(np.max(gaps)),
        'Mediana': round(np.median(gaps), 1),
        'Odch_std': round(np.std(gaps), 2),
        'Wystąpień': len(indices),
        'Aktualny_gap': aktualny_gap
    })

    for g in gaps[:200]:
        gap_detail.append({'Liczba': num, 'Gap': g})

gap_stats_df = pd.DataFrame(gap_stats).sort_values('Śr_odstęp')
gap_stats_df.index = range(1, len(gap_stats_df) + 1)
gap_stats_df.index.name = 'Ranking'

gap_detail_df = pd.DataFrame(gap_detail)

fig, axes = plt.subplots(2, 1, figsize=(14, 12))

ax1 = axes[0]
colors_gap = ['crimson' if s < 7 else 'steelblue' if s < 9 else 'gray'
              for s in gap_stats_df['Śr_odstęp']]
bars = ax1.bar(gap_stats_df['Liczba'].astype(str), gap_stats_df['Śr_odstęp'],
               color=colors_gap, alpha=0.75)
ax1.axhline(y=49/6, color='red', linestyle='--', linewidth=1.5,
            label=f'Oczekiwany średni odstęp (~{49/6:.1f})')
ax1.set_xlabel('Liczba', fontsize=11)
ax1.set_ylabel('Średni odstęp (losowań)', fontsize=11)
ax1.set_title('Średni odstęp między wystąpieniami dla każdej liczby', fontsize=12, fontweight='bold')
ax1.legend()
ax1.grid(axis='y', alpha=0.3)
ax1.tick_params(axis='x', rotation=90)

ax2 = axes[1]
x = range(len(gap_stats_df))
srednie = gap_stats_df['Śr_odstęp'].values
minima = gap_stats_df['Min_odstęp'].values
maksima = gap_stats_df['Max_odstęp'].values
ax2.errorbar(x, srednie,
             yerr=[srednie - minima, maksima - srednie],
             fmt='o', color='steelblue', ecolor='lightcoral',
             elinewidth=1.5, capsize=3, markersize=4, alpha=0.8)
ax2.set_xticks(list(x))
ax2.set_xticklabels(gap_stats_df['Liczba'].astype(str), rotation=90, fontsize=8)
ax2.set_xlabel('Liczba (posortowane wg średniego odstępu)', fontsize=11)
ax2.set_ylabel('Odstęp (losowań)', fontsize=11)
ax2.set_title('Min / Średnia / Max odstęp dla każdej liczby', fontsize=12, fontweight='bold')
ax2.axhline(y=49/6, color='red', linestyle='--', linewidth=1.5, alpha=0.7)
ax2.grid(alpha=0.3)

plt.tight_layout()
plt.savefig(out('wykres_gap_analysis.png'), dpi=150, bbox_inches='tight')
plt.close()
print("   ✓ Wygenerowano wykres_gap_analysis.png")

plt.figure(figsize=(12, 5))
all_gaps = gap_detail_df['Gap'].values
plt.hist(all_gaps, bins=range(1, int(all_gaps.max())+2), color='teal',
         alpha=0.7, edgecolor='black')
plt.axvline(x=49/6, color='red', linestyle='--', linewidth=2,
            label=f'Oczekiwany: ~{49/6:.1f}')
plt.xlabel('Odstęp między wystąpieniami (liczba losowań)', fontsize=12)
plt.ylabel('Liczba przypadków', fontsize=12)
plt.title('Rozkład wszystkich odstępów (dla wszystkich liczb 1-49)', fontsize=12, fontweight='bold')
plt.legend()
plt.grid(axis='y', alpha=0.3)
plt.tight_layout()
plt.savefig(out('wykres_gap_histogram.png'), dpi=150, bbox_inches='tight')
plt.close()
print("   ✓ Wygenerowano wykres_gap_histogram.png")

# ============================================================
# ZAKŁADKA 23: ANTY-TRÓJKI — najrzadziej/nigdy nie wystąpiły
# ============================================================
print("\n23. Generowanie anty-trójek...")

wszystkie_trojki = set(
    f"{a}-{b}-{c}"
    for a, b, c in combinations(range(1, 50), 3)
)

antytrojki_df = pd.DataFrame([
    {'Trójka': t, 'Wystąpienia': counter_trojki.get(t, 0)}
    for t in wszystkie_trojki
]).sort_values('Wystąpienia').reset_index(drop=True)
antytrojki_df.index += 1
antytrojki_df.index.name = 'Ranking'

nigdy = antytrojki_df[antytrojki_df['Wystąpienia'] == 0]
raz   = antytrojki_df[antytrojki_df['Wystąpienia'] == 1]
dwa   = antytrojki_df[antytrojki_df['Wystąpienia'] == 2]

print(f"   ✓ Nigdy nie wystąpiły:       {len(nigdy):>6} trójek ({len(nigdy)/18424*100:.1f}%)")
print(f"   ✓ Wystąpiły dokładnie 1 raz:  {len(raz):>6} trójek ({len(raz)/18424*100:.1f}%)")
print(f"   ✓ Wystąpiły dokładnie 2 razy: {len(dwa):>6} trójek ({len(dwa)/18424*100:.1f}%)")

rozklad_trojki = antytrojki_df.groupby('Wystąpienia').size().reset_index(name='Liczba_trojek')

plt.figure(figsize=(14, 6))
plt.bar(rozklad_trojki['Wystąpienia'], rozklad_trojki['Liczba_trojek'],
        color='darkviolet', alpha=0.7, edgecolor='black')
plt.xlabel('Liczba wystąpień', fontsize=12)
plt.ylabel('Liczba trójek', fontsize=12)
plt.title(f'Rozkład wystąpień wszystkich {18424} możliwych trójek\n'
          f'(nigdy: {len(nigdy)}, raz: {len(raz)}, dwa razy: {len(dwa)})',
          fontsize=12, fontweight='bold')
plt.axvline(x=146240/18424, color='red', linestyle='--', linewidth=2,
            label=f'Oczekiwana średnia: {146240/18424:.1f}')
plt.legend()
plt.grid(axis='y', alpha=0.3)
plt.tight_layout()
plt.savefig(out('wykres_antytrojki_rozklad.png'), dpi=150, bbox_inches='tight')
plt.close()
print("   ✓ Wygenerowano wykres_antytrojki_rozklad.png")

# ============================================================
# ZAPISANIE DO PLIKU XLSX
# ============================================================
print("\n24. Zapisywanie danych do pliku Excel...")
output_file = out('statystyki_lotto.xlsx')

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    freq_df.to_excel(writer, sheet_name='1_Czestotliwosc', index=False)
    top_20.to_excel(writer, sheet_name='2_Hot_Numbers')
    bottom_20.to_excel(writer, sheet_name='3_Cold_Numbers')
    rozklad_parz[['Rozkład', 'Liczba_losowan', 'Procent']].to_excel(writer, sheet_name='4_Parzyste_Nieparzyste', index=False)
    rozklad_nh[['Rozkład', 'Liczba_losowan', 'Procent']].to_excel(writer, sheet_name='5_Niskie_Wysokie', index=False)
    pary_df.to_excel(writer, sheet_name='6_TOP50_Par')
    spread_summary.to_excel(writer, sheet_name='7_Spread_Statystyki', index=False)
    spread_detail.to_excel(writer, sheet_name='8_Spread_Rozklad')
    suma_top.to_excel(writer, sheet_name='9_Suma_TOP50')
    validation_summary.to_excel(writer, sheet_name='10_Walidacja_Regul', index=False)
    hot_recent.to_excel(writer, sheet_name='11_Hot_Ostatnie50')
    recent_stats.to_excel(writer, sheet_name='12_Stats_Ostatnie50', index=False)
    cold_streaks_df.to_excel(writer, sheet_name='13_Cold_Streaks')
    trojki_df.to_excel(writer, sheet_name='14_TOP50_Trojek')
    antypary_df.head(50).to_excel(writer, sheet_name='15_Antypary', index=False)
    rolling_export.to_excel(writer, sheet_name='16_Rolling_Freq', index=False)
    heatmap_df.to_excel(writer, sheet_name='17_Heatmapa_Rok')
    pozycje_stats.to_excel(writer, sheet_name='18_Stat_Pozycyjne')
    pozycja_df.to_excel(writer, sheet_name='18b_Liczby_Pozycje', index=False)
    r_stats.to_excel(writer, sheet_name='19_Roznice_R1R5')
    prawie_dup_df.to_excel(writer, sheet_name='20_Prawie_Duplikaty')
    sliding_df.to_excel(writer, sheet_name='21_Sliding_Window', index=False)
    gap_stats_df.to_excel(writer, sheet_name='22_Gap_Statystyki')
    gap_detail_df.to_excel(writer, sheet_name='22b_Gap_Szczegoly', index=False)
    antytrojki_df.head(100).to_excel(writer, sheet_name='23_Anty_Trojki_TOP100', index=True)
    nigdy.to_excel(writer, sheet_name='23b_Trojki_Nigdy', index=False)
    rozklad_trojki.to_excel(writer, sheet_name='23c_Trojki_Rozklad', index=False)

print(f"   ✓ Zapisano plik: {output_file}")
print("\n✅ Wszystkie statystyki zostały wygenerowane i zapisane do pliku Excel!")
