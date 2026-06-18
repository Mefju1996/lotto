#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generator statystyk Lotto - WERSJA Z WYKRESAMI W EXCELU
Analizuje bazę historycznych wyników i generuje raport XLSX z wykresami wbudowanymi

Użycie:
    python generate_lotto_stats_with_charts.py

Wymagania:
    - wyniki_lotto.xlsx w tym samym katalogu
    - pandas, matplotlib, openpyxl, pillow
"""

import pandas as pd
import matplotlib.pyplot as plt
from collections import Counter
from itertools import combinations
import numpy as np
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
import io

ROOT_DIR = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT_DIR / 'data'

print("="*70)
print("GENERATOR STATYSTYK LOTTO - WERSJA Z WYKRESAMI")
print("="*70)

# ============================================================
# WCZYTANIE DANYCH
# ============================================================
print("\n1. Wczytywanie danych...")
file = DATA_DIR / 'wyniki_lotto.xlsx'
df = pd.read_excel(file, sheet_name='Arkusz1')
print(f"   ✓ Wczytano {len(df)} losowań")

liczby_cols = ['pierwsza', 'druga', 'trzecia', 'czwarta', 'piąta', 'szósta']
wszystkie_liczby = df[liczby_cols].values.flatten()

# Obliczenia pomocnicze
def count_parzyste(row):
    return sum(1 for x in row if x % 2 == 0)

def count_niskie(row):
    return sum(1 for x in row if x <= 24)

df['spread'] = df['szósta'] - df['pierwsza']
df['parzyste'] = df[liczby_cols].apply(count_parzyste, axis=1)
df['nieparzyste'] = 6 - df['parzyste']
df['niskie'] = df[liczby_cols].apply(count_niskie, axis=1)
df['wysokie'] = 6 - df['niskie']

df_recent = df.head(50).copy()
wszystkie_liczby_recent = df_recent[liczby_cols].values.flatten()

# ============================================================
# FUNKCJA DO KONWERSJI WYKRESU MATPLOTLIB NA OBRAZ
# ============================================================
def fig_to_img():
    """Konwertuje aktualny wykres matplotlib na obiekt Image openpyxl"""
    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight')
    buf.seek(0)
    return XLImage(buf)

# ============================================================
# GENEROWANIE DANYCH (bez zmian)
# ============================================================
print("\n2. Generowanie statystyk...")

counter = Counter(wszystkie_liczby)
freq_df = pd.DataFrame([
    {'Liczba': int(num), 'Wystąpienia': count, 'Procent': round(count/len(df)*100, 2)}
    for num, count in sorted(counter.items())
])

top_20 = freq_df.nlargest(20, 'Wystąpienia').reset_index(drop=True)
top_20.index += 1
top_20.index.name = 'Ranking'

bottom_20 = freq_df.nsmallest(20, 'Wystąpienia').reset_index(drop=True)
bottom_20.index += 1
bottom_20.index.name = 'Ranking'

rozklad_parz = df.groupby('parzyste').size().reset_index(name='Liczba_losowan')
rozklad_parz['Rozkład'] = rozklad_parz['parzyste'].astype(str) + 'P-' + (6 - rozklad_parz['parzyste']).astype(str) + 'N'
rozklad_parz['Procent'] = round(rozklad_parz['Liczba_losowan'] / len(df) * 100, 2)

rozklad_nh = df.groupby('niskie').size().reset_index(name='Liczba_losowan')
rozklad_nh['Rozkład'] = rozklad_nh['niskie'].astype(str) + 'L-' + (6 - rozklad_nh['niskie']).astype(str) + 'H'
rozklad_nh['Procent'] = round(rozklad_nh['Liczba_losowan'] / len(df) * 100, 2)

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

spread_stats = df['spread'].describe()
spread_grouped = df.groupby('spread').size().reset_index(name='Liczba_losowan')
spread_grouped['Procent'] = round(spread_grouped['Liczba_losowan'] / len(df) * 100, 2)

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

suma_grouped = df.groupby('Suma').size().reset_index(name='Liczba_losowan')
suma_grouped['Procent'] = round(suma_grouped['Liczba_losowan'] / len(df) * 100, 2)
suma_top = suma_grouped.sort_values('Liczba_losowan', ascending=False).head(50).reset_index(drop=True)
suma_top.index += 1
suma_top.index.name = 'Ranking'

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
    'Reguła': ['Suma 130-205', 'Max różnica ≤17', 'Suma różnic 27-47', 'WSZYSTKIE'],
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

try:
    najczestszy_parz_recent = df_recent.groupby(['parzyste', 'nieparzyste']).size().idxmax()
    najczestszy_parz_all = df.groupby(['parzyste', 'nieparzyste']).size().idxmax()
    najczestszy_nh_recent = df_recent.groupby(['niskie', 'wysokie']).size().idxmax()
    najczestszy_nh_all = df.groupby(['niskie', 'wysokie']).size().idxmax()
except:
    najczestszy_parz_recent = 'N/A'
    najczestszy_parz_all = 'N/A'
    najczestszy_nh_recent = 'N/A'
    najczestszy_nh_all = 'N/A'

recent_stats = pd.DataFrame({
    'Statystyka': ['Średnia suma', 'Średni spread', 'Rozkład P/N', 'Rozkład L/H'],
    'Ostatnie_50': [
        round(df_recent['Suma'].mean(), 2),
        round(df_recent['spread'].mean(), 2),
        str(najczestszy_parz_recent),
        str(najczestszy_nh_recent)
    ],
    'Cała_historia': [
        round(df['Suma'].mean(), 2),
        round(df['spread'].mean(), 2),
        str(najczestszy_parz_all),
        str(najczestszy_nh_all)
    ]
})

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

# ============================================================
# ZAPISANIE DANYCH DO EXCELA
# ============================================================
print("\n3. Zapisywanie danych do Excel...")
output_file = ROOT_DIR / 'analysis' / 'statystyki_lotto_z_wykresami.xlsx'

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

print("   ✓ Dane zapisane")

# ============================================================
# WSTAWIANIE WYKRESÓW DO EXCELA
# ============================================================
print("\n4. Generowanie i wstawianie wykresów do arkuszy...")

wb = load_workbook(output_file)

# Wykres 1: Częstotliwość -> zakładka 1_Czestotliwosc
print("   • Wykres 1: Częstotliwość liczb...")
plt.figure(figsize=(12, 5))
plt.bar(freq_df['Liczba'], freq_df['Wystąpienia'], color='steelblue', alpha=0.7)
plt.xlabel('Liczba', fontsize=10)
plt.ylabel('Wystąpienia', fontsize=10)
plt.title(f'Częstotliwość liczb 1-49 ({len(df)} losowań)', fontsize=11, fontweight='bold')
plt.xticks(range(1, 50), fontsize=7)
plt.grid(axis='y', alpha=0.3)
plt.tight_layout()
img1 = fig_to_img()
ws1 = wb['1_Czestotliwosc']
ws1.add_image(img1, 'E2')
plt.close()

# Wykres 2: Hot vs Cold -> zakładka 2_Hot_Numbers
print("   • Wykres 2: Hot vs Cold Numbers...")
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 5))
top_10 = freq_df.nlargest(10, 'Wystąpienia')
ax1.barh(top_10['Liczba'].astype(str), top_10['Wystąpienia'], color='crimson', alpha=0.7)
ax1.set_xlabel('Wystąpienia', fontsize=9)
ax1.set_ylabel('Liczba', fontsize=9)
ax1.set_title('TOP 10 Hot', fontsize=10, fontweight='bold')
ax1.invert_yaxis()
ax1.grid(axis='x', alpha=0.3)

bottom_10 = freq_df.nsmallest(10, 'Wystąpienia')
ax2.barh(bottom_10['Liczba'].astype(str), bottom_10['Wystąpienia'], color='dodgerblue', alpha=0.7)
ax2.set_xlabel('Wystąpienia', fontsize=9)
ax2.set_ylabel('Liczba', fontsize=9)
ax2.set_title('TOP 10 Cold', fontsize=10, fontweight='bold')
ax2.invert_yaxis()
ax2.grid(axis='x', alpha=0.3)
plt.tight_layout()
img2 = fig_to_img()
ws2 = wb['2_Hot_Numbers']
ws2.add_image(img2, 'E2')
plt.close()

# Wykres 3: Parzystość -> zakładka 4_Parzyste_Nieparzyste
print("   • Wykres 3: Rozkład parzystości...")
plt.figure(figsize=(10, 5))
plt.bar(rozklad_parz['Rozkład'], rozklad_parz['Liczba_losowan'], color='mediumseagreen', alpha=0.7)
plt.xlabel('Rozkład (P=Parzyste, N=Nieparzyste)', fontsize=10)
plt.ylabel('Liczba losowań', fontsize=10)
plt.title('Rozkład parzystości', fontsize=11, fontweight='bold')
for i, row in rozklad_parz.iterrows():
    plt.text(i, row['Liczba_losowan'] + 50, f"{row['Procent']}%", ha='center', fontsize=9)
plt.grid(axis='y', alpha=0.3)
plt.tight_layout()
img3 = fig_to_img()
ws3 = wb['4_Parzyste_Nieparzyste']
ws3.add_image(img3, 'E2')
plt.close()

# Wykres 4: Niskie/Wysokie -> zakładka 5_Niskie_Wysokie
print("   • Wykres 4: Rozkład niskie/wysokie...")
plt.figure(figsize=(10, 5))
plt.bar(rozklad_nh['Rozkład'], rozklad_nh['Liczba_losowan'], color='darkorange', alpha=0.7)
plt.xlabel('Rozkład (L=Niskie 1-24, H=Wysokie 25-49)', fontsize=10)
plt.ylabel('Liczba losowań', fontsize=10)
plt.title('Rozkład niskie/wysokie', fontsize=11, fontweight='bold')
for i, row in rozklad_nh.iterrows():
    plt.text(i, row['Liczba_losowan'] + 50, f"{row['Procent']}%", ha='center', fontsize=9)
plt.grid(axis='y', alpha=0.3)
plt.tight_layout()
img4 = fig_to_img()
ws4 = wb['5_Niskie_Wysokie']
ws4.add_image(img4, 'E2')
plt.close()

# Wykres 5: Top pary -> zakładka 6_TOP50_Par
print("   • Wykres 5: TOP 15 par liczb...")
top_15_pary = pary_df.head(15)
plt.figure(figsize=(10, 6))
plt.barh(top_15_pary['Para'], top_15_pary['Wystąpienia'], color='purple', alpha=0.6)
plt.xlabel('Wystąpienia', fontsize=10)
plt.ylabel('Para liczb', fontsize=9)
plt.title('TOP 15 najczęstszych par', fontsize=11, fontweight='bold')
plt.gca().invert_yaxis()
plt.grid(axis='x', alpha=0.3)
plt.tight_layout()
img5 = fig_to_img()
ws5 = wb['6_TOP50_Par']
ws5.add_image(img5, 'D2')
plt.close()

# Wykres 6: Spread -> zakładka 7_Spread_Statystyki
print("   • Wykres 6: Histogram spreadu...")
plt.figure(figsize=(10, 5))
plt.hist(df['spread'], bins=40, color='teal', alpha=0.7, edgecolor='black')
plt.xlabel('Spread', fontsize=10)
plt.ylabel('Liczba losowań', fontsize=10)
plt.title(f'Rozkład spreadu (śr: {spread_stats["mean"]:.1f})', fontsize=11, fontweight='bold')
plt.axvline(spread_stats['mean'], color='red', linestyle='--', linewidth=2, label=f'Średnia: {spread_stats["mean"]:.1f}')
plt.legend()
plt.grid(axis='y', alpha=0.3)
plt.tight_layout()
img6 = fig_to_img()
ws6 = wb['7_Spread_Statystyki']
ws6.add_image(img6, 'D2')
plt.close()

# Wykres 7: Suma -> zakładka 9_Suma_TOP50
print("   • Wykres 7: Histogram sum...")
plt.figure(figsize=(12, 5))
plt.hist(df['Suma'], bins=50, color='navy', alpha=0.6, edgecolor='black')
plt.xlabel('Suma 6 liczb', fontsize=10)
plt.ylabel('Liczba losowań', fontsize=10)
plt.title(f'Rozkład sum (śr: {df["Suma"].mean():.1f})', fontsize=11, fontweight='bold')
plt.axvline(df['Suma'].mean(), color='red', linestyle='--', linewidth=2, label=f'Średnia: {df["Suma"].mean():.1f}')
plt.axvline(130, color='green', linestyle=':', linewidth=2, label='Zakres algorytmu: 130-205')
plt.axvline(205, color='green', linestyle=':', linewidth=2)
plt.legend(fontsize=8)
plt.grid(axis='y', alpha=0.3)
plt.tight_layout()
img7 = fig_to_img()
ws7 = wb['9_Suma_TOP50']
ws7.add_image(img7, 'E2')
plt.close()

# Wykres 8: Walidacja -> zakładka 10_Walidacja_Regul
print("   • Wykres 8: Walidacja reguł...")
plt.figure(figsize=(10, 5))
colors = ['green' if p > 50 else 'orange' if p > 30 else 'red' for p in validation_summary['Procent']]
bars = plt.bar(validation_summary['Reguła'], validation_summary['Procent'], color=colors, alpha=0.7)
plt.ylabel('Procent losowań spełniających regułę', fontsize=10)
plt.title('Walidacja reguł algorytmu', fontsize=11, fontweight='bold')
plt.ylim(0, 105)
for i, (bar, row) in enumerate(zip(bars, validation_summary.itertuples())):
    plt.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 2, 
             f"{row.Procent}%\n({row.Losowań_spełnia})", ha='center', fontsize=9, fontweight='bold')
plt.xticks(rotation=15, ha='right', fontsize=9)
plt.grid(axis='y', alpha=0.3)
plt.tight_layout()
img8 = fig_to_img()
ws8 = wb['10_Walidacja_Regul']
ws8.add_image(img8, 'E2')
plt.close()

# Wykres 9: Hot Recent -> zakładka 11_Hot_Ostatnie50
print("   • Wykres 9: Hot Numbers ostatnie 50...")
plt.figure(figsize=(10, 6))
colors_recent = ['darkred' if r > 5 else 'crimson' if r > 2 else 'orange' if r > 0 else 'gray' for r in hot_recent['Różnica']]
plt.barh(hot_recent['Liczba'].astype(str), hot_recent['Wystąpienia_ostatnie_50'], color=colors_recent, alpha=0.7)
plt.xlabel('Wystąpienia w ostatnich 50', fontsize=10)
plt.ylabel('Liczba', fontsize=9)
plt.title('TOP 20 Hot Numbers - ostatnie 50', fontsize=11, fontweight='bold')
plt.gca().invert_yaxis()
plt.grid(axis='x', alpha=0.3)
plt.tight_layout()
img9 = fig_to_img()
ws9 = wb['11_Hot_Ostatnie50']
ws9.add_image(img9, 'G2')
plt.close()

# Wykres 10: Cold Streaks -> zakładka 13_Cold_Streaks
print("   • Wykres 10: Cold Streaks...")
top_cold_streaks = cold_streaks_df.head(10)
plt.figure(figsize=(10, 6))
plt.barh(top_cold_streaks['Liczba'].astype(str), top_cold_streaks['Losowań_temu'], color='dodgerblue', alpha=0.7)
plt.xlabel('Losowań od ostatniego wystąpienia', fontsize=10)
plt.ylabel('Liczba', fontsize=9)
plt.title('TOP 10 Cold Streaks', fontsize=11, fontweight='bold')
plt.gca().invert_yaxis()
plt.grid(axis='x', alpha=0.3)
plt.tight_layout()
img10 = fig_to_img()
ws10 = wb['13_Cold_Streaks']
ws10.add_image(img10, 'E2')
plt.close()

# Zapisanie pliku z wykresami
wb.save(output_file)
print("   ✓ Wszystkie wykresy wstawione")

# ============================================================
# PODSUMOWANIE
# ============================================================
print("\n" + "="*70)
print("✅ GENERATOR ZAKOŃCZONY POMYŚLNIE")
print("="*70)
print(f"\n📊 Wygenerowano plik: {output_file}")
print("   Zawiera 13 zakładek z danymi + 10 wykresów wbudowanych!\n")
print("📈 Wykresy wstawione w zakładkach:")
print("   1. Częstotliwość liczb (zakładka 1)")
print("   2. Hot vs Cold (zakładka 2)")
print("   3. Rozkład parzystości (zakładka 4)")
print("   4. Rozkład niskie/wysokie (zakładka 5)")
print("   5. TOP 15 par (zakładka 6)")
print("   6. Histogram spreadu (zakładka 7)")
print("   7. Histogram sum (zakładka 9)")
print("   8. Walidacja reguł (zakładka 10)")
print("   9. Hot Numbers ostatnie 50 (zakładka 11)")
print("  10. Cold Streaks (zakładka 13)")
print("\n✨ Otwórz plik w Excelu - wykresy są już w środku!")
