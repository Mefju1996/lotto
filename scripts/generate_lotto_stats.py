#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generator statystyk Lotto
Analizuje bazę historycznych wyników i generuje raport XLSX z wykresami PNG
"""

import pandas as pd
import matplotlib.pyplot as plt
from collections import Counter
from itertools import combinations
import numpy as np
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
df = pd.read_excel(file, sheet_name='Arkusz1')
print(f"   ✓ Wczytano {len(df)} losowań")

liczby_cols = ['pierwsza', 'druga', 'trzecia', 'czwarta', 'piąta', 'szósta']
wszystkie_liczby = df[liczby_cols].values.flatten()

# Ostatnie 50 losowań (najnowsze są na górze!)
df_recent = df.head(50).copy()
wszystkie_liczby_recent = df_recent[liczby_cols].values.flatten()

print(f"   ✓ Najnowsze losowanie: {df.iloc[0][liczby_cols].tolist()}")
print(f"   ✓ Najstarsze losowanie w bazie: {df.iloc[-1][liczby_cols].tolist()}")

# ============================================================
# ZAKŁADKA 1: CZĘSTOTLIWOŚĆ LICZB (CAŁA HISTORIA)
# ============================================================
print("\n2. Generowanie statystyk częstotliwości...")
counter = Counter(wszystkie_liczby)
freq_df = pd.DataFrame([
    {'Liczba': int(num), 'Wystąpienia': count, 'Procent': round(count/len(df)*100, 2)}
    for num, count in sorted(counter.items())
])

# Wykres 1: Częstotliwość wszystkich liczb
plt.figure(figsize=(14, 6))
plt.bar(freq_df['Liczba'], freq_df['Wystąpienia'], color='steelblue', alpha=0.7)
plt.xlabel('Liczba', fontsize=12)
plt.ylabel('Liczba wystąpień', fontsize=12)
plt.title(f'Częstotliwość wylosowania liczb 1-49 ({len(df)} losowań)', fontsize=14, fontweight='bold')
plt.xticks(range(1, 50))
plt.grid(axis='y', alpha=0.3)
plt.tight_layout()
plt.savefig('wykres_czestotliwosc.png', dpi=150, bbox_inches='tight')
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

# Wykres 2: Top 10 Hot vs Cold
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
plt.savefig('wykres_hot_cold.png', dpi=150, bbox_inches='tight')
plt.close()
print("   ✓ Wygenerowano wykres_hot_cold.png")

# ============================================================
# ZAKŁADKA 4: ROZKŁAD PARZYSTOŚĆ/NIEPARZYSTOŚĆ
# ============================================================
print("\n3. Generowanie rozkładów parzystość/nieparzystość...")
def count_parzyste(row):
    return sum(1 for x in row if x % 2 == 0)

df['parzyste'] = df[liczby_cols].apply(count_parzyste, axis=1)
df['nieparzyste'] = 6 - df['parzyste']

rozklad_parz = df.groupby('parzyste').size().reset_index(name='Liczba_losowan')
rozklad_parz['Rozkład'] = rozklad_parz['parzyste'].astype(str) + 'P-' + (6 - rozklad_parz['parzyste']).astype(str) + 'N'
rozklad_parz['Procent'] = round(rozklad_parz['Liczba_losowan'] / len(df) * 100, 2)

# Wykres 3: Rozkład parzystość
plt.figure(figsize=(10, 6))
plt.bar(rozklad_parz['Rozkład'], rozklad_parz['Liczba_losowan'], color='mediumseagreen', alpha=0.7)
plt.xlabel('Rozkład (P=Parzyste, N=Nieparzyste)', fontsize=12)
plt.ylabel('Liczba losowań', fontsize=12)
plt.title('Rozkład parzystości w losowaniach', fontsize=14, fontweight='bold')
for i, row in rozklad_parz.iterrows():
    plt.text(i, row['Liczba_losowan'] + 50, f"{row['Procent']}%", ha='center', fontsize=10)
plt.grid(axis='y', alpha=0.3)
plt.tight_layout()
plt.savefig('wykres_parzyste.png', dpi=150, bbox_inches='tight')
plt.close()
print("   ✓ Wygenerowano wykres_parzyste.png")

# ============================================================
# ZAKŁADKA 5: ROZKŁAD NISKIE/WYSOKIE (1-24 vs 25-49)
# ============================================================
print("\n4. Generowanie rozkładów niskie/wysokie...")
def count_niskie(row):
    return sum(1 for x in row if x <= 24)

df['niskie'] = df[liczby_cols].apply(count_niskie, axis=1)
df['wysokie'] = 6 - df['niskie']

rozklad_nh = df.groupby('niskie').size().reset_index(name='Liczba_losowan')
rozklad_nh['Rozkład'] = rozklad_nh['niskie'].astype(str) + 'L-' + (6 - rozklad_nh['niskie']).astype(str) + 'H'
rozklad_nh['Procent'] = round(rozklad_nh['Liczba_losowan'] / len(df) * 100, 2)

# Wykres 4: Rozkład niskie/wysokie
plt.figure(figsize=(10, 6))
plt.bar(rozklad_nh['Rozkład'], rozklad_nh['Liczba_losowan'], color='darkorange', alpha=0.7)
plt.xlabel('Rozkład (L=Niskie 1-24, H=Wysokie 25-49)', fontsize=12)
plt.ylabel('Liczba losowań', fontsize=12)
plt.title('Rozkład niskie/wysokie w losowaniach', fontsize=14, fontweight='bold')
for i, row in rozklad_nh.iterrows():
    plt.text(i, row['Liczba_losowan'] + 50, f"{row['Procent']}%", ha='center', fontsize=10)
plt.grid(axis='y', alpha=0.3)
plt.tight_layout()
plt.savefig('wykres_niskie_wysokie.png', dpi=150, bbox_inches='tight')
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

# Wykres 5: Top 15 par
top_15_pary = pary_df.head(15)
plt.figure(figsize=(12, 7))
plt.barh(top_15_pary['Para'], top_15_pary['Wystąpienia'], color='purple', alpha=0.6)
plt.xlabel('Wystąpienia', fontsize=12)
plt.ylabel('Para liczb', fontsize=12)
plt.title('TOP 15 najczęstszych par liczb', fontsize=12, fontweight='bold')
plt.gca().invert_yaxis()
plt.grid(axis='x', alpha=0.3)
plt.tight_layout()
plt.savefig('wykres_pary.png', dpi=150, bbox_inches='tight')
plt.close()
print("   ✓ Wygenerowano wykres_pary.png")

# ============================================================
# ZAKŁADKA 7 & 8: SPREAD (rozpiętość)
# ============================================================
print("\n6. Generowanie statystyk spreadu...")
df['spread'] = df['szósta'] - df['pierwsza']
spread_stats = df['spread'].describe()

spread_grouped = df.groupby('spread').size().reset_index(name='Liczba_losowan')
spread_grouped['Procent'] = round(spread_grouped['Liczba_losowan'] / len(df) * 100, 2)

# Wykres 6: Histogram spreadu
plt.figure(figsize=(12, 6))
plt.hist(df['spread'], bins=40, color='teal', alpha=0.7, edgecolor='black')
plt.xlabel('Spread (różnica między max a min)', fontsize=12)
plt.ylabel('Liczba losowań', fontsize=12)
plt.title(f'Rozkład spreadu (średnia: {spread_stats["mean"]:.1f})', fontsize=14, fontweight='bold')
plt.axvline(spread_stats['mean'], color='red', linestyle='--', linewidth=2, label=f'Średnia: {spread_stats["mean"]:.1f}')
plt.legend()
plt.grid(axis='y', alpha=0.3)
plt.tight_layout()
plt.savefig('wykres_spread.png', dpi=150, bbox_inches='tight')
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
suma_grouped = df.groupby('Suma').size().reset_index(name='Liczba_losowan')
suma_grouped['Procent'] = round(suma_grouped['Liczba_losowan'] / len(df) * 100, 2)

# Wykres 7: Histogram sum
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
plt.savefig('wykres_suma.png', dpi=150, bbox_inches='tight')
plt.close()
print("   ✓ Wygenerowano wykres_suma.png")

suma_top = suma_grouped.sort_values('Liczba_losowan', ascending=False).head(50).reset_index(drop=True)
suma_top.index += 1
suma_top.index.name = 'Ranking'

# ============================================================
# ZAKŁADKA 11: WALIDACJA REGUŁ ALGORYTMU
# ============================================================
print("\n8. Walidacja reguł algorytmu...")
def check_rules(row):
    nums = row[liczby_cols].values
    suma = row['Suma']

    # Reguła 1: Suma 130-205
    rule1 = 130 <= suma <= 205

    # Reguła 2: Max różnica między parami <= 17
    diffs = [abs(nums[i] - nums[i-1]) for i in range(1, 6)]
    rule2 = all(d <= 17 for d in diffs)

    # Reguła 3: Suma różnic 27-47
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

# Wykres 8: Walidacja reguł
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
plt.savefig('wykres_walidacja.png', dpi=150, bbox_inches='tight')
plt.close()
print("   ✓ Wygenerowano wykres_walidacja.png")

# ============================================================
# ZAKŁADKA 12: OSTATNIE 50 LOSOWAŃ - STATYSTYKI
# ============================================================
print("\n9. Generowanie statystyk dla ostatnich 50 losowań...")

# Częstotliwość w ostatnich 50
counter_recent = Counter(wszystkie_liczby_recent)
freq_recent_df = pd.DataFrame([
    {
        'Liczba': int(num), 
        'Wystąpienia_ostatnie_50': counter_recent.get(num, 0),
        'Wystąpienia_ogółem': counter[num],
        'Różnica': counter_recent.get(num, 0) - (counter[num] / len(df) * 50)
    }
    for num in range(1, 50)
])
freq_recent_df['Różnica'] = freq_recent_df['Różnica'].round(2)
freq_recent_df = freq_recent_df.sort_values('Wystąpienia_ostatnie_50', ascending=False).reset_index(drop=True)
freq_recent_df.index += 1
freq_recent_df.index.name = 'Ranking'

# Hot numbers w ostatnich 50
hot_recent = freq_recent_df.head(20)

# Wykres 9: Hot numbers ostatnie 50 losowań
plt.figure(figsize=(12, 7))
colors_recent = ['darkred' if r > 5 else 'crimson' if r > 2 else 'orange' for r in hot_recent['Różnica']]
plt.barh(hot_recent['Liczba'].astype(str), hot_recent['Wystąpienia_ostatnie_50'], color=colors_recent, alpha=0.7)
plt.xlabel('Wystąpienia w ostatnich 50 losowaniach', fontsize=12)
plt.ylabel('Liczba', fontsize=12)
plt.title('TOP 20 Hot Numbers - ostatnie 50 losowań', fontsize=12, fontweight='bold')
plt.gca().invert_yaxis()
plt.grid(axis='x', alpha=0.3)
plt.tight_layout()
plt.savefig('wykres_hot_recent50.png', dpi=150, bbox_inches='tight')
plt.close()
print("   ✓ Wygenerowano wykres_hot_recent50.png")

# Statystyki porównawcze
recent_stats = pd.DataFrame({
    'Statystyka': [
        'Średnia suma',
        'Średni spread',
        'Średnia suma różnic',
        'Najczęstszy rozkład P/N',
        'Najczęstszy rozkład L/H'
    ],
    'Ostatnie_50': [
        round(df_recent['Suma'].mean(), 2),
        round(df_recent['spread'].mean(), 2),
        round(df_recent.iloc[:, -1].mean(), 2),  # Bezpieczny dostęp do ostatniej kolumny
        str(df_recent.groupby(['parzyste', 'nieparzyste']).size().idxmax()),
        str(df_recent.groupby(['niskie', 'wysokie']).size().idxmax())
    ],
    'Cała_historia': [
        round(df['Suma'].mean(), 2),
        round(df['spread'].mean(), 2),
        round(df.iloc[:, -1].mean(), 2),
        str(df.groupby(['parzyste', 'nieparzyste']).size().idxmax()),
        str(df.groupby(['niskie', 'wysokie']).size().idxmax())
    ]
})

# ============================================================
# ZAKŁADKA 13: COLD STREAKS (liczby nie wylosowane od X losowań)
# ============================================================
print("\n10. Generowanie statystyk cold streaks...")

def find_last_occurrence(liczba):
    """Znajduje, ile losowań temu liczba została ostatnio wylosowana"""
    for idx, row in df[liczby_cols].iterrows():
        if liczba in row.values:
            return idx
    return len(df)

cold_streaks = []
for num in range(1, 50):
    last_idx = find_last_occurrence(num)
    cold_streaks.append({
        'Liczba': num,
        'Ostatnio_losowana': last_idx,
        'Losowań_temu': last_idx,
        'Status': 'HOT' if last_idx < 10 else 'MEDIUM' if last_idx < 30 else 'COLD'
    })

cold_streaks_df = pd.DataFrame(cold_streaks).sort_values('Losowań_temu', ascending=False).reset_index(drop=True)
cold_streaks_df.index += 1
cold_streaks_df.index.name = 'Ranking'

# TOP 10 najbardziej "zimnych" liczb
top_cold_streaks = cold_streaks_df.head(10)

# Wykres 10: Cold streaks
plt.figure(figsize=(10, 7))
plt.barh(top_cold_streaks['Liczba'].astype(str), top_cold_streaks['Losowań_temu'], color='dodgerblue', alpha=0.7)
plt.xlabel('Liczba losowań od ostatniego wystąpienia', fontsize=12)
plt.ylabel('Liczba', fontsize=12)
plt.title('TOP 10 Cold Streaks - najdłużej nie wylosowane liczby', fontsize=12, fontweight='bold')
plt.gca().invert_yaxis()
plt.grid(axis='x', alpha=0.3)
plt.tight_layout()
plt.savefig('wykres_cold_streaks.png', dpi=150, bbox_inches='tight')
plt.close()
print("   ✓ Wygenerowano wykres_cold_streaks.png")

# ============================================================
# ZAPISANIE DO PLIKU XLSX
# ============================================================
print("\n11. Zapisywanie danych do pliku Excel...")
output_file = ROOT_DIR / 'analysis' / 'statystyki_lotto.xlsx'

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

print(f"   ✓ Zapisano plik: {output_file}")

# ============================================================
# PODSUMOWANIE
# ============================================================
print("\n" + "="*70)
print("✅ GENERATOR STATYSTYK LOTTO - ZAKOŃCZONO POMYŚLNIE")
print("="*70)
print(f"\n📊 Wygenerowano plik: {output_file}")
print(f"   Zawiera 13 zakładek ze statystykami\n")

print("📈 Zakładki:")
print("   1. Częstotliwość liczb 1-49")
print("   2. TOP 20 Hot Numbers")
print("   3. TOP 20 Cold Numbers")
print("   4. Rozkład parzystość/nieparzystość")
print("   5. Rozkład niskie/wysokie")
print("   6. TOP 50 par liczb")
print("   7. Statystyki spreadu")
print("   8. Rozkład spreadu (TOP 30)")
print("   9. TOP 50 sum")
print("  10. Walidacja reguł algorytmu ⭐")
print("  11. Hot Numbers - ostatnie 50 losowań 🔥")
print("  12. Statystyki porównawcze ostatnie 50 vs cała historia 🆕")
print("  13. Cold Streaks - najdłużej nie wylosowane ❄️")

print(f"\n🖼️  Wygenerowano 10 wykresów PNG:")
wykres_list = [
    'wykres_czestotliwosc.png',
    'wykres_hot_cold.png',
    'wykres_parzyste.png',
    'wykres_niskie_wysokie.png',
    'wykres_pary.png',
    'wykres_spread.png',
    'wykres_suma.png',
    'wykres_walidacja.png',
    'wykres_hot_recent50.png',
    'wykres_cold_streaks.png'
]
for w in wykres_list:
    print(f"   • {w}")

print("\n✨ Gotowe do analizy!")
