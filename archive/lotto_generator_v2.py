import random
import pandas as pd


def is_valid(numbers):
    """
    Walidacja kombinacji zgodnie z analizą statystyk (Opcja 1 + dodatkowe reguły)

    Pokrywa ~50-55% historycznych losowań Lotto
    """

    # REGUŁA 1: Suma - rozszerzona (110-185)
    # Pokrycie: ~75% historycznych losowań
    if not (110 <= sum(numbers) <= 185):
        return False

    # Oblicz różnice między kolejnymi liczbami
    differences = [abs(numbers[i] - numbers[i - 1]) for i in range(1, len(numbers))]

    # REGUŁA 2: Max różnica między parami ≤20
    # Pokrycie: ~84% historycznych losowań
    if any(diff > 20 for diff in differences):
        return False

    # REGUŁA 3: Suma różnic (27-47) - NAJLEPSZA REGUŁA!
    # Pokrycie: ~86% historycznych losowań
    if not (27 <= sum(differences) <= 47):
        return False

    # REGUŁA 4: Unikaj skrajnej parzystości (0P-6N lub 6P-0N)
    # Te przypadki to tylko 2.6% historycznych losowań
    parzyste = sum(1 for n in numbers if n % 2 == 0)
    if parzyste == 0 or parzyste == 6:
        return False

    # REGUŁA 5: Unikaj skrajnego rozkładu niskie/wysokie (0L-6H lub 6L-0H)
    # Te przypadki to tylko 2.4% historycznych losowań
    niskie = sum(1 for n in numbers if n <= 24)
    if niskie == 0 or niskie == 6:
        return False

    return True


def check_combination_exists(numbers, file_path, sheet_name):
    """Sprawdza, czy kombinacja już występowała w historii"""
    try:
        # Wczytanie danych z pliku
        existing_data = pd.read_excel(file_path, engine='odf', sheet_name=sheet_name, usecols="A:F")
        # Sprawdzenie, czy kombinacja istnieje
        return any((existing_data.values == numbers).all(axis=1))
    except FileNotFoundError:
        print(f"⚠️ Plik {file_path} nie został znaleziony. Sprawdzanie pomijane.")
        return False


def generate_numbers():
    """Generuje liczby spełniające wszystkie reguły walidacji"""
    while True:
        # Losowanie 6 unikalnych liczb z przedziału 1-49
        numbers = sorted(random.sample(range(1, 50), 6))

        if is_valid(numbers):
            return numbers


def display_stats(numbers):
    """Wyświetla statystyki dla wylosowanej kombinacji"""
    print("\n" + "="*60)
    print("🎲 WYLOSOWANE LICZBY (posortowane):", numbers)
    print("="*60)

    # Podstawowe statystyki
    suma = sum(numbers)
    differences = [abs(numbers[i] - numbers[i - 1]) for i in range(1, len(numbers))]
    suma_roznic = sum(differences)
    spread = numbers[-1] - numbers[0]

    print(f"📊 Suma:                {suma}")
    print(f"📏 Spread (rozpiętość): {spread}")
    print(f"🔢 Różnice:             {differences}")
    print(f"➕ Suma różnic:         {suma_roznic}")

    # Rozkład parzystości
    parzyste = sum(1 for n in numbers if n % 2 == 0)
    nieparzyste = 6 - parzyste
    print(f"⚖️ Rozkład P/N:          {parzyste}P-{nieparzyste}N")

    # Rozkład niskie/wysokie
    niskie = sum(1 for n in numbers if n <= 24)
    wysokie = 6 - niskie
    print(f"📊 Rozkład L/H:          {niskie}L-{wysokie}H")

    print("="*60)


# ============================================================
# GŁÓWNA PĘTLA PROGRAMU
# ============================================================

file_path = 'C:/Users/mathe/Desktop/lotto/wyniki_lotto.ods'
sheet_name = 'Arkusz1'

print("\n" + "="*60)
print("🎰 GENERATOR LICZB LOTTO - WERSJA ULEPSZONA")
print("="*60)
print("Reguły walidacji:")
print("  ✓ Suma: 110-185 (pokrycie ~75%)")
print("  ✓ Max różnica: ≤20 (pokrycie ~84%)")
print("  ✓ Suma różnic: 27-47 (pokrycie ~86%)")
print("  ✓ Brak skrajnej parzystości (0P/6P)")
print("  ✓ Brak skrajnego rozkładu L/H (0L/6H)")
print("="*60 + "\n")

while True:
    # Generowanie liczb
    result = generate_numbers()

    # Sprawdzenie, czy kombinacja istnieje w pliku
    exists = check_combination_exists(result, file_path, sheet_name)

    if exists:
        print("⚠️ UWAGA: Ta kombinacja już występowała w historii!")
        print("   Liczby:", result)
        print("\n   Mimo to pokazuję statystyki:")

    # Wyświetl statystyki
    display_stats(result)

    # Opcje użytkownika
    print("\n💡 Opcje:")
    print("   [Enter] - Wylosuj nową kombinację")
    print("   [W]     - Wyjdź z programu")

    user_input = input("\nTwój wybór: ")

    if user_input.strip().upper() == 'W':
        print("\n👋 Zakończono proces. Powodzenia!")
        break

    print("\n")  # Odstęp przed następnym losowaniem
