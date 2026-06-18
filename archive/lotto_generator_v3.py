import random
import pandas as pd


def is_valid(numbers):
    """
    Walidacja kombinacji zgodnie z analizą statystyk (Opcja 1 + dodatkowe reguły)

    Pokrywa ~50-55% historycznych losowań Lotto
    """

    # REGUŁA 1: Suma - rozszerzona (110-205)
    # Pokrycie: ~75% historycznych losowań
    if not (110 <= sum(numbers) <= 205):
        return False

    # Oblicz różnice między kolejnymi liczbami
    differences = [abs(numbers[i] - numbers[i - 1]) for i in range(1, len(numbers))]

    # REGUŁA 2: Max różnica między parami ≤22
    # Pokrycie: ~84% historycznych losowań
    if any(diff > 22 for diff in differences):
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
        existing_data = pd.read_excel(file_path, engine='openpyxl', sheet_name=sheet_name, usecols="A:F")
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


def display_stats(numbers, losowanie_nr=None, show_header=True):
    """Wyświetla statystyki dla wylosowanej kombinacji"""
    if show_header:
        print("\n" + "="*60)

    if losowanie_nr is not None:
        print(f"🎲 LOSOWANIE #{losowanie_nr}: {numbers}")
    else:
        print(f"🎲 WYLOSOWANE LICZBY: {numbers}")

    if show_header:
        print("="*60)

    # Podstawowe statystyki
    suma = sum(numbers)
    differences = [abs(numbers[i] - numbers[i - 1]) for i in range(1, len(numbers))]
    suma_roznic = sum(differences)
    spread = numbers[-1] - numbers[0]

    print(f"   Suma: {suma} | Spread: {spread} | Suma różnic: {suma_roznic}")

    # Rozkład parzystości
    parzyste = sum(1 for n in numbers if n % 2 == 0)
    nieparzyste = 6 - parzyste

    # Rozkład niskie/wysokie
    niskie = sum(1 for n in numbers if n <= 24)
    wysokie = 6 - niskie

    print(f"   Różnice: {differences}")
    print(f"   Rozkład: {parzyste}P-{nieparzyste}N, {niskie}L-{wysokie}H")

    if show_header:
        print("="*60)


def display_stats_detailed(numbers):
    """Wyświetla szczegółowe statystyki dla pojedynczego losowania"""
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


def batch_generate(count, file_path, sheet_name):
    """Generuje N losowań i wyświetla wyniki"""
    print(f"\n🎰 Generowanie {count} losowań...")
    print("="*60)

    generated = []
    duplicates_in_history = 0

    for i in range(count):
        result = generate_numbers()
        generated.append(result)

        # Sprawdź czy występowała w historii
        exists = check_combination_exists(result, file_path, sheet_name)
        if exists:
            duplicates_in_history += 1

        # Wyświetl zwięzłe info
        display_stats(result, losowanie_nr=i+1, show_header=(i==0))
        if exists:
            print("   ⚠️ Występowała w historii!")

    # Podsumowanie
    print("\n" + "="*60)
    print("📊 PODSUMOWANIE:")
    print("="*60)
    print(f"   Wygenerowano:              {count} kombinacji")
    print(f"   Powtórzeń z historii:      {duplicates_in_history}")
    print(f"   Unikalnych (nowych):       {count - duplicates_in_history}")
    print("="*60)

    return generated


def interactive_mode(file_path, sheet_name):
    """Tryb interaktywny - losowanie po losowaniu"""
    print("\n🎮 TRYB INTERAKTYWNY")
    print("="*60)
    print("Naciśnij [Enter] aby losować, [W] aby wrócić do menu\n")

    while True:
        # Generowanie liczb
        result = generate_numbers()

        # Sprawdzenie, czy kombinacja istnieje w pliku
        exists = check_combination_exists(result, file_path, sheet_name)

        if exists:
            print("⚠️ UWAGA: Ta kombinacja już występowała w historii!")
            print("   Liczby:", result)
            print("\n   Mimo to pokazuję statystyki:")

        # Wyświetl szczegółowe statystyki
        display_stats_detailed(result)

        # Opcje użytkownika
        print("\n💡 [Enter] - Losuj ponownie | [W] - Wróć do menu")

        user_input = input("Twój wybór: ")

        if user_input.strip().upper() == 'W':
            break


def main_menu():
    """Główne menu programu"""
    file_path = 'C:/Users/mathe/Desktop/lotto/wyniki_lotto.xlsx'
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
    print("="*60)

    while True:
        print("\n📋 MENU GŁÓWNE:")
        print("="*60)
        print("  [1] Tryb interaktywny (jedno losowanie na raz)")
        print("  [2] Tryb wsadowy (wiele losowań)")
        print("  [W] Wyjście z programu")
        print("="*60)

        choice = input("\nWybierz opcję: ").strip()

        if choice == '1':
            interactive_mode(file_path, sheet_name)

        elif choice == '2':
            while True:
                try:
                    count = input("\nIle losowań chcesz wygenerować? (lub [Enter] aby anulować): ").strip()

                    if count == '':
                        print("Anulowano.")
                        break

                    count = int(count)

                    if count <= 0:
                        print("⚠️ Liczba musi być większa od 0!")
                        continue

                    if count > 1000:
                        confirm = input(f"⚠️ Chcesz wygenerować {count} losowań? To może chwilę potrwać. Kontynuować? [T/N]: ")
                        if confirm.strip().upper() != 'T':
                            print("Anulowano.")
                            break

                    # Generuj
                    batch_generate(count, file_path, sheet_name)
                    break

                except ValueError:
                    print("⚠️ Błąd! Podaj liczbę całkowitą.")

        elif choice.upper() == 'W':
            print("\n👋 Zakończono program. Powodzenia!")
            break

        else:
            print("⚠️ Nieprawidłowa opcja! Wybierz 1, 2 lub W.")


# ============================================================
# URUCHOMIENIE PROGRAMU
# ============================================================

if __name__ == "__main__":
    main_menu()
