import random
import pandas as pd

def is_valid(numbers):
    # Suma - rozszerzona
    if not (110 <= sum(numbers) <= 185):
        return False
    
    # Max różnica - luzowanie
    differences = [abs(numbers[i] - numbers[i - 1]) for i in range(1, len(numbers))]
    if any(diff > 20 for diff in differences):
        return False
    
    # Suma różnic - bez zmian (najlepsza reguła!)
    if not (27 <= sum(differences) <= 47):
        return False
    
    # Unikaj skrajnej parzystości
    parzyste = sum(1 for n in numbers if n % 2 == 0)
    if parzyste == 0 or parzyste == 6:
        return False
    
    # Unikaj skrajnego rozkładu niskie/wysokie
    niskie = sum(1 for n in numbers if n <= 24)
    if niskie == 0 or niskie == 6:
        return False
    
    return True


file_path = 'C:/Users/mathe/Desktop/lotto/wyniki_lotto.ods'
sheet_name = 'Arkusz1'

while True:  # Pętla, która pozwala na wielokrotne losowanie
    # Generowanie liczb
    result = generate_numbers()
    
    # Sprawdzenie, czy kombinacja istnieje w pliku
    exists = check_combination_exists(result, file_path, sheet_name)
    if exists:
        print("Wylosowana kombinacja już istnieje w pliku:", result)
    else:
        print("Wylosowane liczby (posortowane):", result)  # Liczby są już posortowane
        print("Suma:", sum(result))
        print("Różnice:", [abs(result[i] - result[i - 1]) for i in range(1, len(result))])
        print("Suma różnic:", sum([abs(result[i] - result[i - 1]) for i in range(1, len(result))]))

    user_input = input("Naciśnij Enter, aby wylosować ponownie, lub 'W', aby zakończyć: ")
    if user_input.strip().upper() == 'W':  # Sprawdzenie, czy użytkownik chce zakończyć
        print("Zakończono proces.")
        break  # Zakończenie pętli