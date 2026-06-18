import random

def is_valid(numbers):
    # Sprawdzenie, czy suma liczb mieści się w przedziale 130-168
    if not (130 <= sum(numbers) <= 205):
        return False
    
    # Sprawdzenie różnicy pomiędzy parami
    differences = [abs(numbers[i] - numbers[i - 1]) for i in range(1, len(numbers))]
    if any(diff > 17 for diff in differences):
        return False
    
    # Sprawdzenie sumy różnic
    if not (27 <= sum(differences) <= 47):
        return False
    
    return True

def generate_numbers():
    while True:
        # Losowanie 6 unikalnych liczb z przedziału 1-49
        numbers = sorted(random.sample(range(1, 50), 6))
        
        if is_valid(numbers):
            return numbers

while True:  # Pętla, która pozwala na wielokrotne losowanie
    # Generowanie liczb
    result = generate_numbers()
    print("Wylosowane liczby (posortowane):", result)  # Liczby są już posortowane
    print("Suma:", sum(result))
    print("Różnice:", [abs(result[i] - result[i - 1]) for i in range(1, len(result))])
    print("Suma różnic:", sum([abs(result[i] - result[i - 1]) for i in range(1, len(result))]))

    user_input = input("Naciśnij Enter, aby wylosować ponownie, lub 'W', aby zakończyć: ")
    if user_input.strip().upper() == 'W':  # Sprawdzenie, czy użytkownik chce zakończyć
        print("Zakończono proces.")
        break  # Zakończenie pętli
