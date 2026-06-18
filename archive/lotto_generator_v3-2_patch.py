import random
from pathlib import Path
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox


BASE_DIR = Path(__file__).resolve().parent
HISTORY_FILE = BASE_DIR / 'statystyki_lotto.xlsx'
DEFAULT_HISTORY_SHEET = 'Arkusz1'
STATISTICS_FILE = BASE_DIR / 'statystyki_lotto.xlsx'


def select_history_file(current_path):
    prompt = f"Ścieżka pliku z historią losowań
[Enter] = zostaw bieżący
Aktualnie: {current_path}
Nowa ścieżka: "
    user_value = input(prompt).strip()
    if user_value:
        new_path = Path(user_value).expanduser()
        return str(new_path)
    return str(current_path)


def select_history_sheet(current_sheet):
    user_value = input(f"Nazwa arkusza z historią [Enter = {current_sheet}]: ").strip()
    return user_value or current_sheet


def select_history_file_gui(current_path):
    selected = filedialog.askopenfilename(
        title='Wybierz plik z historią losowań',
        initialdir=str(Path(current_path).resolve().parent),
        initialfile=Path(current_path).name,
        filetypes=[('Excel', '*.xlsx *.xls'), ('Wszystkie pliki', '*.*')]
    )
    return selected or str(current_path)


class LottoStatistics:
    def __init__(self, xlsx_path):
        self.xlsx_path = Path(xlsx_path)
        self.frequency = {}
        self.hot_rank = {}
        self.cold_rank = {}
        self.top_pairs = {}
        self.rolling = {}
        self.year_heatmap = {}
        self.years = []
        self._load()

    def _safe_read(self, sheet_name):
        return pd.read_excel(self.xlsx_path, sheet_name=sheet_name, engine='openpyxl')

    def _load(self):
        freq_df = self._safe_read('1_Czestotliwosc')
        for _, row in freq_df.iterrows():
            number = int(row['Liczba'])
            self.frequency[number] = {
                'wystapienia': int(row['Wystąpienia']),
                'procent': float(row['Procent']),
            }

        hot_df = self._safe_read('2_Hot_Numbers')
        for _, row in hot_df.iterrows():
            self.hot_rank[int(row['Liczba'])] = int(row['Ranking'])

        cold_df = self._safe_read('3_Cold_Numbers')
        for _, row in cold_df.iterrows():
            self.cold_rank[int(row['Liczba'])] = int(row['Ranking'])

        pairs_df = self._safe_read('6_TOP50_Par')
        for _, row in pairs_df.iterrows():
            a, b = map(int, str(row['Para']).split('-'))
            pair_entry = {'para': f'{a}-{b}', 'wystapienia': int(row['Wystąpienia'])}
            self.top_pairs.setdefault(a, []).append(pair_entry)
            self.top_pairs.setdefault(b, []).append(pair_entry)

        rolling_df = self._safe_read('16RollingFreq')
        for col in rolling_df.columns:
            if col == 'Indekslosowania':
                continue
            if str(col).startswith('Liczba') and str(col).endswith('roll100'):
                number = int(str(col).replace('Liczba', '').replace('roll100', ''))
                series = rolling_df[col].dropna().tolist()
                self.rolling[number] = [int(v) for v in series]

        heatmap_df = self._safe_read('17_Heatmapa_Rok')
        self.years = [str(col) for col in heatmap_df.columns if str(col) != 'Liczba']
        for _, row in heatmap_df.iterrows():
            number = int(row['Liczba'])
            self.year_heatmap[number] = {year: int(row[year]) for year in self.years if pd.notna(row[year])}

    def get_status(self, number):
        if number in self.hot_rank:
            return f'HOT #{self.hot_rank[number]}'
        if number in self.cold_rank:
            return f'COLD #{self.cold_rank[number]}'
        return 'NEUTRALNA'

    def get_number_stats(self, number):
        freq = self.frequency.get(number, {'wystapienia': 0, 'procent': 0.0})
        rolling = self.rolling.get(number, [])
        yearly = self.year_heatmap.get(number, {})
        pairs = sorted(self.top_pairs.get(number, []), key=lambda x: x['wystapienia'], reverse=True)[:5]
        return {
            'number': number,
            'wystapienia': freq['wystapienia'],
            'procent': freq['procent'],
            'status': self.get_status(number),
            'rolling': rolling,
            'rolling_current': rolling[-1] if rolling else None,
            'rolling_min': min(rolling) if rolling else None,
            'rolling_max': max(rolling) if rolling else None,
            'pairs': pairs,
            'years': yearly,
            'last_years': list(yearly.items())[-8:] if yearly else [],
        }


def is_valid(numbers):
    if not (110 <= sum(numbers) <= 185):
        return False
    differences = [abs(numbers[i] - numbers[i - 1]) for i in range(1, len(numbers))]
    if any(diff > 20 for diff in differences):
        return False
    if not (27 <= sum(differences) <= 47):
        return False
    parzyste = sum(1 for n in numbers if n % 2 == 0)
    if parzyste == 0 or parzyste == 6:
        return False
    niskie = sum(1 for n in numbers if n <= 24)
    if niskie == 0 or niskie == 6:
        return False
    return True


def check_combination_exists(numbers, file_path, sheet_name):
    try:
        existing_data = pd.read_excel(file_path, engine='openpyxl', sheet_name=sheet_name, usecols='A:F')
        return any((existing_data.values == numbers).all(axis=1))
    except FileNotFoundError:
        print(f'⚠️ Plik {file_path} nie został znaleziony. Sprawdzanie pomijane.')
        return False
    except ValueError:
        return False


def generate_numbers():
    while True:
        numbers = sorted(random.sample(range(1, 50), 6))
        if is_valid(numbers):
            return numbers


def calculate_draw_stats(numbers):
    suma = sum(numbers)
    differences = [abs(numbers[i] - numbers[i - 1]) for i in range(1, len(numbers))]
    suma_roznic = sum(differences)
    spread = numbers[-1] - numbers[0]
    parzyste = sum(1 for n in numbers if n % 2 == 0)
    nieparzyste = 6 - parzyste
    niskie = sum(1 for n in numbers if n <= 24)
    wysokie = 6 - niskie
    return {
        'suma': suma,
        'spread': spread,
        'roznice': differences,
        'suma_roznic': suma_roznic,
        'parzyste': parzyste,
        'nieparzyste': nieparzyste,
        'niskie': niskie,
        'wysokie': wysokie,
    }


def display_stats(numbers, losowanie_nr=None, show_header=True):
    if show_header:
        print('\n' + '=' * 60)
    if losowanie_nr is not None:
        print(f'🎲 LOSOWANIE #{losowanie_nr}: {numbers}')
    else:
        print(f'🎲 WYLOSOWANE LICZBY: {numbers}')
    if show_header:
        print('=' * 60)
    stats = calculate_draw_stats(numbers)
    print(f" Suma: {stats['suma']} | Spread: {stats['spread']} | Suma różnic: {stats['suma_roznic']}")
    print(f" Różnice: {stats['roznice']}")
    print(f" Rozkład: {stats['parzyste']}P-{stats['nieparzyste']}N, {stats['niskie']}L-{stats['wysokie']}H")
    if show_header:
        print('=' * 60)


def display_stats_detailed(numbers):
    print('\n' + '=' * 60)
    print('🎲 WYLOSOWANE LICZBY (posortowane):', numbers)
    print('=' * 60)
    stats = calculate_draw_stats(numbers)
    print(f"📊 Suma: {stats['suma']}")
    print(f"📏 Spread (rozpiętość): {stats['spread']}")
    print(f"🔢 Różnice: {stats['roznice']}")
    print(f"➕ Suma różnic: {stats['suma_roznic']}")
    print(f"⚖️ Rozkład P/N: {stats['parzyste']}P-{stats['nieparzyste']}N")
    print(f"📊 Rozkład L/H: {stats['niskie']}L-{stats['wysokie']}H")
    print('=' * 60)


class LottoApp:
    def __init__(self, master, stats_provider, history_file=HISTORY_FILE, history_sheet=DEFAULT_HISTORY_SHEET):
        self.master = master
        self.stats_provider = stats_provider
        self.history_file = history_file
        self.history_sheet = history_sheet
        self.current_numbers = []
        self.number_buttons = []
        self.master.title('Generator Lotto + statystyki liczb')
        self.master.geometry('1180x700')
        self.master.minsize(1000, 620)
        self._build_ui()
        self.draw_numbers()

    def _build_ui(self):
        self.master.configure(bg='#0f172a')
        main = tk.Frame(self.master, bg='#0f172a')
        main.pack(fill='both', expand=True, padx=16, pady=16)

        left = tk.Frame(main, bg='#111827')
        left.pack(side='left', fill='both', expand=True, padx=(0, 10))

        right = tk.Frame(main, bg='#111827', width=360)
        right.pack(side='right', fill='y')
        right.pack_propagate(False)

        header = tk.Label(left, text='Generator Lotto', font=('Segoe UI', 22, 'bold'), fg='#f8fafc', bg='#111827')
        header.pack(anchor='w', padx=20, pady=(20, 8))

        sub = tk.Label(left, text='Kliknij wylosowaną liczbę, aby zobaczyć statystyki z pliku Excel.', font=('Segoe UI', 11), fg='#cbd5e1', bg='#111827')
        sub.pack(anchor='w', padx=20)

        controls = tk.Frame(left, bg='#111827')
        controls.pack(fill='x', padx=20, pady=18)
        tk.Button(controls, text='🎲 Losuj', command=self.draw_numbers, font=('Segoe UI', 12, 'bold'), bg='#0ea5e9', fg='white', activebackground='#0284c7', relief='flat', padx=18, pady=10).pack(side='left')

        self.history_label = tk.Label(controls, text='', font=('Segoe UI', 11), fg='#fbbf24', bg='#111827')
        self.history_label.pack(side='left', padx=20)

        balls_frame = tk.Frame(left, bg='#111827')
        balls_frame.pack(fill='x', padx=20, pady=(10, 16))
        for i in range(6):
            btn = tk.Button(
                balls_frame,
                text='--',
                font=('Segoe UI', 18, 'bold'),
                width=4,
                height=2,
                relief='flat',
                bg='#1d4ed8',
                fg='white',
                activebackground='#2563eb',
                cursor='hand2',
                command=lambda idx=i: self.show_number_details(self.current_numbers[idx] if self.current_numbers else None)
            )
            btn.pack(side='left', padx=8, pady=4)
            self.number_buttons.append(btn)

        self.draw_stats_var = tk.StringVar(value='')
        stats_box = tk.Label(left, textvariable=self.draw_stats_var, justify='left', anchor='nw', font=('Consolas', 13), fg='#e2e8f0', bg='#0b1220', padx=16, pady=16)
        stats_box.pack(fill='x', padx=20, pady=(0, 16))

        self.canvas = tk.Canvas(left, bg='#0b1220', height=130, highlightthickness=0)
        self.canvas.pack(fill='x', padx=20, pady=(0, 20))

        panel_title = tk.Label(right, text='Szczegóły liczby', font=('Segoe UI', 18, 'bold'), fg='#f8fafc', bg='#111827')
        panel_title.pack(anchor='w', padx=18, pady=(20, 8))

        self.details_var = tk.StringVar(value='Kliknij jedną z wylosowanych liczb.')
        details = tk.Label(right, textvariable=self.details_var, justify='left', anchor='nw', font=('Segoe UI', 11), fg='#e2e8f0', bg='#111827', wraplength=320)
        details.pack(fill='x', padx=18, pady=(0, 12))

        self.spark = tk.Canvas(right, bg='#0b1220', height=120, highlightthickness=0)
        self.spark.pack(fill='x', padx=18, pady=(0, 12))

        self.year_var = tk.StringVar(value='')
        year_label = tk.Label(right, textvariable=self.year_var, justify='left', anchor='nw', font=('Consolas', 10), fg='#cbd5e1', bg='#0b1220', padx=12, pady=12, wraplength=320)
        year_label.pack(fill='x', padx=18, pady=(0, 12))

        self.pairs_var = tk.StringVar(value='')
        pairs_label = tk.Label(right, textvariable=self.pairs_var, justify='left', anchor='nw', font=('Segoe UI', 10), fg='#cbd5e1', bg='#0b1220', padx=12, pady=12, wraplength=320)
        pairs_label.pack(fill='x', padx=18, pady=(0, 18))

    def draw_numbers(self):
        numbers = generate_numbers()
        self.current_numbers = numbers
        exists = check_combination_exists(numbers, self.history_file, self.history_sheet)
        self.history_label.config(text='⚠️ kombinacja była już w historii' if exists else '✅ nowa kombinacja')
        for btn, num in zip(self.number_buttons, numbers):
            btn.config(text=str(num))
        stats = calculate_draw_stats(numbers)
        self.draw_stats_var.set(
            f"Suma: {stats['suma']}\n"
            f"Spread: {stats['spread']}\n"
            f"Różnice: {stats['roznice']}\n"
            f"Suma różnic: {stats['suma_roznic']}\n"
            f"Parzyste/Nieparzyste: {stats['parzyste']}P-{stats['nieparzyste']}N\n"
            f"Niskie/Wysokie: {stats['niskie']}L-{stats['wysokie']}H"
        )
        self._draw_distribution(numbers)
        self.show_number_details(numbers[0])

    def _draw_distribution(self, numbers):
        self.canvas.delete('all')
        w = max(self.canvas.winfo_width(), 700)
        h = int(self.canvas['height'])
        margin = 26
        self.canvas.create_text(margin, 16, text='Rozkład liczb 1–49', anchor='w', fill='#93c5fd', font=('Segoe UI', 11, 'bold'))
        y = 80
        self.canvas.create_line(margin, y, w - margin, y, fill='#334155', width=2)
        for num in range(1, 50):
            x = margin + (w - 2 * margin) * (num - 1) / 48
            color = '#22c55e' if num in numbers else '#64748b'
            radius = 10 if num in numbers else 4
            self.canvas.create_oval(x - radius, y - radius, x + radius, y + radius, fill=color, outline='')
            if num in numbers:
                self.canvas.create_text(x, y - 24, text=str(num), fill='#f8fafc', font=('Segoe UI', 10, 'bold'))

    def show_number_details(self, number):
        if number is None:
            return
        stats = self.stats_provider.get_number_stats(number)
        rolling = stats['rolling']
        rolling_text = 'brak danych'
        if rolling:
            rolling_text = f"{stats['rolling_current']} (min {stats['rolling_min']}, max {stats['rolling_max']})"
        self.details_var.set(
            f"Liczba: {stats['number']}\n"
            f"Status: {stats['status']}\n"
            f"Wystąpienia historyczne: {stats['wystapienia']}\n"
            f"Udział procentowy: {stats['procent']:.2f}%\n"
            f"Rolling100: {rolling_text}"
        )
        last_years = stats['last_years']
        if last_years:
            year_lines = ['Ostatnie lata:']
            year_lines.extend([f"{year}: {value}" for year, value in last_years])
            self.year_var.set('\n'.join(year_lines))
        else:
            self.year_var.set('Brak danych rocznych.')

        if stats['pairs']:
            pair_lines = ['TOP pary z tą liczbą:']
            pair_lines.extend([f"• {p['para']}  —  {p['wystapienia']} razy" for p in stats['pairs']])
            self.pairs_var.set('\n'.join(pair_lines))
        else:
            self.pairs_var.set('Brak tej liczby w TOP50 par.')

        self._draw_sparkline(rolling)

    def _draw_sparkline(self, values):
        self.spark.delete('all')
        w = max(self.spark.winfo_width(), 320)
        h = int(self.spark['height'])
        self.spark.create_text(12, 14, text='Trend rolling100', anchor='w', fill='#93c5fd', font=('Segoe UI', 10, 'bold'))
        if not values or len(values) < 2:
            self.spark.create_text(w / 2, h / 2, text='Brak danych do wykresu', fill='#94a3b8', font=('Segoe UI', 11))
            return
        points = values[-80:]
        vmin, vmax = min(points), max(points)
        if vmin == vmax:
            vmax += 1
        margin_x = 14
        top = 28
        bottom = h - 14
        width = w - 2 * margin_x
        self.spark.create_rectangle(0, 0, w, h, outline='')
        coords = []
        for i, value in enumerate(points):
            x = margin_x + width * i / (len(points) - 1)
            y = bottom - (value - vmin) * (bottom - top) / (vmax - vmin)
            coords.extend([x, y])
        self.spark.create_line(*coords, fill='#22c55e', width=2, smooth=True)
        self.spark.create_text(margin_x, bottom + 2, text=str(points[0]), anchor='sw', fill='#94a3b8', font=('Segoe UI', 8))
        self.spark.create_text(w - margin_x, bottom + 2, text=str(points[-1]), anchor='se', fill='#94a3b8', font=('Segoe UI', 8))
        self.spark.create_text(w - margin_x, 14, text=f'max {max(points)}', anchor='ne', fill='#94a3b8', font=('Segoe UI', 8))
        self.spark.create_text(margin_x, 14, text=f'min {min(points)}', anchor='nw', fill='#94a3b8', font=('Segoe UI', 8))


def launch_gui(history_file, history_sheet):
    if not STATISTICS_FILE.exists():
        raise FileNotFoundError(f'Brak pliku statystyk: {STATISTICS_FILE}')
    root = tk.Tk()
    root.withdraw()
    chosen_history_file = select_history_file_gui(history_file)
    root.deiconify()
    try:
        stats_provider = LottoStatistics(STATISTICS_FILE)
    except Exception as exc:
        messagebox.showerror('Błąd wczytywania statystyk', str(exc))
        root.destroy()
        return history_file, history_sheet
    app = LottoApp(root, stats_provider, chosen_history_file, history_sheet)
    root.mainloop()
    return chosen_history_file, history_sheet


def interactive_mode(file_path, sheet_name):
    print('\n🎮 TRYB INTERAKTYWNY')
    print('=' * 60)
    print('Naciśnij [Enter] aby losować, [W] aby wrócić do menu\n')
    while True:
        result = generate_numbers()
        exists = check_combination_exists(result, file_path, sheet_name)
        if exists:
            print('⚠️ UWAGA: Ta kombinacja już występowała w historii!')
            print(' Liczby:', result)
            print('\n Mimo to pokazuję statystyki:')
        display_stats_detailed(result)
        print('\n💡 [Enter] - Losuj ponownie | [W] - Wróć do menu')
        user_input = input('Twój wybór: ')
        if user_input.strip().upper() == 'W':
            break


def batch_generate(count, file_path, sheet_name):
    print(f'\n🎰 Generowanie {count} losowań...')
    print('=' * 60)
    generated = []
    duplicates_in_history = 0
    for i in range(count):
        result = generate_numbers()
        generated.append(result)
        exists = check_combination_exists(result, file_path, sheet_name)
        if exists:
            duplicates_in_history += 1
        display_stats(result, losowanie_nr=i + 1, show_header=(i == 0))
        if exists:
            print(' ⚠️ Występowała w historii!')
    print('\n' + '=' * 60)
    print('📊 PODSUMOWANIE:')
    print('=' * 60)
    print(f' Wygenerowano: {count} kombinacji')
    print(f' Powtórzeń z historii: {duplicates_in_history}')
    print(f' Unikalnych (nowych): {count - duplicates_in_history}')
    print('=' * 60)
    return generated


def main_menu():
    file_path = str(HISTORY_FILE)
    sheet_name = DEFAULT_HISTORY_SHEET
    print('\n' + '=' * 60)
    print('🎰 GENERATOR LICZB LOTTO - WERSJA ULEPSZONA')
    print('=' * 60)
    print('Reguły walidacji:')
    print(' ✓ Suma: 110-185 (pokrycie ~75%)')
    print(' ✓ Max różnica: ≤20 (pokrycie ~84%)')
    print(' ✓ Suma różnic: 27-47 (pokrycie ~86%)')
    print(' ✓ Brak skrajnej parzystości (0P/6P)')
    print(' ✓ Brak skrajnego rozkładu L/H (0L/6H)')
    print('=' * 60)
    print(f'Plik historii: {file_path}')
    print(f'Arkusz historii: {sheet_name}')
    while True:
        print('\n📋 MENU GŁÓWNE:')
        print('=' * 60)
        print(' [1] Tryb interaktywny (konsola)')
        print(' [2] Tryb wsadowy (wiele losowań)')
        print(' [3] Tryb okienkowy z klikaniem liczb')
        print(' [4] Wybierz plik/arkusz historii')
        print(' [W] Wyjście z programu')
        print('=' * 60)
        choice = input('\nWybierz opcję: ').strip()
        if choice == '1':
            interactive_mode(file_path, sheet_name)
        elif choice == '2':
            while True:
                try:
                    count = input('\nIle losowań chcesz wygenerować? (lub [Enter] aby anulować): ').strip()
                    if count == '':
                        print('Anulowano.')
                        break
                    count = int(count)
                    if count <= 0:
                        print('⚠️ Liczba musi być większa od 0!')
                        continue
                    if count > 1000:
                        confirm = input(f'⚠️ Chcesz wygenerować {count} losowań? To może chwilę potrwać. Kontynuować? [T/N]: ')
                        if confirm.strip().upper() != 'T':
                            print('Anulowano.')
                            break
                    batch_generate(count, file_path, sheet_name)
                    break
                except ValueError:
                    print('⚠️ Błąd! Podaj liczbę całkowitą.')
        elif choice == '3':
            file_path, sheet_name = launch_gui(file_path, sheet_name)
        elif choice == '4':
            file_path = select_history_file(file_path)
            sheet_name = select_history_sheet(sheet_name)
            print(f'✅ Ustawiono plik historii: {file_path}')
            print(f'✅ Ustawiono arkusz historii: {sheet_name}')
        elif choice.upper() == 'W':
            print('\n👋 Zakończono program. Powodzenia!')
            break
        else:
            print('⚠️ Nieprawidłowa opcja! Wybierz 1, 2, 3, 4 lub W.')


if __name__ == '__main__':
    main_menu()
