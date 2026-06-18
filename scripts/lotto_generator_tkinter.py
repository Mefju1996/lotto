import sys
import random
import sqlite3
import json
from pathlib import Path
from tkinter import *
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import subprocess
import threading
import time

APP_TITLE = "Generator Lotto - Tkinter"
ROOT_DIR = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT_DIR / "data"
DEFAULT_STATS = ROOT_DIR / "analysis" / "statystyki_lotto.xlsx"
DEFAULT_HISTORY_SHEET = "Arkusz1"
DEFAULT_HISTORY_DB = DATA_DIR / "lotto_history.db"


def is_valid(numbers):
    if not (110 <= sum(numbers) <= 185):
        return False
    diffs = [abs(numbers[i] - numbers[i - 1]) for i in range(1, len(numbers))]
    if any(d > 20 for d in diffs):
        return False
    if not (27 <= sum(diffs) <= 47):
        return False
    p = sum(1 for n in numbers if n % 2 == 0)
    if p == 0 or p == 6:
        return False
    lo = sum(1 for n in numbers if n <= 24)
    if lo == 0 or lo == 6:
        return False
    return True


def generate_numbers():
    while True:
        nums = sorted(random.sample(range(1, 50), 6))
        if is_valid(nums):
            return nums


def draw_stats(numbers):
    diffs = [abs(numbers[i] - numbers[i - 1]) for i in range(1, len(numbers))]
    p = sum(1 for n in numbers if n % 2 == 0)
    lo = sum(1 for n in numbers if n <= 24)
    return {
        "suma": sum(numbers),
        "spread": numbers[-1] - numbers[0],
        "diffs": diffs,
        "suma_roznic": sum(diffs),
        "parzyste": p,
        "nieparzyste": 6 - p,
        "niskie": lo,
        "wysokie": 6 - lo,
    }


class LottoStatistics:
    REQUIRED = {"freq": ["1_Czestotliwosc"], "hot": ["2_Hot_Numbers"], "cold": ["3_Cold_Numbers"]}
    OPTIONAL = {
        "pairs":   ["6_TOP50_Par"],
        "streak":  ["13_Cold_Streaks"],
        "rolling": ["16RollingFreq", "16_RollingFreq", "21_Sliding_Window", "Sliding_Window"],
        "years":   ["17_Heatmapa_Rok", "17_HeatmapaRok", "Heatmapa_Rok"],
    }

    def __init__(self, xlsx_path):
        self.path = Path(xlsx_path)
        self.frequency = {}
        self.hot_rank = {}
        self.cold_rank = {}
        self.top_pairs = {}
        self.rolling = {}
        self.year_heatmap = {}
        self.cold_streak = {}
        self._load()

    def _read(self, candidates):
        for name in candidates:
            try:
                return pd.read_excel(self.path, sheet_name=name, engine="openpyxl")
            except Exception:
                pass
        return None

    def _sheet_names(self):
        try:
            return pd.ExcelFile(self.path, engine="openpyxl").sheet_names
        except Exception:
            return []

    def _load(self):
        freq_df = self._read(self.REQUIRED["freq"])
        if freq_df is None:
            sheets = ", ".join(self._sheet_names())
            raise ValueError(f"Brak arkusza czestotliwosci.\nDostepne arkusze: {sheets}")
        for _, row in freq_df.iterrows():
            n = int(row["Liczba"])
            self.frequency[n] = {"wystapienia": int(row["Wystapienia"] if "Wystapienia" in row else row.iloc[1]),
                                  "procent": float(row["Procent"] if "Procent" in row else row.iloc[2])}

        hot_df = self._read(self.REQUIRED["hot"])
        if hot_df is not None:
            for _, row in hot_df.iterrows():
                try:
                    self.hot_rank[int(row["Liczba"])] = int(row["Ranking"])
                except Exception:
                    pass

        cold_df = self._read(self.REQUIRED["cold"])
        if cold_df is not None:
            for _, row in cold_df.iterrows():
                try:
                    self.cold_rank[int(row["Liczba"])] = int(row["Ranking"])
                except Exception:
                    pass

        pairs_df = self._read(self.OPTIONAL["pairs"])
        if pairs_df is not None:
            try:
                for _, row in pairs_df.iterrows():
                    parts = str(row.iloc[1]).split("-")
                    if len(parts) == 2:
                        a, b = int(parts[0]), int(parts[1])
                        entry = {"para": f"{a}-{b}", "wystapienia": int(row.iloc[2])}
                        self.top_pairs.setdefault(a, []).append(entry)
                        self.top_pairs.setdefault(b, []).append(entry)
            except Exception:
                pass

        streak_df = self._read(self.OPTIONAL["streak"])
        if streak_df is not None:
            try:
                for _, row in streak_df.iterrows():
                    self.cold_streak[int(row["Liczba"])] = {
                        "losowan_temu": int(row.iloc[2]),
                        "status": str(row.iloc[3]),
                    }
            except Exception:
                pass

        rolling_df = self._read(self.OPTIONAL["rolling"])
        if rolling_df is not None:
            try:
                for col in rolling_df.columns:
                    col_s = str(col).strip()
                    num = None
                    if col_s.isdigit():
                        num = int(col_s)
                    else:
                        try:
                            num = int(col_s)
                        except ValueError:
                            pass
                    if num is not None and 1 <= num <= 49:
                        series = pd.to_numeric(rolling_df[col], errors="coerce").dropna().tolist()
                        self.rolling[num] = [float(v) for v in series[-120:]]
            except Exception:
                pass

        years_df = self._read(self.OPTIONAL["years"])
        if years_df is not None:
            try:
                year_cols = [c for c in years_df.columns if str(c) != "Liczba"]
                for _, row in years_df.iterrows():
                    n = int(row["Liczba"])
                    self.year_heatmap[n] = {
                        str(y): int(row[y]) for y in year_cols if pd.notna(row.get(y, None))
                    }
            except Exception:
                pass

    def status_short(self, n):
        if n in self.hot_rank:
            return f"H#{self.hot_rank[n]}"
        if n in self.cold_rank:
            return f"C#{self.cold_rank[n]}"
        return ""

    def status_long(self, n):
        if n in self.hot_rank:
            return f"HOT #{self.hot_rank[n]}"
        if n in self.cold_rank:
            return f"COLD #{self.cold_rank[n]}"
        return "NEUTRALNA"

    def number_stats(self, n):
        freq = self.frequency.get(n, {"wystapienia": 0, "procent": 0.0})
        rolling = self.rolling.get(n, [])
        pairs = sorted(self.top_pairs.get(n, []), key=lambda x: x["wystapienia"], reverse=True)[:5]
        yearly = self.year_heatmap.get(n, {})
        streak = self.cold_streak.get(n)
        return {
            "n": n,
            "status": self.status_long(n),
            "wystapienia": freq["wystapienia"],
            "procent": freq["procent"],
            "rolling": rolling,
            "rolling_current": rolling[-1] if rolling else None,
            "rolling_min": min(rolling) if rolling else None,
            "rolling_max": max(rolling) if rolling else None,
            "pairs": pairs,
            "last_years": list(yearly.items())[-8:],
            "streak": streak,
        }


class LottoApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1200x800")
        self.root.configure(bg="#0b1220")
        
        # Variables
        self.stats = None
        self.stats_path = None
        self.freq_map = {i: 0.0 for i in range(1, 50)}
        self.status_map = {i: "" for i in range(1, 50)}
        self.history_db = None
        self.history_path = None
        self.history_sheet = DEFAULT_HISTORY_SHEET
        self.current_numbers = []
        
        # Build UI
        self._build_ui()
        self._auto_load()
    
    def _build_ui(self):
        # Title
        title = Label(self.root, text="Generator Lotto - Tkinter", 
                     font=("Arial", 20, "bold"), bg="#0b1220", fg="white")
        title.pack(pady=10)
        
        # Control buttons
        btn_frame = Frame(self.root, bg="#0b1220")
        btn_frame.pack(pady=5, fill=X, padx=10)
        
        Button(btn_frame, text="Losuj", command=self._draw, 
               bg="#0ea5e9", fg="white", font=("Arial", 10, "bold"), width=15).pack(side=LEFT, padx=5)
        Button(btn_frame, text="Wczytaj statystyki", command=self._pick_stats, 
               bg="#7c3aed", fg="white", font=("Arial", 10), width=20).pack(side=LEFT, padx=5)
        Button(btn_frame, text="Plik historii", command=self._pick_history, 
               bg="#334155", fg="white", font=("Arial", 10), width=15).pack(side=LEFT, padx=5)
        Button(btn_frame, text="Historia", command=self._show_history, 
               bg="#334155", fg="white", font=("Arial", 10), width=15).pack(side=LEFT, padx=5)
        
        # Status labels
        self.lbl_stats = Label(self.root, text="Statystyki: brak", 
                              bg="#0b1220", fg="#a78bfa", font=("Arial", 9))
        self.lbl_stats.pack(pady=2)
        
        self.lbl_history = Label(self.root, text="Historia: brak", 
                                bg="#0b1220", fg="#fbbf24", font=("Arial", 9))
        self.lbl_history.pack(pady=2)
        
        # Main content
        content_frame = Frame(self.root, bg="#0b1220")
        content_frame.pack(fill=BOTH, expand=True, padx=10, pady=10)
        
        # Left panel
        left_frame = Frame(content_frame, bg="#0b1220")
        left_frame.pack(side=LEFT, fill=BOTH, expand=True, padx=5)
        
        # Balls (numbers display)
        balls_frame = Frame(left_frame, bg="#0b1220")
        balls_frame.pack(pady=10)
        
        self.ball_labels = []
        for i in range(6):
            lbl = Label(balls_frame, text="--", font=("Arial", 18, "bold"), 
                       bg="#2563eb", fg="white", width=6, relief=RAISED, cursor="hand2")
            lbl.pack(side=LEFT, padx=5)
            lbl.bind("<Button-1>", lambda e, idx=i: self._on_ball_clicked(idx))
            self.ball_labels.append(lbl)
        
        # Stats cards
        cards_frame = Frame(left_frame, bg="#0b1220")
        cards_frame.pack(pady=10, fill=X)
        
        self.stat_cards = {}
        for title in ["Suma", "Spread", "Suma roznic", "P/N", "L/H", "Historia"]:
            card = Frame(cards_frame, bg="#1f2937", relief=RIDGE, bd=1)
            card.pack(side=LEFT, padx=5, fill=BOTH, expand=True)
            Label(card, text=title, bg="#1f2937", fg="#9ca3af", font=("Arial", 8)).pack()
            val = Label(card, text="--", bg="#1f2937", fg="white", font=("Arial", 14, "bold"))
            val.pack(pady=5)
            self.stat_cards[title] = val
        
        # Range bars
        for title, vmin, vmax, gmin, gmax in [
            ("Suma", 80, 220, 110, 185),
            ("Spread", 8, 48, 31, 42)
        ]:
            frame = Frame(left_frame, bg="#1f2937", relief=RIDGE, bd=1)
            frame.pack(pady=5, fill=X)
            Label(frame, text=f"{title}: --", bg="#1f2937", fg="white", 
                 font=("Arial", 9, "bold")).pack(anchor=W, padx=5, pady=2)
            self.stat_cards[f"{title}_label"] = frame.winfo_children()[0]
        
        # Differences
        diff_frame = Frame(left_frame, bg="#1f2937", relief=RIDGE, bd=1)
        diff_frame.pack(pady=5, fill=BOTH, expand=True)
        Label(diff_frame, text="Różnice między kolejnymi liczbami", 
             bg="#1f2937", fg="white", font=("Arial", 9, "bold")).pack(anchor=W, padx=5, pady=2)
        self.diff_text = Text(diff_frame, height=4, bg="#0b1220", fg="#e5e7eb", 
                             font=("Courier", 9), relief=FLAT, wrap=WORD)
        self.diff_text.pack(fill=BOTH, expand=True, padx=5, pady=5)
        self.diff_text.config(state=DISABLED)
        
        # Right panel - details
        right_frame = Frame(content_frame, bg="#111827", relief=RIDGE, bd=1)
        right_frame.pack(side=RIGHT, fill=BOTH, expand=True, padx=5)
        
        Label(right_frame, text="Szczegóły liczby", bg="#111827", fg="white", 
             font=("Arial", 14, "bold")).pack(anchor=W, padx=10, pady=5)
        
        self.lbl_num = Label(right_frame, text="--", bg="#111827", fg="#ffd166", 
                            font=("Arial", 32, "bold"))
        self.lbl_num.pack(pady=5)
        
        self.details_text = Text(right_frame, height=12, bg="#0b1220", fg="#e5e7eb", 
                                font=("Courier", 10), relief=FLAT, wrap=WORD)
        self.details_text.pack(fill=BOTH, expand=True, padx=10, pady=5)
        self.details_text.config(state=DISABLED)
    
    def _update_labels(self):
        sp = str(self.stats_path) if self.stats_path else "brak"
        ok = "OK" if self.stats else "brak danych"
        self.lbl_stats.config(text=f"Statystyki: {sp}  |  {ok}")
        hp = str(self.history_path) if self.history_path else "brak"
        self.lbl_history.config(text=f"Historia: {hp}  |  Arkusz: {self.history_sheet}")
    
    def _generate_stats_in_background(self):
        """Generuj statystyki w osobnym wątku"""
        def run():
            try:
                self.lbl_stats.config(text="Statystyki: generowanie...", fg="#fbbf24")
                self.root.update()
                
                # Uruchom skrypt generowania statystyk
                script_path = ROOT_DIR / "scripts" / "generate_lotto_stats.py"
                result = subprocess.run(
                    [sys.executable, str(script_path)],
                    capture_output=True,
                    text=True,
                    timeout=120
                )
                
                if result.returncode == 0:
                    print("✓ Statystyki wygenerowane pomyślnie")
                    # Załaduj statystyki
                    if DEFAULT_STATS.exists():
                        self.root.after(100, lambda: self._load_stats(DEFAULT_STATS))
                else:
                    print(f"✗ Błąd generowania: {result.stderr}")
                    self.lbl_stats.config(text="Statystyki: błąd generowania", fg="#ef4444")
            except subprocess.TimeoutExpired:
                print("✗ Timeout przy generowaniu statystyk")
                self.lbl_stats.config(text="Statystyki: timeout", fg="#ef4444")
            except Exception as e:
                print(f"✗ Błąd: {e}")
                self.lbl_stats.config(text=f"Statystyki: {str(e)}", fg="#ef4444")
        
        thread = threading.Thread(target=run, daemon=True)
        thread.start()
    
    def _auto_load(self):
        # Załaduj historię jeśli istnieje
        if DEFAULT_HISTORY_DB.exists():
            self.history_path = DEFAULT_HISTORY_DB
            self._load_history()
        
        # Załaduj statystyki jeśli istnieją, inaczej generuj je
        if DEFAULT_STATS.exists():
            self._load_stats(DEFAULT_STATS)
        else:
            print("Statystyki nie znalezione, generuję...")
            self._generate_stats_in_background()
        
        self._update_labels()
        self._draw()
    
    def _load_stats(self, path):
        try:
            self.stats = LottoStatistics(path)
            self.stats_path = Path(path)
            self.freq_map = {k: v["procent"] for k, v in self.stats.frequency.items()}
            self.status_map = {i: self.stats.status_short(i) for i in range(1, 50)}
        except Exception as exc:
            self.stats = None
            self.freq_map = {i: 0.0 for i in range(1, 50)}
            self.status_map = {i: "" for i in range(1, 50)}
            messagebox.showerror("Błąd statystyk", str(exc))
        self._update_labels()
    
    def _load_history(self):
        if self.history_db:
            self.history_db.close()
            self.history_db = None
        
        if not self.history_path or not self.history_path.exists():
            return
        try:
            self.history_db = sqlite3.connect(str(self.history_path))
            # Test connection
            cursor = self.history_db.cursor()
            cursor.execute("SELECT COUNT(*) FROM draws")
            count = cursor.fetchone()[0]
            print(f"Historia: {count} wylosowań w bazie")
        except Exception as e:
            print(f"Błąd bazy: {e}")
            self.history_db = None
    
    def _pick_stats(self):
        start = str(self.stats_path.parent if self.stats_path else ROOT_DIR)
        path = filedialog.askopenfilename(
            parent=self.root, title="Wybierz plik statystyk", 
            initialdir=start, filetypes=[("Excel", "*.xlsx *.xls")])
        if path:
            self._load_stats(path)
            if not self.history_path:
                self.history_path = Path(path)
                self._load_history()
            self._update_labels()
            if self.current_numbers:
                self._show_number(self.current_numbers[0])
    
    def _pick_history(self):
        start = str(self.history_path.parent if self.history_path else ROOT_DIR)
        path = filedialog.askopenfilename(
            parent=self.root, title="Wybierz plik historii", 
            initialdir=start, filetypes=[("All", "*.*"), ("Excel", "*.xlsx *.xls"), ("SQLite", "*.db")])
        if path:
            self.history_path = Path(path)
            self._load_history()
            self._update_labels()
    
    def _combination_exists(self, numbers):
        if not self.history_db:
            return False
        try:
            numbers_json = json.dumps(sorted(numbers))
            cursor = self.history_db.execute(
                "SELECT COUNT(*) FROM draws WHERE numbers = ?",
                (numbers_json,)
            )
            return cursor.fetchone()[0] > 0
        except Exception:
            return False
    
    def _draw(self):
        numbers = generate_numbers()
        self.current_numbers = numbers
        s = draw_stats(numbers)
        hot = self.stats.hot_rank if self.stats else {}
        cold = self.stats.cold_rank if self.stats else {}
        
        # Update balls
        for btn, n in zip(self.ball_labels, numbers):
            if n in hot:
                color = "#ef4444"
            elif n in cold:
                color = "#475569"
            else:
                color = "#2563eb"
            btn.config(text=str(n), bg=color)
        
        # Update stat cards
        self.stat_cards["Suma"].config(text=str(s["suma"]))
        self.stat_cards["Spread"].config(text=str(s["spread"]))
        self.stat_cards["Suma roznic"].config(text=str(s["suma_roznic"]))
        self.stat_cards["P/N"].config(text=f"{s['parzyste']}P-{s['nieparzyste']}N")
        self.stat_cards["L/H"].config(text=f"{s['niskie']}L-{s['wysokie']}H")
        self.stat_cards["Historia"].config(text="BYŁA" if self._combination_exists(numbers) else "NOWA")
        
        # Update range bars
        suma_label = self.root.winfo_children()
        # Find and update labels (simplified)
        for i, card_key in enumerate(["Suma", "Spread"]):
            if card_key == "Suma":
                self.stat_cards["Suma_label"] = s["suma"]
            else:
                self.stat_cards["Spread_label"] = s["spread"]
        
        # Update differences
        self.diff_text.config(state=NORMAL)
        self.diff_text.delete("1.0", END)
        self.diff_text.insert(END, "Różnice: " + " - ".join(str(d) for d in s["diffs"]))
        self.diff_text.config(state=DISABLED)
        
        self._show_number(numbers[0])
    
    def _on_ball_clicked(self, idx):
        if idx < len(self.current_numbers):
            self._show_number(self.current_numbers[idx])
    
    def _get_history_draws(self, limit=20):
        """Pobierz ostatnie wylosowania z bazy (najstarsze na górze)"""
        if not self.history_db:
            return []
        try:
            cursor = self.history_db.cursor()
            cursor.execute(
                "SELECT draw_date, numbers FROM draws ORDER BY id DESC LIMIT ?",
                (limit,)
            )
            draws = []
            for row in cursor.fetchall():
                date, numbers_json = row
                try:
                    numbers = json.loads(numbers_json)
                    draws.append((date, numbers))
                except:
                    pass
            draws.reverse()  # Odwróć aby najstarsze były na górze
            return draws
        except Exception as e:
            print(f"Błąd pobierania historii: {e}")
            return []
    
    def _show_history(self):
        """Wyświetl okno z historią wylosowań"""
        draws = self._get_history_draws(50)
        
        if not draws:
            messagebox.showinfo("Historia", "Brak danych historycznych w bazie")
            return
        
        # Utwórz nowe okno
        hist_win = Toplevel(self.root)
        hist_win.title("Historia wylosowań")
        hist_win.geometry("600x500")
        hist_win.configure(bg="#0b1220")
        
        Label(hist_win, text="Ostatnie wylosowania", 
             font=("Arial", 12, "bold"), bg="#0b1220", fg="white").pack(pady=5)
        
        # Text widget z historią
        text = Text(hist_win, bg="#0b1220", fg="#e5e7eb", font=("Courier", 10), 
                   relief=FLAT, wrap=WORD)
        text.pack(fill=BOTH, expand=True, padx=10, pady=5)
        
        scrollbar = ttk.Scrollbar(text)
        text.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=text.yview)
        scrollbar.pack(side=RIGHT, fill=Y)
        
        # Dodaj dane
        content = "Data                | Liczby\n"
        content += "-" * 50 + "\n"
        for date, numbers in draws:
            numbers_str = " ".join(f"{n:2d}" for n in numbers)
            content += f"{date:20s} | {numbers_str}\n"
        
        text.insert(END, content)
        text.config(state=DISABLED)
        
        Button(hist_win, text="Zamknij", command=hist_win.destroy,
              bg="#334155", fg="white").pack(pady=5)
    
    def _show_number(self, n):
        self.lbl_num.config(text=str(n))
        
        if not self.stats:
            details = "Brak wczytanych statystyk.\nUżyj przycisku 'Wczytaj statystyki'."
        else:
            d = self.stats.number_stats(n)
            streak_txt = "brak danych"
            if d["streak"]:
                streak_txt = f"{d['streak']['losowan_temu']} los. temu | {d['streak']['status']}"
            
            roll_txt = "brak danych"
            if d["rolling_current"] is not None:
                roll_txt = f"{d['rolling_current']:.1f} (min {d['rolling_min']:.1f}, max {d['rolling_max']:.1f})"
            
            details = (
                f"Status:      {d['status']}\n"
                f"Wystąpienia: {d['wystapienia']}\n"
                f"Udział:      {d['procent']:.2f} %\n"
                f"Rolling100:  {roll_txt}\n"
                f"Cold streak: {streak_txt}\n\n"
            )
            
            if d["last_years"]:
                details += "Ostatnie lata:\n"
                details += "  ".join(f"{y}: {v}" for y, v in d["last_years"]) + "\n\n"
            else:
                details += "Brak danych rocznych.\n\n"
            
            if d["pairs"]:
                details += "TOP pary:\n"
                details += "\n".join(f"  {p['para']} - {p['wystapienia']} razy" for p in d["pairs"])
            else:
                details += "Brak danych o parach."
        
        self.details_text.config(state=NORMAL)
        self.details_text.delete("1.0", END)
        self.details_text.insert(END, details)
        self.details_text.config(state=DISABLED)


def main():
    root = Tk()
    app = LottoApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
