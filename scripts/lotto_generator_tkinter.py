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


def get_root_dir() -> Path:
    """Zwraca ROOT_DIR zarówno dla .py jak i .exe (PyInstaller).
    Jeśli .exe znajduje się w dist/, cofa się do katalogu nadrzędnego.
    """
    if getattr(sys, 'frozen', False):
        exe_dir = Path(sys.executable).parent
        # PyInstaller domyślnie buduje do dist/ — sprawdź czy scripts/ jest obok lub poziom wyżej
        if (exe_dir / "scripts").exists():
            return exe_dir
        parent = exe_dir.parent
        if (parent / "scripts").exists():
            return parent
        # Ostatnia deska ratunku — katalog exe
        return exe_dir
    return Path(__file__).resolve().parents[1]


def get_python_executable() -> str:
    """Zwraca ścieżkę do python.exe — działa i dla .py i dla .exe (PyInstaller)."""
    if getattr(sys, 'frozen', False):
        import shutil
        py = shutil.which("python") or shutil.which("python3")
        if py:
            return py
        raise RuntimeError(
            "Nie znaleziono python.exe w PATH!\n"
            "Zainstaluj Python i upewnij się, że jest w PATH."
        )
    return sys.executable


ROOT_DIR = get_root_dir()
DATA_DIR = ROOT_DIR / "data"
ANALYSIS_DIR = ROOT_DIR / "analysis"
DEFAULT_HISTORY_SHEET = "Arkusz1"
DEFAULT_HISTORY_DB = DATA_DIR / "lotto_history.db"

# Kolory status bara
_STATUS_COLORS = {
    "info":    "#e5e7eb",
    "success": "#4ade80",
    "warning": "#fbbf24",
    "error":   "#ef4444",
}


def find_latest_stats() -> Path | None:
    """Zwraca najnowszy plik statystyk z folderu analysis/ (według mtime)."""
    if not ANALYSIS_DIR.exists():
        return None
    candidates = list(ANALYSIS_DIR.glob("statystyki_lotto*.xlsx"))
    if not candidates:
        candidates = list(ANALYSIS_DIR.glob("*.xlsx"))
    if not candidates:
        return None
    return max(candidates, key=lambda p: p.stat().st_mtime)


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


def _clamp(v, lo=0, hi=255):
    return max(lo, min(hi, int(v)))


def _hex_to_rgb(hex_color: str):
    hex_color = hex_color.lstrip("#")
    return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))


def _rgb_to_hex(rgb):
    r, g, b = (_clamp(x) for x in rgb)
    return f"#{r:02x}{g:02x}{b:02x}"


def _mix_colors(c1: str, c2: str, t: float):
    t = max(0.0, min(1.0, float(t)))
    r1, g1, b1 = _hex_to_rgb(c1)
    r2, g2, b2 = _hex_to_rgb(c2)
    return _rgb_to_hex((
        r1 + (r2 - r1) * t,
        g1 + (g2 - g1) * t,
        b1 + (b2 - b1) * t,
    ))


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
        self._status_reset_id = None

        # Build UI
        self._build_ui()
        self._auto_load()

    # ------------------------------------------------------------------
    # STATUS BAR
    # ------------------------------------------------------------------
    def _number_color_from_cold_streak(self, n: int) -> str:
        neutral = "#475569"
        red = "#ef4444"
        orange = "#f59e0b"
        blue = "#2563eb"

        if not self.stats:
            return neutral

        try:
            data = self.stats.number_stats(n)
            streak = data.get("streak") or {}
            missing = streak.get("losowan_temu")
            if missing is None:
                return neutral

            missing = int(missing)

            if missing <= 5:
                return _mix_colors(red, orange, missing / 5)
            if missing <= 25:
                return _mix_colors(orange, blue, (missing - 5) / 20)
            return blue
        except Exception:
            return neutral

    def set_status(self, msg: str, level: str = "info", auto_reset: bool = True) -> None:
        color = _STATUS_COLORS.get(level, _STATUS_COLORS["info"])
        self.status_bar.config(text=f"  {msg}", foreground=color)

        if self._status_reset_id is not None:
            self.root.after_cancel(self._status_reset_id)
            self._status_reset_id = None

        if auto_reset and level in ("info", "success"):
            self._status_reset_id = self.root.after(
                5000, lambda: self.set_status("Gotowy", "info", auto_reset=False)
            )

    def _build_ui(self):
        Label(self.root, text="Generator Lotto - Tkinter",
              font=("Arial", 20, "bold"), bg="#0b1220", fg="white").pack(pady=10)

        btn_frame = Frame(self.root, bg="#0b1220")
        btn_frame.pack(pady=5, fill=X, padx=10)

        Button(btn_frame, text="\U0001f3b1  Losuj", command=self._draw,
               bg="#0ea5e9", fg="white", font=("Arial", 10, "bold"), width=14,
               relief=RAISED, bd=2).pack(side=LEFT, padx=5)

        Button(btn_frame, text="\u27f3  Aktualizuj", command=self._update_results,
               bg="#059669", fg="white", font=("Arial", 10, "bold"), width=14,
               relief=RAISED, bd=2).pack(side=LEFT, padx=5)

        ttk.Separator(btn_frame, orient=VERTICAL).pack(side=LEFT, fill=Y, padx=8, pady=4)

        self._data_menu = Menu(self.root, tearoff=0, bg="#1f2937", fg="white",
                               activebackground="#374151", activeforeground="white",
                               relief=FLAT, bd=0)
        self._data_menu.add_command(label="\U0001f4ca  Wczytaj statystyki",
                                    command=self._pick_stats)
        self._data_menu.add_command(label="\U0001f4c2  Plik historii \u2014 zmie\u0144",
                                    command=self._pick_history)
        self._data_menu.add_command(label="\U0001f4dc  Historia losowa\u0144",
                                    command=self._show_history)
        self._data_menu.add_separator()
        self._data_menu.add_command(label="\U0001f5c4  Baza danych",
                                    command=self._show_database)

        def _open_data_menu():
            btn = self._btn_data
            x = btn.winfo_rootx()
            y = btn.winfo_rooty() + btn.winfo_height()
            self._data_menu.tk_popup(x, y)

        self._btn_data = Button(
            btn_frame, text="\u25be  Dane", command=_open_data_menu,
            bg="#334155", fg="white", font=("Arial", 10), width=12,
            relief=RAISED, bd=1
        )
        self._btn_data.pack(side=LEFT, padx=5)

        content_frame = Frame(self.root, bg="#0b1220")
        content_frame.pack(fill=BOTH, expand=True, padx=10, pady=10)

        left_frame = Frame(content_frame, bg="#0b1220")
        left_frame.pack(side=LEFT, fill=BOTH, expand=True, padx=5)

        balls_frame = Frame(left_frame, bg="#0b1220")
        balls_frame.pack(pady=10)

        self.ball_labels = []
        for i in range(6):
            lbl = Label(balls_frame, text="--", font=("Arial", 18, "bold"),
                       bg="#2563eb", fg="white", width=6, relief=RAISED, cursor="hand2")
            lbl.pack(side=LEFT, padx=5)
            lbl.bind("<Button-1>", lambda e, idx=i: self._on_ball_clicked(idx))
            self.ball_labels.append(lbl)

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

        for title, vmin, vmax, gmin, gmax in [
            ("Suma", 80, 220, 110, 185),
            ("Spread", 8, 48, 31, 42)
        ]:
            frame = Frame(left_frame, bg="#1f2937", relief=RIDGE, bd=1)
            frame.pack(pady=5, fill=X)
            Label(frame, text=f"{title}: --", bg="#1f2937", fg="white",
                 font=("Arial", 9, "bold")).pack(anchor=W, padx=5, pady=2)
            self.stat_cards[f"{title}_label"] = frame.winfo_children()[0]

        diff_frame = Frame(left_frame, bg="#1f2937", relief=RIDGE, bd=1)
        diff_frame.pack(pady=5, fill=BOTH, expand=True)
        Label(diff_frame, text="R\u00f3\u017cnice mi\u0119dzy kolejnymi liczbami",
             bg="#1f2937", fg="white", font=("Arial", 9, "bold")).pack(anchor=W, padx=5, pady=2)
        self.diff_text = Text(diff_frame, height=4, bg="#0b1220", fg="#e5e7eb",
                             font=("Courier", 9), relief=FLAT, wrap=WORD)
        self.diff_text.pack(fill=BOTH, expand=True, padx=5, pady=5)
        self.diff_text.config(state=DISABLED)

        right_frame = Frame(content_frame, bg="#111827", relief=RIDGE, bd=1)
        right_frame.pack(side=RIGHT, fill=BOTH, expand=True, padx=5)

        Label(right_frame, text="Szczeg\u00f3\u0142y liczby", bg="#111827", fg="white",
             font=("Arial", 14, "bold")).pack(anchor=W, padx=10, pady=5)

        self.lbl_num = Label(right_frame, text="--", bg="#111827", fg="#ffd166",
                            font=("Arial", 32, "bold"))
        self.lbl_num.pack(pady=5)

        self.details_text = Text(right_frame, height=12, bg="#0b1220", fg="#e5e7eb",
                                font=("Courier", 10), relief=FLAT, wrap=WORD)
        self.details_text.pack(fill=BOTH, expand=True, padx=10, pady=5)
        self.details_text.config(state=DISABLED)

        ttk.Separator(self.root, orient=HORIZONTAL).pack(fill=X, side=BOTTOM)
        self.status_bar = ttk.Label(
            self.root,
            text="  Gotowy",
            relief=FLAT,
            anchor=W,
            font=("Arial", 9),
            foreground=_STATUS_COLORS["info"],
            background="#111827",
            padding=(4, 3),
        )
        self.status_bar.pack(side=BOTTOM, fill=X)

    def _update_labels(self):
        if self.stats and self.history_db:
            self.set_status(
                f"Statystyki: {self.stats_path.name}  |  Historia: {self.history_path.name}",
                "info", auto_reset=False
            )
        elif self.stats:
            self.set_status(
                f"Statystyki: {self.stats_path.name}  |  Historia: brak",
                "warning", auto_reset=False
            )
        else:
            self.set_status("Brak statystyk \u2014 u\u017cyj menu 'Dane' \u2192 Wczytaj statystyki",
                            "warning", auto_reset=False)

    def _generate_stats_in_background(self):
        def run():
            try:
                self.root.after(0, lambda: self.set_status(
                    "Generowanie statystyk...", "warning", auto_reset=False))

                script_path = ROOT_DIR / "scripts" / "generate_lotto_stats_final.py"
                if not script_path.exists():
                    script_path = ROOT_DIR / "scripts" / "generate_lotto_stats.py"

                if not script_path.exists():
                    self.root.after(0, lambda: self.set_status(
                        f"Brak skryptu statystyk: {script_path}", "error", auto_reset=False))
                    return

                result = subprocess.run(
                    [get_python_executable(), str(script_path)],
                    capture_output=True,
                    text=True,
                    timeout=300
                )

                if result.returncode == 0:
                    latest = find_latest_stats()
                    if latest:
                        self.root.after(100, lambda p=latest: self._load_stats(p))
                        self.root.after(200, lambda: self.set_status(
                            "Statystyki wygenerowane pomy\u015blnie", "success"))
                    else:
                        self.root.after(0, lambda: self.set_status(
                            "Statystyki: plik nie znaleziony po generowaniu", "error", auto_reset=False))
                else:
                    err = result.stderr[:200] if result.stderr else "Nieznany b\u0142\u0105d"
                    self.root.after(0, lambda e=err: self.set_status(
                        f"B\u0142\u0105d generowania statystyk: {e}", "error", auto_reset=False))
            except subprocess.TimeoutExpired:
                self.root.after(0, lambda: self.set_status(
                    "Timeout przy generowaniu statystyk", "error", auto_reset=False))
            except Exception as e:
                self.root.after(0, lambda ex=e: self.set_status(
                    f"B\u0142\u0105d statystyk: {ex}", "error", auto_reset=False))

        threading.Thread(target=run, daemon=True).start()

    def _update_results(self):
        def run():
            try:
                self.root.after(0, lambda: self.set_status(
                    "Aktualizacja wynik\u00f3w lotto...", "warning", auto_reset=False))

                script_path = ROOT_DIR / "scripts" / "scraper_megalotto.py"
                if not script_path.exists():
                    script_path = ROOT_DIR / "scripts" / "update_lotto_results.py"

                if not script_path.exists():
                    self.root.after(0, lambda: self.set_status(
                        f"Brak skryptu aktualizacji: {script_path}", "error", auto_reset=False))
                    return

                result = subprocess.run(
                    [get_python_executable(), str(script_path), "--update-xlsx"],
                    capture_output=True,
                    text=True,
                    timeout=120
                )

                if result.returncode == 0:
                    self.root.after(0, self._reload_history_main_thread)
                    self.root.after(100, lambda: self.set_status(
                        "Wyniki lotto zaktualizowane pomy\u015blnie", "success"))
                    self.root.after(300, self._generate_stats_in_background)
                else:
                    error_msg = result.stderr[:200] if result.stderr else "Nieznany b\u0142\u0105d"
                    self.root.after(0, self._update_labels)
                    self.root.after(0, lambda e=error_msg: self.set_status(
                        f"B\u0142\u0105d aktualizacji: {e}", "error", auto_reset=False))
                    self.root.after(100, lambda e=error_msg: messagebox.showerror(
                        "B\u0142\u0105d", f"Nie uda\u0142o si\u0119 zaktualizowa\u0107 wynik\u00f3w:\n{e}"))
            except subprocess.TimeoutExpired:
                self.root.after(0, self._update_labels)
                self.root.after(0, lambda: self.set_status(
                    "Timeout \u2014 aktualizacja trwa\u0142a zbyt d\u0142ugo", "error", auto_reset=False))
                self.root.after(100, lambda: messagebox.showerror(
                    "Timeout", "Aktualizacja trwa\u0142a zbyt d\u0142ugo"))
            except Exception as e:
                self.root.after(0, self._update_labels)
                self.root.after(0, lambda ex=e: self.set_status(
                    f"B\u0142\u0105d aktualizacji: {ex}", "error", auto_reset=False))
                self.root.after(100, lambda ex=e: messagebox.showerror(
                    "B\u0142\u0105d", f"B\u0142\u0105d aktualizacji: {ex}"))

        threading.Thread(target=run, daemon=True).start()

    def _reload_history_main_thread(self):
        self._load_history()
        self._update_labels()

    def _show_database(self):
        if not self.history_db:
            self.set_status("Baza danych nie jest za\u0142adowana", "warning")
            messagebox.showwarning("Baza danych", "Baza danych nie jest za\u0142adowana")
            return

        try:
            cursor = self.history_db.cursor()
            cursor.execute("SELECT id, draw_date, numbers FROM draws ORDER BY id DESC LIMIT 100")
            rows = cursor.fetchall()
        except Exception as e:
            self.set_status(f"B\u0142\u0105d odczytu bazy: {e}", "error")
            messagebox.showerror("B\u0142\u0105d", f"Nie mog\u0119 odczyta\u0107 bazy: {e}")
            return

        self.set_status(f"Otwarto widok bazy ({len(rows)} rekord\u00f3w)", "info")

        db_win = Toplevel(self.root)
        db_win.title("Baza danych - Wyniki Lotto")
        db_win.geometry("900x600")
        db_win.configure(bg="#0b1220")

        Label(Frame(db_win, bg="#0b1220").pack(pady=5) or db_win,
              text=f"Wy\u015bwietlono 100 ostatnich losowa\u0144 (total: {len(rows)})",
              font=("Arial", 11, "bold"), bg="#0b1220", fg="white").pack()

        tree_frame = Frame(db_win, bg="#0b1220")
        tree_frame.pack(fill=BOTH, expand=True, padx=5, pady=5)

        vsb = ttk.Scrollbar(tree_frame, orient=VERTICAL)
        hsb = ttk.Scrollbar(tree_frame, orient=HORIZONTAL)

        tree = ttk.Treeview(tree_frame, columns=("ID", "Data", "Liczby"), height=25,
                           yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.config(command=tree.yview)
        hsb.config(command=tree.xview)

        tree.heading("#0", text="Lp.")
        tree.heading("ID", text="ID")
        tree.heading("Data", text="Data")
        tree.heading("Liczby", text="Liczby")
        tree.column("#0", width=40)
        tree.column("ID", width=60)
        tree.column("ID", width=60)
        tree.column("Data", width=150)
        tree.column("Liczby", width=600)

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", background="#111827", foreground="white",
                       fieldbackground="#1f2937", borderwidth=0)
        style.configure("Treeview.Heading", background="#1f2937", foreground="white")
        style.map("Treeview", background=[("selected", "#374151")])

        for idx, (draw_id, date, numbers_json) in enumerate(rows, 1):
            try:
                numbers_str = " ".join(f"{n:2d}" for n in json.loads(numbers_json))
            except Exception:
                numbers_str = "ERROR"
            tree.insert("", 0, text=str(idx), values=(draw_id, date[:16], numbers_str))

        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        def export_to_csv():
            try:
                path = filedialog.asksaveasfilename(
                    parent=db_win, defaultextension=".csv",
                    filetypes=[("CSV", "*.csv"), ("All", "*.*")])
                if not path:
                    return
                with open(path, "w", encoding="utf-8") as f:
                    f.write("ID,Data,Liczby\n")
                    for draw_id, date, numbers_json in rows:
                        try:
                            nums = " ".join(str(n) for n in json.loads(numbers_json))
                        except Exception:
                            nums = "ERROR"
                        f.write(f"{draw_id},{date},{nums}\n")
                self.set_status(f"Dane wyeksportowane: {path}", "success")
                messagebox.showinfo("Sukces", f"Dane wyeksportowane do:\n{path}")
            except Exception as e:
                self.set_status(f"B\u0142\u0105d eksportu: {e}", "error")
                messagebox.showerror("B\u0142\u0105d", f"Nie mog\u0119 eksportowa\u0107: {e}")

        Button(db_win, text="Eksportuj do CSV", command=export_to_csv,
               bg="#334155", fg="white", font=("Arial", 10)).pack(pady=5)

    def _auto_load(self):
        if DEFAULT_HISTORY_DB.exists():
            self.history_path = DEFAULT_HISTORY_DB
            self._load_history()

        latest = find_latest_stats()
        if latest:
            self._load_stats(latest)
        else:
            self.set_status("Statystyki nie znalezione \u2014 generuj\u0119...", "warning", auto_reset=False)
            self._generate_stats_in_background()

        self._update_labels()
        self._draw()

    def _load_stats(self, path):
        try:
            self.stats = LottoStatistics(path)
            self.stats_path = Path(path)
            self.freq_map = {k: v["procent"] for k, v in self.stats.frequency.items()}
            self.status_map = {i: self.stats.status_short(i) for i in range(1, 50)}
            self.set_status(f"Za\u0142adowano statystyki: {Path(path).name}", "success")
        except Exception as exc:
            self.stats = None
            self.freq_map = {i: 0.0 for i in range(1, 50)}
            self.status_map = {i: "" for i in range(1, 50)}
            self.set_status(f"B\u0142\u0105d statystyk: {exc}", "error", auto_reset=False)
            messagebox.showerror("B\u0142\u0105d statystyk", str(exc))
        self._update_labels()

    def _load_history(self):
        if self.history_db:
            try:
                self.history_db.close()
            except Exception:
                pass
            self.history_db = None

        if not self.history_path or not self.history_path.exists():
            return
        try:
            self.history_db = sqlite3.connect(
                str(self.history_path), check_same_thread=False)
            count = self.history_db.execute(
                "SELECT COUNT(*) FROM draws").fetchone()[0]
            self.set_status(f"Historia za\u0142adowana: {count} losowa\u0144", "success")
        except Exception as e:
            self.history_db = None
            self.set_status(f"B\u0142\u0105d bazy historii: {e}", "error", auto_reset=False)

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
            initialdir=start,
            filetypes=[("All", "*.*"), ("Excel", "*.xlsx *.xls"), ("SQLite", "*.db")])
        if path:
            self.history_path = Path(path)
            self._load_history()
            self._update_labels()

    def _combination_exists(self, numbers):
        if not self.history_db:
            return False
        try:
            numbers_json = json.dumps(sorted(numbers))
            return self.history_db.execute(
                "SELECT COUNT(*) FROM draws WHERE numbers = ?",
                (numbers_json,)).fetchone()[0] > 0
        except Exception:
            return False

    def _draw(self):
        numbers = generate_numbers()
        self.current_numbers = numbers
        s = draw_stats(numbers)

        for btn, n in zip(self.ball_labels, numbers):
            color = self._number_color_from_cold_streak(n)
            btn.config(
                text=str(n),
                bg=color,
                activebackground=color,
                fg="white",
                activeforeground="white",
            )

        self.stat_cards["Suma"].config(text=str(s["suma"]))
        self.stat_cards["Spread"].config(text=str(s["spread"]))
        self.stat_cards["Suma roznic"].config(text=str(s["suma_roznic"]))
        self.stat_cards["P/N"].config(text=f"{s['parzyste']}P-{s['nieparzyste']}N")
        self.stat_cards["L/H"].config(text=f"{s['niskie']}L-{s['wysokie']}H")
        self.stat_cards["Historia"].config(
            text="BY\u0141A" if self._combination_exists(numbers) else "NOWA")

        self.diff_text.config(state=NORMAL)
        self.diff_text.delete("1.0", END)
        self.diff_text.insert(END, "R\u00f3\u017cnice: " + " - ".join(str(d) for d in s["diffs"]))
        self.diff_text.config(state=DISABLED)

        self._show_number(numbers[0])

    def _on_ball_clicked(self, idx):
        if idx < len(self.current_numbers):
            self._show_number(self.current_numbers[idx])

    def _get_history_draws(self, limit=20):
        if not self.history_db:
            return []
        try:
            rows = self.history_db.execute(
                "SELECT draw_date, numbers FROM draws ORDER BY id DESC LIMIT ?",
                (limit,)).fetchall()
            draws = []
            for date, nj in rows:
                try:
                    draws.append((date, json.loads(nj)))
                except Exception:
                    pass
            draws.reverse()
            return draws
        except Exception as e:
            self.set_status(f"B\u0142\u0105d pobierania historii: {e}", "error")
            return []

    def _show_history(self):
        draws = self._get_history_draws(50)
        if not draws:
            self.set_status("Brak danych historycznych w bazie", "warning")
            messagebox.showinfo("Historia", "Brak danych historycznych w bazie")
            return

        self.set_status(f"Otwarto histori\u0119 ({len(draws)} losowa\u0144)", "info")

        hist_win = Toplevel(self.root)
        hist_win.title("Historia wylosowa\u0144")
        hist_win.geometry("600x500")
        hist_win.configure(bg="#0b1220")

        Label(hist_win, text="Ostatnie wylosowania",
             font=("Arial", 12, "bold"), bg="#0b1220", fg="white").pack(pady=5)

        text = Text(hist_win, bg="#0b1220", fg="#e5e7eb", font=("Courier", 10),
                   relief=FLAT, wrap=WORD)
        text.pack(fill=BOTH, expand=True, padx=10, pady=5)

        scrollbar = ttk.Scrollbar(text)
        text.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=text.yview)
        scrollbar.pack(side=RIGHT, fill=Y)

        content = "Data                | Liczby\n" + "-" * 50 + "\n"
        for date, numbers in draws:
            content += f"{date:20s} | {' '.join(f'{n:2d}' for n in numbers)}\n"

        text.insert(END, content)
        text.config(state=DISABLED)
        Button(hist_win, text="Zamknij", command=hist_win.destroy,
               bg="#334155", fg="white").pack(pady=5)

    def _show_number(self, n):
        self.lbl_num.config(text=str(n))

        if not self.stats:
            details = "Brak wczytanych statystyk.\nU\u017cyj menu 'Dane' \u2192 Wczytaj statystyki."
        else:
            d = self.stats.number_stats(n)
            streak_txt = (f"{d['streak']['losowan_temu']} los. temu | {d['streak']['status']}"
                          if d["streak"] else "brak danych")
            roll_txt = (f"{d['rolling_current']:.1f} (min {d['rolling_min']:.1f}, max {d['rolling_max']:.1f})"
                        if d["rolling_current"] is not None else "brak danych")

            details = (
                f"Status:      {d['status']}\n"
                f"Wyst\u0105pienia: {d['wystapienia']}\n"
                f"Udzia\u0142:      {d['procent']:.2f} %\n"
                f"Rolling100:  {roll_txt}\n"
                f"Cold streak: {streak_txt}\n\n"
            )
            if d["last_years"]:
                details += "Ostatnie lata:\n" + "  ".join(f"{y}: {v}" for y, v in d["last_years"]) + "\n\n"
            else:
                details += "Brak danych rocznych.\n\n"
            if d["pairs"]:
                details += "TOP pary:\n" + "\n".join(
                    f"  {p['para']} - {p['wystapienia']} razy" for p in d["pairs"])
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
