import sys
import random
import sqlite3
import json
import statistics
import itertools
from pathlib import Path
from tkinter import *
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import subprocess
import threading
import time

import matplotlib
matplotlib.use("TkAgg")
import numpy as np
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

APP_TITLE = "Generator Lotto - Tkinter"


def get_root_dir() -> Path:
    """Zwraca ROOT_DIR zarowno dla .py jak i .exe (PyInstaller).
    Jesli .exe znajduje sie w dist/, cofa sie do katalogu nadrzednego.
    """
    if getattr(sys, 'frozen', False):
        exe_dir = Path(sys.executable).parent
        if (exe_dir / "scripts").exists():
            return exe_dir
        parent = exe_dir.parent
        if (parent / "scripts").exists():
            return parent
        return exe_dir
    return Path(__file__).resolve().parents[1]


def get_python_executable() -> str:
    """Zwraca sciezke do python.exe - dziala i dla .py i dla .exe (PyInstaller)."""
    if getattr(sys, 'frozen', False):
        import shutil
        py = shutil.which("python") or shutil.which("python3")
        if py:
            return py
        raise RuntimeError(
            "Nie znaleziono python.exe w PATH!\n"
            "Zainstaluj Python i upewnij sie, ze jest w PATH."
        )
    return sys.executable


def run_script(script_path, extra_args=None, timeout=300):
    """
    Uruchamia skrypt Pythona z wymuszonym UTF-8 (PYTHONUTF8=1 + -X utf8).
    Zwraca CompletedProcess.
    """
    import os
    env = os.environ.copy()
    env["PYTHONUTF8"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"

    cmd = [get_python_executable(), "-X", "utf8", str(script_path)]
    if extra_args:
        cmd += extra_args

    return subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=timeout,
        env=env,
    )


ROOT_DIR = get_root_dir()
DATA_DIR = ROOT_DIR / "data"
ANALYSIS_DIR = ROOT_DIR / "analysis"
DEFAULT_HISTORY_SHEET = "Arkusz1"
DEFAULT_HISTORY_DB = DATA_DIR / "lotto_history.db"

_STATUS_COLORS = {
    "info":    "#e5e7eb",
    "success": "#4ade80",
    "warning": "#fbbf24",
    "error":   "#ef4444",
}


def find_latest_stats() -> Path | None:
    """Zwraca najnowszy plik statystyk z folderu analysis/ (wedlug mtime)."""
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

        # Statystyki odstepow (historyczne, per liczba) + wykres biezacego losowania
        self.number_gap_stats = {}
        self.chart_canvas = None
        self.chart_figure = None
        self.chart_ax = None
        self.chart_ax2 = None
        self.chart_host = None
        self.gap_stats_label = None

        # Build UI
        self._build_ui()
        self._auto_load()

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

    # ------------------------------------------------------------------
    # STATYSTYKI ODSTEPOW MIEDZY WYSTAPIENIAMI LICZBY (historyczne, per liczba)
    # ------------------------------------------------------------------
    def _get_all_draws(self):
        if not self.history_db:
            return []
        try:
            rows = self.history_db.execute(
                "SELECT id, draw_date, numbers FROM draws ORDER BY id ASC"
            ).fetchall()
            draws = []
            for draw_id, draw_date, numbers_json in rows:
                try:
                    numbers = json.loads(numbers_json)
                    if isinstance(numbers, list):
                        draws.append({
                            "id": draw_id,
                            "draw_date": str(draw_date or ""),
                            "numbers": [int(x) for x in numbers],
                        })
                except Exception:
                    pass
            return draws
        except Exception:
            return []

    def _compute_gap_stats_for_all_numbers(self):
        """Liczy dla kazdej liczby 1-49 odstepy (w liczbie losowan) miedzy
        kolejnymi wystapieniami oraz srednia, mediane i dominanty tych odstepow."""
        draws = self._get_all_draws()
        positions = {n: [] for n in range(1, 50)}

        for idx, draw in enumerate(draws):
            for n in draw["numbers"]:
                if 1 <= n <= 49:
                    positions[n].append(idx)

        total_draws = len(draws)
        result = {}
        for n in range(1, 50):
            pos = positions[n]
            gaps = [pos[i] - pos[i - 1] for i in range(1, len(pos))]
            current_gap = (total_draws - 1 - pos[-1]) if pos else None

            if gaps:
                result[n] = {
                    "count": len(pos),
                    "gaps": gaps,
                    "mean_gap": round(statistics.mean(gaps), 2),
                    "median_gap": statistics.median(gaps),
                    "mode_gaps": statistics.multimode(gaps),
                    "min_gap": min(gaps),
                    "max_gap": max(gaps),
                    "current_gap": current_gap,
                }
            else:
                result[n] = {
                    "count": len(pos),
                    "gaps": [],
                    "mean_gap": None,
                    "median_gap": None,
                    "mode_gaps": [],
                    "min_gap": None,
                    "max_gap": None,
                    "current_gap": current_gap,
                }

        self.number_gap_stats = result

    # ------------------------------------------------------------------
    # WYKRES BIEZACEGO LOSOWANIA: "ile losowan temu" dla 6 liczb
    # + porownanie z historyczna srednia/mediana + suma skumulowana
    # ------------------------------------------------------------------
    def _current_draw_gaps(self):
        """Zwraca liste (liczba, ile_losowan_temu) dla aktualnie wylosowanych
        liczb. Najpierw probuje danych z arkusza statystyk (Cold Streaks),
        a jesli ich nie ma - liczy to samo na podstawie bazy historii."""
        result = []
        for n in self.current_numbers:
            gap = None
            if self.stats:
                streak = self.stats.cold_streak.get(n)
                if streak:
                    gap = streak.get("losowan_temu")
            if gap is None:
                hist = self.number_gap_stats.get(n)
                if hist:
                    gap = hist.get("current_gap")
            result.append((n, gap))
        return result

    def _draw_gap_summary_stats(self, gaps):
        values = [g for _, g in gaps if g is not None]
        if not values:
            return None
        return {
            "mean": round(statistics.mean(values), 2),
            "median": statistics.median(values),
            "mode": statistics.multimode(values),
            "sum": sum(values),
            "cumsum": list(itertools.accumulate(values)),
        }

    def _build_chart_panel(self, parent):
        wrapper = Frame(parent, bg="#111827", relief=RIDGE, bd=1)
        wrapper.pack(fill=BOTH, expand=True, padx=10, pady=(0, 10))

        Label(
            wrapper,
            text="Ile losowan temu (biezace) vs historyczna srednia + suma skumulowana",
            bg="#111827",
            fg="white",
            font=("Arial", 10, "bold"),
        ).pack(anchor=W, padx=10, pady=(8, 4))

        self.chart_host = Frame(wrapper, bg="#0b1220")
        self.chart_host.pack(fill=BOTH, expand=True, padx=10, pady=(0, 4))

        self.chart_figure = Figure(figsize=(5, 2.6), dpi=100)
        self.chart_ax = self.chart_figure.add_subplot(111)
        self.chart_figure.patch.set_facecolor("#0b1220")
        self.chart_ax.set_facecolor("#0b1220")

        self.chart_canvas = FigureCanvasTkAgg(self.chart_figure, master=self.chart_host)
        self.chart_canvas.draw()
        self.chart_canvas.get_tk_widget().pack(fill=BOTH, expand=True)

        self.gap_stats_label = Label(
            wrapper,
            text="Srednia: --   Mediana: --   Dominanta: --   Suma: --",
            bg="#111827",
            fg="#9ca3af",
            font=("Arial", 9),
            justify=LEFT,
        )
        self.gap_stats_label.pack(anchor=W, padx=10, pady=(0, 8))

    def _update_draw_gap_chart(self):
        if not self.chart_ax or not self.chart_canvas:
            return

        gaps = self._current_draw_gaps()
        labels = [str(n) for n, _ in gaps]
        current_vals = [g if g is not None else 0 for _, g in gaps]
        historical_vals = []
        for n, _ in gaps:
            hist = self.number_gap_stats.get(n)
            if hist and hist.get("mean_gap") is not None:
                historical_vals.append(hist["mean_gap"])
            else:
                historical_vals.append(0)

        has_current = any(g is not None for _, g in gaps)
        has_history = any(v > 0 for v in historical_vals)

        self.chart_ax.clear()
        if self.chart_ax2 is not None:
            self.chart_ax2.remove()
            self.chart_ax2 = None

        self.chart_ax.set_facecolor("#0b1220")
        self.chart_figure.patch.set_facecolor("#0b1220")

        if not has_current or not labels:
            self.chart_ax.text(
                0.5, 0.5, "Brak danych 'ile losowan temu' dla biezacych liczb",
                ha="center", va="center", color="white", fontsize=9,
                transform=self.chart_ax.transAxes, wrap=True
            )
            self.chart_ax.set_xticks([])
            self.chart_ax.set_yticks([])
            self.chart_canvas.draw()
            if self.gap_stats_label:
                self.gap_stats_label.config(text="Srednia: --   Mediana: --   Dominanta: --   Suma: --")
            return

        x = np.arange(len(labels))
        width = 0.38

        self.chart_ax.bar(x - width / 2, current_vals, width,
                           label="Biezace (ile los. temu)",
                           color="#2563eb", edgecolor="#60a5fa")
        if has_history:
            self.chart_ax.bar(x + width / 2, historical_vals, width,
                               label="Historyczna srednia",
                               color="#10b981", edgecolor="#34d399")

        self.chart_ax.set_xticks(x)
        self.chart_ax.set_xticklabels(labels)
        self.chart_ax.set_ylabel("Losowan temu", color="#93c5fd", fontsize=8)
        self.chart_ax.tick_params(axis="x", colors="white", labelsize=9)
        self.chart_ax.tick_params(axis="y", colors="#93c5fd", labelsize=8)
        self.chart_ax.grid(axis="y", color="#334155", alpha=0.35, linestyle="--", linewidth=0.7)
        for spine in self.chart_ax.spines.values():
            spine.set_color("#475569")

        cumsum = list(itertools.accumulate(current_vals))
        ax2 = self.chart_ax.twinx()
        self.chart_ax2 = ax2
        ax2.plot(x, cumsum, color="#f59e0b", linewidth=2, marker="o", label="Suma skumulowana")
        ax2.set_ylabel("Suma skumulowana", color="#f59e0b", fontsize=8)
        ax2.tick_params(axis="y", colors="#f59e0b", labelsize=8)
        for spine in ax2.spines.values():
            spine.set_color("#475569")
        max_cum = max(cumsum) if cumsum else 1
        ax2.set_ylim(0, max_cum * 1.15 if max_cum > 0 else 1)

        handles1, labels1 = self.chart_ax.get_legend_handles_labels()
        handles2, labels2 = ax2.get_legend_handles_labels()
        legend = self.chart_ax.legend(
            handles1 + handles2, labels1 + labels2,
            loc="upper left", fontsize=7, framealpha=0.2, facecolor="#111827",
            labelcolor="white",
        )

        self.chart_ax.set_title("Biezace vs historyczne 'ile losowan temu' + suma skumulowana",
                                 color="white", fontsize=9)

        self.chart_figure.tight_layout()
        self.chart_canvas.draw()

        stats = self._draw_gap_summary_stats(gaps)
        if stats and self.gap_stats_label:
            mode_txt = ", ".join(str(x) for x in stats["mode"])
            self.gap_stats_label.config(
                text=(f"Srednia: {stats['mean']}   Mediana: {stats['median']}   "
                      f"Dominanta: {mode_txt}   Suma: {stats['sum']}")
            )

    def _build_ui(self):
        Label(self.root, text="Generator Lotto - Tkinter",
              font=("Arial", 20, "bold"), bg="#0b1220", fg="white").pack(pady=10)

        btn_frame = Frame(self.root, bg="#0b1220")
        btn_frame.pack(pady=5, fill=X, padx=10)

        Button(btn_frame, text="Losuj", command=self._draw,
               bg="#0ea5e9", fg="white", font=("Arial", 10, "bold"), width=14,
               relief=RAISED, bd=2).pack(side=LEFT, padx=5)

        Button(btn_frame, text="Aktualizuj", command=self._update_results,
               bg="#059669", fg="white", font=("Arial", 10, "bold"), width=14,
               relief=RAISED, bd=2).pack(side=LEFT, padx=5)

        ttk.Separator(btn_frame, orient=VERTICAL).pack(side=LEFT, fill=Y, padx=8, pady=4)

        self._data_menu = Menu(self.root, tearoff=0, bg="#1f2937", fg="white",
                               activebackground="#374151", activeforeground="white",
                               relief=FLAT, bd=0)
        self._data_menu.add_command(label="Wczytaj statystyki",
                                    command=self._pick_stats)
        self._data_menu.add_command(label="Plik historii - zmien",
                                    command=self._pick_history)
        self._data_menu.add_command(label="Historia losowan",
                                    command=self._show_history)
        self._data_menu.add_separator()
        self._data_menu.add_command(label="Baza danych",
                                    command=self._show_database)

        def _open_data_menu():
            btn = self._btn_data
            x = btn.winfo_rootx()
            y = btn.winfo_rooty() + btn.winfo_height()
            self._data_menu.tk_popup(x, y)

        self._btn_data = Button(
            btn_frame, text="Dane", command=_open_data_menu,
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

        diff_frame = Frame(left_frame, bg="#1f2937", relief=RIDGE, bd=1)
        diff_frame.pack(pady=5, fill=BOTH, expand=True)
        Label(diff_frame, text="Roznice miedzy kolejnymi liczbami",
             bg="#1f2937", fg="white", font=("Arial", 9, "bold")).pack(anchor=W, padx=5, pady=2)
        self.diff_text = Text(diff_frame, height=4, bg="#0b1220", fg="#e5e7eb",
                             font=("Courier", 9), relief=FLAT, wrap=WORD)
        self.diff_text.pack(fill=BOTH, expand=True, padx=5, pady=5)
        self.diff_text.config(state=DISABLED)

        right_frame = Frame(content_frame, bg="#111827", relief=RIDGE, bd=1)
        right_frame.pack(side=RIGHT, fill=BOTH, expand=True, padx=5)

        Label(right_frame, text="Szczegoly liczby", bg="#111827", fg="white",
             font=("Arial", 14, "bold")).pack(anchor=W, padx=10, pady=5)

        self.lbl_num = Label(right_frame, text="--", bg="#111827", fg="#ffd166",
                            font=("Arial", 32, "bold"))
        self.lbl_num.pack(pady=5)

        self.details_text = Text(right_frame, height=12, bg="#0b1220", fg="#e5e7eb",
                                font=("Courier", 10), relief=FLAT, wrap=WORD)
        self.details_text.pack(fill=BOTH, expand=True, padx=10, pady=5)
        self.details_text.config(state=DISABLED)

        self._build_chart_panel(right_frame)

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
            self.set_status("Brak statystyk - uzyj menu 'Dane' -> Wczytaj statystyki",
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

                result = run_script(script_path, timeout=300)

                if result.returncode == 0:
                    latest = find_latest_stats()
                    if latest:
                        self.root.after(100, lambda p=latest: self._load_stats(p))
                        self.root.after(200, lambda: self.set_status(
                            "Statystyki wygenerowane pomyslnie", "success"))
                    else:
                        self.root.after(0, lambda: self.set_status(
                            "Statystyki: plik nie znaleziony po generowaniu", "error", auto_reset=False))
                else:
                    err = result.stderr[:300] if result.stderr else "Nieznany blad"
                    self.root.after(0, lambda e=err: self.set_status(
                        f"Blad generowania statystyk: {e}", "error", auto_reset=False))
            except subprocess.TimeoutExpired:
                self.root.after(0, lambda: self.set_status(
                    "Timeout przy generowaniu statystyk", "error", auto_reset=False))
            except Exception as e:
                self.root.after(0, lambda ex=e: self.set_status(
                    f"Blad statystyk: {ex}", "error", auto_reset=False))

        threading.Thread(target=run, daemon=True).start()

    def _update_results(self):
        def run():
            try:
                self.root.after(0, lambda: self.set_status(
                    "Aktualizacja wynikow lotto...", "warning", auto_reset=False))

                script_path = ROOT_DIR / "scripts" / "scraper_megalotto.py"
                if not script_path.exists():
                    script_path = ROOT_DIR / "scripts" / "update_lotto_results.py"

                if not script_path.exists():
                    self.root.after(0, lambda: self.set_status(
                        f"Brak skryptu aktualizacji: {script_path}", "error", auto_reset=False))
                    return

                result = run_script(script_path, extra_args=["--update-xlsx"], timeout=120)

                if result.returncode == 0:
                    self.root.after(0, self._reload_history_main_thread)
                    self.root.after(100, lambda: self.set_status(
                        "Wyniki lotto zaktualizowane pomyslnie", "success"))
                    self.root.after(300, self._generate_stats_in_background)
                else:
                    error_msg = result.stderr[:300] if result.stderr else "Nieznany blad"
                    self.root.after(0, self._update_labels)
                    self.root.after(0, lambda e=error_msg: self.set_status(
                        f"Blad aktualizacji: {e}", "error", auto_reset=False))
                    self.root.after(100, lambda e=error_msg: messagebox.showerror(
                        "Blad", f"Nie udalo sie zaktualizowac wynikow:\n{e}"))
            except subprocess.TimeoutExpired:
                self.root.after(0, self._update_labels)
                self.root.after(0, lambda: self.set_status(
                    "Timeout - aktualizacja trwala zbyt dlugo", "error", auto_reset=False))
                self.root.after(100, lambda: messagebox.showerror(
                    "Timeout", "Aktualizacja trwala zbyt dlugo"))
            except Exception as e:
                self.root.after(0, self._update_labels)
                self.root.after(0, lambda ex=e: self.set_status(
                    f"Blad aktualizacji: {ex}", "error", auto_reset=False))
                self.root.after(100, lambda ex=e: messagebox.showerror(
                    "Blad", f"Blad aktualizacji: {ex}"))

        threading.Thread(target=run, daemon=True).start()

    def _reload_history_main_thread(self):
        self._load_history()
        self._update_labels()

    def _show_database(self):
        if not self.history_db:
            self.set_status("Baza danych nie jest zaladowana", "warning")
            messagebox.showwarning("Baza danych", "Baza danych nie jest zaladowana")
            return

        try:
            cursor = self.history_db.cursor()
            cursor.execute("SELECT id, draw_date, numbers FROM draws ORDER BY id DESC LIMIT 100")
            rows = cursor.fetchall()
        except Exception as e:
            self.set_status(f"Blad odczytu bazy: {e}", "error")
            messagebox.showerror("Blad", f"Nie moge odczytac bazy: {e}")
            return

        self.set_status(f"Otwarto widok bazy ({len(rows)} rekordow)", "info")

        db_win = Toplevel(self.root)
        db_win.title("Baza danych - Wyniki Lotto")
        db_win.geometry("900x600")
        db_win.configure(bg="#0b1220")

        Label(Frame(db_win, bg="#0b1220").pack(pady=5) or db_win,
              text=f"Wyswietlono 100 ostatnich losowan (total: {len(rows)})",
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
                self.set_status(f"Blad eksportu: {e}", "error")
                messagebox.showerror("Blad", f"Nie moge eksportowac: {e}")

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
            self.set_status("Statystyki nie znalezione - generuje...", "warning", auto_reset=False)
            self._generate_stats_in_background()

        self._update_labels()
        self._draw()

    def _load_stats(self, path):
        try:
            self.stats = LottoStatistics(path)
            self.stats_path = Path(path)
            self.freq_map = {k: v["procent"] for k, v in self.stats.frequency.items()}
            self.status_map = {i: self.stats.status_short(i) for i in range(1, 50)}
            self.set_status(f"Zaladowano statystyki: {Path(path).name}", "success")
        except Exception as exc:
            self.stats = None
            self.freq_map = {i: 0.0 for i in range(1, 50)}
            self.status_map = {i: "" for i in range(1, 50)}
            self.set_status(f"Blad statystyk: {exc}", "error", auto_reset=False)
            messagebox.showerror("Blad statystyk", str(exc))
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
            self.set_status(f"Historia zaladowana: {count} losowan", "success")
            self._compute_gap_stats_for_all_numbers()
        except Exception as e:
            self.history_db = None
            self.set_status(f"Blad bazy historii: {e}", "error", auto_reset=False)

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
            self._update_draw_gap_chart()

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
            self._update_draw_gap_chart()

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
            text="BYLA" if self._combination_exists(numbers) else "NOWA")

        self.diff_text.config(state=NORMAL)
        self.diff_text.delete("1.0", END)
        self.diff_text.insert(END, "Roznice: " + " - ".join(str(d) for d in s["diffs"]))
        self.diff_text.config(state=DISABLED)

        self._show_number(numbers[0])
        self._update_draw_gap_chart()

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
            self.set_status(f"Blad pobierania historii: {e}", "error")
            return []

    def _show_history(self):
        draws = self._get_history_draws(50)
        if not draws:
            self.set_status("Brak danych historycznych w bazie", "warning")
            messagebox.showinfo("Historia", "Brak danych historycznych w bazie")
            return

        self.set_status(f"Otwarto historie ({len(draws)} losowan)", "info")

        hist_win = Toplevel(self.root)
        hist_win.title("Historia wylosowan")
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
            details = "Brak wczytanych statystyk.\nUzyj menu 'Dane' -> Wczytaj statystyki."
        else:
            d = self.stats.number_stats(n)
            streak_txt = (f"{d['streak']['losowan_temu']} los. temu | {d['streak']['status']}"
                          if d["streak"] else "brak danych")
            roll_txt = (f"{d['rolling_current']:.1f} (min {d['rolling_min']:.1f}, max {d['rolling_max']:.1f})"
                        if d["rolling_current"] is not None else "brak danych")

            gap = self.number_gap_stats.get(n, {})
            mode_gaps = gap.get("mode_gaps") or []
            mode_txt = ", ".join(str(x) for x in mode_gaps) if mode_gaps else "brak danych"
            mean_txt = f"{gap['mean_gap']:.2f}" if gap.get("mean_gap") is not None else "brak danych"
            median_txt = str(gap["median_gap"]) if gap.get("median_gap") is not None else "brak danych"
            minmax_txt = (
                f"{gap['min_gap']} / {gap['max_gap']}"
                if gap.get("min_gap") is not None and gap.get("max_gap") is not None
                else "brak danych"
            )
            count_txt = str(gap["count"]) if gap.get("count") is not None else "brak danych"

            current_streak = d["streak"]["losowan_temu"] if d["streak"] else gap.get("current_gap")
            delta_vs_median = (
                current_streak - gap["median_gap"]
                if current_streak is not None and gap.get("median_gap") is not None
                else None
            )
            delta_txt = f"{delta_vs_median:+g}" if delta_vs_median is not None else "brak danych"
            current_txt = str(current_streak) if current_streak is not None else "brak danych"

            details_lines = [
                f"Status:      {d['status']}",
                f"Wystapienia: {d['wystapienia']}",
                f"Udzial:      {d['procent']:.2f} %",
                f"Rolling100:  {roll_txt}",
                f"Cold streak: {streak_txt}",
                "",
                "Odstepy miedzy wystapieniami (historia bazy):",
                f"  Liczba wystapien w historii: {count_txt}",
                f"  Biezacy odstep:      {current_txt}",
                f"  Srednia historyczna: {mean_txt}",
                f"  Mediana historyczna: {median_txt}",
                f"  Dominanta:           {mode_txt}",
                f"  Min / Max:           {minmax_txt}",
                f"  Biezacy vs mediana:  {delta_txt}",
                "",
            ]

            if d["last_years"]:
                details_lines.append("Ostatnie lata:")
                details_lines.append("  ".join(f"{y}: {v}" for y, v in d["last_years"]))
                details_lines.append("")
            else:
                details_lines.append("Brak danych rocznych.")
                details_lines.append("")

            if d["pairs"]:
                details_lines.append("TOP pary:")
                for p in d["pairs"]:
                    details_lines.append(f"  {p['para']} - {p['wystapienia']} razy")
            else:
                details_lines.append("Brak danych o parach.")

            details = "\n".join(details_lines)

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