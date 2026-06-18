import sys
import random
import sqlite3
import json
from pathlib import Path

import pandas as pd
from PySide6.QtCore import Qt, QRectF,Signal
from PySide6.QtGui import QColor, QPainter, QPen, QBrush, QFont
from PySide6.QtWidgets import (
    QApplication, QWidget, QMainWindow, QLabel, QPushButton, QVBoxLayout,
    QHBoxLayout, QGridLayout, QFrame, QMessageBox, QFileDialog, QScrollArea,
    QInputDialog
)

APP_TITLE = "Generator Lotto - PySide6 v5"
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


class HeatmapWidget(QWidget):
    numberClicked = Signal(int)

    def __init__(self):
        super().__init__()
        self.freq = {i: 0.0 for i in range(1, 50)}
        self.selected = []
        self.status = {}
        self._rects = {}
        self.setMinimumHeight(260)
        self.setCursor(Qt.PointingHandCursor)

    def set_data(self, freq, selected, status):
        self.freq = freq or self.freq
        self.selected = selected or []
        self.status = status or {}
        self.update()

    def _color(self, value, selected):
        if selected:
            return QColor("#ffd166")
        mn = min(self.freq.values(), default=0)
        mx = max(self.freq.values(), default=1)
        ratio = 0 if mx == mn else (value - mn) / (mx - mn)
        c = QColor("#274060")
        h = QColor("#d1495b")
        return QColor(
            int(c.red() + (h.red() - c.red()) * ratio),
            int(c.green() + (h.green() - c.green()) * ratio),
            int(c.blue() + (h.blue() - c.blue()) * ratio),
        )

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.fillRect(self.rect(), QColor("#111827"))
        cols, margin, gap = 7, 12, 8
        rows = 7
        w = (self.width() - 2 * margin - gap * (cols - 1)) / cols
        h = (self.height() - 2 * margin - gap * (rows - 1)) / rows
        self._rects = {}
        for idx, num in enumerate(range(1, 50)):
            r, col = idx // cols, idx % cols
            x = margin + col * (w + gap)
            y = margin + r * (h + gap)
            rect = QRectF(x, y, w, h)
            self._rects[num] = rect
            sel = num in self.selected
            painter.setBrush(QBrush(self._color(self.freq.get(num, 0), sel)))
            painter.setPen(QPen(QColor("#374151"), 1))
            painter.drawRoundedRect(rect, 10, 10)
            painter.setPen(QColor("#111827") if sel else QColor("white"))
            font = QFont("Arial", 10)
            font.setBold(True)
            painter.setFont(font)
            painter.drawText(rect.adjusted(0, 4, 0, -12), Qt.AlignCenter, str(num))
            painter.setFont(QFont("Arial", 7))
            painter.setPen(QColor("#111827") if sel else QColor("#94a3b8"))
            painter.drawText(rect.adjusted(2, int(h / 2) - 2, -2, -2),
                             Qt.AlignHCenter | Qt.AlignBottom,
                             self.status.get(num, ""))

    def mousePressEvent(self, event):
        try:
            pos = event.position()
        except AttributeError:
            pos = event.posF()
        for num, rect in self._rects.items():
            if rect.contains(pos):
                self.numberClicked.emit(num)
                return
        super().mousePressEvent(event)


class SparklineWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.values = []
        self.setMinimumHeight(130)

    def set_values(self, values):
        self.values = values or []
        self.update()

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        p.fillRect(self.rect(), QColor("#111827"))
        p.setPen(QColor("white"))
        p.setFont(QFont("Arial", 10, QFont.Bold))
        p.drawText(12, 20, "Trend rolling100")
        vals = self.values[-80:]
        if len(vals) < 2:
            p.setPen(QColor("#94a3b8"))
            p.drawText(self.rect(), Qt.AlignCenter, "Brak danych do wykresu")
            return
        L, T, R, B = 16, 34, self.width() - 16, self.height() - 18
        mn, mx = min(vals), max(vals)
        if mn == mx:
            mx += 1
        pts = [(L + (R - L) * i / (len(vals) - 1),
                B - (v - mn) * (B - T) / (mx - mn)) for i, v in enumerate(vals)]
        p.setPen(QPen(QColor("#22c55e"), 2))
        for i in range(1, len(pts)):
            p.drawLine(int(pts[i - 1][0]), int(pts[i - 1][1]), int(pts[i][0]), int(pts[i][1]))
        p.setPen(QColor("#94a3b8"))
        p.setFont(QFont("Arial", 8))
        p.drawText(12, self.height() - 4, f"min {mn:.1f}")
        p.drawText(self.width() - 70, 16, f"max {mx:.1f}")


class RangeBarWidget(QWidget):
    def __init__(self, title, vmin, vmax, gmin, gmax):
        super().__init__()
        self.title = title
        self.vmin, self.vmax, self.gmin, self.gmax = vmin, vmax, gmin, gmax
        self.value = vmin
        self.setMinimumHeight(72)

    def set_value(self, v):
        self.value = v
        self.update()

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        p.fillRect(self.rect(), QColor("#1f2937"))
        p.setPen(QColor("white"))
        p.setFont(QFont("Arial", 10, QFont.Bold))
        p.drawText(12, 20, f"{self.title}: {self.value}")
        bar = QRectF(12, 32, self.width() - 24, 20)
        p.setPen(Qt.NoPen)
        p.setBrush(QColor("#374151"))
        p.drawRoundedRect(bar, 10, 10)
        scale = bar.width() / (self.vmax - self.vmin)
        gx = bar.x() + (self.gmin - self.vmin) * scale
        gw = (self.gmax - self.gmin) * scale
        p.setBrush(QColor("#10b981"))
        p.drawRoundedRect(QRectF(gx, bar.y(), gw, bar.height()), 10, 10)
        raw_x = (max(self.vmin, min(self.vmax, self.value)) - self.vmin) * scale
        mx = bar.x() + raw_x
        p.setBrush(QColor("#ffd166"))
        p.drawEllipse(QRectF(mx - 7, bar.y() - 4, 14, 28))
        p.setPen(QColor("#9ca3af"))
        p.setFont(QFont("Arial", 8))
        p.drawText(12, 66, f"zakres: {self.vmin}-{self.vmax}  preferowane: {self.gmin}-{self.gmax}")


class DifferencesWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.diffs = []
        self.setMinimumHeight(140)

    def set_differences(self, diffs):
        self.diffs = diffs or []
        self.update()

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        p.fillRect(self.rect(), QColor("#111827"))
        p.setPen(QColor("white"))
        p.setFont(QFont("Arial", 10, QFont.Bold))
        p.drawText(12, 20, "Roznice miedzy kolejnymi liczbami")
        if not self.diffs:
            return
        L, B, T = 16, self.height() - 20, 36
        cw = self.width() - 32
        ch = B - T
        mx = max(20, max(self.diffs))
        bw = cw / max(len(self.diffs), 1)
        for i, v in enumerate(self.diffs):
            x = L + i * bw + 6
            h = (v / mx) * (ch - 8)
            y = B - h
            color = QColor("#60a5fa") if v <= 20 else QColor("#ef4444")
            p.setBrush(color)
            p.setPen(Qt.NoPen)
            p.drawRoundedRect(QRectF(x, y, bw - 12, h), 6, 6)
            p.setPen(QColor("#e5e7eb"))
            p.setFont(QFont("Arial", 8))
            p.drawText(QRectF(x, B + 2, bw - 12, 14), Qt.AlignCenter, f"d{i+1}")
            p.drawText(QRectF(x, y - 16, bw - 12, 14), Qt.AlignCenter, str(v))


class BallButton(QPushButton):
    def __init__(self):
        super().__init__("--")
        self.num = 0
        self.setFixedSize(72, 72)
        self.setCursor(Qt.PointingHandCursor)
        self._apply_style("#2563eb")

    def set_number(self, n, color="#2563eb"):
        self.num = int(n)
        self.setText(str(n))
        self._apply_style(color)

    def _apply_style(self, color):
        self.setStyleSheet(
            f"background:{color}; color:white; border:none; "
            f"border-radius:36px; font-size:26px; font-weight:700;"
        )


class StatCard(QFrame):
    def __init__(self, title):
        super().__init__()
        self.setStyleSheet("background:#1f2937; border:1px solid #374151; border-radius:14px;")
        lay = QVBoxLayout(self)
        t = QLabel(title)
        t.setStyleSheet("color:#9ca3af; font-size:12px;")
        self.val = QLabel("--")
        self.val.setStyleSheet("color:white; font-size:24px; font-weight:700;")
        lay.addWidget(t)
        lay.addWidget(self.val)

    def set_value(self, v):
        self.val.setText(str(v))


class LottoWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.resize(1340, 940)
        self.stats = None
        self.stats_path = None
        self.freq_map = {i: 0.0 for i in range(1, 50)}
        self.status_map = {i: "" for i in range(1, 50)}
        self.history_db = None
        self.history_path = None
        self.history_sheet = DEFAULT_HISTORY_SHEET
        self.current_numbers = []
        self._build_ui()
        self._auto_load()

    @staticmethod
    def _lbl(text, style=""):
        lbl = QLabel(text)
        if style:
            lbl.setStyleSheet(style)
        return lbl

    @staticmethod
    def _add_btn(layout, text, color, slot, bold=False):
        b = QPushButton(text)
        fw = "700" if bold else "400"
        b.setStyleSheet(
            f"background:{color}; color:white; padding:10px 16px; "
            f"border:none; border-radius:10px; font-weight:{fw};"
        )
        b.clicked.connect(slot)
        layout.addWidget(b)

    def _build_ui(self):
        root = QWidget()
        self.setCentralWidget(root)
        root.setStyleSheet("background:#0b1220; color:white;")
        main = QVBoxLayout(root)
        main.setSpacing(6)
        main.addWidget(self._lbl("Generator Lotto  -  PySide6 v5",
                                  "font-size:24px; font-weight:800;"))
        main.addWidget(self._lbl("Kliknij kulke lub pole heatmapy.",
                                  "font-size:12px; color:#94a3b8;"))

        ctrl = QHBoxLayout()
        ctrl.setSpacing(8)
        self._add_btn(ctrl, "Losuj",               "#0ea5e9", self._draw,         bold=True)
        self._add_btn(ctrl, "Wczytaj statystyki",  "#7c3aed", self._pick_stats,   bold=True)
        self._add_btn(ctrl, "Plik historii",        "#334155", self._pick_history)
        self._add_btn(ctrl, "Arkusz",               "#334155", self._pick_sheet)
        ctrl.addStretch()
        main.addLayout(ctrl)

        self.lbl_stats   = self._lbl("Statystyki: brak", "color:#a78bfa; font-size:11px;")
        self.lbl_history = self._lbl("Historia: brak",   "color:#fbbf24; font-size:11px;")
        main.addWidget(self.lbl_stats)
        main.addWidget(self.lbl_history)

        content = QHBoxLayout()
        main.addLayout(content, 1)

        left_scroll = QScrollArea()
        left_scroll.setWidgetResizable(True)
        left_scroll.setFrameShape(QFrame.NoFrame)
        left_scroll.setStyleSheet("background:transparent;")
        left_wrap = QWidget()
        left_wrap.setStyleSheet("background:transparent;")
        left = QVBoxLayout(left_wrap)
        left.setSpacing(10)
        left_scroll.setWidget(left_wrap)
        content.addWidget(left_scroll, 3)

        balls_row = QHBoxLayout()
        balls_row.setSpacing(12)
        self.balls = []
        for _ in range(6):
            btn = BallButton()
            btn.clicked.connect(self._on_ball_clicked)
            self.balls.append(btn)
            balls_row.addWidget(btn)
        balls_row.addStretch()
        left.addLayout(balls_row)

        cards = QGridLayout()
        cards.setSpacing(8)
        self.c_sum    = StatCard("Suma")
        self.c_spread = StatCard("Spread")
        self.c_dsum   = StatCard("Suma roznic")
        self.c_pn     = StatCard("P / N")
        self.c_lh     = StatCard("L / H")
        self.c_hist   = StatCard("Historia")
        for i, c in enumerate([self.c_sum, self.c_spread, self.c_dsum,
                                self.c_pn,  self.c_lh,    self.c_hist]):
            cards.addWidget(c, i // 3, i % 3)
        left.addLayout(cards)

        self.bar_sum    = RangeBarWidget("Suma",   80, 220, 110, 185)
        self.bar_spread = RangeBarWidget("Spread",  8,  48,  31,  42)
        left.addWidget(self.bar_sum)
        left.addWidget(self.bar_spread)

        self.heatmap = HeatmapWidget()
        self.heatmap.numberClicked.connect(self._show_number)
        left.addWidget(self.heatmap)

        self.diffs_widget = DifferencesWidget()
        left.addWidget(self.diffs_widget)

        right_frame = QFrame()
        right_frame.setStyleSheet(
            "background:#111827; border:1px solid #1f2937; border-radius:16px;")
        right = QVBoxLayout(right_frame)
        right.setSpacing(10)
        right.setContentsMargins(16, 16, 16, 16)
        content.addWidget(right_frame, 2)

        right.addWidget(self._lbl("Szczegoly liczby", "font-size:18px; font-weight:700;"))
        self.lbl_num = self._lbl("--",
                                  "font-size:40px; font-weight:800; color:#ffd166;")
        self.lbl_details = self._lbl("Kliknij kulke lub pole heatmapy.",
                                      "font-size:13px; color:#e5e7eb;")
        self.lbl_details.setWordWrap(True)
        right.addWidget(self.lbl_num)
        right.addWidget(self.lbl_details)

        self.sparkline = SparklineWidget()
        right.addWidget(self.sparkline)

        self.lbl_years = self._lbl("",
            "background:#0b1220; border-radius:10px; padding:10px; "
            "color:#cbd5e1; font-size:12px;")
        self.lbl_years.setWordWrap(True)
        right.addWidget(self.lbl_years)

        self.lbl_pairs = self._lbl("",
            "background:#0b1220; border-radius:10px; padding:10px; "
            "color:#cbd5e1; font-size:12px;")
        self.lbl_pairs.setWordWrap(True)
        right.addWidget(self.lbl_pairs)
        right.addStretch()

    def _update_labels(self):
        sp = str(self.stats_path) if self.stats_path else "brak"
        ok = "OK" if self.stats else "brak danych"
        self.lbl_stats.setText(f"Statystyki: {sp}  |  {ok}")
        hp = str(self.history_path) if self.history_path else "brak"
        self.lbl_history.setText(f"Historia: {hp}  |  Arkusz: {self.history_sheet}")

    def _auto_load(self):
        if DEFAULT_STATS.exists():
            self._load_stats(DEFAULT_STATS)
        if DEFAULT_HISTORY_DB.exists():
            self.history_path = DEFAULT_HISTORY_DB
            self._load_history()
        self._update_labels()
        self._draw()

    def _load_stats(self, path):
        try:
            self.stats = LottoStatistics(path)
            self.stats_path = Path(path)
            self.freq_map   = {k: v["procent"]
                               for k, v in self.stats.frequency.items()}
            self.status_map = {i: self.stats.status_short(i) for i in range(1, 50)}
        except Exception as exc:
            self.stats = None
            self.freq_map   = {i: 0.0 for i in range(1, 50)}
            self.status_map = {i: "" for i in range(1, 50)}
            QMessageBox.critical(self, "Blad statystyk", str(exc))
        self._update_labels()

    def _load_history(self):
        if self.history_db:
            self.history_db.close()
            self.history_db = None
        
        if not self.history_path or not self.history_path.exists():
            return
        try:
            self.history_db = sqlite3.connect(str(self.history_path))
        except Exception:
            self.history_db = None

    def _pick_stats(self):
        start = str(self.stats_path.parent if self.stats_path else BASE_DIR)
        path, _ = QFileDialog.getOpenFileName(
            self, "Wybierz plik statystyk", start, "Excel (*.xlsx *.xls)")
        if not path:
            return
        self._load_stats(path)
        if not self.history_path:
            self.history_path = Path(path)
            self._load_history()
        self._update_labels()
        if self.current_numbers:
            self._show_number(self.current_numbers[0])

    def _pick_history(self):
        start = str(self.history_path.parent if self.history_path else BASE_DIR)
        path, _ = QFileDialog.getOpenFileName(
            self, "Wybierz plik historii", start, "All (*.xlsx *.xls *.db);;Excel (*.xlsx *.xls);;SQLite (*.db)")
        if path:
            self.history_path = Path(path)
            self._load_history()
            self._update_labels()

    def _pick_sheet(self):
        if not self.history_path or not self.history_path.exists():
            QMessageBox.warning(self, "Brak pliku", "Najpierw wybierz plik historii.")
            return
        
        # Check if it's a SQLite database
        if str(self.history_path).endswith('.db'):
            QMessageBox.information(self, "Baza SQLite", "Baza danych SQLite nie wymaga wyboru arkusza.")
            return
        
        try:
            sheets = pd.ExcelFile(self.history_path, engine="openpyxl").sheet_names
        except Exception:
            QMessageBox.warning(self, "Blad", "Nie mozna odczytac arkuszy.")
            return
        idx = sheets.index(self.history_sheet) if self.history_sheet in sheets else 0
        sheet, ok = QInputDialog.getItem(
            self, "Wybierz arkusz", "Arkusz historii:", sheets, idx, False)
        if ok and sheet:
            self.history_sheet = sheet
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
        hot  = self.stats.hot_rank  if self.stats else {}
        cold = self.stats.cold_rank if self.stats else {}
        for btn, n in zip(self.balls, numbers):
            if n in hot:
                color = "#ef4444"
            elif n in cold:
                color = "#475569"
            else:
                color = "#2563eb"
            btn.set_number(n, color)
        self.c_sum.set_value(s["suma"])
        self.c_spread.set_value(s["spread"])
        self.c_dsum.set_value(s["suma_roznic"])
        self.c_pn.set_value(f"{s['parzyste']}P-{s['nieparzyste']}N")
        self.c_lh.set_value(f"{s['niskie']}L-{s['wysokie']}H")
        self.c_hist.set_value("BYLA" if self._combination_exists(numbers) else "NOWA")
        self.bar_sum.set_value(s["suma"])
        self.bar_spread.set_value(s["spread"])
        self.diffs_widget.set_differences(s["diffs"])
        self.heatmap.set_data(self.freq_map, numbers, self.status_map)
        self._show_number(numbers[0])

    def _on_ball_clicked(self):
        btn = self.sender()
        if isinstance(btn, BallButton) and btn.num:
            self._show_number(btn.num)

    def _show_number(self, n):
        self.lbl_num.setText(str(n))
        if not self.stats:
            self.lbl_details.setText(
                "Brak wczytanych statystyk.\n"
                "Uzyj przycisku  Wczytaj statystyki.")
            self.sparkline.set_values([])
            self.lbl_years.setText("")
            self.lbl_pairs.setText("")
            return
        d = self.stats.number_stats(n)
        streak_txt = "brak danych"
        if d["streak"]:
            streak_txt = (f"{d['streak']['losowan_temu']} los. temu  |  "
                          f"{d['streak']['status']}")
        roll_txt = "brak danych"
        if d["rolling_current"] is not None:
            roll_txt = (f"{d['rolling_current']:.1f}  "
                        f"(min {d['rolling_min']:.1f}, max {d['rolling_max']:.1f})")
        self.lbl_details.setText(
            f"Status:      {d['status']}\n"
            f"Wystapienia: {d['wystapienia']}\n"
            f"Udzial:      {d['procent']:.2f} %\n"
            f"Rolling100:  {roll_txt}\n"
            f"Cold streak: {streak_txt}"
        )
        self.sparkline.set_values(d["rolling"])
        if d["last_years"]:
            yt = "Ostatnie lata:\n" + "  ".join(f"{y}: {v}" for y, v in d["last_years"])
        else:
            yt = "Brak danych rocznych."
        self.lbl_years.setText(yt)
        if d["pairs"]:
            pt = "TOP pary:\n" + "\n".join(
                f"  {p['para']}  -  {p['wystapienia']} razy" for p in d["pairs"])
        else:
            pt = "Brak danych o parach."
        self.lbl_pairs.setText(pt)
        self.heatmap.set_data(self.freq_map, self.current_numbers, self.status_map)

    def closeEvent(self, event):
        """Close SQLite connection before closing application."""
        if self.history_db:
            self.history_db.close()
        event.accept()


def main():
    app = QApplication(sys.argv)
    win = LottoWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
