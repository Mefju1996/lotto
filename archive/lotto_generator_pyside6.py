import sys
import random
from pathlib import Path

import pandas as pd
from PySide6.QtCore import Qt, QRectF
from PySide6.QtGui import QColor, QPainter, QPen, QBrush, QFont
from PySide6.QtWidgets import (
    QApplication, QWidget, QMainWindow, QLabel, QPushButton, QVBoxLayout,
    QHBoxLayout, QGridLayout, QFrame, QMessageBox, QFileDialog, QScrollArea
)

APP_TITLE = "Lotto Generator PySide6"
DEFAULT_STATS = Path("statystyki_lotto.xlsx")


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


def generate_numbers():
    while True:
        numbers = sorted(random.sample(range(1, 50), 6))
        if is_valid(numbers):
            return numbers


class HeatmapWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.freq = {i: 0.0 for i in range(1, 50)}
        self.selected = []
        self.status = {}
        self.setMinimumHeight(250)

    def set_data(self, freq, selected, status):
        self.freq = freq or self.freq
        self.selected = selected or []
        self.status = status or {}
        self.update()

    def color_for_value(self, value, selected=False):
        min_v = min(self.freq.values()) if self.freq else 0
        max_v = max(self.freq.values()) if self.freq else 1
        ratio = 0 if max_v == min_v else (value - min_v) / (max_v - min_v)
        cold = QColor("#274060")
        hot = QColor("#d1495b")
        r = int(cold.red() + (hot.red() - cold.red()) * ratio)
        g = int(cold.green() + (hot.green() - cold.green()) * ratio)
        b = int(cold.blue() + (hot.blue() - cold.blue()) * ratio)
        base = QColor(r, g, b)
        if selected:
            return QColor("#ffd166")
        return base

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.fillRect(self.rect(), QColor("#111827"))
        cols = 7
        rows = 7
        margin = 12
        gap = 8
        w = (self.width() - 2 * margin - gap * (cols - 1)) / cols
        h = (self.height() - 2 * margin - gap * (rows - 1)) / rows

        for idx, number in enumerate(range(1, 50)):
            row = idx // cols
            col = idx % cols
            x = margin + col * (w + gap)
            y = margin + row * (h + gap)
            rect = QRectF(x, y, w, h)
            selected = number in self.selected
            color = self.color_for_value(self.freq.get(number, 0), selected)
            painter.setBrush(QBrush(color))
            painter.setPen(QPen(QColor("#374151"), 1))
            painter.drawRoundedRect(rect, 10, 10)

            status = self.status.get(number, "")
            text_color = QColor("#111827") if selected else QColor("white")
            painter.setPen(text_color)
            font = QFont("Arial", 10)
            font.setBold(True)
            painter.setFont(font)
            painter.drawText(rect.adjusted(0, 4, 0, -12), Qt.AlignCenter, str(number))

            small = QFont("Arial", 7)
            painter.setFont(small)
            painter.drawText(rect.adjusted(2, h/2 - 2, -2, -2), Qt.AlignHCenter | Qt.AlignBottom, status)


class RangeBarWidget(QWidget):
    def __init__(self, title, min_value, max_value, good_min, good_max):
        super().__init__()
        self.title = title
        self.min_value = min_value
        self.max_value = max_value
        self.good_min = good_min
        self.good_max = good_max
        self.value = min_value
        self.setMinimumHeight(72)

    def set_value(self, value):
        self.value = value
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

        scale = bar.width() / (self.max_value - self.min_value)
        good_x = bar.x() + (self.good_min - self.min_value) * scale
        good_w = (self.good_max - self.good_min) * scale
        p.setBrush(QColor("#10b981"))
        p.drawRoundedRect(QRectF(good_x, bar.y(), good_w, bar.height()), 10, 10)

        marker_x = bar.x() + (self.value - self.min_value) * scale
        p.setBrush(QColor("#ffd166"))
        p.drawEllipse(QRectF(marker_x - 7, bar.y() - 4, 14, 28))

        p.setPen(QColor("#d1d5db"))
        p.setFont(QFont("Arial", 8))
        p.drawText(12, 66, f"zakres: {self.min_value}-{self.max_value} | preferowane: {self.good_min}-{self.good_max}")


class DifferencesWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.differences = []
        self.setMinimumHeight(140)

    def set_differences(self, diffs):
        self.differences = diffs or []
        self.update()

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        p.fillRect(self.rect(), QColor("#111827"))
        p.setPen(QColor("white"))
        p.setFont(QFont("Arial", 10, QFont.Bold))
        p.drawText(12, 20, "Różnice między kolejnymi liczbami")

        if not self.differences:
            return

        left, bottom, top = 16, self.height() - 20, 36
        chart_w = self.width() - 32
        chart_h = bottom - top
        max_v = max(20, max(self.differences))
        bar_w = chart_w / max(len(self.differences), 1)

        for i, value in enumerate(self.differences):
            x = left + i * bar_w + 6
            h = (value / max_v) * (chart_h - 8)
            y = bottom - h
            color = QColor("#60a5fa") if value <= 20 else QColor("#ef4444")
            p.setBrush(color)
            p.setPen(Qt.NoPen)
            p.drawRoundedRect(QRectF(x, y, bar_w - 12, h), 6, 6)
            p.setPen(QColor("#e5e7eb"))
            p.setFont(QFont("Arial", 8))
            p.drawText(QRectF(x, bottom + 2, bar_w - 12, 14), Qt.AlignCenter, f"d{i+1}")
            p.drawText(QRectF(x, y - 16, bar_w - 12, 14), Qt.AlignCenter, str(value))


class BallLabel(QLabel):
    def __init__(self, text="--"):
        super().__init__(text)
        self.setAlignment(Qt.AlignCenter)
        self.setFixedSize(72, 72)
        self.setStyleSheet(
            "background:#2563eb; color:white; border-radius:36px; font-size:26px; font-weight:700;"
        )


class StatCard(QFrame):
    def __init__(self, title):
        super().__init__()
        self.setStyleSheet("background:#1f2937; border:1px solid #374151; border-radius:14px;")
        layout = QVBoxLayout(self)
        self.title = QLabel(title)
        self.title.setStyleSheet("color:#9ca3af; font-size:12px;")
        self.value = QLabel("--")
        self.value.setStyleSheet("color:white; font-size:26px; font-weight:700;")
        layout.addWidget(self.title)
        layout.addWidget(self.value)

    def set_value(self, value):
        self.value.setText(str(value))


class LottoWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.resize(1300, 900)
        self.freq_map = {i: 0.0 for i in range(1, 50)}
        self.hot_set = set()
        self.cold_set = set()
        self.history_df = None
        self.stats_path = DEFAULT_STATS if DEFAULT_STATS.exists() else None
        self.setup_ui()
        self.load_stats(self.stats_path)
        self.generate_and_update()

    def setup_ui(self):
        root = QWidget()
        self.setCentralWidget(root)
        root.setStyleSheet("background:#0b1220; color:white;")
        main = QVBoxLayout(root)

        top = QHBoxLayout()
        title = QLabel("Generator Lotto - wizualizacja PySide6")
        title.setStyleSheet("font-size:28px; font-weight:800;")
        top.addWidget(title)
        top.addStretch()
        self.path_label = QLabel("Statystyki: brak")
        self.path_label.setStyleSheet("color:#9ca3af;")
        btn_load = QPushButton("Wczytaj XLSX")
        btn_load.clicked.connect(self.pick_stats_file)
        btn_draw = QPushButton("Losuj")
        btn_draw.clicked.connect(self.generate_and_update)
        btn_draw.setStyleSheet("background:#2563eb; color:white; padding:10px 18px; border-radius:10px; font-weight:700;")
        top.addWidget(self.path_label)
        top.addWidget(btn_load)
        top.addWidget(btn_draw)
        main.addLayout(top)

        balls_row = QHBoxLayout()
        self.ball_labels = []
        for _ in range(6):
            ball = BallLabel()
            self.ball_labels.append(ball)
            balls_row.addWidget(ball)
        balls_row.addStretch()
        self.history_badge = QLabel("Historia: --")
        self.history_badge.setStyleSheet("background:#1f2937; border-radius:10px; padding:12px; font-size:16px; font-weight:600;")
        balls_row.addWidget(self.history_badge)
        main.addLayout(balls_row)

        cards = QGridLayout()
        self.card_sum = StatCard("Suma")
        self.card_spread = StatCard("Spread")
        self.card_diffsum = StatCard("Suma różnic")
        self.card_pn = StatCard("Parzyste / Nieparzyste")
        self.card_lh = StatCard("Niskie / Wysokie")
        self.card_hot = StatCard("Hot w zestawie")
        for idx, card in enumerate([self.card_sum, self.card_spread, self.card_diffsum, self.card_pn, self.card_lh, self.card_hot]):
            cards.addWidget(card, idx // 3, idx % 3)
        main.addLayout(cards)

        middle = QHBoxLayout()
        left_col = QVBoxLayout()
        self.heatmap = HeatmapWidget()
        left_col.addWidget(self.wrap_group("Heatmapa 1-49", self.heatmap))
        self.diff_widget = DifferencesWidget()
        left_col.addWidget(self.wrap_group("Struktura zestawu", self.diff_widget))
        middle.addLayout(left_col, 3)

        right_col = QVBoxLayout()
        self.sum_bar = RangeBarWidget("Suma", 80, 220, 110, 185)
        self.spread_bar = RangeBarWidget("Spread", 8, 48, 31, 42)
        self.info_label = QLabel()
        self.info_label.setWordWrap(True)
        self.info_label.setStyleSheet("background:#1f2937; border:1px solid #374151; border-radius:14px; padding:14px; color:#e5e7eb;")
        right_col.addWidget(self.wrap_group("Pozycja względem zakresów", self.sum_bar))
        right_col.addWidget(self.wrap_group("Rozpiętość liczb", self.spread_bar))
        right_col.addWidget(self.wrap_group("Opis zestawu", self.info_label))
        right_col.addStretch()
        middle.addLayout(right_col, 2)
        main.addLayout(middle)

    def wrap_group(self, title, widget):
        frame = QFrame()
        frame.setStyleSheet("background:#111827; border:1px solid #1f2937; border-radius:16px;")
        layout = QVBoxLayout(frame)
        label = QLabel(title)
        label.setStyleSheet("font-size:16px; font-weight:700; color:#f3f4f6; padding:4px;")
        layout.addWidget(label)
        layout.addWidget(widget)
        return frame

    def pick_stats_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Wybierz plik statystyk", str(Path.cwd()), "Excel (*.xlsx *.xls)")
        if path:
            self.load_stats(Path(path))
            self.generate_and_update()

    def load_stats(self, path):
        if not path or not Path(path).exists():
            self.path_label.setText("Statystyki: brak pliku, tryb uproszczony")
            return
        try:
            freq_df = pd.read_excel(path, sheet_name="1_Czestotliwosc")
            hot_df = pd.read_excel(path, sheet_name="2_Hot_Numbers")
            cold_df = pd.read_excel(path, sheet_name="3_Cold_Numbers")
            self.freq_map = {int(row["Liczba"]): float(row["Wystąpienia"]) for _, row in freq_df.iterrows()}
            self.hot_set = set(int(x) for x in hot_df["Liczba"].head(20).tolist())
            self.cold_set = set(int(x) for x in cold_df["Liczba"].head(20).tolist())
            self.stats_path = Path(path)
            self.path_label.setText(f"Statystyki: {self.stats_path.name}")
        except Exception as e:
            QMessageBox.warning(self, "Błąd", f"Nie udało się wczytać statystyk: {e}")
            self.path_label.setText("Statystyki: błąd odczytu")

    def check_combination_exists(self, numbers):
        if not self.stats_path or not self.stats_path.exists():
            return False
        try:
            existing_data = pd.read_excel(self.stats_path, engine="openpyxl", sheet_name=0, usecols="A:F")
            return any((existing_data.values == numbers).all(axis=1))
        except Exception:
            return False

    def build_status_map(self):
        status = {}
        for i in range(1, 50):
            if i in self.hot_set:
                status[i] = "HOT"
            elif i in self.cold_set:
                status[i] = "COLD"
            else:
                status[i] = "MID"
        return status

    def generate_and_update(self):
        numbers = generate_numbers()
        differences = [abs(numbers[i] - numbers[i - 1]) for i in range(1, len(numbers))]
        suma = sum(numbers)
        spread = numbers[-1] - numbers[0]
        suma_roznic = sum(differences)
        parzyste = sum(1 for n in numbers if n % 2 == 0)
        nieparzyste = 6 - parzyste
        niskie = sum(1 for n in numbers if n <= 24)
        wysokie = 6 - niskie
        hot_hits = sum(1 for n in numbers if n in self.hot_set)
        exists = self.check_combination_exists(numbers)

        for label, value in zip(self.ball_labels, numbers):
            color = "#ef4444" if value in self.hot_set else "#2563eb"
            if value in self.cold_set:
                color = "#475569"
            label.setText(str(value))
            label.setStyleSheet(f"background:{color}; color:white; border-radius:36px; font-size:26px; font-weight:700;")

        self.card_sum.set_value(suma)
        self.card_spread.set_value(spread)
        self.card_diffsum.set_value(suma_roznic)
        self.card_pn.set_value(f"{parzyste}P / {nieparzyste}N")
        self.card_lh.set_value(f"{niskie}L / {wysokie}H")
        self.card_hot.set_value(hot_hits)
        self.history_badge.setText("Historia: była już" if exists else "Historia: nowa kombinacja")
        self.history_badge.setStyleSheet(
            f"background:{'#7f1d1d' if exists else '#064e3b'}; border-radius:10px; padding:12px; font-size:16px; font-weight:600;"
        )

        self.sum_bar.set_value(suma)
        self.spread_bar.set_value(spread)
        self.diff_widget.set_differences(differences)
        self.heatmap.set_data(self.freq_map, numbers, self.build_status_map())

        hot_list = [n for n in numbers if n in self.hot_set]
        cold_list = [n for n in numbers if n in self.cold_set]
        self.info_label.setText(
            f"Liczby: {numbers}\n"
            f"Różnice: {differences}\n"
            f"Status hot: {hot_list if hot_list else 'brak'}\n"
            f"Status cold: {cold_list if cold_list else 'brak'}\n"
            f"Ocena reguł: suma {'OK' if 110 <= suma <= 185 else 'poza'}; "
            f"spread {'typowy' if 31 <= spread <= 42 else 'mniej typowy'}; "
            f"suma różnic {'OK' if 27 <= suma_roznic <= 47 else 'poza'}"
        )


def main():
    app = QApplication(sys.argv)
    win = LottoWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
