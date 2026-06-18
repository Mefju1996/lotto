import sys
import random
from pathlib import Path

import pandas as pd
from PySide6.QtCore import Qt, QRectF, Signal
from PySide6.QtGui import QColor, QPainter, QPen, QBrush, QFont
from PySide6.QtWidgets import (
    QApplication, QWidget, QMainWindow, QLabel, QPushButton, QVBoxLayout,
    QHBoxLayout, QGridLayout, QFrame, QMessageBox, QFileDialog, QScrollArea,
    QLineEdit, QInputDialog
)

APP_TITLE = "Lotto Generator PySide6"
BASE_DIR = Path(__file__).resolve().parent
DEFAULT_STATS = BASE_DIR / "statystyki_lotto.xlsx"
DEFAULT_HISTORY_SHEET = "Arkusz1"


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


def calculate_draw_stats(numbers):
    differences = [abs(numbers[i] - numbers[i - 1]) for i in range(1, len(numbers))]
    parzyste = sum(1 for n in numbers if n % 2 == 0)
    niskie = sum(1 for n in numbers if n <= 24)
    return {
        'suma': sum(numbers),
        'spread': numbers[-1] - numbers[0],
        'differences': differences,
        'suma_roznic': sum(differences),
        'parzyste': parzyste,
        'nieparzyste': 6 - parzyste,
        'niskie': niskie,
        'wysokie': 6 - niskie,
    }


class LottoStatistics:
    def __init__(self, xlsx_path):
        self.xlsx_path = Path(xlsx_path)
        self.frequency = {}
        self.hot_rank = {}
        self.cold_rank = {}
        self.top_pairs = {}
        self.rolling = {}
        self.year_heatmap = {}
        self.cold_streak = {}
        self.years = []
        self.load()

    def read_sheet(self, sheet_name):
        return pd.read_excel(self.xlsx_path, sheet_name=sheet_name, engine='openpyxl')

    def load(self):
        freq_df = self.read_sheet('1_Czestotliwosc')
        for _, row in freq_df.iterrows():
            number = int(row['Liczba'])
            self.frequency[number] = {
                'wystapienia': int(row['Wystąpienia']),
                'procent': float(row['Procent']),
            }

        hot_df = self.read_sheet('2_Hot_Numbers')
        for _, row in hot_df.iterrows():
            self.hot_rank[int(row['Liczba'])] = int(row['Ranking'])

        cold_df = self.read_sheet('3_Cold_Numbers')
        for _, row in cold_df.iterrows():
            self.cold_rank[int(row['Liczba'])] = int(row['Ranking'])

        pairs_df = self.read_sheet('6_TOP50_Par')
        for _, row in pairs_df.iterrows():
            a, b = map(int, str(row['Para']).split('-'))
            pair_entry = {'para': f'{a}-{b}', 'wystapienia': int(row['Wystąpienia'])}
            self.top_pairs.setdefault(a, []).append(pair_entry)
            self.top_pairs.setdefault(b, []).append(pair_entry)

        rolling_df = self.read_sheet('16RollingFreq')
        for col in rolling_df.columns:
            col_str = str(col)
            if col_str.startswith('Liczba') and col_str.endswith('roll100'):
                number = int(col_str.replace('Liczba', '').replace('roll100', ''))
                self.rolling[number] = [int(v) for v in rolling_df[col].dropna().tolist()]

        heatmap_df = self.read_sheet('17_Heatmapa_Rok')
        self.years = [str(c) for c in heatmap_df.columns if str(c) != 'Liczba']
        for _, row in heatmap_df.iterrows():
            number = int(row['Liczba'])
            self.year_heatmap[number] = {year: int(row[year]) for year in self.years if pd.notna(row[year])}

        streak_df = self.read_sheet('13_Cold_Streaks')
        for _, row in streak_df.iterrows():
            self.cold_streak[int(row['Liczba'])] = {
                'losowan_temu': int(row['Losowań_temu']),
                'status': str(row['Status']),
            }

    def get_status_short(self, number):
        if number in self.hot_rank:
            return f'H#{self.hot_rank[number]}'
        if number in self.cold_rank:
            return f'C#{self.cold_rank[number]}'
        return ''

    def get_status_long(self, number):
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
        streak = self.cold_streak.get(number)
        return {
            'number': number,
            'wystapienia': freq['wystapienia'],
            'procent': freq['procent'],
            'status': self.get_status_long(number),
            'status_short': self.get_status_short(number),
            'rolling': rolling,
            'rolling_current': rolling[-1] if rolling else None,
            'rolling_min': min(rolling) if rolling else None,
            'rolling_max': max(rolling) if rolling else None,
            'pairs': pairs,
            'years': yearly,
            'last_years': list(yearly.items())[-8:] if yearly else [],
            'streak': streak,
        }


class HeatmapWidget(QWidget):
    numberClicked = Signal(int)

    def __init__(self):
        super().__init__()
        self.freq = {i: 0.0 for i in range(1, 50)}
        self.selected = []
        self.status = {}
        self.rects = {}
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
        cold = QColor('#274060')
        hot = QColor('#d1495b')
        r = int(cold.red() + (hot.red() - cold.red()) * ratio)
        g = int(cold.green() + (hot.green() - cold.green()) * ratio)
        b = int(cold.blue() + (hot.blue() - cold.blue()) * ratio)
        base = QColor(r, g, b)
        if selected:
            return QColor('#ffd166')
        return base

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.fillRect(self.rect(), QColor('#111827'))
        cols = 7
        margin = 12
        gap = 8
        width = (self.width() - 2 * margin - gap * (cols - 1)) / cols
        height = (self.height() - 2 * margin - gap * (cols - 1)) / cols
        self.rects = {}
        for idx, number in enumerate(range(1, 50)):
            row = idx // cols
            col = idx % cols
            x = margin + col * (width + gap)
            y = margin + row * (height + gap)
            rect = QRectF(x, y, width, height)
            self.rects[number] = rect
            selected = number in self.selected
            color = self.color_for_value(self.freq.get(number, 0), selected)
            painter.setBrush(QBrush(color))
            painter.setPen(QPen(QColor('#374151'), 1))
            painter.drawRoundedRect(rect, 10, 10)
            status = self.status.get(number, '')
            text_color = QColor('#111827') if selected else QColor('white')
            painter.setPen(text_color)
            font = QFont('Arial', 10)
            font.setBold(True)
            painter.setFont(font)
            painter.drawText(rect.adjusted(0, 4, 0, -12), Qt.AlignCenter, str(number))
            small = QFont('Arial', 7)
            painter.setFont(small)
            painter.drawText(rect.adjusted(2, int(height / 2) - 2, -2, -2), Qt.AlignHCenter | Qt.AlignBottom, status)

    def mousePressEvent(self, event):
        pos = event.position()
        for number, rect in self.rects.items():
            if rect.contains(pos):
                self.numberClicked.emit(number)
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
        p.fillRect(self.rect(), QColor('#111827'))
        p.setPen(QColor('white'))
        p.setFont(QFont('Arial', 10, QFont.Bold))
        p.drawText(12, 20, 'Trend rolling100')
        if len(self.values) < 2:
            p.setPen(QColor('#94a3b8'))
            p.drawText(self.rect(), Qt.AlignCenter, 'Brak danych do wykresu')
            return
        values = self.values[-80:]
        left, top, right, bottom = 16, 34, self.width() - 16, self.height() - 18
        min_v, max_v = min(values), max(values)
        if min_v == max_v:
            max_v += 1
        points = []
        for i, value in enumerate(values):
            x = left + (right - left) * i / (len(values) - 1)
            y = bottom - (value - min_v) * (bottom - top) / (max_v - min_v)
            points.append((x, y))
        p.setPen(QPen(QColor('#22c55e'), 2))
        for i in range(1, len(points)):
            p.drawLine(int(points[i - 1][0]), int(points[i - 1][1]), int(points[i][0]), int(points[i][1]))
        p.setPen(QColor('#94a3b8'))
        p.setFont(QFont('Arial', 8))
        p.drawText(12, self.height() - 4, f'min {min_v}')
        p.drawText(self.width() - 60, 16, f'max {max_v}')


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
        p.fillRect(self.rect(), QColor('#1f2937'))
        p.setPen(QColor('white'))
        p.setFont(QFont('Arial', 10, QFont.Bold))
        p.drawText(12, 20, f'{self.title}: {self.value}')
        bar = QRectF(12, 32, self.width() - 24, 20)
        p.setPen(Qt.NoPen)
        p.setBrush(QColor('#374151'))
        p.drawRoundedRect(bar, 10, 10)
        scale = bar.width() / (self.max_value - self.min_value)
        good_x = bar.x() + (self.good_min - self.min_value) * scale
        good_w = (self.good_max - self.good_min) * scale
        p.setBrush(QColor('#10b981'))
        p.drawRoundedRect(QRectF(good_x, bar.y(), good_w, bar.height()), 10, 10)
        marker_x = bar.x() + (self.value - self.min_value) * scale
        p.setBrush(QColor('#ffd166'))
        p.drawEllipse(QRectF(marker_x - 7, bar.y() - 4, 14, 28))
        p.setPen(QColor('#d1d5db'))
        p.setFont(QFont('Arial', 8))
        p.drawText(12, 66, f'zakres: {self.min_value}-{self.max_value} | preferowane: {self.good_min}-{self.good_max}')


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
        p.fillRect(self.rect(), QColor('#111827'))
        p.setPen(QColor('white'))
        p.setFont(QFont('Arial', 10, QFont.Bold))
        p.drawText(12, 20, 'Różnice między kolejnymi liczbami')
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
            color = QColor('#60a5fa') if value <= 20 else QColor('#ef4444')
            p.setBrush(color)
            p.setPen(Qt.NoPen)
            p.drawRoundedRect(QRectF(x, y, bar_w - 12, h), 6, 6)
            p.setPen(QColor('#e5e7eb'))
            p.setFont(QFont('Arial', 8))
            p.drawText(QRectF(x, bottom + 2, bar_w - 12, 14), Qt.AlignCenter, f'd{i+1}')
            p.drawText(QRectF(x, y - 16, bar_w - 12, 14), Qt.AlignCenter, str(value))


class BallButton(QPushButton):
    def __init__(self, number=0):
        super().__init__('--' if number == 0 else str(number))
        self.setFixedSize(72, 72)
        self.setCursor(Qt.PointingHandCursor)
        self.setStyleSheet('background:#2563eb; color:white; border:none; border-radius:36px; font-size:26px; font-weight:700;')

    def set_number(self, number):
        self.setText(str(number))


class StatCard(QFrame):
    def __init__(self, title):
        super().__init__()
        self.setStyleSheet('background:#1f2937; border:1px solid #374151; border-radius:14px;')
        layout = QVBoxLayout(self)
        self.title = QLabel(title)
        self.title.setStyleSheet('color:#9ca3af; font-size:12px;')
        self.value = QLabel('--')
        self.value.setStyleSheet('color:white; font-size:26px; font-weight:700;')
        layout.addWidget(self.title)
        layout.addWidget(self.value)

    def set_value(self, value):
        self.value.setText(str(value))


class LottoWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.resize(1320, 920)
        self.freq_map = {i: 0.0 for i in range(1, 50)}
        self.status_map = {}
        self.history_df = None
        self.stats_provider = None
        self.stats_path = DEFAULT_STATS if DEFAULT_STATS.exists() else None
        self.history_path = self.stats_path
        self.history_sheet = DEFAULT_HISTORY_SHEET
        self.current_numbers = []
        self.setup_ui()
        self.load_stats(self.stats_path)
        self.select_history_file(initial=True)
        self.generate_and_update()

    def setup_ui(self):
        root = QWidget()
        self.setCentralWidget(root)
        root.setStyleSheet('background:#0b1220; color:white;')
        main = QVBoxLayout(root)

        title = QLabel('Generator Lotto - PySide6')
        title.setStyleSheet('font-size:26px; font-weight:700; color:white;')
        subtitle = QLabel('Kliknij kulkę lub pole na heatmapie, aby zobaczyć statystyki liczby.')
        subtitle.setStyleSheet('font-size:13px; color:#cbd5e1;')
        main.addWidget(title)
        main.addWidget(subtitle)

        controls = QHBoxLayout()
        self.draw_button = QPushButton('🎲 Losuj')
        self.draw_button.setStyleSheet('background:#0ea5e9; color:white; padding:10px 18px; border:none; border-radius:10px; font-weight:700;')
        self.draw_button.clicked.connect(self.generate_and_update)
        controls.addWidget(self.draw_button)

        self.choose_history_button = QPushButton('📂 Wybierz plik historii')
        self.choose_history_button.setStyleSheet('background:#334155; color:white; padding:10px 18px; border:none; border-radius:10px;')
        self.choose_history_button.clicked.connect(self.select_history_file)
        controls.addWidget(self.choose_history_button)

        self.choose_sheet_button = QPushButton('🧾 Wybierz arkusz')
        self.choose_sheet_button.setStyleSheet('background:#334155; color:white; padding:10px 18px; border:none; border-radius:10px;')
        self.choose_sheet_button.clicked.connect(self.select_history_sheet)
        controls.addWidget(self.choose_sheet_button)

        controls.addStretch(1)
        main.addLayout(controls)

        self.path_label = QLabel('')
        self.path_label.setStyleSheet('color:#fbbf24; font-size:12px;')
        main.addWidget(self.path_label)

        content = QHBoxLayout()
        main.addLayout(content, 1)

        left_scroll = QScrollArea()
        left_scroll.setWidgetResizable(True)
        left_scroll.setFrameShape(QFrame.NoFrame)
        left_wrap = QWidget()
        left = QVBoxLayout(left_wrap)
        left_scroll.setWidget(left_wrap)
        content.addWidget(left_scroll, 3)

        right_panel = QFrame()
        right_panel.setStyleSheet('background:#111827; border:1px solid #1f2937; border-radius:16px;')
        right = QVBoxLayout(right_panel)
        content.addWidget(right_panel, 2)

        balls_layout = QHBoxLayout()
        self.ball_buttons = []
        for _ in range(6):
            btn = BallButton()
            btn.clicked.connect(self.handle_ball_click)
            self.ball_buttons.append(btn)
            balls_layout.addWidget(btn)
        left.addLayout(balls_layout)

        cards = QGridLayout()
        self.card_sum = StatCard('Suma')
        self.card_spread = StatCard('Spread')
        self.card_diff_sum = StatCard('Suma różnic')
        self.card_pn = StatCard('P/N')
        self.card_lh = StatCard('L/H')
        self.card_history = StatCard('Historia')
        cards.addWidget(self.card_sum, 0, 0)
        cards.addWidget(self.card_spread, 0, 1)
        cards.addWidget(self.card_diff_sum, 0, 2)
        cards.addWidget(self.card_pn, 1, 0)
        cards.addWidget(self.card_lh, 1, 1)
        cards.addWidget(self.card_history, 1, 2)
        left.addLayout(cards)

        self.range_sum = RangeBarWidget('Suma', 80, 220, 110, 185)
        self.range_spread = RangeBarWidget('Spread', 8, 48, 31, 42)
        left.addWidget(self.range_sum)
        left.addWidget(self.range_spread)

        self.heatmap = HeatmapWidget()
        self.heatmap.numberClicked.connect(self.show_number_stats)
        left.addWidget(self.heatmap)

        self.diff_widget = DifferencesWidget()
        left.addWidget(self.diff_widget)

        details_title = QLabel('Szczegóły liczby')
        details_title.setStyleSheet('font-size:20px; font-weight:700; color:white;')
        right.addWidget(details_title)

        self.number_title = QLabel('Kliknij liczbę')
        self.number_title.setStyleSheet('font-size:34px; font-weight:800; color:#ffd166;')
        right.addWidget(self.number_title)

        self.details_label = QLabel('Brak wybranej liczby.')
        self.details_label.setWordWrap(True)
        self.details_label.setStyleSheet('font-size:13px; color:#e5e7eb;')
        right.addWidget(self.details_label)

        self.spark = SparklineWidget()
        right.addWidget(self.spark)

        self.years_label = QLabel('')
        self.years_label.setWordWrap(True)
        self.years_label.setStyleSheet('background:#0b1220; border-radius:12px; padding:12px; color:#cbd5e1;')
        right.addWidget(self.years_label)

        self.pairs_label = QLabel('')
        self.pairs_label.setWordWrap(True)
        self.pairs_label.setStyleSheet('background:#0b1220; border-radius:12px; padding:12px; color:#cbd5e1;')
        right.addWidget(self.pairs_label)

        right.addStretch(1)

    def load_stats(self, path):
        if not path or not Path(path).exists():
            QMessageBox.warning(self, 'Brak statystyk', 'Nie znaleziono pliku statystyki_lotto.xlsx.')
            return
        try:
            self.stats_provider = LottoStatistics(path)
            self.freq_map = {k: v['procent'] for k, v in self.stats_provider.frequency.items()}
            self.status_map = {i: self.stats_provider.get_status_short(i) for i in range(1, 50)}
        except Exception as exc:
            QMessageBox.critical(self, 'Błąd statystyk', str(exc))

    def select_history_file(self, initial=False):
        start = str(self.history_path.parent if self.history_path else BASE_DIR)
        selected, _ = QFileDialog.getOpenFileName(self, 'Wybierz plik historii losowań', start, 'Excel (*.xlsx *.xls);;Wszystkie pliki (*)')
        if selected:
            self.history_path = Path(selected)
        elif initial and self.stats_path:
            self.history_path = self.stats_path
        self.load_history_dataframe()
        self.update_path_label()

    def load_history_dataframe(self):
        if not self.history_path or not self.history_path.exists():
            self.history_df = None
            return
        try:
            self.history_df = pd.read_excel(self.history_path, sheet_name=self.history_sheet, engine='openpyxl', usecols='A:F')
        except Exception:
            self.history_df = None

    def available_sheets(self):
        if not self.history_path or not self.history_path.exists():
            return []
        try:
            xls = pd.ExcelFile(self.history_path, engine='openpyxl')
            return xls.sheet_names
        except Exception:
            return []

    def select_history_sheet(self):
        sheets = self.available_sheets()
        if not sheets:
            QMessageBox.warning(self, 'Brak arkuszy', 'Nie udało się odczytać listy arkuszy z wybranego pliku.')
            return
        current_index = sheets.index(self.history_sheet) if self.history_sheet in sheets else 0
        sheet, ok = QInputDialog.getItem(self, 'Wybierz arkusz', 'Arkusz historii:', sheets, current_index, False)
        if ok and sheet:
            self.history_sheet = sheet
            self.load_history_dataframe()
            self.update_path_label()

    def update_path_label(self):
        file_txt = str(self.history_path) if self.history_path else 'brak'
        self.path_label.setText(f'Plik historii: {file_txt} | Arkusz: {self.history_sheet}')

    def combination_exists(self, numbers):
        if self.history_df is None:
            return False
        try:
            return any((self.history_df.values == numbers).all(axis=1))
        except Exception:
            return False

    def generate_and_update(self):
        numbers = generate_numbers()
        self.current_numbers = numbers
        stats = calculate_draw_stats(numbers)
        for btn, number in zip(self.ball_buttons, numbers):
            btn.set_number(number)
        self.card_sum.set_value(stats['suma'])
        self.card_spread.set_value(stats['spread'])
        self.card_diff_sum.set_value(stats['suma_roznic'])
        self.card_pn.set_value(f"{stats['parzyste']}P-{stats['nieparzyste']}N")
        self.card_lh.set_value(f"{stats['niskie']}L-{stats['wysokie']}H")
        self.card_history.set_value('BYŁA' if self.combination_exists(numbers) else 'NOWA')
        self.range_sum.set_value(stats['suma'])
        self.range_spread.set_value(stats['spread'])
        self.diff_widget.set_differences(stats['differences'])
        self.heatmap.set_data(self.freq_map, numbers, self.status_map)
        self.show_number_stats(numbers[0])

    def handle_ball_click(self):
        sender = self.sender()
        try:
            number = int(sender.text())
        except Exception:
            return
        self.show_number_stats(number)

    def show_number_stats(self, number):
        if not self.stats_provider:
            return
        stats = self.stats_provider.get_number_stats(number)
        self.number_title.setText(str(number))
        streak = stats['streak']
        streak_text = 'brak danych'
        if streak:
            streak_text = f"{streak['losowan_temu']} los. temu | {streak['status']}"
        rolling_text = 'brak danych'
        if stats['rolling_current'] is not None:
            rolling_text = f"{stats['rolling_current']} (min {stats['rolling_min']}, max {stats['rolling_max']})"
        self.details_label.setText(
            f"Status: {stats['status']}
"
            f"Wystąpienia: {stats['wystapienia']}
"
            f"Udział: {stats['procent']:.2f}%
"
            f"Rolling100: {rolling_text}
"
            f"Cold streak: {streak_text}"
        )
        self.spark.set_values(stats['rolling'])
        if stats['last_years']:
            years_text = 'Ostatnie lata:
' + '
'.join([f"{year}: {value}" for year, value in stats['last_years']])
        else:
            years_text = 'Brak danych rocznych.'
        self.years_label.setText(years_text)
        if stats['pairs']:
            pairs_text = 'TOP pary:
' + '
'.join([f"• {p['para']} — {p['wystapienia']} razy" for p in stats['pairs']])
        else:
            pairs_text = 'Brak tej liczby w TOP50 par.'
        self.pairs_label.setText(pairs_text)
        self.heatmap.set_data(self.freq_map, self.current_numbers, self.status_map)


def main():
    app = QApplication(sys.argv)
    window = LottoWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
