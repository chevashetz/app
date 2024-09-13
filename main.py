import csv
import logging
import os
import re
import sys

import numpy as np
import pandas as pd
from PyQt6 import uic
from PyQt6.QtCore import Qt, pyqtSignal, QRectF, QLineF, QStringListModel
from PyQt6.QtGui import (QAction, QUndoStack, QUndoCommand, QKeySequence, QTextDocument, QFont, QPixmap, QPainter, QPen,
                         QBrush, QColor, QPainterPath)
from PyQt6.QtSql import QSqlDatabase, QSqlQuery, QSqlTableModel
from PyQt6.QtWidgets import (QApplication, QMainWindow, QLineEdit, QPushButton, QStackedWidget, QHeaderView,
                             QTableWidget, QTableWidgetItem, QComboBox, QFileDialog, QDialog, QInputDialog, QVBoxLayout,
                             QMenu, QGraphicsScene, QGraphicsView, QUndoView, QWidget, QHBoxLayout, QLabel, QMessageBox,
                             QScrollArea, QGraphicsPathItem, QListView, QStyledItemDelegate)
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

logging.basicConfig(level=logging.ERROR, format='%(asctime)s - %(levelname)s - %(message)s')
path1 = "сsv_files/"
path2 = "db_files/"
path3 = "images/"
path4 = "msh_files/"


class CenteredItemDelegate(QStyledItemDelegate):
    def initStyleOption(self, option, index):
        super().initStyleOption(option, index)
        # Устанавливаем выравнивание по центру
        option.displayAlignment = Qt.AlignmentFlag.AlignCenter


class ShadingDrawer:
    def __init__(self, lengths, ends, diameter_offsets, diameter_hole, x_offset, scene=None, reverse=False):
        self.horizontal_offset = 30
        self.reverse = reverse
        self.x_offset = x_offset
        self.diameter_hole = diameter_hole
        self.lengths = lengths
        self.scene = scene
        self.ends = [ends[0]] + [max(ends[i] - ends[i - 1], 0) for i in range(1, len(ends))]
        self.rectangles = [[end, max(self.horizontal_offset, offset)] for end, offset in
                           zip(self.ends, diameter_offsets)]

        self.offsets = [min(self.horizontal_offset, offset) for offset in diameter_offsets]
        self.rectangles[0][0] += 5

        if reverse:
            for i in range(0, len(self.offsets)):
                self.offsets[i] *= -1
                self.rectangles[i][1] *= -1

        self.pen_hole = QPen(Qt.GlobalColor.black, 2)

    def build_curve(self, path, index, x, y):

        # Текущие размеры и отступы блока
        height, width = self.rectangles[index]
        offset = self.offsets[index]

        x += offset

        # Начальная точка в верхнем левом углу блока
        if index == 0:
            path.moveTo(x, y)
        # Слева направо
        path.lineTo(x, y)
        # Сверху-вниз линия
        path.lineTo(x, y + height)

        if index < len(self.rectangles) - 1:
            self.build_curve(path, index + 1, x, y + height)

        # Линия вверх
        path.lineTo(x + width, y + height)
        path.lineTo(x + width, y)
        # Влево
        path.lineTo(x + width - offset, y)

        self.scene.addLine(QLineF(x + width, y + height, x + width, y), self.pen_hole)

        if index == 0:
            # Верхняя линия
            path.lineTo(x, y)
            self.scene.addLine(QLineF(x + width, y, x, y), self.pen_hole)
        else:
            self.scene.addLine(QLineF(x + width, y, x + width - offset, y), self.pen_hole)

    def draw_curve(self):
        path = QPainterPath()
        self.build_curve(path, 0,
                         self.x_offset - self.horizontal_offset if not self.reverse else self.x_offset + self.horizontal_offset,
                         self.ends[0] - self.lengths[0])  # Начальная точка
        path.closeSubpath()

        # Добавление пути на сцену
        path_item = QGraphicsPathItem(path)
        # Установка прозрачного пера
        transparent_pen = QPen(QColor(0, 0, 0, 0))  # Черный цвет с альфа-прозрачностью 50
        path_item.setPen(transparent_pen)
        # Установка штриховки для заливки
        hatch_brush = QBrush(Qt.BrushStyle.DiagCrossPattern)
        path_item.setBrush(hatch_brush)
        self.scene.addItem(path_item)


class MplCanvas(FigureCanvas):
    def __init__(self, parent=None, width=4.5, height=1.5, dpi=100):
        self.fig = Figure(figsize=(width, height), dpi=dpi)
        self.axes = self.fig.add_subplot(111)
        super().__init__(self.fig)
        self.setParent(parent)
        self.setFixedSize(int(width * dpi), int(height * dpi))


class DatabaseManager:
    def __init__(self, db_path=None):
        self.database = QSqlDatabase.addDatabase("QSQLITE")
        if db_path:
            self.database.setDatabaseName(db_path)
            self.open_database()

    def open_database(self):
        if not self.database.open():
            QMessageBox.critical(None, "Ошибка", "Не удалось открыть файл базы данных.")
            return False
        return True

    def get_tables(self):
        return self.database.tables()

    def load_table(self, table_name):
        model = QSqlTableModel()
        model.setTable(table_name)
        model.select()

        if model.lastError().isValid():
            QMessageBox.critical(None, "Ошибка", f"Не удалось загрузить данные: {model.lastError().text()}")
            return None
        return model

    def close_database(self):
        self.database.close()

    custom_headers = ["Название", "Индексация", "Верх", "Низ", "Коэффициент ", "Плотность"]

    def preview_table(self, table_name, custom_headers=None, parent=None):
        model = self.load_table(table_name)
        if not model:
            return None

        dialog = QDialog(parent)
        dialog.setWindowTitle(f"Предварительный просмотр - {table_name}")
        dialog.setGeometry(30, 150, 1410, 420)

        layout = QVBoxLayout()

        table_widget = QTableWidget()
        table_widget.setRowCount(model.rowCount())
        table_widget.setColumnCount(model.columnCount() - 1)

        if custom_headers:
            headers = custom_headers
        else:
            headers = [model.headerData(i, Qt.Orientation.Horizontal) for i in
                       range(1, model.columnCount())]

        table_widget.setHorizontalHeaderLabels(headers)

        for row in range(model.rowCount()):
            for col in range(1, model.columnCount()):
                data = model.data(model.index(row, col))
                item = QTableWidgetItem(str(data))
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                table_widget.setItem(row, col - 1, item)

        header = table_widget.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

        layout.addWidget(table_widget)

        button_layout = QHBoxLayout()

        insert_button = QPushButton("Вставить в таблицу")
        return_button = QPushButton("Вернуться к выбору таблицы")

        insert_button.setFixedSize(685, 40)
        return_button.setFixedSize(685, 40)

        button_layout.addWidget(insert_button)
        button_layout.addWidget(return_button)
        layout.addLayout(button_layout)

        dialog.setLayout(layout)

        # Связываем кнопки с действиями
        insert_button.clicked.connect(dialog.accept)
        return_button.clicked.connect(dialog.reject)

        if dialog.exec() == QDialog.DialogCode.Accepted:
            return model
        return None

class PasteCommand(QUndoCommand):
    def __init__(self, tableWidget, text_data, start_row, start_col, description, parent=None):
        super().__init__(description, parent)
        self.tableWidget = tableWidget
        self.text_data = text_data
        self.start_row = start_row
        self.start_col = start_col
        self.old_data = []
        self.new_rows_needed = 0

    def undo(self):
        for row, col_data in self.old_data:
            for col, data in col_data.items():
                item = QTableWidgetItem()
                item.setData(Qt.ItemDataRole.DisplayRole, data)
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)  # Центрация текста
                self.tableWidget.setItem(row, col, item)
        # Удаление добавленных строк
        for _ in range(self.new_rows_needed):
            self.tableWidget.removeRow(self.tableWidget.rowCount() - 1)

    def redo(self):
        rows = self.text_data
        self.old_data = []

        total_cells_needed = sum(len(row) for row in rows)
        current_cells_available = ((self.tableWidget.rowCount() - self.start_row) * self.tableWidget.columnCount()
                                   - self.start_col)
        self.new_rows_needed = max(0,
                                   (total_cells_needed - current_cells_available + self.tableWidget.columnCount() - 1)
                                   // self.tableWidget.columnCount())

        for _ in range(self.new_rows_needed):
            self.tableWidget.insertRow(self.tableWidget.rowCount())

        doc = QTextDocument()

        current_row = self.start_row
        for row_data in rows:
            columns = row_data
            old_row_data = {}

            if current_row >= self.tableWidget.rowCount():
                self.tableWidget.insertRow(self.tableWidget.rowCount())

            current_col = self.start_col
            for col_index, value in enumerate(columns):
                if current_col >= self.tableWidget.columnCount():
                    current_row += 1
                    current_col = 0
                    if current_row >= self.tableWidget.rowCount():
                        self.tableWidget.insertRow(self.tableWidget.rowCount())

                item = self.tableWidget.item(current_row, current_col)
                old_row_data[current_col] = item.text() if item else ""

                item = QTableWidgetItem()
                doc.setHtml(value)
                item.setData(Qt.ItemDataRole.DisplayRole, doc.toPlainText())
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)  # Центрация текста
                self.tableWidget.setItem(current_row, current_col, item)
                current_col += 1

            self.old_data.append((current_row, old_row_data))
            current_row += 1

class ComboHeader(QHeaderView):
    def __init__(self, parent=None):
        super(ComboHeader, self).__init__(Qt.Orientation.Horizontal, parent)
        self.setStretchLastSection(True)
        self.combobox = QComboBox(self)
        self.combobox.addItems(["Азимут (град)", "Азимут маг(град)", "Азимут дир(град)"])
        self.combobox.setStyleSheet("QComboBox { text-align: center; }")
        for i in range(self.combobox.count()):
            self.combobox.setItemData(i, Qt.AlignmentFlag.AlignCenter, Qt.ItemDataRole.TextAlignmentRole)
        self.setSectionsClickable(True)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self.combobox:
            index = 2
            x = self.sectionViewportPosition(index)
            w = self.sectionSize(index)
            self.combobox.setGeometry(x, 0, w, self.height())

class UpdateTableCommand(QUndoCommand):
    def __init__(self, knbk_table_instance, old_data, new_data, description="загрузку КНБК"):
        super().__init__(description)
        self.knbk_table_instance = knbk_table_instance
        self.old_data = old_data
        self.new_data = new_data

    def undo(self):
        self.knbk_table_instance.update_table_widget(self.old_data)

    def redo(self):
        self.knbk_table_instance.update_table_widget(self.new_data)

    def update_table_widget(self, data):
        self.knbk_table_instance.clear_images()

        self.knbk_table_instance.tbl_KNBK.clearContents()
        self.knbk_table_instance.tbl_KNBK.setRowCount(len(data))

        for row_index, row_data in enumerate(data):
            for col_index, value in enumerate(row_data):
                if value is None:
                    value = ""
                item = QTableWidgetItem(value)
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.knbk_table_instance.tbl_KNBK.setItem(row_index, col_index, item)
            self.knbk_table_instance.file_key = self.knbk_table_instance.tbl_KNBK.item(row_index, 0).text()
            self.knbk_table_instance.add_image()

        # Восстановление начального состояния
        self.knbk_table_instance.restore_initial_state()
        self.knbk_table_instance.add_label(f"КНБК - {self.knbk_table_instance.tbl_KNBK.item(0, 4).text()} мм")

class CsvTableDialog(QDialog):
    data_selected = pyqtSignal(list, str)

    def __init__(self, file_name, load_table=False, initial_sort_value_KNBK=None, sort_value_casing_srings=None,
                 parent=None):
        super().__init__(parent)
        self.file_name = file_name
        self.load_table = load_table
        self.initial_sort_value_KNBK = initial_sort_value_KNBK
        self.sort_value_casing_srings = sort_value_casing_srings
        self.sort_order = Qt.SortOrder.AscendingOrder
        self.sort_column = -1
        self.initUI()

    def initUI(self):
        self.setWindowTitle('CSV Data')
        self.setGeometry(60, 100, 1400, 700)
        layout = QVBoxLayout()

        self.tableWidget = QTableWidget(self)
        layout.addWidget(self.tableWidget)

        self.setLayout(layout)
        self.load_csv()

        self.tableWidget.setSortingEnabled(False)
        self.tableWidget.horizontalHeader().setSortIndicatorShown(True)
        self.tableWidget.horizontalHeader().sectionClicked.connect(self.on_header_clicked)

        # Initial sort if values are provided
        if self.load_table == False:
            if self.sort_value_casing_srings is not None:
                print("Sorting by sort_value_casing_srings")
                self.sort_column = 3
                self.sort_by_column_and_value(self.sort_column, self.sort_value_casing_srings)
            elif self.initial_sort_value_KNBK is not None:
                print(f"Sorting by initial_sort_value_KNBK: {self.initial_sort_value_KNBK}")
                self.sort_column = 5
                self.sort_by_column_and_value_custom(self.sort_column, self.initial_sort_value_KNBK)

    def load_csv(self):
        try:
            with open(self.file_name, newline='', encoding='utf-8') as csvfile:
                csvreader = csv.reader(csvfile)
                data = list(csvreader)

                if data:
                    headers = data[0]
                    self.tableWidget.setColumnCount(len(headers))
                    self.tableWidget.setHorizontalHeaderLabels(headers)
                    self.tableWidget.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)

                    for row_data in data[1:]:
                        row = self.tableWidget.rowCount()
                        self.tableWidget.insertRow(row)
                        for col, cell_data in enumerate(row_data):
                            item = QTableWidgetItem(cell_data.strip())
                            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                            self.tableWidget.setItem(row, col, item)

                    if not self.load_table:
                        self.tableWidget.cellDoubleClicked.connect(self.cell_was_double_clicked)
                    else:
                        self.tableWidget.cellDoubleClicked.connect(self.cell_was_double_clicked_2)
                        self.sort_by_column_and_value(0, self.sort_value_casing_srings)
                        print("Connected cellDoubleClicked signal to cell_was_double_clicked_2")
                else:
                    print("No data found in the file.")
        except Exception as e:
            print(f"Error loading CSV: {e}")

    def sort_by_column_and_value(self, column, value):
        data = []
        for row in range(self.tableWidget.rowCount()):
            row_data = []
            for col in range(self.tableWidget.columnCount()):
                item = self.tableWidget.item(row, col)
                row_data.append(item.text() if item else "")
            data.append(row_data)

        matching_rows = [row for row in data if row[column] == value]
        non_matching_rows = [row for row in data if row[column] != value]

        sorted_data = matching_rows + non_matching_rows

        self.update_table_with_sorted_data(sorted_data)

    def sort_by_column_and_value_custom(self, column, value):
        try:
            value_numeric = float(value.split('З-')[1])
            print(value_numeric)
        except (IndexError, ValueError):
            value_numeric = float('inf')

        data = []
        for row in range(self.tableWidget.rowCount()):
            row_data = []
            for col in range(self.tableWidget.columnCount()):
                item = self.tableWidget.item(row, col)
                row_data.append(item.text() if item else "")
            data.append(row_data)

        def extract_numeric(text):
            if text.startswith('З-'):
                try:
                    return float(text.split('З-')[1])
                except (IndexError, ValueError):
                    return float('inf')
            return float('inf')

        matching_rows = [row for row in data if extract_numeric(row[column]) == value_numeric]
        non_matching_rows = [row for row in data if extract_numeric(row[column]) != value_numeric]

        sorted_data = matching_rows + non_matching_rows

        self.update_table_with_sorted_data(sorted_data)

    def on_header_clicked(self, logical_index):
        if self.sort_column == logical_index:
            self.sort_order = Qt.SortOrder.DescendingOrder if self.sort_order == Qt.SortOrder.AscendingOrder else Qt.SortOrder.AscendingOrder
        else:
            self.sort_order = Qt.SortOrder.AscendingOrder
        self.sort_column = logical_index
        self.sort_table()

    def sort_table(self):
        data = []
        for row in range(self.tableWidget.rowCount()):
            row_data = []
            for column in range(self.tableWidget.columnCount()):
                item = self.tableWidget.item(row, column)
                row_data.append(item.text() if item else "")
            data.append(row_data)

        data.sort(key=lambda row: self.custom_sort_key(row[self.sort_column]),
                  reverse=self.sort_order == Qt.SortOrder.DescendingOrder)

        self.update_table_with_sorted_data(data)

    def update_table_with_sorted_data(self, sorted_data):
        self.tableWidget.setRowCount(0)
        for row_data in sorted_data:
            row = self.tableWidget.rowCount()
            self.tableWidget.insertRow(row)
            for column, item in enumerate(row_data):
                table_item = QTableWidgetItem(item)
                table_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.tableWidget.setItem(row, column, table_item)

    def custom_sort_key(self, text):
        try:
            return float(text)
        except ValueError:
            if text.startswith('З-'):
                try:
                    return int(text.split('З-')[1])
                except ValueError:
                    return text.lower()
            return text.lower()

    def cell_was_double_clicked(self, row, column):
        try:
            row_data = []
            for col in range(self.tableWidget.columnCount()):
                item = self.tableWidget.item(row, col)
                if item:
                    row_data.append(item.text())
                else:
                    row_data.append('')
            self.data_selected.emit(row_data, "")
            self.accept()
        except Exception as e:
            print(f"Error in cell_was_double_clicked: {e}")

    def cell_was_double_clicked_2(self, row, column):
        try:
            column = 1
            item = self.tableWidget.item(row, column)
            if item:
                data = item.text()
                parts = data.split(';')

                row_data = []
                keys = []
                for part in parts:
                    part = part.strip()
                    if not part:
                        continue
                    if '_' in part:
                        key, name = part.split('_', 1)
                        key = key.strip()
                        name = name.strip()
                        keys.append(key)
                        csv_path = f'{path1}{key}.csv'
                        try:
                            with open(csv_path, "r", encoding='utf-8') as csvfile:
                                csv_reader = csv.reader(csvfile)
                                for csv_row in csv_reader:
                                    if name == csv_row[0].strip():
                                        row_data.append([key] + csv_row)
                                        break
                        except FileNotFoundError:
                            print(f"File not found: {csv_path}")
                        except Exception as e:
                            print(f"Error reading {csv_path}: {e}")
                    else:
                        print(f"No '_' found in part: {part}")
                self.data_selected.emit(row_data, ", ".join(keys))
                self.accept()
            else:
                print("Error: item is None")
        except Exception as e:
            print(f"Error in cell_was_double_clicked_2: {e}")

class KNBK_Table(QWidget):
    def __init__(self, index, sort_key=None, parent=None):
        super(KNBK_Table, self).__init__(parent)
        uic.loadUi('table.ui', self)

        self.sort_key = sort_key
        self.setup_ui()
        self.labels = []
        self.current_y = 700
        self.max_height = 700

        self.scroll_area = QScrollArea(self)
        self.scroll_area.setGeometry(1445, 50, 75, 700)
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.scroll_area.hide()

        self.image_container = QWidget(self.scroll_area)
        self.image_container.setFixedWidth(67)
        self.image_container_layout = QVBoxLayout(self.image_container)
        self.image_container_layout.setSpacing(0)
        self.image_container_layout.setContentsMargins(0, 0, 0, 40)
        self.scroll_area.setWidget(self.image_container)

        header = self.tbl_KNBK.horizontalHeaderItem(0)
        if header is not None:
            header.setTextAlignment(Qt.AlignmentFlag.AlignCenter)

        self.add_image(mode="static", static_path=path3 + 'Долото.png')

    def setup_ui(self):
        self.label = None
        self.undo_stack = QUndoStack(self)
        self.undo_view = QUndoView(self.undo_stack)
        self.tbl_KNBK: QTableWidget = self.findChild(QTableWidget, 'table_KNBK')
        self.tbl_KNBK.itemChanged.connect(self.update_label)
        self.tbl_KNBK.itemChanged.connect(self.center_text_in_item)

        self.btn_add_row_KNBK: QPushButton = self.findChild(QPushButton, 'pushButton_add_row')
        self.btn_delete_row_KNBK: QPushButton = self.findChild(QPushButton, 'pushButton_delete_row')
        self.btn_load_table: QPushButton = self.findChild(QPushButton, 'pushButton_load_table')
        self.btn_row_up: QPushButton = self.findChild(QPushButton, 'pushButton_row_up')
        self.btn_row_down: QPushButton = self.findChild(QPushButton, 'pushButton_row_down')
        self.btn_add_page: QPushButton = self.findChild(QPushButton, 'pushButton_add_page')
        self.btn_delete_page: QPushButton = self.findChild(QPushButton, 'pushButton_delete_page')

        self.open_file_act: QAction = self.findChild(QAction, 'actionOpen')

        self.tbl_KNBK.cellDoubleClicked.connect(self.open_csv_table_dialog)
        self.tbl_KNBK.cellDoubleClicked.connect(self.open_fixed_path_csv_dialog)
        self.tbl_KNBK.cellClicked.connect(self.add_QCombobox_cell_clicked)

        item_0_0_KNBK = QTableWidgetItem("Долото")
        item_0_0_KNBK.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        self.tbl_KNBK.setItem(0, 0, item_0_0_KNBK)

        self.btn_add_row_KNBK.clicked.connect(self.add_row_KNBK)
        self.btn_delete_row_KNBK.clicked.connect(self.delete_row_KNBK)
        self.btn_load_table.clicked.connect(self.load_table)
        self.btn_row_up.clicked.connect(self.row_up)
        self.btn_row_down.clicked.connect(self.row_down)
        self.btn_add_page.clicked.connect(self.parentWidget().add_page_2)
        self.btn_delete_page.clicked.connect(self.parentWidget().delete_page)

        self.set_column_width(0, 150)
        self.set_column_width(1, 150)
        self.set_column_width(2, 150)
        self.set_column_width(3, 150)
        self.set_column_width(4, 170)
        self.set_column_width(5, 100)
        self.set_column_width(6, 80)
        self.set_column_width(7, 110)
        self.set_column_width(8, 80)
        self.set_column_width(9, 60)
        self.set_column_width(10, 60)

    def set_column_width(self, column, width):
        self.tbl_KNBK.setColumnWidth(column, width)

    def add_row_KNBK(self):
        row_count2_1 = self.tbl_KNBK.rowCount()
        self.tbl_KNBK.setRowCount(row_count2_1 + 1)
        for column in range(self.tbl_KNBK.columnCount()):
            if column == 0:
                self.add_QCombobox(row_count=row_count2_1, column=column)
            else:
                item = QTableWidgetItem("")
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.tbl_KNBK.setItem(row_count2_1, column, item)
        self.tbl_KNBK.resizeRowsToContents()

    def add_QCombobox(self, row_count, column):
        combo = QComboBox()
        combo.setModel(QStringListModel([
            "<Не выбрано>", "ВЗД", "РУС", "Бурильные трубы", "Переводник", "УБТ", "Телеметрия", "Ясс",
            "Калибратор спиральный", "Обратный клапан", "Центратор прямой",
            "Центратор спиральный", "Предохранительный переводник"]))

        listView = QListView()

        listView.setWordWrap(True)

        combo.setView(listView)

        showPopup = combo.showPopup

        def on_popup_show():
            showPopup()
            combo.view().reset()  # Сброс состояния представления для предотвращения наслаивания

        combo.showPopup = on_popup_show

        listView.setItemDelegate(CenteredItemDelegate(listView))
        combo.setEditable(True)
        line_edit: QLineEdit = combo.lineEdit()
        line_edit.setAlignment(Qt.AlignmentFlag.AlignCenter)
        line_edit.setReadOnly(False)

        combo.currentIndexChanged.connect(
            lambda index, combo=combo, row=row_count: self.handle_combo_change(row, combo))

        self.tbl_KNBK.setCellWidget(row_count, column, combo)

    def handle_combo_change(self, row, combo):
        text = combo.currentText()
        item = QTableWidgetItem(text)
        item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        self.tbl_KNBK.setItem(row, 0, item)
        self.tbl_KNBK.removeCellWidget(row, 0)
        self.tbl_KNBK.resizeRowsToContents()

        self.file_key = text
        self.add_image(row=row)
        self.tbl_KNBK.viewport().update()

    def add_image(self, mode="dynamic", static_path=None, row=None):
        if mode == "static" and static_path:
            image_path = static_path
        elif mode == "dynamic":
            file_name = self.get_file_name()
            if file_name:
                image_path = file_name[1]
            else:
                return
        else:
            print("Incorrect mode or missing static path")
            return

        if not os.path.exists(image_path):
            print(f"Image file does not exist: {image_path}")
            return

        pixmap = QPixmap(image_path)
        new_label = QLabel(self)
        new_label.setPixmap(pixmap)
        new_label.setFixedWidth(67)
        new_label.setScaledContents(True)
        new_label.show()

        image_height = pixmap.height()
        total_height = self.current_y - image_height

        if total_height < 0:
            # Если превышает лимит, перемещаем все изображения в ScrollArea
            self.move_images_to_scroll_area()

            new_label.setParent(self.image_container)
            self.image_container_layout.insertWidget(0 if row is None else len(self.labels) - row, new_label)
            self.scroll_area.show()
        else:
            if row is None:
                new_label.move(1445, self.current_y - image_height)
            else:
                image_total_height = sum(label.height() for label in self.labels[:row])
                new_label.move(1445, self.max_height - image_total_height - image_height)
                for label in self.labels[row:]:
                    label.move(1445, self.max_height - image_total_height - image_height - label.height())

        self.current_y -= image_height
        if row is None:
            self.labels.append(new_label)
        else:
            self.labels.insert(row, new_label)


    def move_images_to_scroll_area(self):
        if not self.scroll_area.isVisible():
            # Перебираем все лейблы в обычном порядке, добавляя сначала старые элементы
            for label in self.labels[::-1]:
                label.setParent(self.image_container)

                # self.image_container_layout.addWidget(label)
            self.scroll_area.show()

    def move_images_back_to_page(self):
        if self.scroll_area.isVisible():
            self.current_y = self.max_height
            for label in self.labels:
                self.image_container_layout.removeWidget(label)
                label.setParent(self)
                label.move(1445, self.current_y - label.pixmap().height())
                label.show()
                self.current_y -= label.pixmap().height()
            if self.current_y >= 0:
                self.scroll_area.hide()

    def delete_row_KNBK(self):
        row_count2_1 = self.tbl_KNBK.currentRow()
        self.delete_image_KNBK(row_count=row_count2_1)
        self.tbl_KNBK.removeRow(row_count2_1)

    def delete_image_KNBK(self, row_count):
        if row_count > 0:
            if 0 <= row_count < len(self.labels):
                label_to_remove = self.labels.pop(row_count)
                if label_to_remove:
                    if self.scroll_area.isVisible():
                        self.image_container_layout.removeWidget(label_to_remove)
                    else:
                        self.current_y += label_to_remove.height()
                    label_to_remove.deleteLater()

            if not self.scroll_area.isVisible():
                self.current_y = self.max_height
                for label in self.labels:
                    self.current_y -= label.height()
                    label.move(1445, self.current_y)

            # Проверяем, нужно ли переместить изображения обратно на страницу
            total_height = sum(label.height() for label in self.labels)
            if total_height <= self.max_height and self.scroll_area.isVisible():
                self.move_images_back_to_page()

    def load_table(self):
        dialog = CsvTableDialog(path1 + 'КНБК.csv', load_table=True, initial_sort_value_KNBK=None,
                                sort_value_casing_srings=self.sort_key, parent=self)
        dialog.data_selected.connect(self.update_table_data_list_2)
        dialog.exec()

    def row_up(self):
        current_row = self.tbl_KNBK.currentRow()
        if current_row > 1:
            self.swap_rows(current_row, current_row - 1)
            self.tbl_KNBK.setCurrentCell(current_row - 1, 0)
            self.update_labels_after_swap(current_row, current_row - 1)

    def row_down(self):
        current_row = self.tbl_KNBK.currentRow()
        if (current_row < self.tbl_KNBK.rowCount() - 1) and current_row > 0:
            self.swap_rows(current_row, current_row + 1)
            self.tbl_KNBK.setCurrentCell(current_row + 1, 0)
            self.update_labels_after_swap(current_row, current_row + 1)

    def swap_rows(self, row1, row2):
        for column in range(self.tbl_KNBK.columnCount()):
            item1 = self.tbl_KNBK.takeItem(row1, column)
            item2 = self.tbl_KNBK.takeItem(row2, column)
            if item1:
                self.tbl_KNBK.setItem(row2, column, item1)
            if item2:
                self.tbl_KNBK.setItem(row1, column, item2)

        if self.scroll_area.isVisible():
            self.swap_images_in_scroll_area(row1, row2)
        else:
            self.swap_images_in_labels(row1, row2)

    def swap_images_in_scroll_area(self, row1, row2):
        label1 = self.labels[row1]
        label2 = self.labels[row2]
        idx1 = self.image_container_layout.indexOf(label1)
        idx2 = self.image_container_layout.indexOf(label2)
        self.image_container_layout.insertWidget(idx1, label2)
        self.image_container_layout.insertWidget(idx2, label1)

    def swap_images_in_labels(self, row1, row2):
        label1 = self.labels[row1]
        label2 = self.labels[row2]
        if label1 and label2:
            # Получаем текущие координаты и высоты для обеих меток
            y1 = label1.y()
            h1 = label1.height()
            y2 = label2.y()
            h2 = label2.height()

            if row2 < row1:  # Move up
                label1.move(label1.x(), y2 - h1 + h2)
                label2.move(label2.x(), y2 - h1)
            else:  # Move down
                label1.move(label1.x(), y2)
                label2.move(label2.x(), y1 - h2 + h1)

    def update_labels_after_swap(self, row1, row2):
        self.labels[row1], self.labels[row2] = self.labels[row2], self.labels[row1]

    def get_file_name(self):
        files = {
            "ВЗД": [path1 + "ВЗД.csv", path3 + "ВЗД.png"],
            "РУС": [path1 + "РУС.csv", path3 + "РУС.png"],
            "Бурильные трубы": [path1 + "Бурильные трубы.csv", path3 + "Бурильные трубы.png"],
            "Переводник": [path1 + "Переводник.csv", path3 + "Переводник.png"],
            "Предохранительный переводник": [path1 + "Предохранительный переводник.csv",
                                             path3 + "Предохранительный переводник.png"],
            "Обратный клапан": [path1 + "Обратный клапан.csv", path3 + "Обратный клапан.png"],
            "Ясс": [path1 + "Ясс.csv", path3 + "Ясс.png"],
            "Калибратор": [path1 + "Калибратор.csv", path3 + "Калибратор.png"],
            "УБТ": [path1 + "УБТ.csv", path3 + "УБТ.png"],
            "Телеметрия": [path1 + "Телеметрия.csv", path3 + "Телеметрия.png"]
        }

        if self.file_key in files:
            return files[self.file_key]
        else:
            return None

    def open_csv_table_dialog(self, row, column):
        try:
            if column == 1 and row != 0:
                file_name = self.get_file_name()
                if file_name:
                    self.current_file_path = file_name[0]

                    if not self.current_file_path:
                        raise ValueError("Путь к файлу не может быть пустым")

                    initial_sort_value_KNBK = self.get_initial_sort_value()
                    dialog = CsvTableDialog(self.current_file_path, load_table=False, sort_value_casing_srings=None,
                                            initial_sort_value_KNBK=initial_sort_value_KNBK, parent=self)
                    dialog.data_selected.connect(lambda data: self.update_table_data(data, row, column))
                    dialog.rejected.connect(lambda: self.csv_dialog_rejected(row, column))
                    dialog.exec()
                else:
                    raise ValueError("Файл не был выбран")
        except ValueError as e:
            print(f"Ошибка: {e}")
        except Exception as e:
            print(f"Произошла ошибка: {e}")

    def open_fixed_path_csv_dialog(self, row, column):
        if column == 1 and row == 0:
            fixed_file_path = path1 + "Долото.csv"
            self.current_file_path = fixed_file_path
            dialog = CsvTableDialog(fixed_file_path, load_table=False, initial_sort_value_KNBK=None,
                                    sort_value_casing_srings=self.sort_key, parent=self)
            dialog.data_selected.connect(lambda data: self.update_table_data(data, row, column))
            dialog.rejected.connect(lambda: self.csv_dialog_rejected_2(row, column))
            dialog.exec()

    def csv_dialog_rejected(self, row, column):
        item = self.tbl_KNBK.item(row, column)
        if item is not None and item.text() == "":
            for col in range(self.tbl_KNBK.columnCount()):
                if col == 5:
                    combo2 = QComboBox()
                    combo2.addItems(["Ниппель", "Муфта"])
                    combo2.setStyleSheet("QComboBox { text-align: center; }")
                    for i in range(combo2.count()):
                        combo2.setItemData(i, Qt.AlignmentFlag.AlignCenter, Qt.ItemDataRole.TextAlignmentRole)
                    self.tbl_KNBK.setCellWidget(row, col, combo2)
                elif col == 6:
                    combo3 = QComboBox()
                    combo3.addItems(
                        ["З-76", "З-86", "З-88", "З-94", "З-101", "З-102", "З-108", "З-118", "З-121", "З-122",
                         "З-133", "З-140", "З-147", "З-152", "З-161", "З-163", "З-171"])
                    combo3.setStyleSheet("QComboBox { text-align: center; }")
                    for i in range(combo3.count()):
                        combo3.setItemData(i, Qt.AlignmentFlag.AlignCenter, Qt.ItemDataRole.TextAlignmentRole)
                    self.tbl_KNBK.setCellWidget(row, col, combo3)
                elif col == 7:
                    combo4 = QComboBox()
                    combo4.addItems(["Ниппель", "Муфта"])
                    combo4.setStyleSheet("QComboBox { text-align: center; }")
                    for i in range(combo4.count()):
                        combo4.setItemData(i, Qt.AlignmentFlag.AlignCenter, Qt.ItemDataRole.TextAlignmentRole)
                    self.tbl_KNBK.setCellWidget(row, col, combo4)
                elif col == 8:
                    combo5 = QComboBox()
                    combo5.addItems(
                        ["З-76", "З-86", "З-88", "З-94", "З-101", "З-102", "З-108", "З-118", "З-121", "З-122",
                         "З-133", "З-140", "З-147", "З-152", "З-161", "З-163", "З-171"])
                    combo5.setStyleSheet("QComboBox { text-align: center; }")
                    for i in range(combo5.count()):
                        combo5.setItemData(i, Qt.AlignmentFlag.AlignCenter, Qt.ItemDataRole.TextAlignmentRole)
                    self.tbl_KNBK.setCellWidget(row, col, combo5)

    def csv_dialog_rejected_2(self, row, column):
        item = self.tbl_KNBK.item(row, column)
        if item is not None and item.text() == "":
            for col in range(self.tbl_KNBK.columnCount()):
                if col == 7:
                    combo4 = QComboBox()
                    combo4.addItems(["Ниппель", "Муфта"])
                    combo4.setStyleSheet("QComboBox { text-align: center; }")
                    for i in range(combo4.count()):
                        combo4.setItemData(i, Qt.AlignmentFlag.AlignCenter, Qt.ItemDataRole.TextAlignmentRole)
                    self.tbl_KNBK.setCellWidget(row, col, combo4)
                elif col == 8:
                    combo5 = QComboBox()
                    combo5.addItems(
                        ["З-76", "З-86", "З-88", "З-94", "З-101", "З-102", "З-108", "З-118", "З-121", "З-122",
                         "З-133", "З-140", "З-147", "З-152", "З-161", "З-163", "З-171"])
                    combo5.setStyleSheet("QComboBox { text-align: center; }")
                    for i in range(combo5.count()):
                        combo5.setItemData(i, Qt.AlignmentFlag.AlignCenter, Qt.ItemDataRole.TextAlignmentRole)
                    self.tbl_KNBK.setCellWidget(row, col, combo5)

    def contextMenuEvent(self, event):
        if self.childAt(event.pos()) == self.tbl_KNBK.viewport():
            contextMenu = QMenu(self)

            saveAstemplate_act = QAction("Сохранить строку", self)
            saveAstemplate_act.triggered.connect(self.saveAstemplate)
            contextMenu.addAction(saveAstemplate_act)

            SaveAsall_act = QAction("Сохранить КНБК", self)
            SaveAsall_act.triggered.connect(self.saveAsall)
            contextMenu.addAction(SaveAsall_act)

            copyRow_act = QAction("Копировать строку", self)
            copyRow_act.triggered.connect(self.copyRow)
            contextMenu.addAction(copyRow_act)

            undo_act = self.undo_stack.createUndoAction(self, "Отменить")
            contextMenu.addAction(undo_act)

            redo_act = self.undo_stack.createRedoAction(self, "Повторить")
            contextMenu.addAction(redo_act)

            contextMenu.exec(self.mapToGlobal(event.pos()))

    def saveAstemplate(self):
        if self.tbl_KNBK.hasFocus():
            row = self.tbl_KNBK.currentRow()
            data = []
            for column in range(1, self.tbl_KNBK.columnCount()):
                widget = self.tbl_KNBK.cellWidget(row, column)
                if widget is not None and isinstance(widget, QComboBox):
                    text = widget.currentText()
                else:
                    item = self.tbl_KNBK.item(row, column)
                    if item is not None:
                        text = item.text()
                    else:
                        text = ""
                data.append(text)
            self.write_csv_2(data)

    def saveAsall(self):
        if self.tbl_KNBK.hasFocus():
            data = []

            item_0_4 = self.tbl_KNBK.item(0, 4)
            text_0_4 = item_0_4.text() if item_0_4 is not None else ''
            data.append(f"{text_0_4},")
            item_0_0 = self.tbl_KNBK.item(0, 0)
            text_0_0 = item_0_0.text() if item_0_0 is not None else ''
            data.append(f" {text_0_0}")
            item_0_1 = self.tbl_KNBK.item(0, 1)
            text_0_1 = item_0_1.text() if item_0_1 is not None else ''
            data.append(f"_{text_0_1};")

            for row in range(1, self.tbl_KNBK.rowCount()):
                item_0 = self.tbl_KNBK.item(row, 0)
                item_1 = self.tbl_KNBK.item(row, 1)

                text_0 = item_0.text() if item_1 is not None else ''
                text_1 = item_1.text() if item_1 is not None else ''

                combined_text = f" {text_0}_{text_1};"
                data.append(combined_text)

            data_str = ''.join(data)

            self.write_csv(data_str)

    def copyRow(self):
        if self.tbl_KNBK.hasFocus():
            row = self.tbl_KNBK.currentRow()

        if row > 0:
            data = []
            for column in range(self.tbl_KNBK.columnCount()):
                item = self.tbl_KNBK.item(row, column)
                text = item.text() if item is not None else ''
                data.append(text)
            self.add_image()

            newRow = self.tbl_KNBK.rowCount()
            self.tbl_KNBK.insertRow(newRow)
            for column, text in enumerate(data):
                newItem = QTableWidgetItem(text)
                newItem.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.tbl_KNBK.setItem(newRow, column, newItem)

    def write_csv(self, data):
        fixed_file_path = path1 + "КНБК.csv"
        try:
            with open(fixed_file_path, 'a', newline='', encoding='utf-8') as f:
                f.write('\n' + data)
        except Exception as e:
            print(f"Error opening file: {e}")

    def write_csv_2(self, data):
        if self.current_file_path is None:
            print("Error: No file path set for writing data.")
            return

        try:
            with open(self.current_file_path, 'a', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(data)
            print(f"Data written to {self.current_file_path}")
        except Exception as e:
            print(f"Error opening file: {e}")

    def get_initial_sort_value(self):
        row_count = self.tbl_KNBK.rowCount()
        column_count = self.tbl_KNBK.columnCount()
        print(f"Row count: {row_count}, Column count: {column_count}")

        if row_count > 1:
            penultimate_row_index = row_count - 2
            item = self.tbl_KNBK.item(penultimate_row_index, 8)
            if item:
                value = item.text().strip()
                print(f"Value in the 8th column of the penultimate row: {value}")
                return value
            else:
                print(f"No item found in row {penultimate_row_index}, column 8")
        else:
            print("Not enough rows in the table")

        return None

    def remove_widgets_from_row(self, table, row):
        for column in range(1, table.columnCount()):
            widget = table.cellWidget(row, column)
            if widget:
                table.removeCellWidget(row, column)
                widget.deleteLater()

    def update_table_data(self, data, row, column):
        self.remove_widgets_from_row(self.tbl_KNBK, row)

        if isinstance(data, list):
            offset = 0
            for col_index, value in enumerate(data):
                item = QTableWidgetItem(value)
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.tbl_KNBK.setItem(row, column + col_index + offset, item)
        else:
            item = QTableWidgetItem(data)
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.tbl_KNBK.setItem(row, column, item)
        self.restore_initial_state()

    def update_table_data_list_2(self, data, key):
        try:
            print(f"Received data: {data} with key: {key}")

            old_data = []
            for row in range(self.tbl_KNBK.rowCount()):
                row_data = []
                for col in range(self.tbl_KNBK.columnCount()):
                    item = self.tbl_KNBK.item(row, col)
                    if not item:
                        item = QTableWidgetItem("")
                        self.tbl_KNBK.setItem(row, col, item)
                    item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    row_data.append(item.text())
                old_data.append(row_data)
            self.undo_stack.push(UpdateTableCommand(self, old_data, data))

            self.update_table_widget(data)
            self.add_label(f"КНБК - {self.tbl_KNBK.item(0, 4).text()} мм")

        except Exception as e:
            print(f"Error in update_table_data_list_2: {e}")

    def update_table_widget(self, data):
        self.clear_images()

        self.tbl_KNBK.clearContents()
        self.tbl_KNBK.setRowCount(len(data))

        self.add_image(mode="static", static_path=path3 + 'Долото.png')

        # self.restore_initial_state()
        for row_index, row_data in enumerate(data):
            for col_index, value in enumerate(row_data):
                if value is None:
                    value = ""
                item = QTableWidgetItem(value)
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.tbl_KNBK.setItem(row_index, col_index, item)
            self.file_key = self.tbl_KNBK.item(row_index, 0).text()
            self.add_image()

    def clear_images(self):
        if self.scroll_area.isVisible():
            for label in self.labels:
                self.image_container_layout.removeWidget(label)
                label.deleteLater()
        else:
            for label in self.labels:
                label.deleteLater()

        self.labels.clear()
        self.current_y = self.max_height
        self.scroll_area.hide()

    def restore_initial_state(self):
        item_0_0_KNBK = QTableWidgetItem("Долото")
        item_0_0_KNBK.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        self.tbl_KNBK.setItem(0, 0, item_0_0_KNBK)

    def update_label(self, item):
        if item.row() == 0 and item.column() == 4:
            self.add_label(f"КНБК - {item.text()} мм")

    def add_label(self, text):
        if self.label is None:
            self.label = QLabel(text, self)
            self.label.setGeometry(660, 10, 200, 50)
            font = QFont('MS Shell Dlg 2', 14)
            self.label.setFont(font)
            self.label.show()
        else:
            self.label.setText(text)

    def center_text_in_item(self, item):
        item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)

    def add_QCombobox_cell_clicked(self, row, column):
        try:
            if row>0 and column==0:
                self.delete_image_KNBK(row_count=row)
                self.add_QCombobox(row_count=row, column=column)
        except Exception as e:
            print(f"Произошла ошибка: {e}")

class MainWindow(QMainWindow):
    def __init__(self, *args, **kwargs):
        super(MainWindow, self).__init__(*args, **kwargs)
        uic.loadUi('app.ui', self)

        # Проверка загрузки файла .ui
        print("app.ui loaded successfully")

        self.undo_stack = QUndoStack(self)
        self.undo_view = QUndoView(self.undo_stack)
        self.file_paths = {}
        self.current_file_path = None
        self.setup_ui()
        self.lines_on_pressure_chart = {}
        self.lines_on_frac_chart = {}
        self.lines_on_gradient_pressure_chart = {}
        self.lines_on_gradient_frac_chart = {}

    def setup_ui(self):
        self.stackedWidget: QStackedWidget = self.findChild(QStackedWidget, 'stackedWidget')
        if self.stackedWidget is None:
            print("Error: QStackedWidget not found in the .ui file")
            return

        self.stackedWidget.setCurrentIndex(0)
        self.stackedWidget.insertWidget(4, KNBK_Table(index=4, parent=self))

        self.tbl_profile: QTableWidget = self.findChild(QTableWidget, 'tableWidget_profile')
        self.tbl_stratigraphy: QTableWidget = self.findChild(QTableWidget, 'tableWidget_stratigraphy')
        self.tbl_casing_strings: QTableWidget = self.findChild(QTableWidget, 'tableWidget_casing_strings')
        self.tbl_drilling_fluids: QTableWidget = self.findChild(QTableWidget, 'tableWidget_drilling_fluids')
        self.tbl_pressure: QTableWidget = self.findChild(QTableWidget, 'tableWidget_pressure')

        self.tbl_profile.itemChanged.connect(self.center_text_in_item)
        self.tbl_pressure.itemChanged.connect(self.center_text_in_item)
        self.tbl_casing_strings.itemChanged.connect(self.center_text_in_item)

        self.tbl_pressure.itemChanged.connect(self.on_item_changed_tbl_pressure)
        self.tbl_casing_strings.itemChanged.connect(self.on_item_changed_tbl_casing_strings)
        self.tbl_profile.itemChanged.connect(self.on_item_changed_tbl_profile)

        combo_1 = QComboBox()
        combo_1.addItems(["Направление", "Кондуктор", "Промежуточная", "Промежуточная 1", "Промежуточная 2",
                          "Потайная", "Эксплуатационная", "Хвостовик", "Райзер", "Фильтр", "Не определена"])
        combo_1.setStyleSheet("QComboBox { text-align: center; }")
        for i in range(combo_1.count()):
            combo_1.setItemData(i, Qt.AlignmentFlag.AlignCenter, Qt.ItemDataRole.TextAlignmentRole)
        self.tbl_casing_strings.setCellWidget(0, 0, combo_1)

        allowed_tables = [self.tbl_profile, self.tbl_stratigraphy, self.tbl_pressure, self.tbl_casing_strings,
                          self.tbl_drilling_fluids]

        for table in allowed_tables:
            table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
            table.customContextMenuRequested.connect(lambda pos, tbl=table: self.create_context_menu(pos, tbl))

            paste_action = QAction('Вставить', self)
            paste_action.setShortcut(QKeySequence.StandardKey.Paste)
            paste_action.triggered.connect(lambda checked, tbl=table: self.paste_from_clipboard(tbl))

            table.addAction(paste_action)

        self.btn_go_to_next_page: QPushButton = self.findChild(QPushButton, 'pushButton_next_page')
        self.btn_go_to_previous_page: QPushButton = self.findChild(QPushButton, 'pushButton_previous_page')
        self.btn_clear_table: QPushButton = self.findChild(QPushButton, 'pushButton_clear_table')
        self.btn_add_row_stratigraphy: QPushButton = self.findChild(QPushButton, 'pushButton_add_row_stratigraphy')
        self.btn_delete_row_stratigraphy: QPushButton = self.findChild(QPushButton,
                                                                       'pushButton_delete_row_stratigraphy')
        self.btn_add_row_casing_strings: QPushButton = self.findChild(QPushButton, 'pushButton_add_row_casing_strings')
        self.btn_delete_row_casing_strings: QPushButton = self.findChild(QPushButton,
                                                                         'pushButton_delete_row_casing_strings')
        self.btn_add_row_drilling_fluids: QPushButton = self.findChild(QPushButton,
                                                                       'pushButton_add_row_drilling_fluids')
        self.btn_delete_row_drilling_fluids: QPushButton = self.findChild(QPushButton,
                                                                          'pushButton_delete_row_drilling_fluids')
        self.btn_load_stratigraphy: QPushButton = self.findChild(QPushButton, 'pushButton_load_stratigraphy')
        self.btn_form_KNBK: QPushButton = self.findChild(QPushButton, 'pushButton_form_KNBK')
        self.btn_form_fluids: QPushButton = self.findChild(QPushButton, 'pushButton_form_fluids')
        self.btn_load_stratigraphic_intervals: QPushButton = self.findChild(QPushButton,
                                                                            'pushButton_load_stratigraphic_intervals')
        self.btn_add_row_pressure: QPushButton = self.findChild(QPushButton, 'pushButton_add_row_pressure')
        self.btn_delete_row_pressure: QPushButton = self.findChild(QPushButton, 'pushButton_delete_row_pressure')
        self.lineEdit1: QLineEdit = self.findChild(QLineEdit, 'lineEdit_1')
        self.lineEdit2: QLineEdit = self.findChild(QLineEdit, 'lineEdit_2')
        self.lineEdit3: QLineEdit = self.findChild(QLineEdit, 'lineEdit_3')
        self.lineEdit4: QLineEdit = self.findChild(QLineEdit, 'lineEdit_4')

        self.graphicsView_profile = self.findChild(QGraphicsView, 'graphicsView_profile')
        self.graphicsView_profile.setScene(QGraphicsScene())

        self.graphicsView_pressure = self.findChild(QGraphicsView, 'graphicsView_pressure')
        self.graphicsView_gradient_pressure = self.findChild(QGraphicsView, 'graphicsView_gradient_pressure')

        self.graphicsView_casing_strings = self.findChild(QGraphicsView, 'graphicsView_casing_strings')
        self.graphicsView_casing_strings.setScene(QGraphicsScene())

        self.open_file_act: QAction = self.findChild(QAction, 'actionOpen')

        if not self.open_file_act:
            self.open_file_act = QAction('Open', self)
            menubar = self.menuBar()
            fileMenu = menubar.addMenu('File')
            self.open_file_act.setShortcut('Ctrl+O')
            self.open_file_act.setStatusTip('Open new file')
            self.open_file_act.triggered.connect(self.open_file)
            fileMenu.addAction(self.open_file_act)
        else:
            self.open_file_act.setShortcut('Ctrl+O')
            self.open_file_act.setStatusTip('Open new file')
            self.open_file_act.triggered.connect(self.open_file)

        self.stackedWidget.currentChanged.connect(self.on_current_index_changed)
        self.btn_go_to_next_page.clicked.connect(self.go_to_next_page)
        self.btn_go_to_previous_page.clicked.connect(self.go_to_previous_page)
        self.btn_clear_table.clicked.connect(self.clear_table)
        self.btn_add_row_stratigraphy.clicked.connect(self.add_row_stratigraphy)
        self.btn_delete_row_stratigraphy.clicked.connect(self.delete_row_stratigraphy)
        self.btn_add_row_casing_strings.clicked.connect(self.add_row_casing_strings)
        self.btn_delete_row_casing_strings.clicked.connect(self.delete_row_casing_strings)
        self.btn_add_row_drilling_fluids.clicked.connect(self.add_row_drilling_fluids)
        self.btn_delete_row_drilling_fluids.clicked.connect(self.delete_row_drilling_fluids)
        self.btn_form_fluids.clicked.connect(self.form_fluids)
        self.btn_form_KNBK.clicked.connect(self.form_KNBK)
        self.btn_form_KNBK.clicked.connect(self.process_first_element)
        self.btn_load_stratigraphy.clicked.connect(self.open_db)
        self.btn_load_stratigraphic_intervals.clicked.connect(self.load_stratigraphic_intervals)
        self.btn_add_row_pressure.clicked.connect(self.add_row_pressure)
        self.btn_delete_row_pressure.clicked.connect(self.delete_row_pressure)

        self.tbl_profile.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.tbl_stratigraphy.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.tbl_casing_strings.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.tbl_drilling_fluids.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

        for col_stratigraphy in range(self.tbl_stratigraphy.columnCount()):
            item_0_col_stratigraphy = QTableWidgetItem("")
            item_0_col_stratigraphy.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.tbl_stratigraphy.setItem(0, col_stratigraphy, item_0_col_stratigraphy)

        for col_casing_strings in range(1, self.tbl_casing_strings.columnCount()):
            item_0_col_casing_strings = QTableWidgetItem("")
            item_0_col_casing_strings.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.tbl_casing_strings.setItem(0, col_casing_strings, item_0_col_casing_strings)

        for col_drilling_fluids in range(self.tbl_drilling_fluids.columnCount()):
            item_0_col_drilling_fluids = QTableWidgetItem("")
            item_0_col_drilling_fluids.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.tbl_drilling_fluids.setItem(0, col_drilling_fluids, item_0_col_drilling_fluids)

        self.btn_go_to_previous_page.setVisible(False)

        self.undo_action = self.undo_stack.createUndoAction(self, "Отменить")
        self.undo_action.setShortcut(QKeySequence.StandardKey.Undo)

        self.redo_action = self.undo_stack.createRedoAction(self, "Повторить")
        self.redo_action.setShortcut(QKeySequence.StandardKey.Redo)

        self.addAction(self.undo_action)
        self.addAction(self.redo_action)

        self.save_action = QAction("Сохранить как", self)
        self.save_action.setShortcut(QKeySequence.StandardKey.Save)
        self.addAction(self.save_action)
        self.save_action.triggered.connect(self.save_to_db)

        self.merge_columns_1(0, 1, 3)
        self.merge_columns_2(0, 4, 7)
        self.merge_columns_3(0, 8, 12)
        self.set_column_width(0, 28)
        self.set_column_width(1, 130)

        cells_data = [
            (1, 1, "Индекс стратиграфического подразделения"),
            (1, 2, "От (верт.) ,м"),
            (1, 3, "До (верт.) ,м"),
            (1, 4, "Пласт. в начале интервала"),
            (1, 5, "Пласт. в конце интервала"),
            (1, 6, "Гидроразр. в начале интервала"),
            (1, 7, "Гидроразр. в конце интервала"),
            (1, 8, "Пласт. в начале интервала"),
            (1, 9, "Пласт. в конце интервала"),
            (1, 10, "Гидроразр. в начале интервала"),
            (1, 11, "Гидроразр. в конце интервала"),
        ]
        self.insert_text_in_cells(cells_data)
        self.tbl_pressure.resizeRowsToContents()
        self.init_graphics_views()
        self.disable_editing_for_rows()
        self.graphicsView_casing_strings.setBackgroundBrush(Qt.GlobalColor.white)
        self.graphicsView_profile.setBackgroundBrush(Qt.GlobalColor.white)

    def disable_editing_for_rows(self):
        for row in [0, 1]:
            for column in range(self.tbl_pressure.columnCount()):
                item = self.tbl_pressure.item(row, column)
                if item:
                    item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)

    def set_column_width(self, column, width):
        self.tbl_pressure.setColumnWidth(column, width)

    def merge_columns_1(self, row, start_col, end_col):
        text = "Интервал"
        for col in range(start_col, end_col + 1):
            item = self.tbl_pressure.takeItem(row, col)
            if item:
                text += item.text() + " "

        merged_item = QTableWidgetItem(text.strip())
        merged_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        self.tbl_pressure.setItem(row, start_col, merged_item)

        self.tbl_pressure.setSpan(row, start_col, 1, end_col - start_col + 1)

        for col in range(start_col + 1, end_col + 1):
            self.tbl_pressure.setItem(row, col, QTableWidgetItem(""))

    def merge_columns_2(self, row, start_col, end_col):
        text = "Давление, кгс/см^2"
        for col in range(start_col, end_col + 1):
            item = self.tbl_pressure.takeItem(row, col)
            if item:
                text += item.text() + " "

        merged_item = QTableWidgetItem(text.strip())
        merged_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        self.tbl_pressure.setItem(row, start_col, merged_item)

        self.tbl_pressure.setSpan(row, start_col, 1, end_col - start_col + 1)

        for col in range(start_col + 1, end_col + 1):
            self.tbl_pressure.setItem(row, col, QTableWidgetItem(""))

    def merge_columns_3(self, row, start_col, end_col):
        text = "Градиент давления, кгс/см^2/м"
        for col in range(start_col, end_col + 1):
            item = self.tbl_pressure.takeItem(row, col)
            if item:
                text += item.text() + " "

        merged_item = QTableWidgetItem(text.strip())
        merged_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        self.tbl_pressure.setItem(row, start_col, merged_item)

        self.tbl_pressure.setSpan(row, start_col, 1, end_col - start_col + 1)

        for col in range(start_col + 1, end_col + 1):
            self.tbl_pressure.setItem(row, col, QTableWidgetItem(""))

    def insert_text_in_cells(self, cells_data):
        for row, col, text in cells_data:
            item = QTableWidgetItem(text)
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.tbl_pressure.setItem(row, col, item)

    def create_context_menu(self, pos, table):
        context_menu = QMenu(self)

        context_menu.addAction(self.undo_action)
        context_menu.addAction(self.redo_action)

        paste_action = QAction('Вставить', self)
        paste_action.setShortcut(QKeySequence.StandardKey.Paste)
        paste_action.triggered.connect(lambda checked, tbl=table: self.paste_from_clipboard(tbl))
        context_menu.addAction(self.save_action)
        context_menu.addAction(paste_action)

        context_menu.exec(table.viewport().mapToGlobal(pos))

    def get_page_count(self):
        return self.stackedWidget.count()

    def on_current_index_changed(self, index):
        total_pages = self.get_page_count()

        if index == 0:
            self.btn_go_to_previous_page.setVisible(False)
        else:
            self.btn_go_to_previous_page.setVisible(True)

        if index == total_pages - 1:
            self.btn_go_to_next_page.setVisible(False)
        else:
            self.btn_go_to_next_page.setVisible(True)

    def go_to_next_page(self):
        current_index = self.stackedWidget.currentIndex()
        self.stackedWidget.setCurrentIndex(current_index + 1)

    def go_to_previous_page(self):
        current_index = self.stackedWidget.currentIndex()
        if current_index > 0:
            self.stackedWidget.setCurrentIndex(current_index - 1)

    def center_text_in_item(self, item):
        item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)

    def paste_from_clipboard(self, tableWidget):
        allowed_tables = [self.tbl_profile, self.tbl_stratigraphy, self.tbl_casing_strings, self.tbl_drilling_fluids]

        if tableWidget not in allowed_tables:
            return

        try:
            clipboard = QApplication.clipboard()
            mime_data = clipboard.mimeData()

            if mime_data.hasHtml():
                text_data = mime_data.html()
                text_data = self.convert_html_to_plain_text(text_data)
            elif mime_data.hasText():
                text_data = mime_data.text()
            else:
                print("Clipboard does not contain text data.")
                return

            current_row = tableWidget.currentRow()
            current_column = tableWidget.currentColumn()

            command = PasteCommand(tableWidget, text_data, current_row, current_column, "вставку")
            self.undo_stack.push(command)
        except Exception as e:
            print(f"Error pasting from clipboard: {e}")

    def convert_html_to_plain_text(self, html):
        from bs4 import BeautifulSoup
        data = []
        soup = BeautifulSoup(html, "lxml")
        table = soup.find('table')
        rows = table.find_all('tr')
        for row in rows:
            cols = row.find_all('td') + row.find_all('th')
            cols = [ele.text.strip() for ele in cols]
            data.append([ele for ele in cols if ele])
        return data

    def clear_table(self):
        self.tbl_profile.clearContents()

        headers = ["Глубина по стволу (м)", "Зенитный угол (град)", "Азимут (град)", "Азимут маг(град)",
                   "Азимут дир(град)", "Глубина по верт(м)"]

        current_column_count = self.tbl_profile.columnCount()
        if current_column_count < len(headers):
            self.tbl_profile.setColumnCount(len(headers))

        self.tbl_profile.setHorizontalHeader(QHeaderView(Qt.Orientation.Horizontal))

        self.tbl_profile.setHorizontalHeaderLabels(headers)
        self.tbl_profile.horizontalHeader().setVisible(True)
        self.tbl_profile.horizontalHeader().repaint()
        self.tbl_profile.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.graphicsView_profile.setScene(QGraphicsScene())

    def add_row_stratigraphy(self):
        row_count2_2 = self.tbl_stratigraphy.rowCount()
        self.tbl_stratigraphy.setRowCount(row_count2_2 + 1)
        for column in range(self.tbl_stratigraphy.columnCount()):
            item = QTableWidgetItem("")
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.tbl_stratigraphy.setItem(row_count2_2, column, item)

    def add_row_casing_strings(self):
        row_count2_3 = self.tbl_casing_strings.rowCount()
        self.tbl_casing_strings.setRowCount(row_count2_3 + 1)
        for column in range(self.tbl_casing_strings.columnCount()):
            if column == 0:
                combo = QComboBox()
                combo.addItems(["Направление", "Кондуктор", "Промежуточная", "Промежуточная 1", "Промежуточная 2",
                                "Потайная", "Эксплуатационная", "Хвостовик", "Райзер", "Фильтр", "Не определена"])
                combo.setStyleSheet("QComboBox { text-align: center; }")
                for i in range(combo.count()):
                    combo.setItemData(i, Qt.AlignmentFlag.AlignCenter, Qt.ItemDataRole.TextAlignmentRole)
                self.tbl_casing_strings.setCellWidget(row_count2_3, column, combo)
            else:
                item = QTableWidgetItem("")
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.tbl_casing_strings.setItem(row_count2_3, column, item)

    def add_row_drilling_fluids(self):
        row_count2_4 = self.tbl_drilling_fluids.rowCount()
        self.tbl_drilling_fluids.setRowCount(row_count2_4 + 1)
        for column in range(self.tbl_drilling_fluids.columnCount()):
            item = QTableWidgetItem("")
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.tbl_drilling_fluids.setItem(row_count2_4, column, item)

    def add_row_pressure(self):
        row_count2_5 = self.tbl_pressure.rowCount()
        self.tbl_pressure.setRowCount(row_count2_5 + 1)
        for column in range(self.tbl_pressure.columnCount()):
            item = QTableWidgetItem("")
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.tbl_pressure.setItem(row_count2_5, column, item)
        k = QTableWidgetItem(str(row_count2_5 - 1))
        k.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        self.tbl_pressure.setItem(row_count2_5, 0, k)

    def open_file(self):
        try:
            file_path, _ = QFileDialog.getOpenFileName(self, 'Open file', 'D:/Программа бурения/',
                                                       'Excel Files (*.xlsx)')
            if not file_path:
                return

            data = pd.read_excel(file_path)
            self.process_excel_data(data)

        except Exception as e:
            logging.error(f"Error opening file: {e}")

    def on_item_changed_tbl_profile(self, item):
        row = item.row()
        column = item.column()
        if column in [0, 1, 2] and self.is_row_complete(row, [0, 1, 2], self.tbl_profile):
            if not hasattr(self, 'selected_data'):
                self.selected_data = []

            # Извлечение данных из таблицы
            data = []
            for r in range(self.tbl_profile.rowCount()):
                if self.is_row_complete(r, [0, 1, 2], self.tbl_profile):
                    L = self.extract_number(self.tbl_profile.item(r, 0).text())
                    zenith_angle = np.radians(self.extract_number(self.tbl_profile.item(r, 1).text()))
                    azimuth_angle = np.radians(self.extract_number(self.tbl_profile.item(r, 2).text()))
                    data.append([L, zenith_angle, azimuth_angle])

            data = np.array(data)

            # Начальные углы и координаты
            current_zenith_angle = data[0, 1]
            current_azimuth_angle = data[0, 2]
            current_coordinates = np.array([0, 0, 0], dtype=np.float64)

            self.selected_data = [current_coordinates.copy()]

            # Выполняем расчет для каждой строки, начиная с первой
            for i in range(1, len(data)):
                delta_L = data[i, 0] - data[i - 1, 0]
                delta_zenith_angle = data[i, 1] - data[i - 1, 1]
                delta_azimuth_angle = data[i, 2] - data[i - 1, 2]

                next_zenith_angle = current_zenith_angle + delta_zenith_angle
                next_azimuth_angle = current_azimuth_angle + delta_azimuth_angle

                delta_x = delta_L * np.sin(next_zenith_angle) * np.cos(next_azimuth_angle)
                delta_y = delta_L * np.sin(next_zenith_angle) * np.sin(next_azimuth_angle)
                delta_z = delta_L * np.cos(next_zenith_angle)

                current_coordinates += np.array([delta_x, delta_y, delta_z])
                self.selected_data.append(current_coordinates.copy())

                current_zenith_angle = next_zenith_angle
                current_azimuth_angle = next_azimuth_angle

            # Convert the list of coordinates to an array and update the graph
            selected_data = np.array(self.selected_data)
            self.plot_graph(selected_data)

            if self.is_row_complete(row, [0, 1, 2], self.tbl_profile):
                delta_z = current_coordinates[2]
                self.tbl_profile.setItem(row, 5, QTableWidgetItem(str(round(delta_z, 2))))
            else:
                self.tbl_profile.setItem(row, 5, QTableWidgetItem(""))

    def process_excel_data(self, data):
        try:
            if self.has_headers(data):
                data.columns = data.iloc[0]
                data = data[1:].reset_index(drop=True)

            self.validate_data(data)
            vertical_depths = self.calculate_vertical_depth(data)
            data['Глубина по верт(м)'] = vertical_depths

            self.populate_table(data)
            self.process_and_plot_data(data)
        except Exception as e:
            logging.error(f"Error processing Excel data: {e}")

    def has_headers(self, df):
        first_row = df.iloc[0]
        return all(isinstance(val, str) for val in first_row) and len(set(first_row)) == len(first_row)

    def validate_data(self, df):
        if df.shape[1] < 3:
            raise ValueError("The file must have at least 3 columns for calculations.")

    def calculate_vertical_depth(self, df):
        vertical_depth = [0.0]
        for i in range(1, len(df)):
            delta_L = df.iloc[i, 0] - df.iloc[i - 1, 0]
            current_zenith_angle = np.radians(df.iloc[i, 1])
            delta_z = delta_L * np.cos(current_zenith_angle)
            vertical_depth.append(vertical_depth[-1] + delta_z)
        return vertical_depth

    def populate_table(self, df):
        self.tbl_profile.blockSignals(True)
        num_rows, num_cols = df.shape
        self.tbl_profile.setRowCount(num_rows)
        self.tbl_profile.setColumnCount(num_cols)

        headers = ["Глубина по стволу (м)", "Зенитный угол (град)", "Азимут (град)", "Азимут маг(град)",
                   "Азимут дир(град)", "Глубина по верт(м)"]

        if len(df.columns) == 4:
            text = ""
            header_item = QTableWidgetItem(text)
            self.tbl_profile.setHorizontalHeaderItem(2, header_item)
            self.tbl_profile.setHorizontalHeader(ComboHeader(self))
            text = "Глубина по верт(м)"
            header_item = QTableWidgetItem(text)
            self.tbl_profile.setHorizontalHeaderItem(3, header_item)
            self.tbl_profile.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

        if self.has_headers(df):
            headers = df.columns.astype(str).tolist()
            logging.info(f"Headers from file: {headers}")
            self.tbl_profile.setHorizontalHeaderLabels(headers)

        for index, row in df.iterrows():
            for col_index, value in enumerate(row):
                item = QTableWidgetItem(str(round(value,2)))
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.tbl_profile.setItem(index, col_index, item)
        self.tbl_profile.blockSignals(False)
        self.tbl_profile.horizontalHeader().setVisible(True)

    def process_and_plot_data(self, data):
        current_coordinates = np.array([0.0, 0.0, 0.0], dtype=np.float64)
        current_zenith_angle = np.radians(0)
        current_azimuth_angle = np.radians(0)
        selected_data = [current_coordinates.copy()]

        for i in range(1, len(data)):
            delta_L = data.iloc[i, 0] - data.iloc[i - 1, 0]
            delta_zenith_angle = np.radians(data.iloc[i, 1] - data.iloc[i - 1, 1])
            delta_azimuth_angle = np.radians(data.iloc[i, 2] - data.iloc[i - 1, 2])

            next_zenith_angle = current_zenith_angle + delta_zenith_angle
            next_azimuth_angle = current_azimuth_angle + delta_azimuth_angle

            delta_x = delta_L * np.sin(next_zenith_angle) * np.cos(next_azimuth_angle)
            delta_y = delta_L * np.sin(next_zenith_angle) * np.sin(next_azimuth_angle)
            delta_z = delta_L * np.cos(next_zenith_angle)

            current_coordinates += np.array([delta_x, delta_y, delta_z])
            selected_data.append(current_coordinates.copy())

            current_zenith_angle = next_zenith_angle
            current_azimuth_angle = next_azimuth_angle

        selected_data = np.array(selected_data)
        self.plot_graph(selected_data)

    def plot_graph(self, data):
        fig = Figure()
        ax = fig.add_subplot(111, projection='3d')
        ax.plot(data[:, 0], data[:, 1], data[:, 2], marker='o', linewidth=0.25)
        ax.set_xlim([min(data[:, 0]), max(data[:, 0])])
        ax.set_ylim([min(data[:, 1]), max(data[:, 1])])
        ax.set_zlim([max(data[:, 2]), min(data[:, 2])])
        ax.view_init(elev=20, azim=35)
        ax.set_box_aspect([1, 1, 1])

        scene = QGraphicsScene()
        canvas = FigureCanvas(fig)
        canvas.setGeometry(0, 0, 475, 615)
        scene.addWidget(canvas)

        self.graphicsView_profile.setScene(scene)

    def init_graphics_views(self):
        self.canvas_pressure = MplCanvas(self.graphicsView_pressure, width=10.40, height=1.8, dpi=100)
        self.canvas_gradient = MplCanvas(self.graphicsView_gradient_pressure, width=3.6, height=1.8, dpi=100)

        self.canvas_pressure.axes.invert_yaxis()
        self.canvas_pressure.axes.set_title(" Давления", fontsize=10)
        self.canvas_pressure.axes.set_xlabel("Давления, кгс/см2", fontsize=8)
        self.canvas_pressure.axes.set_ylabel("Глубина по вертикали, м", fontsize=8)
        self.canvas_pressure.axes.tick_params(axis='x', labelsize=8)
        self.canvas_pressure.axes.tick_params(axis='y', labelsize=8)
        self.canvas_pressure.fig.tight_layout()
        self.canvas_pressure.axes.grid(True)

        self.canvas_gradient.axes.invert_yaxis()
        self.canvas_gradient.axes.set_title("Градиент давления", fontsize=10)
        self.canvas_gradient.axes.set_xlabel("Градиент давления, кгс/см2/м", fontsize=8)
        self.canvas_gradient.axes.set_ylabel("Глубина по вертикали, м", fontsize=8)
        self.canvas_gradient.axes.tick_params(axis='x', labelsize=8)
        self.canvas_gradient.axes.tick_params(axis='y', labelsize=8)
        self.canvas_gradient.fig.tight_layout()
        self.canvas_gradient.axes.grid(True)

        self.add_canvas_to_view(self.canvas_pressure, self.graphicsView_pressure)
        self.add_canvas_to_view(self.canvas_gradient, self.graphicsView_gradient_pressure)

    def add_canvas_to_view(self, canvas, graphics_view):
        scene = QGraphicsScene()
        scene.addWidget(canvas)
        graphics_view.setScene(scene)
        graphics_view.setRenderHint(QPainter.RenderHint.Antialiasing)
        graphics_view.fitInView(scene.itemsBoundingRect(), Qt.AspectRatioMode.KeepAspectRatio)

    def on_item_changed_tbl_pressure(self, item):
        row = item.row()
        column = item.column()
        if row > 1:
            if column == 8:
                gradient_pressure = self.tbl_pressure.item(row, 9)
                if gradient_pressure is not None:
                    self.tbl_pressure.blockSignals(True)
                    gradient_pressure.setText(item.text())
                    self.tbl_pressure.blockSignals(False)

            if column == 9:
                gradient_pressure = self.tbl_pressure.item(row, 8)
                if gradient_pressure is not None:
                    self.tbl_pressure.blockSignals(True)
                    gradient_pressure.setText(item.text())
                    self.tbl_pressure.blockSignals(False)

            if column == 10:
                gradient_frac = self.tbl_pressure.item(row, 11)
                if gradient_frac is not None:
                    self.tbl_pressure.blockSignals(True)
                    gradient_frac.setText(item.text())
                    self.tbl_pressure.blockSignals(False)

            if column == 11:
                gradient_frac = self.tbl_pressure.item(row, 10)
                if gradient_frac is not None:
                    self.tbl_pressure.blockSignals(True)
                    gradient_frac.setText(item.text())
                    self.tbl_pressure.blockSignals(False)

            if column in [2, 3, 4, 5] and self.is_row_complete(row, [2, 3, 4, 5], self.tbl_pressure):
                x_1 = self.extract_number(self.tbl_pressure.item(row, 2).text())
                x_2 = self.extract_number(self.tbl_pressure.item(row, 3).text())
                pressure_end = self.extract_number(self.tbl_pressure.item(row, 5).text())
                pressure_start = self.extract_number(self.tbl_pressure.item(row, 4).text())
                gradient = round((pressure_end - pressure_start) / (x_2 - x_1), 2)
                self.tbl_pressure.blockSignals(True)
                item_gradient_1 = QTableWidgetItem(str(gradient))
                item_gradient_1.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                item_gradient_2 = QTableWidgetItem(str(gradient))
                item_gradient_2.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.tbl_pressure.setItem(row, 8, item_gradient_1)
                self.tbl_pressure.setItem(row, 9, item_gradient_2)
                self.tbl_pressure.blockSignals(False)

            if column in [2, 3, 6, 7] and self.is_row_complete(row, [2, 3, 6, 7], self.tbl_pressure):
                x_1 = self.extract_number(self.tbl_pressure.item(row, 2).text())
                x_2 = self.extract_number(self.tbl_pressure.item(row, 3).text())
                pressure_end = self.extract_number(self.tbl_pressure.item(row, 7).text())
                pressure_start = self.extract_number(self.tbl_pressure.item(row, 6).text())
                gradient = round((pressure_end - pressure_start) / (x_2 - x_1), 2)
                self.tbl_pressure.blockSignals(True)
                item_gradient_1 = QTableWidgetItem(str(gradient))
                item_gradient_1.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                item_gradient_2 = QTableWidgetItem(str(gradient))
                item_gradient_2.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.tbl_pressure.setItem(row, 10, item_gradient_1)
                self.tbl_pressure.setItem(row, 11, item_gradient_2)
                self.tbl_pressure.blockSignals(False)

            if column in [2, 3, 8, 9] and self.is_row_complete(row, [2, 3, 8, 9], self.tbl_pressure):
                x_1 = self.extract_number(self.tbl_pressure.item(row, 2).text())
                x_2 = self.extract_number(self.tbl_pressure.item(row, 3).text())
                gradient = self.extract_number(self.tbl_pressure.item(row, 8).text())
                pressure_start = self.extract_number(self.tbl_pressure.item(row, 4).text())
                pressure_end = round(gradient * (x_2 - x_1) + pressure_start, 2)
                self.tbl_pressure.blockSignals(True)
                item_pressure_end = QTableWidgetItem(str(pressure_end))
                item_pressure_end.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.tbl_pressure.setItem(row, 5, item_pressure_end)
                self.tbl_pressure.blockSignals(False)

            if column in [2, 3, 10, 11] and self.is_row_complete(row, [2, 3, 10, 11], self.tbl_pressure):
                x_1 = self.extract_number(self.tbl_pressure.item(row, 2).text())
                x_2 = self.extract_number(self.tbl_pressure.item(row, 3).text())
                gradient = self.extract_number(self.tbl_pressure.item(row, 10).text())
                pressure_start = self.extract_number(self.tbl_pressure.item(row, 6).text())
                pressure_end = round(gradient * (x_2 - x_1) + pressure_start, 2)
                self.tbl_pressure.blockSignals(True)
                item_pressure_end = QTableWidgetItem(str(pressure_end))
                item_pressure_end.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.tbl_pressure.setItem(row, 7, item_pressure_end)
                self.tbl_pressure.blockSignals(False)

            if self.is_row_complete(row, [2, 3, 10, 11], self.tbl_pressure):
                self.update_graph(row, [10, 11], self.canvas_gradient, self.lines_on_gradient_frac_chart,
                                  line_id=2)

            if self.is_row_complete(row, [2, 3, 6, 7], self.tbl_pressure):
                self.update_graph(row, [6, 7], self.canvas_pressure, self.lines_on_frac_chart, line_id=2)

            if self.is_row_complete(row, [2, 3, 8, 9], self.tbl_pressure):
                self.update_graph(row, [8, 9], self.canvas_gradient, self.lines_on_gradient_pressure_chart,
                                  line_id=1)

            if self.is_row_complete(row, [2, 3, 4, 5], self.tbl_pressure):
                self.update_graph(row, [4, 5], self.canvas_pressure, self.lines_on_pressure_chart, line_id=1)

    def is_row_complete(self, row, required_columns, table):
        for col in required_columns:
            table_item = table.item(row, col)
            if not table_item or not table_item.text().strip():
                return False
        return True

    def update_graph(self, row, columns, canvas, chart, line_id):
        try:
            start_depth = float(self.tbl_pressure.item(row, 2).text())
            end_depth = float(self.tbl_pressure.item(row, 3).text())
            start = float(self.tbl_pressure.item(row, columns[0]).text())
            end = float(self.tbl_pressure.item(row, columns[1]).text())
            if row in chart:
                line = chart.pop(row)
                line.remove()

            if line_id == 1:
                line, = canvas.axes.plot(
                    [start, end], [start_depth, end_depth], color='green', linewidth=2)

            else:
                line, = canvas.axes.plot(
                    [start, end], [start_depth, end_depth], color='red', linewidth=2)

            chart[row] = line
            canvas.draw()

        except ValueError as ve:
            print(f"Ошибка преобразования данных в строке {row}: {ve}")
        except KeyError as ke:
            print(f"Ошибка доступа к данным в строке {row}: {ke}")
        except Exception as e:
            print(f"Неизвестная ошибка при обновлении графиков в строке {row}: {e}")

    def delete_row_stratigraphy(self):
        row_count2_2 = self.tbl_stratigraphy.rowCount()
        if row_count2_2 > 0:
            self.tbl_stratigraphy.setRowCount(row_count2_2 - 1)

    def delete_row_casing_strings(self):
        row_count2_3 = self.tbl_casing_strings.rowCount()
        self.tbl_casing_strings.setRowCount(row_count2_3 - 1)
        self.draw_wellbore_diagram()

    def delete_row_drilling_fluids(self):
        row_count2_4 = self.tbl_drilling_fluids.rowCount()
        if row_count2_4 > 0:
            self.tbl_drilling_fluids.setRowCount(row_count2_4 - 1)

    def delete_row_pressure(self):
        row_count2_5 = self.tbl_pressure.rowCount()
        if row_count2_5 > 2:
            self.tbl_pressure.setRowCount(row_count2_5 - 1)

    def on_item_changed_tbl_casing_strings(self, item):
        row = item.row()
        column = item.column()
        if column in [1, 2]:
            self.update_drilling_fluids(row, column)
        if self.is_row_complete(row, [1, 2, 3, 4], self.tbl_casing_strings):
            self.draw_wellbore_diagram()

    def update_drilling_fluids(self, row, column):
        item = self.tbl_casing_strings.item(row, column)
        if item is not None:
            text = item.text().strip().lower()
            if text != "":
                if column == 1:
                    if "до забоя" in text:
                        profile_item = self.tbl_profile.item(self.tbl_profile.rowCount() - 1, 0)
                        new_text = f"до забоя ({profile_item.text()})" if profile_item else "0"
                        new_item = QTableWidgetItem(new_text)
                    elif text.replace('.', '', 1).isdigit():
                        new_item = QTableWidgetItem(item.text())
                    else:
                        new_item = QTableWidgetItem('0')

                    new_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    self.tbl_casing_strings.blockSignals(True)
                    self.tbl_casing_strings.setItem(row, column, new_item)
                    self.tbl_casing_strings.blockSignals(False)

                elif column == 2:
                    if "до устья" in text:
                        profile_item = self.tbl_profile.item(0, 0)
                        if profile_item:
                            new_text = str(self.extract_number(self.tbl_casing_strings.item(row, 1).text()) - float(
                                profile_item.text()))
                        elif self.tbl_casing_strings.item(row, 1).text():
                            new_text = str(self.extract_number(self.tbl_casing_strings.item(row, 1).text()) - 0)
                        else:
                            new_text = "0"
                        new_item = QTableWidgetItem(new_text)
                    elif text.replace('.', '', 1).isdigit():
                        new_item = QTableWidgetItem(item.text())
                    else:
                        new_item = QTableWidgetItem('0')

                    new_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    self.tbl_casing_strings.blockSignals(True)
                    self.tbl_casing_strings.setItem(row, column, new_item)
                    self.tbl_casing_strings.blockSignals(False)

        print(f"Updated row: {row}, column: {column}, text: {text}")

    def extract_number(self, text):
        if not text or text.strip() == "":
            return 0
        text = text.replace(",", ".")
        match = re.search(r'\d+\.\d+', text)
        if match:
            return float(match.group(0))
        try:
            return float(text)
        except ValueError as ve:
            print(f"Ошибка преобразования данных: {ve}")
            return 0

    def draw_wellbore_diagram(self):
        scene = self.graphicsView_casing_strings.scene()
        scene.clear()

        # Константы для размеров и смещений
        view_width = 360
        view_height = 570
        max_hole_width = 940
        scale_factor_x = view_width / max_hole_width
        horizontal_offset = 30
        vertical_padding = 5

        total_height = 2
        row_count = self.tbl_casing_strings.rowCount()

        pen_casing = QPen(Qt.GlobalColor.black)
        brush_casing = QBrush(Qt.GlobalColor.lightGray)
        pen_hole = QPen(Qt.GlobalColor.black, 2)

        x_offset = view_width / 2
        last_end = 0
        last_diameter_hole = 0
        lengths = []
        diameter_offsets = []
        ends = []
        first_diameter_hole = 0
        for row in range(row_count):
            end = self.extract_number(
                self.tbl_casing_strings.item(row, 1).text())
            if row > 0 and end > last_end:
                total_height = self.extract_number(
                    self.tbl_casing_strings.item(row_count - 1, 1).text()) + 2 * vertical_padding
            else:
                total_height = last_end
            last_end = end

        for row in range(row_count):
            scale_factor_y = view_height / total_height if total_height > 0 else 1
            end = self.extract_number(
                self.tbl_casing_strings.item(row, 1).text()) * scale_factor_y
            length = self.extract_number(
                self.tbl_casing_strings.item(row, 2).text()) * scale_factor_y  # Масштабируем длину
            diameter_casing = self.extract_number(self.tbl_casing_strings.item(row, 3).text()) * scale_factor_x
            diameter_hole = self.extract_number(self.tbl_casing_strings.item(row, 4).text()) * scale_factor_x

            lengths.append(length)
            diameter_offsets.append((last_diameter_hole - diameter_hole) / 2)
            ends.append(end)
            if row == 0:
                first_diameter_hole = diameter_hole

            scene.addRect(QRectF(x_offset - diameter_casing / 2, end - length, diameter_casing, length), pen_casing,
                          brush_casing)

            last_diameter_hole = diameter_hole

        scene.setSceneRect(0, 0, view_width, view_height)
        if self.tbl_casing_strings.rowCount() != 0:
            ShadingDrawer(lengths, ends, diameter_offsets, first_diameter_hole, x_offset, scene).draw_curve()
            ShadingDrawer(lengths, ends, diameter_offsets, first_diameter_hole, x_offset, scene,
                          reverse=True).draw_curve()
        self.graphicsView_casing_strings.setScene(scene)
        self.graphicsView_casing_strings.fitInView(scene.sceneRect(), Qt.AspectRatioMode.KeepAspectRatio)
        self.graphicsView_casing_strings.update()

    def form_fluids(self):
        row_count_casing_strings = self.tbl_casing_strings.rowCount()
        row_count_drilling_fluids = self.tbl_drilling_fluids.rowCount()

        if row_count_casing_strings != row_count_drilling_fluids:
            self.tbl_drilling_fluids.setRowCount(row_count_casing_strings)

        for row in range(row_count_casing_strings):
            for column in range(1, 3):
                try:
                    casing_item = self.tbl_casing_strings.item(row, column)
                    if casing_item:
                        text = casing_item.text().strip().lower()
                        if column == 1 and "до забоя" in text:
                            start_index = text.find('(') + 1
                            end_index = text.find(')')
                            profile_value = text[
                                            start_index:end_index] if start_index > 0 and end_index > start_index else ""
                            drilling_fluids_item = QTableWidgetItem(profile_value)
                        elif column == 1:
                            drilling_fluids_item = QTableWidgetItem(text)
                        else:
                            first_value = self.extract_number(self.tbl_casing_strings.item(row, 1).text())
                            second_value = self.extract_number(self.tbl_casing_strings.item(row, 2).text())

                            if first_value is not None and second_value is not None:
                                difference = first_value - second_value
                                drilling_fluids_item = QTableWidgetItem(str(difference))
                            else:
                                drilling_fluids_item = QTableWidgetItem("Ошибка")

                        drilling_fluids_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                        self.tbl_drilling_fluids.setItem(row, column, drilling_fluids_item)
                except Exception as e:
                    print(f"Error updating drilling fluids at row {row}, column {column}: {e}")

    def load_stratigraphic_intervals(self):
        row_count = self.tbl_stratigraphy.rowCount()
        self.tbl_pressure.setRowCount(row_count + 2)

        for row in range(row_count):
            num_item = QTableWidgetItem(str(row + 1))
            num_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.tbl_pressure.setItem(row + 2, 0, num_item)

            for column in range(1, 4):
                item = self.tbl_stratigraphy.item(row, column)
                if item:
                    new_item = QTableWidgetItem(item.text())
                    new_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    self.tbl_pressure.setItem(row + 2, column, new_item)

    def form_KNBK(self):
        if self.tbl_casing_strings is None or self.stackedWidget is None:
            print("Error: Table or StackedWidget is not initialized")
            return

        number_of_pages = self.tbl_casing_strings.rowCount()
        current_index = self.stackedWidget.currentIndex()

        for i in range(1, number_of_pages):
            item = self.tbl_casing_strings.item(i, 4)
            sort_key = item.text()
            if item:
                value = item.text()
            else:
                value = "unknown"
            text = f"КНБК - {value} мм"
            self.add_page(text, sort_key)
        next_index = min(current_index + 1, self.stackedWidget.count() - 1)
        self.stackedWidget.setCurrentIndex(next_index)

    def process_first_element(self):
        if self.tbl_casing_strings is None or self.stackedWidget is None:
            print("Error: Table or StackedWidget is not initialized")
            return

        item_0 = self.tbl_casing_strings.item(0, 4)
        if item_0:
            text_0 = f"КНБК - {item_0.text()} мм"
            sort_key_0 = item_0.text()
            initial_page = self.stackedWidget.widget(4)
            initial_page.add_label(text_0)
            initial_page.sort_key = sort_key_0

    def add_page(self, text, sort_key=None):
        if self.stackedWidget is None:
            return

        num_pages = self.stackedWidget.count()
        insert_index = num_pages - 1

        new_page = KNBK_Table(index=insert_index + 1, sort_key=sort_key, parent=self)
        new_page.add_label(text)
        new_page.sort_key = sort_key
        self.stackedWidget.insertWidget(insert_index, new_page)
        self.stackedWidget.setCurrentWidget(new_page)

    def add_page_2(self):
        if self.stackedWidget is None:
            return
        num_pages = self.stackedWidget.count()
        insert_index = num_pages - 1
        print(f"Current number of pages: {num_pages}, Inserting at index: {insert_index}")
        new_page = KNBK_Table(index=insert_index + 1, parent=self)
        self.stackedWidget.insertWidget(insert_index, new_page)
        self.stackedWidget.setCurrentWidget(new_page)

    def delete_page(self):
        if self.stackedWidget is None:
            return

        current_index = self.stackedWidget.currentIndex()
        num_pages = self.stackedWidget.count()

        if num_pages > 6:
            widget_to_remove = self.stackedWidget.widget(current_index)
            self.stackedWidget.removeWidget(widget_to_remove)

            widget_to_remove.setParent(None)
            widget_to_remove.deleteLater()

            new_index = min(current_index, self.stackedWidget.count() - 1)
            self.stackedWidget.setCurrentIndex(new_index)

            self.on_current_index_changed(new_index)

            print(
                f"Page at index {current_index} deleted. Current number of pages: {self.stackedWidget.count()}")
        else:
            print(
                "Cannot delete the only remaining page.")

    def on_current_index_changed(self, index):
        total_pages = self.stackedWidget.count()

        self.btn_go_to_previous_page.setVisible(index > 0)
        self.btn_go_to_next_page.setVisible(index < total_pages - 1)

    def save_to_db(self):
        stratigraphy_name, ok = QInputDialog.getText(self, "Имя", "Введите имя:")

        if ok and stratigraphy_name.strip():  # Проверка на пустую строку
            database = QSqlDatabase.addDatabase("QSQLITE")  # SQLite version 3
            database.setDatabaseName(path2 + "stratigraphy.db")
            if not database.open():
                QMessageBox.critical(self, "Ошибка", "Не удалось открыть файл базы данных.")
                return  # Выход из функции без завершения приложения

            query = QSqlQuery()
            query.exec(f"DROP TABLE {stratigraphy_name}")
            create_table_query = f"""CREATE TABLE IF NOT EXISTS {stratigraphy_name} (
                                    id INTEGER PRIMARY KEY AUTOINCREMENT UNIQUE NOT NULL,
                                    name VARCHAR(64) NOT NULL,
                                    indexation VARCHAR(64) NOT NULL,
                                    up FLOAT NOT NULL,
                                    down FLOAT NOT NULL,
                                    K_k FLOAT NOT NULL, 
                                    density VARCHAR(64) NOT NULL)"""

            if not query.exec(create_table_query):
                QMessageBox.critical(self, "Ошибка", f"Не удалось создать таблицу: {query.lastError().text()}")
                database.close()
                return

            insert_query = f"""INSERT INTO {stratigraphy_name} (
                              name, indexation, 
                              up, down, K_k, density) 
                              VALUES (?, ?, ?, ?, ?, ?)"""

            if not query.prepare(insert_query):
                QMessageBox.critical(self, "Ошибка", f"Не удалось подготовить запрос: {query.lastError().text()}")
                database.close()
                return

            for row in range(self.tbl_stratigraphy.rowCount()):
                for col in range(self.tbl_stratigraphy.columnCount()):
                    item = self.tbl_stratigraphy.item(row, col)
                    if item is None:
                        query.addBindValue('')
                    else:
                        query.addBindValue(item.text())

                if not query.exec():
                    QMessageBox.critical(self, "Ошибка", f"Не удалось выполнить запрос: {query.lastError().text()}")
                    database.close()
                    return

            database.close()

    def open_db(self):
        db_file = path2 + "stratigraphy.db"
        self.db_manager = DatabaseManager(db_file)
        if not self.db_manager.open_database():
            return

        while True:

            tables = self.db_manager.get_tables()
            if not tables:
                QMessageBox.warning(self, "Предупреждение", "В базе данных нет таблиц.")
                self.db_manager.close_database()
                return

            if 'sqlite_sequence' in tables:
                tables.remove('sqlite_sequence')

            table_name, ok = QInputDialog.getItem(self, "Выбор таблицы", "Выберите таблицу для загрузки:", tables, 0,
                                                  False)
            if ok and table_name:
                custom_headers = ["Название", "Индексация", "Верх(м)", "Низ(м)", "Коэффициент кавернозности",
                                  "Плотность(г/см^3)"]

                model = self.db_manager.preview_table(table_name, custom_headers, self)

                if model:
                    self.tbl_stratigraphy.setRowCount(0)
                    self.tbl_stratigraphy.setColumnCount(model.columnCount() - 1)
                    for row in range(model.rowCount()):
                        self.tbl_stratigraphy.insertRow(row)
                        for col in range(1, model.columnCount()):
                            data = model.data(model.index(row, col))
                            item = QTableWidgetItem(str(data))
                            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                            self.tbl_stratigraphy.setItem(row, col - 1, item)
                    break
            else:
                break

        self.db_manager.close_database()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    mainWindow = MainWindow()
    mainWindow.show()
    sys.exit(app.exec())
