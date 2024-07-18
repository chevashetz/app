import sys
import sqlite3
import pandas as pd
from PyQt6.QtGui import QAction, QUndoStack, QUndoCommand, QKeySequence, QTextDocument, QFontMetrics
from PyQt6.QtWidgets import (QApplication, QMainWindow, QLineEdit, QPushButton, QStackedWidget, QHeaderView,
                             QTableWidget, QTableWidgetItem, QComboBox, QFileDialog, QDialog, QVBoxLayout, QMenu,
                             QGraphicsScene, QGraphicsView, QUndoView, QWidget, QHBoxLayout, QLabel)
from PyQt6.QtCore import Qt, pyqtSignal, QPoint
from PyQt6 import uic
import csv
import numpy as np
from matplotlib.figure import Figure
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from mpl_toolkits.mplot3d import Axes3D
from bs4 import BeautifulSoup

path1 = "D:/program/сsv_files/"

class UpdateTableCommand(QUndoCommand):
    def __init__(self, tableWidget, old_data, new_data, description="загрузку КНБК"):
        super().__init__(description)
        self.tableWidget = tableWidget
        self.old_data = old_data
        self.new_data = new_data

    def undo(self):
        self.set_table_data(self.old_data)

    def redo(self):
        self.set_table_data(self.new_data)

    def set_table_data(self, data):
        self.tableWidget.clearContents()
        self.tableWidget.setRowCount(len(data))
        for row_index, row_data in enumerate(data):
            for col_index, value in enumerate(row_data):
                if value is None:
                    value = ""
                self.tableWidget.setItem(row_index, col_index, QTableWidgetItem(value))

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
                self.tableWidget.setItem(row, col, item)
        for _ in range(self.new_rows_needed):
            self.tableWidget.removeRow(self.tableWidget.rowCount() - 1)

    def redo(self):
        rows = self.text_data
        self.old_data = []

        total_cells_needed = sum(len(row) for row in rows)
        current_cells_available = (self.tableWidget.rowCount() - self.start_row) * self.tableWidget.columnCount() - self.start_col
        self.new_rows_needed = max(0, (total_cells_needed - current_cells_available + self.tableWidget.columnCount() - 1) // self.tableWidget.columnCount())

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
                self.tableWidget.setItem(current_row, current_col, item)
                current_col += 1

            self.old_data.append((current_row, old_row_data))
            current_row += 1


class CsvTableDialog(QDialog):
    data_selected = pyqtSignal(list, str)  # Обновляем сигнал, чтобы принимать два аргумента

    def __init__(self, file_name, load_table=False, initial_sort_column=5, initial_sort_value=None, parent=None):
        super().__init__(parent)
        self.file_name = file_name
        self.load_table = load_table
        self.sort_order = Qt.SortOrder.AscendingOrder
        self.sort_column = initial_sort_column
        self.initial_sort_value = initial_sort_value
        self.initUI()

    def initUI(self):
        self.setWindowTitle('CSV Data')
        self.setGeometry(100, 100, 1100, 700)
        layout = QVBoxLayout()

        self.tableWidget = QTableWidget(self)
        layout.addWidget(self.tableWidget)

        self.setLayout(layout)
        self.load_csv()
        self.tableWidget.setSortingEnabled(False)
        self.tableWidget.horizontalHeader().setSortIndicatorShown(True)
        self.tableWidget.horizontalHeader().sectionClicked.connect(self.on_header_clicked)


    def load_csv(self):
        try:
            with open(self.file_name, newline='', encoding='utf-8') as csvfile:
                csvreader = csv.reader(csvfile)
                data = list(csvreader)

                if data:
                    headers = data[0]
                    self.tableWidget.setColumnCount(len(headers))
                    self.tableWidget.setHorizontalHeaderLabels(headers)
                    self.tableWidget.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

                    for row_data in data[1:]:
                        row = self.tableWidget.rowCount()
                        self.tableWidget.insertRow(row)
                        for col, cell_data in enumerate(row_data):
                            item = QTableWidgetItem(cell_data.strip())
                            self.tableWidget.setItem(row, col, item)

                    if not self.load_table:
                        self.tableWidget.cellDoubleClicked.connect(self.cell_was_double_clicked)
                    else:
                        self.tableWidget.cellDoubleClicked.connect(self.cell_was_double_clicked_2)

                    if self.sort_column is not None and self.initial_sort_value is not None:
                        self.sort_table(self.sort_order, initial_sort=True)
                else:
                    print("No data found in the file.")
        except Exception as e:
            print(f"Error loading CSV: {e}")

    def custom_sort_key(self, row_data):
        text = row_data[self.sort_column]
        if self.sort_column == 0:
            return text  # Текстовые значения для 0-го столбца
        elif self.sort_column in [1, 2, 3, 8, 9]:  # Числовые значения
            try:
                return float(text)
            except ValueError:
                return float('-inf')  # Обработка некорректных числовых значений
        elif self.sort_column in [4, 6]:  # Индексы столбцов, которые имеют только 2 текстовых значения
            return text
        elif self.sort_column in [5, 7]:  # Индексы столбцов вида "З-76"
            if text.startswith('З-'):
                return int(text.split('-')[1])
        return text

    def on_header_clicked(self, logical_index):
        # Устанавливаем индекс столбца для сортировки
        self.sort_column = logical_index
        self.sort_table(self.sort_order)
        # Переключение направления сортировки
        self.sort_order = Qt.SortOrder.DescendingOrder if self.sort_order == Qt.SortOrder.AscendingOrder \
            else Qt.SortOrder.AscendingOrder

    def sort_table(self, order, initial_sort=False):
        data = []
        for row in range(self.tableWidget.rowCount()):
            row_data = []
            for column in range(self.tableWidget.columnCount()):
                item = self.tableWidget.item(row, column)
                row_data.append(item.text() if item else "")
            data.append(row_data)

        if initial_sort and self.initial_sort_value is not None:
            # Первоначальная сортировка по значению из 8-го столбца
            data.sort(key=lambda row: (row[self.sort_column] != self.initial_sort_value, self.custom_sort_key(row)))
        else:
            data.sort(key=self.custom_sort_key, reverse=(order == Qt.SortOrder.DescendingOrder))

        self.tableWidget.setRowCount(0)
        for row_data in data:
            row = self.tableWidget.rowCount()
            self.tableWidget.insertRow(row)
            for column, item in enumerate(row_data):
                self.tableWidget.setItem(row, column, QTableWidgetItem(item))

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
            item = self.tableWidget.item(row, 1)
            if item:
                data = item.text()
                print(f"Double clicked data: {data}")
                row_data = []
                keys = []
                for part in data.split(';'):
                    part = part.strip()
                    print(f"Processing part: {part}")
                    if '_' in part:
                        key, name = part.split('_')
                        key = key.strip()
                        name = name.strip()
                        keys.append(key)
                        csv_path = f'{path1}{key}.csv'
                        print(f"Opening file: {csv_path}")
                        try:
                            with open(csv_path, "r", encoding='utf-8') as csvfile:
                                csv_reader = csv.reader(csvfile)
                                for csv_row in csv_reader:
                                    print(f"Checking row in {csv_path}: {csv_row}")
                                    if name == csv_row[0].strip():
                                        print(f"Found match: {csv_row}")
                                        row_data.append([key] + csv_row)  # Добавляем ключ в начало данных строки
                                        break
                                    else:
                                        print(f"No match for {name} in {csv_row[0].strip()}")
                        except FileNotFoundError:
                            print(f"File not found: {csv_path}")
                        except Exception as e:
                            print(f"Error reading {csv_path}: {e}")
                    else:
                        print(f"No '_' found in part: {part}")
                print(f"Emitting data: {row_data} with keys: {keys}")
                self.data_selected.emit(row_data, ", ".join(keys))  # Передаем данные и ключи
                self.accept()
        except Exception as e:
            print(f"Error in cell_was_double_clicked_2: {e}")

class KNBK_Table(QWidget):
    def __init__(self, index, parent=None):
        super(KNBK_Table,self).__init__(parent)
        uic.loadUi('table.ui', self)
        self.setup_ui()

    def setup_ui(self):
        self.undo_stack = QUndoStack(self)
        self.undo_view = QUndoView(self.undo_stack)
        self.tbl_KNBK: QTableWidget = self.findChild(QTableWidget, 'KNBK')

        self.btn_add_row_KNBK: QPushButton = self.findChild(QPushButton, 'add_row')
        self.btn_delete_row_KNBK: QPushButton = self.findChild(QPushButton, 'delete_row')
        self.btn_load_table: QPushButton = self.findChild(QPushButton, 'load_KNBK')
        self.btn_row_up: QPushButton = self.findChild(QPushButton, 'row_up')
        self.btn_row_down: QPushButton = self.findChild(QPushButton, 'row_down')
        self.btn_add_page_KNBK: QPushButton = self.findChild(QPushButton, 'add_page')

        self.open_file_act: QAction = self.findChild(QAction, 'actionOpen')

        self.tbl_KNBK.cellDoubleClicked.connect(self.open_csv_table_dialog)
        self.tbl_KNBK.cellDoubleClicked.connect(self.open_fixed_path_csv_dialog)
        self.tbl_KNBK.setItem(0, 0, QTableWidgetItem("Долото"))

        self.btn_add_row_KNBK.clicked.connect(self.add_row_KNBK)
        self.btn_delete_row_KNBK.clicked.connect(self.delete_row_KNBK)
        self.btn_load_table.clicked.connect(self.load_table_KNBK)
        self.btn_row_up.clicked.connect(self.row_up_KNBK)
        self.btn_row_down.clicked.connect(self.row_down_KNBK)


    def add_row_KNBK(self):
        row_count2_1 = self.tbl_KNBK.rowCount()
        self.tbl_KNBK.setRowCount(row_count2_1 + 1)
        for column in range(self.tbl_KNBK.columnCount()):
            if column == 0:
                combo = QComboBox()
                combo.addItems(["ВЗД", "РУС", "Бурильные трубы", "Переводник", "Предохранительный переводник", "УБТ",
                                "Телеметрия", "Ясс", "Калибратор", "Обратный клапан"])
                self.tbl_KNBK.setCellWidget(row_count2_1, column, combo)
            else:
                self.tbl_KNBK.setItem(row_count2_1, column, QTableWidgetItem(""))
        #self.tbl_KNBK.resizeRowsToContents()

    def delete_row_KNBK(self):
        row_count2_1 = self.tbl_KNBK.rowCount()
        if row_count2_1 > 0:
            self.tbl_KNBK.setRowCount(row_count2_1 - 1)

    def load_table_KNBK(self):
        dialog = CsvTableDialog(path1 + 'КНБК.csv', load_table=True, parent=self)
        dialog.data_selected.connect(self.update_table_data_list_2)
        dialog.exec()

    def row_up_KNBK(self):
        current_row = self.tbl_KNBK.currentRow()
        if current_row > 0:
            self.swap_rows(current_row, current_row - 1)
            self.tbl_KNBK.setCurrentCell(current_row - 1, 0)

    def row_down_KNBK(self):
        current_row = self.tbl_KNBK.currentRow()
        if current_row < self.tbl_KNBK.rowCount() - 1:
            self.swap_rows(current_row, current_row + 1)
            self.tbl_KNBK.setCurrentCell(current_row + 1, 0)

    def swap_rows(self, row1, row2):
        for column in range(self.tbl_KNBK.columnCount()):
            widget1 = self.tbl_KNBK.cellWidget(row1, column)
            widget2 = self.tbl_KNBK.cellWidget(row2, column)
            item1 = self.tbl_KNBK.item(row1, column)
            item2 = self.tbl_KNBK.item(row2, column)

            if widget1 or widget2:
                if isinstance(widget1, QComboBox):
                    index1 = widget1.currentIndex()
                    text1 = widget1.currentText()
                else:
                    index1 = None
                    text1 = None

                if isinstance(widget2, QComboBox):
                    index2 = widget2.currentIndex()
                    text2 = widget2.currentText()
                else:
                    index2 = None
                    text2 = None

                if text1 is not None:
                    combo1 = QComboBox()
                    combo1.addItems(["ВЗД", "РУС", "Бурильные трубы", "Переводник", "Предохранительный переводник", "УБТ",
                                     "Телеметрия", "Ясс", "Калибратор", "Обратный клапан"])
                    combo1.setCurrentIndex(index1)
                    self.tbl_KNBK.setCellWidget(row2, column, combo1)
                else:
                    self.tbl_KNBK.removeCellWidget(row2, column)

                if text2 is not None:
                    combo2 = QComboBox()
                    combo2.addItems(["ВЗД", "РУС", "Бурильные трубы", "Переводник", "Предохранительный переводник", "УБТ",
                                     "Телеметрия", "Ясс", "Калибратор", "Обратный клапан"])
                    combo2.setCurrentIndex(index2)
                    self.tbl_KNBK.setCellWidget(row1, column, combo2)
                else:
                    self.tbl_KNBK.removeCellWidget(row1, column)

            if item1 or item2:
                text1 = item1.text() if item1 else ''
                text2 = item2.text() if item2 else ''
                self.tbl_KNBK.setItem(row1, column, QTableWidgetItem(text2))
                self.tbl_KNBK.setItem(row2, column, QTableWidgetItem(text1))

    def open_csv_table_dialog(self, row, column):
        if column == 1:
            combo = self.tbl_KNBK.cellWidget(row, 0)
            if combo:
                file_key = combo.currentText()
                print(f"Selected file key: {file_key}")
                files = {
                    "ВЗД": path1 + "ВЗД.csv",
                    "РУС": path1 + "РУС.csv",
                    "Бурильные трубы": path1 + "Бурильные трубы.csv",
                    "Переводник": path1 + "Переводник.csv",
                    "Предохранительный переводник": path1 + "Предохранительный переводник.csv",
                    "Обратный клапан": path1 + "Обратный клапан.csv",
                    "Ясс": path1 + "Ясс.csv",
                    "Калибратор": path1 + "Калибратор.csv",
                    "УБТ": path1 + "УБТ.csv",
                    "Телеметрия": path1 + "Телеметрия.csv"
                }

                if file_key in files:
                    file_name = files[file_key]
                    print(f"Opening file dialog for file: {file_name}")
                    self.current_file_path = file_name
                    # Получаем значение для первоначальной сортировки из 8-го столбца предпоследней строки
                    initial_sort_value = self.get_initial_sort_value()
                    dialog = CsvTableDialog(file_name, load_table=False, initial_sort_column=5,
                                            initial_sort_value=initial_sort_value, parent=self)
                    dialog.data_selected.connect(lambda data: self.update_table_data(data, row, column))
                    dialog.rejected.connect(lambda: self.csv_dialog_rejected(row, column))
                    dialog.exec()

    def open_fixed_path_csv_dialog(self, row, column):
        if column == 1 and row == 0:
            fixed_file_path = path1 + "Долото.csv"
            self.current_file_path = fixed_file_path
            # Получаем значение для первоначальной сортировки из 8-го столбца предпоследней строки
            initial_sort_value = self.get_initial_sort_value()
            dialog = CsvTableDialog(fixed_file_path, load_table=False, initial_sort_column=5,
                                    initial_sort_value=initial_sort_value, parent=self)
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
                    self.tbl_KNBK.setCellWidget(row, col, combo2)
                elif col == 6:
                    combo3 = QComboBox()
                    combo3.addItems(
                        ["З-76", "З-86", "З-88", "З-94", "З-101", "З-102", "З-108", "З-118", "З-121", "З-122",
                         "З-133", "З-140", "З-147", "З-152", "З-161", "З-163", "З-171"])
                    self.tbl_KNBK.setCellWidget(row, col, combo3)
                elif col == 7:
                    combo4 = QComboBox()
                    combo4.addItems(["Ниппель", "Муфта"])
                    self.tbl_KNBK.setCellWidget(row, col, combo4)
                elif col == 8:
                    combo5 = QComboBox()
                    combo5.addItems(
                        ["З-76", "З-86", "З-88", "З-94", "З-101", "З-102", "З-108", "З-118", "З-121", "З-122",
                         "З-133", "З-140", "З-147", "З-152", "З-161", "З-163", "З-171"])
                    self.tbl_KNBK.setCellWidget(row, col, combo5)

    def csv_dialog_rejected_2(self, row, column):
        item = self.tbl_KNBK.item(row, column)
        if item is not None and item.text() == "":
            for col in range(self.tbl_KNBK.columnCount()):
                if col == 7:
                    combo4 = QComboBox()
                    combo4.addItems(["Ниппель", "Муфта"])
                    self.tbl_KNBK.setCellWidget(row, col, combo4)
                elif col == 8:
                    combo5 = QComboBox()
                    combo5.addItems(
                        ["З-76", "З-86", "З-88", "З-94", "З-101", "З-102", "З-108", "З-118", "З-121", "З-122",
                         "З-133", "З-140", "З-147", "З-152", "З-161", "З-163", "З-171"])
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
                combo = self.tbl_KNBK.cellWidget(row, column)
                if combo and isinstance(combo, QComboBox):
                    text = combo.currentText()
                else:
                    item = self.tbl_KNBK.item(row, column)
                    text = item.text() if item is not None else ''
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
            # Обрабатываем остальные строки
            for row in range(1,self.tbl_KNBK.rowCount()):  # Итерируемся по всем строкам
                combo_0 = self.tbl_KNBK.cellWidget(row, 0)  # Первый столбец с QComboBox
                item_1 = self.tbl_KNBK.item(row, 1)  # Второй столбец с обычным элементом

                text_0 = combo_0.currentText() if combo_0 is not None else ''
                text_1 = item_1.text() if item_1 is not None else ''

                combined_text = f" {text_0}_{text_1};"
                data.append(combined_text)

            data_str = ''.join(data)  # Объединяем с пробелом

            self.write_csv(data_str)

    def copyRow(self):
        if self.tbl_KNBK.hasFocus():
            row = self.tbl_KNBK.currentRow()
            if row != -1:
                data = []
                for column in range(self.tbl_KNBK.columnCount()):
                    combo = self.tbl_KNBK.cellWidget(row, column)
                    if combo and isinstance(combo, QComboBox):
                        text = combo.currentText()
                    else:
                        item = self.tbl_KNBK.item(row, column)
                        text = item.text() if item is not None else ''
                    data.append(text)

                newRow = self.tbl_KNBK.rowCount()
                self.tbl_KNBK.insertRow(newRow)
                for column, text in enumerate(data):
                    newItem = QTableWidgetItem(text)
                    self.tbl_KNBK.setItem(newRow, column, newItem)

    def write_csv(self, data):
        fixed_file_path = path1+"КНБК.csv"
        try:
            with open(fixed_file_path, 'a', newline='', encoding='utf-8') as f:
                f.write('\n'+data )
        except Exception as e:
            print(f"Error opening file: {e}")

    def write_csv_2(self, data):
        if self.current_file_path is None:
            print("Error: No file path set for writing data.")
            return

        try:
            # Открываем файл в режиме добавления ('a') с параметром 'newline'
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

        # Проверяем, есть ли строки в главной таблице
        if row_count > 1:
            # Получаем значение из 8-го столбца предпоследней строки
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
                self.tbl_KNBK.setItem(row, column + col_index + offset, QTableWidgetItem(value))
        else:
            self.tbl_KNBK.setItem(row, column, QTableWidgetItem(data))

    def update_table_data_list_2(self, data, key):
        try:
            print(f"Received data: {data} with key: {key}")

            # Сохраняем старые данные перед обновлением
            old_data = []
            for row in range(self.tbl_KNBK.rowCount()):
                row_data = []
                for col in range(self.tbl_KNBK.columnCount()):
                    item = self.tbl_KNBK.item(row, col)
                    row_data.append(item.text() if item else "")
                old_data.append(row_data)

            # Добавляем команду в стек отмены
            self.undo_stack.push(UpdateTableCommand(self.tbl_KNBK, old_data, data))

            # Обновляем данные таблицы (сейчас это будет сделано в UpdateTableCommand)
            self.update_table_widget(self.tbl_KNBK, data)

        except Exception as e:
            print(f"Error in update_table_data_list_2: {e}")

    def update_table_widget(self, tableWidget, data):
        tableWidget.clearContents()
        tableWidget.setRowCount(len(data))
        for row_index, row_data in enumerate(data):
            for col_index, value in enumerate(row_data):
                if value is None:
                    value = ""
                tableWidget.setItem(row_index, col_index, QTableWidgetItem(value))

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

    def setup_ui(self):
        self.stackedWidget: QStackedWidget = self.findChild(QStackedWidget, 'stackedWidget')
        if self.stackedWidget is None:
            print("Error: QStackedWidget not found in the .ui file")
            return

        self.stackedWidget.setCurrentIndex(0)
        self.stackedWidget.insertWidget(4,KNBK_Table(index = 4, parent = self))

        self.tableWidget1: QTableWidget = self.findChild(QTableWidget, 'tableWidget_1')
        self.tableWidget2: QTableWidget = self.findChild(QTableWidget, 'tableWidget_2')
        self.tableWidget3: QTableWidget = self.findChild(QTableWidget, 'tableWidget_3')
        self.tableWidget4: QTableWidget = self.findChild(QTableWidget, 'tableWidget_4')

        combo_1 = QComboBox()
        combo_1.addItems(["Направление", "Кондуктор", "Промежуточная", "Промежуточная 1", "Промежуточная 2",
                          "Потайная", "Эксплуатационная", "Хвостовик", "Райзер", "Фильтр", "Не определена"])
        self.tableWidget3.setCellWidget(0, 0, combo_1)

        allowed_tables = [self.tableWidget1, self.tableWidget2, self.tableWidget3, self.tableWidget4]

        for table in allowed_tables:
            table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
            table.customContextMenuRequested.connect(lambda pos, tbl=table: self.create_context_menu(pos, tbl))

            paste_action = QAction('Вставить', self)
            paste_action.setShortcut(QKeySequence.StandardKey.Paste)
            paste_action.triggered.connect(lambda checked, tbl=table: self.paste_from_clipboard(tbl))

            table.addAction(paste_action)

        self.pushButton2: QPushButton = self.findChild(QPushButton, 'pushButton_2')
        self.pushButton4: QPushButton = self.findChild(QPushButton, 'pushButton_4')
        self.pushButton6: QPushButton = self.findChild(QPushButton, 'pushButton_6')
        self.pushButton7: QPushButton = self.findChild(QPushButton, 'pushButton_7')
        self.pushButton8: QPushButton = self.findChild(QPushButton, 'pushButton_8')
        self.pushButton9: QPushButton = self.findChild(QPushButton, 'pushButton_9')
        self.pushButton10: QPushButton = self.findChild(QPushButton, 'pushButton_10')
        self.pushButton11: QPushButton = self.findChild(QPushButton, 'pushButton')
        self.pushButton12: QPushButton = self.findChild(QPushButton, 'pushButton_11')
        self.pushButton16: QPushButton = self.findChild(QPushButton, 'pushButton_17')

        self.lineEdit1: QLineEdit = self.findChild(QLineEdit, 'lineEdit_1')
        self.lineEdit2: QLineEdit = self.findChild(QLineEdit, 'lineEdit_2')
        self.lineEdit3: QLineEdit = self.findChild(QLineEdit, 'lineEdit_3')
        self.lineEdit4: QLineEdit = self.findChild(QLineEdit, 'lineEdit_4')

        self.graphicsView = self.findChild(QGraphicsView, 'graphicsView')

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

        # Create undo and redo actions

        self.stackedWidget.currentChanged.connect(self.on_current_index_changed)
        self.pushButton2.clicked.connect(self.go_to_next_page)
        self.pushButton4.clicked.connect(self.go_to_previous_page)
        self.pushButton6.clicked.connect(self.clear_table)
        self.pushButton7.clicked.connect(self.add_row_2)
        self.pushButton8.clicked.connect(self.delete_row_2)
        self.pushButton9.clicked.connect(self.delete_row_3)
        self.pushButton10.clicked.connect(self.add_row_3)
        self.pushButton11.clicked.connect(self.delete_row_4)
        self.pushButton12.clicked.connect(self.add_row_4)




        self.tableWidget1.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.tableWidget2.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.tableWidget3.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.tableWidget4.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

        self.tableWidget1.horizontalHeader().setVisible(True)
        self.tableWidget2.horizontalHeader().setVisible(True)
        self.tableWidget3.horizontalHeader().setVisible(True)
        self.tableWidget4.horizontalHeader().setVisible(True)

        self.pushButton4.setVisible(False)

        self.undo_action = self.undo_stack.createUndoAction(self, "Отменить")
        self.undo_action.setShortcut(QKeySequence.StandardKey.Undo)

        self.redo_action = self.undo_stack.createRedoAction(self, "Повторить")
        self.redo_action.setShortcut(QKeySequence.StandardKey.Redo)

        self.addAction(self.undo_action)
        self.addAction(self.redo_action)


    def create_context_menu(self, pos, table):
        context_menu = QMenu(self)

        context_menu.addAction(self.undo_action)
        context_menu.addAction(self.redo_action)

        paste_action = QAction('Вставить', self)
        paste_action.setShortcut(QKeySequence.StandardKey.Paste)
        paste_action.triggered.connect(lambda checked, tbl=table: self.paste_from_clipboard(tbl))

        context_menu.addAction(paste_action)

        context_menu.exec(table.viewport().mapToGlobal(pos))

    def get_page_count(self):
        return self.stackedWidget.count()

    def on_current_index_changed(self, index):
        total_pages = self.get_page_count()

        if index == 0:
            self.pushButton4.setVisible(False)
        else:
            self.pushButton4.setVisible(True)

        if index == total_pages - 1:  # Last page
            self.pushButton2.setVisible(False)
        else:
            self.pushButton2.setVisible(True)

    def go_to_next_page(self):
        current_index = self.stackedWidget.currentIndex()
        self.stackedWidget.setCurrentIndex(current_index + 1)

    def go_to_previous_page(self):
        current_index = self.stackedWidget.currentIndex()
        if current_index > 0:
            self.stackedWidget.setCurrentIndex(current_index - 1)

    def paste_from_clipboard(self, tableWidget):
        allowed_tables = [self.tableWidget1, self.tableWidget2, self.tableWidget3, self.tableWidget4]

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
            data.append([ele for ele in cols if ele])  # Get rid of empty values
        return data

    def clear_table(self):
        self.tableWidget1.clearContents()  # Очищает содержимое, не меняя количество строк и столбцов

        headers = ["Глубина по стволу (м)", "Зенитный угол (град)", "Азимут (град)", "Азимут маг(град)",
                   "Азимут дир(град)", "Глубина по верт(м)"]
        current_column_count = self.tableWidget1.columnCount()

        if current_column_count < len(headers):
            self.tableWidget1.setColumnCount(len(headers))

        self.tableWidget1.setHorizontalHeaderLabels(headers)
        self.tableWidget1.horizontalHeader().setVisible(True)
        self.tableWidget1.horizontalHeader().repaint()

    def add_row_2(self):
        row_count2_2 = self.tableWidget2.rowCount()
        self.tableWidget2.setRowCount(row_count2_2 + 1)
        for column in range(self.tableWidget2.columnCount()):
            self.tableWidget2.setItem(row_count2_2, column, QTableWidgetItem(""))

    def add_row_3(self):
        row_count2_3 = self.tableWidget4.rowCount()
        self.tableWidget4.setRowCount(row_count2_3 + 1)
        for column in range(self.tableWidget4.columnCount()):
            if column == 0:
                combo = QComboBox()
                combo.addItems(["Направление", "Кондуктор", "Промежуточная", "Промежуточная 1", "Промежуточная 2",
                                "Потайная", "Эксплуатационная", "Хвостовик", "Райзер", "Фильтр", "Не определена"])
                self.tableWidget4.setCellWidget(row_count2_3, column, combo)
            else:
                self.tableWidget4.setItem(row_count2_3, column, QTableWidgetItem(""))

    def add_row_4(self):
        row_count2_4 = self.tableWidget4.rowCount()
        self.tableWidget4.setRowCount(row_count2_4 + 1)
        for column in range(self.tableWidget4.columnCount()):
            self.tableWidget4.setItem(row_count2_4, column, QTableWidgetItem(""))

    def open_file(self):
        try:
            xl, _ = QFileDialog.getOpenFileName(self, 'Open file', 'D:/Программа бурения/', 'Excel Files (*.xlsx)')
            if xl:
                # Читаем первые несколько строк для определения наличия заголовков
                sample_data = pd.read_excel(xl, nrows=5)

                def has_headers(df):
                    first_row = df.iloc[0]
                    return all(isinstance(val, str) for val in first_row) and len(set(first_row)) == len(first_row)

                # Определяем наличие заголовков
                has_headers_flag = has_headers(sample_data)

                # Загрузка данных из файла
                k = pd.read_excel(xl, header=0 if has_headers_flag else None)
                num_rows, num_cols = k.shape
                self.tableWidget1.setRowCount(num_rows)
                self.tableWidget1.setColumnCount(num_cols)

                # Предопределенные заголовки
                predefined_headers = ["Глубина по стволу (м)", "Зенитный угол (град)", "Азимут (град)",
                                      "Азимут маг(град)", "Азимут дир(град)", "Глубина по верт(м)"]

                # Установка заголовков, если они есть, или создание их автоматически
                if has_headers_flag:
                    self.tableWidget1.setHorizontalHeaderLabels(k.columns.astype(str).tolist())
                else:
                    self.tableWidget1.setHorizontalHeaderLabels(predefined_headers[:num_cols])

                # Заполнение таблицы данными
                for index, row in k.iterrows():
                    for col_index, value in enumerate(row):
                        self.tableWidget1.setItem(index, col_index, QTableWidgetItem(str(value)))

                self.tableWidget1.horizontalHeader().setVisible(True)

                self.process_and_plot_data(k)
        except Exception as e:
            print(f"Error opening file: {e}")

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
        ax.set_xlim([max(data[:, 0]), min(data[:, 0])])
        ax.set_ylim([max(data[:, 1]), min(data[:, 1])])
        ax.set_zlim([max(data[:, 2]), min(data[:, 2])])
        ax.view_init(elev=20, azim=35)
        ax.set_box_aspect([1, 1, 1])

        scene = QGraphicsScene()
        canvas = FigureCanvas(fig)
        canvas.setGeometry(0, 0, 400, 600)
        scene.addWidget(canvas)

        self.graphicsView.setScene(scene)

    def delete_row_2(self):
        row_count2_2 = self.tableWidget2.rowCount()
        if row_count2_2 > 0:
            self.tableWidget2.setRowCount(row_count2_2 - 1)

    def delete_row_3(self):
        row_count2_3 = self.tableWidget4.rowCount()
        if row_count2_3 > 0:
            self.tableWidget4.setRowCount(row_count2_3 - 1)

    def delete_row_4(self):
        row_count2_4 = self.tableWidget4.rowCount()
        if row_count2_4 > 0:
            self.tableWidget4.setRowCount(row_count2_4 - 1)

    def add_page(self):
        current_index = self.stackedWidget.currentIndex()

        new_page = KNBK_Table(index = current_index+1, parent = self)

        self.stackedWidget.insertWidget(current_index + 1, new_page)

        self.stackedWidget.setCurrentWidget(new_page)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    mainWindow = MainWindow()
    mainWindow.show()
    sys.exit(app.exec())
