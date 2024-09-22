import csv

from PyQt6.QtCore import pyqtSignal, Qt
from PyQt6.QtWidgets import QDialog, QLineEdit, QLabel, QDialogButtonBox, QVBoxLayout, QTableWidget, QHeaderView, \
    QTableWidgetItem

from config import path1


class DualInputDialog(QDialog):
    def __init__(self, parent=None):
        super(DualInputDialog, self).__init__(parent)

        # Устанавливаем заголовок окна
        self.setWindowTitle("Ввод данных")

        # Создаём первый текстовый ввод
        self.first_input = QLineEdit(self)
        self.first_input.setPlaceholderText("Введите первое значение")

        # Создаём второй текстовый ввод
        self.second_input = QLineEdit(self)
        self.second_input.setPlaceholderText("Введите второе значение")

        # Добавляем метки для каждого ввода (опционально)
        self.first_label = QLabel("Первое значение:", self)
        self.second_label = QLabel("Второе значение:", self)

        # Создаём кнопки "ОК" и "Отмена"
        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)

        # Подключаем кнопки к функциям
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)

        # Размещение элементов в макете
        layout = QVBoxLayout()

        layout.addWidget(self.first_label)
        layout.addWidget(self.first_input)

        layout.addWidget(self.second_label)
        layout.addWidget(self.second_input)

        layout.addWidget(self.button_box)

        self.setLayout(layout)

    def get_inputs(self):
        #Возвращает значения, введённые пользователем в оба поля
        return self.first_input.text(), self.second_input.text()


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
