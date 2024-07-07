import sys
import os
import pandas as pd
from PyQt6.QtGui import QAction, QUndoStack, QUndoCommand, QKeySequence, QTextDocument
from PyQt6.QtWidgets import (QApplication, QMainWindow, QLineEdit, QPushButton, QStackedWidget, QHeaderView,
                             QTableWidget, QTableWidgetItem, QComboBox, QFileDialog, QDialog, QVBoxLayout, QTabWidget,
                             QMenu, QGraphicsScene, QGraphicsView)
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6 import uic
import csv
import numpy as np
from matplotlib.figure import Figure
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from mpl_toolkits.mplot3d import Axes3D
from bs4 import BeautifulSoup

path1 = "D:/program/csv_files/"

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
        rows = self.text_data.split('\n')
        self.old_data = []

        total_cells_needed = sum(len(row.split('\t')) for row in rows if row.strip() != "")
        current_cells_available = (self.tableWidget.rowCount() - self.start_row) * self.tableWidget.columnCount() - self.start_col
        self.new_rows_needed = max(0, (total_cells_needed - current_cells_available + self.tableWidget.columnCount() - 1) // self.tableWidget.columnCount())

        for _ in range(self.new_rows_needed):
            self.tableWidget.insertRow(self.tableWidget.rowCount())

        doc = QTextDocument()

        current_row = self.start_row
        for row_data in rows:
            columns = row_data.split('\t')
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

                print(f"Inserting '{value}' at row {current_row}, col {current_col}")  # Debug output
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
    data_selected = pyqtSignal(list)

    def __init__(self, file_name, parent=None):
        super().__init__(parent)
        self.file_name = file_name
        self.initUI()

    def initUI(self):
        self.setWindowTitle('CSV Data')
        self.setGeometry(100, 100, 800, 600)
        layout = QVBoxLayout()

        self.tableWidget = QTableWidget(self)
        layout.addWidget(self.tableWidget)

        self.setLayout(layout)

        self.load_csv()

    def load_csv(self):
        try:
            print(f"Loading file: {self.file_name}")
            with open(self.file_name, newline='', encoding='utf-8') as csvfile:
                csvreader = csv.reader(csvfile)
                data = list(csvreader)

                print("Data loaded from CSV file:")
                for row in data:
                    print(row)

                if data:
                    headers = data[0]
                    self.tableWidget.setColumnCount(len(headers))
                    self.tableWidget.setHorizontalHeaderLabels(headers)
                    self.tableWidget.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

                    for row_data in data[1:]:
                        row = self.tableWidget.rowCount()
                        self.tableWidget.insertRow(row)
                        for col, cell_data in enumerate(row_data):
                            item = QTableWidgetItem(cell_data)
                            self.tableWidget.setItem(row, col, item)

                    self.tableWidget.cellDoubleClicked.connect(self.cell_was_double_clicked)
                else:
                    print("No data found in the file.")
        except FileNotFoundError:
            print(f"File {self.file_name} not found.")
        except Exception as e:
            print(f"An error occurred while reading the file: {e}")

    def cell_was_double_clicked(self, row, column):
        row_data = []
        for col in range(self.tableWidget.columnCount()):
            item = self.tableWidget.item(row, col)
            if item:
                row_data.append(item.text())
            else:
                row_data.append('')
        print(f"Row data: {row_data}")
        self.data_selected.emit(row_data)
        self.accept()

class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
        uic.loadUi('app.ui', self)
        self.file_paths = {}
        self.current_file_path = None
        self.setup_ui()

    def setup_ui(self):
        self.undo_stack = QUndoStack(self)

        self.stackedWidget: QStackedWidget = self.findChild(QStackedWidget, 'stackedWidget')
        self.stackedWidget.setCurrentIndex(0)

        self.tableWidget1: QTableWidget = self.findChild(QTableWidget, 'tableWidget')
        self.tableWidget2: QTableWidget = self.findChild(QTableWidget, 'tableWidget_2')
        self.tableWidget3: QTableWidget = self.findChild(QTableWidget, 'tableWidget_3')
        self.tableWidget4: QTableWidget = self.findChild(QTableWidget, 'tableWidget_4')
        self.tableWidget5: QTableWidget = self.findChild(QTableWidget, 'tableWidget_5')

        allowed_tables = [self.tableWidget1, self.tableWidget3, self.tableWidget4, self.tableWidget5]

        for table in [self.tableWidget1, self.tableWidget3, self.tableWidget4, self.tableWidget5]:
            table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
            table.customContextMenuRequested.connect(lambda pos, tbl=table: self.create_context_menu(pos, tbl))

            paste_action = QAction('Вставить', self)
            paste_action.setShortcut(QKeySequence.StandardKey.Paste)
            if table in allowed_tables:
                paste_action.triggered.connect(lambda checked, tbl=table: self.paste_from_clipboard(tbl))

            table.addAction(paste_action)

        self.pushButton2: QPushButton = self.findChild(QPushButton, 'pushButton_2')
        self.pushButton3: QPushButton = self.findChild(QPushButton, 'pushButton_3')
        self.pushButton4: QPushButton = self.findChild(QPushButton, 'pushButton_4')
        self.pushButton5: QPushButton = self.findChild(QPushButton, 'pushButton_5')
        self.pushButton6: QPushButton = self.findChild(QPushButton, 'pushButton_6')
        self.pushButton7: QPushButton = self.findChild(QPushButton, 'pushButton_7')
        self.pushButton8: QPushButton = self.findChild(QPushButton, 'pushButton_8')
        self.pushButton9: QPushButton = self.findChild(QPushButton, 'pushButton_9')
        self.pushButton10: QPushButton = self.findChild(QPushButton, 'pushButton_10')
        self.pushButton11: QPushButton = self.findChild(QPushButton, 'pushButton')
        self.pushButton12: QPushButton = self.findChild(QPushButton, 'pushButton_11')
        self.pushButton13: QPushButton = self.findChild(QPushButton, 'pushButton_12')

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
        self.undo_action = self.undo_stack.createUndoAction(self, "Отменить")
        self.undo_action.setShortcut(QKeySequence.StandardKey.Undo)

        self.redo_action = self.undo_stack.createRedoAction(self, "Повторить")
        self.redo_action.setShortcut(QKeySequence.StandardKey.Redo)

        self.addAction(self.undo_action)
        self.addAction(self.redo_action)

        self.stackedWidget.currentChanged.connect(self.on_current_index_changed)
        self.pushButton2.clicked.connect(self.go_to_next_page)
        self.pushButton3.clicked.connect(self.add_row)
        self.pushButton4.clicked.connect(self.go_to_previous_page)
        self.pushButton5.clicked.connect(self.delete_row_1)
        self.pushButton6.clicked.connect(self.clear_table)
        self.pushButton7.clicked.connect(self.add_row_2)
        self.pushButton8.clicked.connect(self.delete_row_2)
        self.pushButton9.clicked.connect(self.delete_row_3)
        self.pushButton10.clicked.connect(self.add_row_3)
        self.pushButton11.clicked.connect(self.delete_row_4)
        self.pushButton12.clicked.connect(self.add_row_4)
        self.pushButton13.clicked.connect(self.load_table)

        self.tableWidget1.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.tableWidget2.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.tableWidget3.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.tableWidget4.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.tableWidget5.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

        self.tableWidget1.horizontalHeader().setVisible(True)
        self.tableWidget2.horizontalHeader().setVisible(True)
        self.tableWidget3.horizontalHeader().setVisible(True)
        self.tableWidget4.horizontalHeader().setVisible(True)
        self.tableWidget5.horizontalHeader().setVisible(True)

        self.pushButton4.setVisible(False)

        self.tableWidget2.cellDoubleClicked.connect(self.open_csv_table_dialog)
        self.tableWidget2.cellDoubleClicked.connect(self.open_fixed_path_csv_dialog)

        self.tableWidget2.setItem(0, 0, QTableWidgetItem("Долото"))
        combo_1 = QComboBox()
        combo_1.addItems(["Направление", "Кондуктор", "Промежуточная", "Промежуточная 1", "Промежуточная 2",
                          "Потайная", "Эксплуатационная", "Хвостовик", "Райзер", "Фильтр", "Не определена"])
        self.tableWidget4.setCellWidget(0, 0, combo_1)

        print("Current working directory:", os.getcwd())
        print("Files in directory 'csv_files':")
        csv_files_path = os.path.join(os.getcwd(), "csv_files")
        if os.path.exists(csv_files_path) and os.path.isdir(csv_files_path):
            for filename in os.listdir(csv_files_path):
                print(filename)

    def create_context_menu(self, pos, table):
        context_menu = QMenu(self)

        context_menu.addAction(self.undo_action)
        context_menu.addAction(self.redo_action)

        paste_action = QAction('Вставить', self)
        paste_action.setShortcut(QKeySequence.StandardKey.Paste)
        paste_action.triggered.connect(lambda checked, tbl=table: self.paste_from_clipboard(tbl))

        context_menu.addAction(paste_action)

        context_menu.exec(table.viewport().mapToGlobal(pos))

    def on_current_index_changed(self, index):
        if index == 0:
            self.pushButton4.setVisible(False)
        else:
            self.pushButton4.setVisible(True)
        if index == 4:
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

    def add_row(self):
        print("Adding row...")
        row_count2_1 = self.tableWidget2.rowCount()
        self.tableWidget2.setRowCount(row_count2_1 + 1)
        for column in range(self.tableWidget2.columnCount()):
            if column == 0:
                combo = QComboBox()
                combo.addItems(["ВЗД", "РУС", "Бурильные трубы", "Переводник", "Предохранительный переводник", "УБТ",
                                "Телеметрия", "Ясс", "Калибратор", "Обратный клапан"])
                self.tableWidget2.setCellWidget(row_count2_1, column, combo)
            else:
                self.tableWidget2.setItem(row_count2_1, column, QTableWidgetItem(""))

    def open_csv_table_dialog(self, row, column):
        print(f"open_csv_table_dialog called with row {row} and column {column}")
        if column == 1:
            combo = self.tableWidget2.cellWidget(row, 0)
            if combo:
                file_key = combo.currentText()
                print(f"Selected file key: {file_key}")
                files = {
                    "ВЗД": path1 + "/ВЗД.csv",
                    "РУС": path1 + "/РУС.csv",
                    "Бурильные трубы": path1 + "/Бурильные трубы.csv",
                    "Переводник": path1 + "/Переводник.csv",
                    "Предохранительный переводник": path1 + "/Предохранительный переводник.csv",
                    "Обратный клапан": path1 + "/Обратный клапан.csv",
                    "Ясс": path1 + "/Ясс.csv",
                    "Калибратор": path1 + "/Калибратор.csv",
                    "УБТ": path1 + "/УБТ.csv",
                    "Телеметрия": path1 + "/Телеметрия.csv"
                }

                if file_key in files:
                    file_name = files[file_key]
                    if os.path.exists(file_name):
                        print(f"Opening file dialog for file: {file_name}")
                        self.current_file_path = file_name  # Устанавливаем текущий путь к файлу
                        dialog = CsvTableDialog(file_name, self)
                        dialog.data_selected.connect(lambda data: self.update_table_data(data, row, column))
                        dialog.rejected.connect(lambda: self.csv_dialog_rejected(row, column))
                        dialog.exec()

    def open_fixed_path_csv_dialog(self, row, column):
        print(f"open_fixed_path_csv_dialog called with row {row} and column {column}")
        if column == 1 and row == 0:
            fixed_file_path = path1 + "Долото.csv"  # Укажите здесь ваш фиксированный путь
            if os.path.exists(fixed_file_path):
                print(f"Opening file dialog for fixed file: {fixed_file_path}")
                self.current_file_path = fixed_file_path  # Устанавливаем текущий путь к файлу
                dialog = CsvTableDialog(fixed_file_path, self)
                dialog.data_selected.connect(lambda data: self.update_table_data(data, row, column))
                dialog.rejected.connect(lambda: self.csv_dialog_rejected_2(row, column))
                dialog.exec()

    def store_file_path(self, file_key, file_path):
        self.file_paths[file_key] = file_path
        print(f"File path stored for {file_key}: {file_path}")

    def get_file_path(self, file_key):
        return self.file_paths.get(file_key, None)

    def remove_widgets_from_row(self, table, row):
        for column in range(1, table.columnCount()):
            widget = table.cellWidget(row, column)
            if widget:
                table.removeCellWidget(row, column)
                widget.deleteLater()

    def update_table_data(self, data, row, column):
        print(f"Updating table data at row {row}, column {column} with {data}")
        self.remove_widgets_from_row(self.tableWidget2, row)

        if isinstance(data, list):
            offset = 0
            for col_index, value in enumerate(data):
                self.tableWidget2.setItem(row, column + col_index + offset, QTableWidgetItem(value))
        else:
            self.tableWidget2.setItem(row, column, QTableWidgetItem(data))

    def column_clicked(self, index):
        print(f"Column {index} clicked")

    def paste_from_clipboard(self, tableWidget):
        allowed_tables = [self.tableWidget1, self.tableWidget3, self.tableWidget4, self.tableWidget5]

        if tableWidget not in allowed_tables:
            print("This table is not allowed for pasting.")
            return

        print("Pasting from clipboard...")
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

        soup = BeautifulSoup(html, "lxml")
        return soup.get_text("\t", strip=True)

    def clear_table(self):
        print("Clearing table...")
        self.tableWidget1.clearContents()  # Очищает содержимое, не меняя количество строк и столбцов

        headers = ["Глубина по стволу (м)", "Зенитный угол (град)", "Азимут (град)", "Азимут маг(град)",
                   "Азимут дир(град)", "Глубина по верт(м)"]
        current_column_count = self.tableWidget1.columnCount()

        if current_column_count < len(headers):
            self.tableWidget1.setColumnCount(len(headers))

        self.tableWidget1.setHorizontalHeaderLabels(headers)
        self.tableWidget1.horizontalHeader().setVisible(True)
        self.tableWidget1.horizontalHeader().repaint()

    def load_table(self):
        print(1)

    def add_row_2(self):
        print("Adding row to tableWidget3...")
        row_count2_2 = self.tableWidget3.rowCount()
        self.tableWidget3.setRowCount(row_count2_2 + 1)
        for column in range(self.tableWidget3.columnCount()):
            self.tableWidget3.setItem(row_count2_2, column, QTableWidgetItem(""))

    def add_row_3(self):
        print("Adding row to tableWidget4...")
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
        print("Adding row to tableWidget5...")
        row_count2_4 = self.tableWidget5.rowCount()
        self.tableWidget5.setRowCount(row_count2_4 + 1)
        for column in range(self.tableWidget5.columnCount()):
            self.tableWidget5.setItem(row_count2_4, column, QTableWidgetItem(""))

    def open_file(self):
        try:
            xl, _ = QFileDialog.getOpenFileName(self, 'Open file', 'D:/Программа бурения/', 'Excel Files (*.xlsx)')
            if xl:
                k = pd.read_excel(xl)
                num_rows, num_cols = k.shape
                self.tableWidget1.setRowCount(num_rows)
                self.tableWidget1.setColumnCount(num_cols)
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

    def delete_row_1(self):
        print("Deleting row...")
        row_count2_1 = self.tableWidget2.rowCount()
        if row_count2_1 > 0:
            self.tableWidget2.setRowCount(row_count2_1 - 1)

    def delete_row_2(self):
        print("Deleting row...")
        row_count2_2 = self.tableWidget3.rowCount()
        if row_count2_2 > 0:
            self.tableWidget3.setRowCount(row_count2_2 - 1)

    def delete_row_3(self):
        print("Deleting row...")
        row_count2_3 = self.tableWidget4.rowCount()
        if row_count2_3 > 0:
            self.tableWidget4.setRowCount(row_count2_3 - 1)

    def delete_row_4(self):
        print("Deleting row...")
        row_count2_4 = self.tableWidget5.rowCount()
        if row_count2_4 > 0:
            self.tableWidget5.setRowCount(row_count2_4 - 1)

    def csv_dialog_rejected(self, row, column):
        print(f"CSV dialog closed without selection for row {row}, column {column}")
        item = self.tableWidget2.item(row, column)
        if item is not None and item.text() == "":
            for col in range(self.tableWidget2.columnCount()):
                if col == 5:
                    combo2 = QComboBox()
                    combo2.addItems(["Ниппель", "Муфта"])
                    self.tableWidget2.setCellWidget(row, col, combo2)
                elif col == 6:
                    combo3 = QComboBox()
                    combo3.addItems(
                        ["З-76", "З-86", "З-88", "З-94", "З-101", "З-102", "З-108", "З-118", "З-121", "З-122",
                         "З-133", "З-140", "З-147", "З-152", "З-161", "З-163", "З-171"])
                    self.tableWidget2.setCellWidget(row, col, combo3)
                elif col == 7:
                    combo4 = QComboBox()
                    combo4.addItems(["Ниппель", "Муфта"])
                    self.tableWidget2.setCellWidget(row, col, combo4)
                elif col == 8:
                    combo5 = QComboBox()
                    combo5.addItems(
                        ["З-76", "З-86", "З-88", "З-94", "З-101", "З-102", "З-108", "З-118", "З-121", "З-122",
                         "З-133", "З-140", "З-147", "З-152", "З-161", "З-163", "З-171"])
                    self.tableWidget2.setCellWidget(row, col, combo5)

    def csv_dialog_rejected_2(self, row, column):
        print(f"CSV dialog closed without selection for row {row}, column {column}")
        item = self.tableWidget2.item(row, column)
        if item is not None and item.text() == "":
            for col in range(self.tableWidget2.columnCount()):
                if col == 7:
                    combo4 = QComboBox()
                    combo4.addItems(["Ниппель", "Муфта"])
                    self.tableWidget2.setCellWidget(row, col, combo4)
                elif col == 8:
                    combo5 = QComboBox()
                    combo5.addItems(
                        ["З-76", "З-86", "З-88", "З-94", "З-101", "З-102", "З-108", "З-118", "З-121", "З-122",
                         "З-133", "З-140", "З-147", "З-152", "З-161", "З-163", "З-171"])
                    self.tableWidget2.setCellWidget(row, col, combo5)

    def contextMenuEvent(self, event):
        if self.childAt(event.pos()) == self.tableWidget2.viewport():
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

            contextMenu.exec(self.mapToGlobal(event.pos()))

    def saveAstemplate(self):
        if self.tableWidget2.hasFocus():
            row = self.tableWidget2.currentRow()
            data = []
            for column in range(self.tableWidget2.columnCount()):
                combo = self.tableWidget2.cellWidget(row, column)
                if combo and isinstance(combo, QComboBox):
                    text = combo.currentText()
                else:
                    item = self.tableWidget2.item(row, column)
                    text = item.text() if item is not None else ''
                data.append(text)
            print(data)
            self.write_csv(data)

    def saveAsall(self):
        print(1)

    def copyRow(self):
        if self.tableWidget2.hasFocus():
            row = self.tableWidget2.currentRow()
            if row != -1:
                data = []
                for column in range(self.tableWidget2.columnCount()):
                    combo = self.tableWidget2.cellWidget(row, column)
                    if combo and isinstance(combo, QComboBox):
                        text = combo.currentText()
                    else:
                        item = self.tableWidget2.item(row, column)
                        text = item.text() if item is not None else ''
                    data.append(text)

                newRow = self.tableWidget2.rowCount()
                self.tableWidget2.insertRow(newRow)
                for column, text in enumerate(data):
                    newItem = QTableWidgetItem(text)
                    self.tableWidget2.setItem(newRow, column, newItem)

    def write_csv(self, data):
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


if __name__ == "__main__":
    app = QApplication(sys.argv)
    mainWindow = MainWindow()
    mainWindow.show()
    sys.exit(app.exec())
