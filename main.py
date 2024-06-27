import sys
import os
import pandas as pd
from PyQt6.QtGui import QAction, QKeySequence
from PyQt6.QtWidgets import (QApplication, QMainWindow, QLineEdit, QPushButton, QStackedWidget, QHeaderView,
                             QTableWidget, QTableWidgetItem, QComboBox, QFileDialog, QDialog, QVBoxLayout, QTabWidget)
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6 import uic
import csv

class CsvTableDialog(QDialog):
    data_selected = pyqtSignal(list)  # Измените сигнал для передачи списка строк

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
        # Извлечение данных всей строки
        row_data = []
        for col in range(self.tableWidget.columnCount()):
            item = self.tableWidget.item(row, col)
            if item:
                row_data.append(item.text())
            else:
                row_data.append('')
        # Вывод отладочной информации
        print(f"Row data: {row_data}")
        # Передача данных строки через сигнал
        self.data_selected.emit(row_data)  # Передаем список строк
        self.accept()

class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
        # Загрузка интерфейса из файла ui
        uic.loadUi('app.ui', self)

        self.stackedWidget.setCurrentIndex(0)
        menubar = self.menuBar()
        fileMenu = menubar.addMenu('File')
        self.menuBar().setNativeMenuBar(False)
        self.open_file_act = QAction('Open', self)
        self.open_file_act.setShortcut('Ctrl+O')
        self.open_file_act.setStatusTip('Open new file')
        self.open_file_act.triggered.connect(self.open_file)
        fileMenu.addAction(self.open_file_act)

        self.pushButton1: QPushButton = self.findChild(QPushButton, 'pushButton')
        self.pushButton2: QPushButton = self.findChild(QPushButton, 'pushButton_2')
        self.pushButton3: QPushButton = self.findChild(QPushButton, 'pushButton_3')
        self.pushButton4: QPushButton = self.findChild(QPushButton, 'pushButton_4')
        self.pushButton5: QPushButton = self.findChild(QPushButton, 'pushButton_5')
        self.pushButton6: QPushButton = self.findChild(QPushButton, 'pushButton_6')
        self.pushButton7: QPushButton = self.findChild(QPushButton, 'pushButton_7')
        self.pushButton8: QPushButton = self.findChild(QPushButton, 'pushButton_8')

        self.lineEdit1: QLineEdit = self.findChild(QLineEdit, 'lineEdit_1')
        self.lineEdit2: QLineEdit = self.findChild(QLineEdit, 'lineEdit_2')
        self.lineEdit3: QLineEdit = self.findChild(QLineEdit, 'lineEdit_3')
        self.lineEdit4: QLineEdit = self.findChild(QLineEdit, 'lineEdit_4')

        self.stackedWidget: QStackedWidget = self.findChild(QStackedWidget, 'stackedWidget')

        self.tableWidget1: QTableWidget = self.findChild(QTableWidget, 'tableWidget')
        self.tableWidget2: QTableWidget = self.findChild(QTableWidget, 'tableWidget_2')
        self.tableWidget3: QTableWidget = self.findChild(QTableWidget, 'tableWidget_3')

        self.open_file_act: QAction = self.findChild(QAction, 'actionOpen')

        row_count2_1 = self.tableWidget2.rowCount()
        self.tableWidget2.setRowCount(row_count2_1)
        row_count2_2 = self.tableWidget3.rowCount()
        self.tableWidget3.setRowCount(row_count2_2)

        self.tableWidget2.setItem(0, 0, QTableWidgetItem("Долото"))
        combo_box1 = QComboBox()
        combo_box1.addItems(["PDC Шестилопостное", "PDC Пятилопостное", "PDC Четырехлопостное"])
        self.tableWidget2.setCellWidget(0, 1, combo_box1)
        combo_box2 = QComboBox()
        combo_box2.addItems(["Ниппель", "Муфта"])
        self.tableWidget2.setCellWidget(0, 7, combo_box2)
        combo_box3 = QComboBox()
        combo_box3.addItems(["З-76", "З-86", "З-88", "З-94", "З-101", "З-102", "З-108", "З-118", "З-121", "З-122",
                             "З-133", "З-140", "З-147", "З-152", "З-161", "З-163", "З-171"])
        self.tableWidget2.setCellWidget(0, 8, combo_box3)

        self.tableWidget1.setContextMenuPolicy(Qt.ContextMenuPolicy.ActionsContextMenu)
        paste_action = QAction('Paste', self)
        paste_action.setShortcut(QKeySequence.StandardKey.Paste)
        paste_action.triggered.connect(self.paste_from_clipboard)
        self.tableWidget1.addAction(paste_action)

        self.tableWidget1.horizontalHeader().sectionClicked.connect(self.column_clicked)

        self.stackedWidget.currentChanged.connect(self.on_current_index_changed)
        self.pushButton1.clicked.connect(self.on_button_click)
        self.pushButton2.clicked.connect(self.go_to_next_page)
        self.pushButton3.clicked.connect(self.add_row)
        self.pushButton4.clicked.connect(self.go_to_previous_page)
        self.pushButton5.clicked.connect(self.delete_row_1)
        self.pushButton6.clicked.connect(self.clear_table)
        self.pushButton7.clicked.connect(self.add_row_2)
        self.pushButton8.clicked.connect(self.delete_row_2)

        self.tableWidget1.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.tableWidget2.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.tableWidget3.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.tableWidget1.horizontalHeader().setVisible(True)
        self.tableWidget2.horizontalHeader().setVisible(True)
        self.tableWidget3.horizontalHeader().setVisible(True)
        self.pushButton4.setVisible(False)

        self.tableWidget2.cellDoubleClicked.connect(self.open_csv_table_dialog)

        # Списки для хранения кнопок и связанных строк
        self.buttons = []
        self.button_rows = []

        # Проверка текущего рабочего каталога
        print("Current working directory:", os.getcwd())

        # Вывод всех файлов в директории
        print("Files in directory 'csv_files':")
        csv_files_path = os.path.join(os.getcwd(), "csv_files")
        if os.path.exists(csv_files_path) and os.path.isdir(csv_files_path):
            for filename in os.listdir(csv_files_path):
                print(filename)

    def on_button_click(self):
        q = self.lineEdit1.text()
        q1 = self.lineEdit2.text()
        q2 = self.lineEdit3.text()
        q3 = self.lineEdit4.text()
        print(q, q1, q2, q3)

    def on_current_index_changed(self, index):
        if index == 0:
            self.pushButton4.setVisible(False)
        else:
            self.pushButton4.setVisible(True)
        if index == 3:
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

        self.add_buttons()

    def open_csv_table_dialog(self, row, column):
        print(f"open_csv_table_dialog called with row {row} and column {column}")
        if column == 1:
            combo = self.tableWidget2.cellWidget(row, 0)
            if combo:
                file_key = combo.currentText()
                print(f"Selected file key: {file_key}")
                files = {
                    "ВЗД": "D:/program/сsv_files/ВЗД.csv",
                    "РУС": "D:/program/сsv_files/РУС.csv",
                    "Бурильные трубы": "D:/program/сsv_files/Бурильные трубы.csv",
                    "Переводник": "D:/program/сsv_files/Переводник.csv",
                    "Предохранительный переводник": "D:/program/сsv_files/Предохранительный переводник.csv",
                    "Обратный клапан": "D:/program/сsv_files/Обратный клапан.csv",
                    "Ясс": "D:/program/сsv_files/Ясс.csv",
                    "Калибратор": "D:/program/сsv_files/Калибратор.csv",
                    "УБТ": "D:/program/сsv_files/УБТ.csv",
                    "Телеметрия": "D:/program/сsv_files/Телеметрия.csv"
                }

                if file_key in files:
                    file_name = files[file_key]
                    if os.path.exists(file_name):
                        print(f"Opening file dialog for file: {file_name}")
                        dialog = CsvTableDialog(file_name, self)
                        dialog.data_selected.connect(lambda data: self.update_table_data(data, row, column))
                        dialog.rejected.connect(lambda: self.csv_dialog_rejected(row, column))
                        dialog.exec()

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

    def column_clicked(self, column_index):
        print(f"Column {column_index} clicked")
        try:
            clipboard = QApplication.clipboard()
            mime_data = clipboard.mimeData()
            if mime_data.hasText():
                text_data = mime_data.text()
                rows = text_data.split('\n')
                current_row = 0

                for row_data in rows:
                    columns = row_data.split('\t')
                    if current_row >= self.tableWidget1.rowCount():
                        self.tableWidget1.insertRow(self.tableWidget1.rowCount())
                    for col_index, value in enumerate(columns):
                        if column_index + col_index >= self.tableWidget1.columnCount():
                            self.tableWidget1.insertColumn(self.tableWidget1.columnCount())
                        self.tableWidget1.setItem(current_row, column_index + col_index, QTableWidgetItem(value))
                    current_row += 1
        except Exception as e:
            print(f"Error processing column click: {e}")

    def paste_from_clipboard(self):
        print("Pasting from clipboard...")
        try:
            clipboard = QApplication.clipboard()
            mime_data = clipboard.mimeData()
            if mime_data.hasText():
                text_data = mime_data.text()
                rows = text_data.split('\n')
                current_row = self.tableWidget1.currentRow()
                current_column = self.tableWidget1.currentColumn()

                for row_data in rows:
                    columns = row_data.split('\t')
                    if current_row >= self.tableWidget1.rowCount():
                        self.tableWidget1.insertRow(self.tableWidget1.rowCount())
                    for col_index, value in enumerate(columns):
                        if current_column + col_index >= self.tableWidget1.columnCount():
                            self.tableWidget1.insertColumn(self.tableWidget1.columnCount())
                        self.tableWidget1.setItem(current_row, current_column + col_index, QTableWidgetItem(value))
                    current_row += 1
        except Exception as e:
            print(f"Error pasting from clipboard: {e}")

    def clear_table(self):
        print("Clearing table...")
        self.tableWidget1.clearContents()

    def add_row_2(self):
        print("Adding row to tableWidget3...")
        row_count2_2 = self.tableWidget3.rowCount()
        self.tableWidget3.setRowCount(row_count2_2 + 1)
        for column in range(self.tableWidget3.columnCount()):
            self.tableWidget3.setItem(row_count2_2, column, QTableWidgetItem(""))

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
        except Exception as e:
            print(f"Error opening file: {e}")

    def delete_row_1(self):
        print("Deleting row...")
        selected_row = self.tableWidget2.currentRow()
        if selected_row != -1:
            self.tableWidget2.removeRow(selected_row)
            self.remove_button_for_row(selected_row)
            self.update_button_positions()

    def delete_row_2(self):
        print("Deleting row...")
        row_count2_2 = self.tableWidget3.rowCount()
        if row_count2_2 > 0:
            self.tableWidget3.setRowCount(row_count2_2 - 1)

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

        self.add_buttons()

    def add_buttons(self):
        for i in range(len(self.buttons)):
            self.buttons[i].setParent(None)
        self.buttons = []
        self.button_rows = []

        for i in range(self.tableWidget2.rowCount()):
            combo = self.tableWidget2.cellWidget(i, 0)
            if combo:
                button = QPushButton(f"Button {i + 1}", self.widget_4)
                button.setFixedSize(100, 30)
                button.move(10, 40 * (i + 1) + self.tableWidget2.height())
                button.clicked.connect(lambda checked, r=i: self.button_clicked(r, combo.currentText()))
                button.show()

                self.buttons.append(button)
                self.button_rows.append(i)

    def remove_button_for_row(self, row):
        for i, button_row in enumerate(self.button_rows):
            if button_row == row:
                self.buttons[i].setParent(None)
                del self.buttons[i]
                del self.button_rows[i]
                break

    def update_button_positions(self):
        for i, button in enumerate(self.buttons):
            button.move(10, 40 * (self.button_rows[i] + 1) + self.tableWidget2.height())

    def button_clicked(self, row, combo_text):
        print(f"Button in row {row + 1} clicked! Selected file key: {combo_text}")
        # Дальнейшая обработка нажатия кнопки с использованием текста из ComboBox

if __name__ == "__main__":
    app = QApplication(sys.argv)
    mainWindow = MainWindow()
    mainWindow.show()
    sys.exit(app.exec())
