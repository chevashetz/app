import csv
import sys

from PyQt6 import uic
from PyQt6.QtCore import Qt, QStringListModel
from PyQt6.QtGui import QAction, QUndoStack, QPixmap
from PyQt6.QtWidgets import QWidget, QScrollArea, QVBoxLayout, QUndoView, QLabel, QTableWidget, QPushButton, \
    QTableWidgetItem, QComboBox, QListView, QLineEdit, QMenu, QStyledItemDelegate, QApplication, QMainWindow

from components.dialogs import CsvTableDialog
from config import IMAGE_PATH, CSV_PATH, BASE_DIR
from components.comands import UpdateTableCommand


class CenteredItemDelegate(QStyledItemDelegate):
    def initStyleOption(self, option, index):
        super().initStyleOption(option, index)
        option.displayAlignment = Qt.AlignmentFlag.AlignCenter


class KNBK_Table(QWidget):
    def __init__(self, sort_key=None, parent=None):
        super().__init__(parent)
        uic.loadUi(BASE_DIR / 'table.ui', self)

        self.sort_key = sort_key
        self.setup_ui()
        self.labels = []

        header = self.tbl_KNBK.horizontalHeaderItem(0)
        if header is not None:
            header.setTextAlignment(Qt.AlignmentFlag.AlignCenter)

        self.add_image(image_path=IMAGE_PATH / 'Долото.png')

    def setup_ui(self):
        self.undo_stack = QUndoStack(self)
        self.undo_view = QUndoView(self.undo_stack)

        self.label: QLabel = self.findChild(QLabel, 'label')
        self.image_container = self.findChild(QWidget, 'image_container')
        self.scroll_area = self.findChild(QScrollArea, 'scroll_area')
        self.image_container_layout: QVBoxLayout = self.image_container.layout()


        self.label.setVisible(False)
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
        new_row = self.tbl_KNBK.rowCount()
        self.tbl_KNBK.setRowCount(new_row + 1)
        for column in range(self.tbl_KNBK.columnCount()):
            if column == 0:
                self.add_QCombobox(row_count=new_row, column=column)
            else:
                item = QTableWidgetItem("")
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.tbl_KNBK.setItem(new_row, column, item)
        self.tbl_KNBK.resizeRowsToContents()


    def add_QCombobox(self, row_count, column):
        combo1 = QComboBox()
        combo1.setModel(QStringListModel([
            "<Не выбрано>", "ВЗД", "РУС", "Бурильные трубы", "Переводник", "УБТ", "Телеметрия", "Ясс",
            "Калибратор спиральный", "Обратный клапан", "Центратор прямой",
            "Центратор спиральный", "Предохранительный переводник"]))

        listView = QListView()

        listView.setWordWrap(True)

        combo1.setView(listView)

        showPopup = combo1.showPopup

        def on_popup_show():
            showPopup()
            combo1.view().reset()  # Сброс состояния представления для предотвращения наслаивания

        combo1.showPopup = on_popup_show

        listView.setItemDelegate(CenteredItemDelegate(listView))
        combo1.setEditable(True)
        line_edit: QLineEdit = combo1.lineEdit()
        line_edit.setAlignment(Qt.AlignmentFlag.AlignCenter)
        line_edit.setReadOnly(False)

        combo1.currentIndexChanged.connect(
            lambda index, combo=combo1, row=row_count: self.handle_combo_change(row, combo))

        self.tbl_KNBK.setCellWidget(row_count, column, combo1)

    def handle_combo_change(self, row, combo):
        text = combo.currentText()
        item = QTableWidgetItem(text)
        item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        self.tbl_KNBK.setItem(row, 0, item)
        self.tbl_KNBK.removeCellWidget(row, 0)
        self.tbl_KNBK.resizeRowsToContents()

        self.add_image(text, row=row)

    def add_image(self, file_key="", image_path=None, row=None):
        if not image_path:
            name = self.get_file_name(file_key)
            if name is None:
                return
            image_path = IMAGE_PATH / name[1]

        label = QLabel()
        pixmap = QPixmap(str(image_path))
        label.setPixmap(pixmap)
        label.setFixedWidth(67)
        label.setScaledContents(True)

        row = row or (self.image_container_layout.count() - 1)
        row = min(row, len(self.labels))

        self.labels.insert(row, label)
        row = self.image_container_layout.count() - row

        self.image_container_layout.insertWidget(row, label, alignment=Qt.AlignmentFlag.AlignBottom)


    def remove_image(self, table_index):
        label_to_remove = self.labels.pop(table_index)
        self.image_container_layout.removeWidget(label_to_remove)
        label_to_remove.deleteLater()

    def delete_row_KNBK(self):
        current_row = self.tbl_KNBK.currentRow()
        if current_row > 0:
            if not isinstance(self.tbl_KNBK.cellWidget(current_row, 0), QComboBox):
                self.remove_image(current_row)
            self.tbl_KNBK.removeRow(current_row)


    def load_table(self):
        dialog = CsvTableDialog(str(CSV_PATH / 'КНБК.csv'), load_table=True, initial_sort_value_KNBK=None,
                                sort_value_casing_srings=self.sort_key, parent=self)
        dialog.data_selected.connect(self.update_table_data_list_2)
        dialog.exec()

    def row_up(self):
        current_row = self.tbl_KNBK.currentRow()
        if current_row > 1:
            self.swap_rows(current_row, current_row - 1)

    def row_down(self):
        current_row = self.tbl_KNBK.currentRow()
        if (current_row < self.tbl_KNBK.rowCount() - 1) and current_row > 0:
            self.swap_rows(current_row, current_row + 1)

    def swap_rows(self, row1, row2):
        if (not isinstance(self.tbl_KNBK.cellWidget(row1, 0), QComboBox) and
                not isinstance(self.tbl_KNBK.cellWidget(row2, 0), QComboBox)):
            for column in range(self.tbl_KNBK.columnCount()):
                item1 = self.tbl_KNBK.takeItem(row1, column)
                item2 = self.tbl_KNBK.takeItem(row2, column)
                if item1:
                    self.tbl_KNBK.setItem(row2, column, item1)
                if item2:
                    self.tbl_KNBK.setItem(row1, column, item2)

            self.tbl_KNBK.setCurrentCell(row2, 0)

            comboBoxCount = 0
            for row in range(row1):
                if isinstance(self.tbl_KNBK.cellWidget(row, 0), QComboBox):
                    comboBoxCount += 1

            row1 -= comboBoxCount
            row2 -= comboBoxCount

            for label in self.labels:
                self.image_container_layout.removeWidget(label)

            self.labels[row1], self.labels[row2] = self.labels[row2], self.labels[row1]

            for label in self.labels:
                self.image_container_layout.insertWidget(1, label)

    def get_file_name(self, file_key):
        files = {
            "ВЗД": ["ВЗД.csv", "ВЗД.png"],
            "РУС": ["РУС.csv", "РУС.png"],
            "Бурильные трубы": ["Бурильные трубы.csv", "Бурильные трубы.png"],
            "Переводник": ["Переводник.csv", "Переводник.png"],
            "Предохранительный переводник": ["Предохранительный переводник.csv",
                                             "Предохранительный переводник.png"],
            "Обратный клапан": ["Обратный клапан.csv", "Обратный клапан.png"],
            "Ясс": ["Ясс.csv", "Ясс.png"],
            "Калибратор спиральный": ["Калибратор спиральный.csv", "Калибратор спиральный.png"],
            "УБТ": ["УБТ.csv", "УБТ.png"],
            "Телеметрия": ["Телеметрия.csv", "Телеметрия.png"]
        }

        return files.get(file_key)

    def open_csv_table_dialog(self, row, column):
        try:
            if column == 1 and row != 0:
                file_name = self.get_file_name("")
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
            fixed_file_path = str(CSV_PATH / "Долото.csv")
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
        fixed_file_path = str(CSV_PATH / "КНБК.csv")
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
            self.set_label(f"КНБК - {self.tbl_KNBK.item(0, 4).text()} мм")

        except Exception as e:
            print(f"Error in update_table_data_list_2: {e}")

    def update_table_widget(self, data):
        self.clear_images()

        self.tbl_KNBK.clearContents()
        self.tbl_KNBK.setRowCount(len(data))

        self.add_image(image_path=str(IMAGE_PATH / 'Долото.png'))

        for row_index, row_data in enumerate(data):
            for col_index, value in enumerate(row_data):
                if value is None:
                    value = ""
                item = QTableWidgetItem(value)
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.tbl_KNBK.setItem(row_index, col_index, item)
            file_key = self.tbl_KNBK.item(row_index, 0).text()
            self.add_image(file_key)

    def clear_images(self):
        for label in self.labels:
            self.image_container_layout.removeWidget(label)
            label.deleteLater()

    def restore_initial_state(self):
        item_0_0_KNBK = QTableWidgetItem("Долото")
        item_0_0_KNBK.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        self.tbl_KNBK.setItem(0, 0, item_0_0_KNBK)

    def update_label(self, item):
        if item.row() == 0 and item.column() == 4:
            self.set_label(f"КНБК - {item.text()} мм")

    def set_label(self, text):
        self.label.setVisible(True)
        self.label.setText(text)

    def center_text_in_item(self, item):
        item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)

    def add_QCombobox_cell_clicked(self, row, column):
        try:
            if row > 0 and column == 0:
                self.remove_image(row)
                self.add_QCombobox(row_count=row, column=column)
        except Exception as e:
            print(f"Произошла ошибка: {e}")


class TestKNBKTable(QMainWindow):
    """ТЕСТОВЫЙ КЛАСС, ЧТОБЫ ЗАПУСКАЛАСЬ KNBK_Table"""
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        ex = KNBK_Table(parent=self)
        self.setCentralWidget(ex)

    def add_page_2(self):
        pass

    def delete_page(self):
        pass


# Тестирование KNBK
if __name__ == '__main__':
    app = QApplication(sys.argv)
    widget = TestKNBKTable()
    widget.show()
    sys.exit(app.exec())
