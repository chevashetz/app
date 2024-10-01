import csv
import os

from PyQt6 import uic
from PyQt6.QtCore import Qt, QStringListModel
from PyQt6.QtGui import QAction, QUndoStack, QPixmap
from PyQt6.QtWidgets import QWidget, QScrollArea, QVBoxLayout, QUndoView, QLabel, QTableWidget, QPushButton, \
    QTableWidgetItem, QComboBox, QListView, QLineEdit, QMenu, QStyledItemDelegate

from components.dialogs import CsvTableDialog
from config import path3, path1
from components.comands import UpdateTableCommand


class CenteredItemDelegate(QStyledItemDelegate):
    def initStyleOption(self, option, index):
        super().initStyleOption(option, index)
        option.displayAlignment = Qt.AlignmentFlag.AlignCenter
class KNBK_Table(QWidget):
    def __init__(self, sort_key=None, parent=None):
        super(KNBK_Table, self).__init__(parent)
        uic.loadUi('table.ui', self)

        self.sort_key = sort_key
        self.setup_ui()
        self.labels = []
        self.current_y = 700
        self.max_height = 700

        header = self.tbl_KNBK.horizontalHeaderItem(0)
        if header is not None:
            header.setTextAlignment(Qt.AlignmentFlag.AlignCenter)

        self.add_image(mode="static", static_path=path3 + 'Долото.png')

    def setup_ui(self):

        self.open_file_act: QAction = self.findChild(QAction, 'actionOpen')
        self.undo_stack = QUndoStack(self)
        self.undo_view = QUndoView(self.undo_stack)

        self.label: QLabel = self.findChild(QLabel, 'label')
        #self.label_image: QLabel = self.findChild(QLabel, 'label_image')
        self.image_container = self.findChild(QWidget, 'image_container')
        self.scroll_area = self.findChild(QScrollArea, 'scroll_area')
        self.image_container_layout = self.image_container.layout()
        self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        #self.scroll_area.hide()

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

            for label in self.labels[::-1]:
                label.setParent(self.image_container)

                self.image_container_layout.addWidget(label)
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
            "Калибратор спиральный": [path1 + "Калибратор спиральный.csv", path3 + "Калибратор спиральный.png"],
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
            self.set_label(f"КНБК - {self.tbl_KNBK.item(0, 4).text()} мм")

        except Exception as e:
            print(f"Error in update_table_data_list_2: {e}")

    def update_table_widget(self, data):
        self.clear_images()

        self.tbl_KNBK.clearContents()
        self.tbl_KNBK.setRowCount(len(data))

        self.add_image(mode="static", static_path=path3 + 'Долото.png')

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
            self.set_label(f"КНБК - {item.text()} мм")

    def set_label(self, text):
        self.label.setVisible(True)
        self.label.setText(text)

    def center_text_in_item(self, item):
        item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)

    def add_QCombobox_cell_clicked(self, row, column):
        try:
            if row > 0 and column == 0:
                self.delete_image_KNBK(row_count=row)
                self.add_QCombobox(row_count=row, column=column)
        except Exception as e:
            print(f"Произошла ошибка: {e}")
