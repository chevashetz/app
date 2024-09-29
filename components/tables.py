
import logging
import os
import re
import shutil
from pathlib import Path

import numpy as np
import pandas as pd
import xlsxwriter
from PyQt6 import uic
from PyQt6.QtCore import Qt, QRectF, pyqtSignal
from PyQt6.QtGui import (QAction, QKeySequence, QPainter, QPen,
                         QBrush, QUndoStack)
from PyQt6.QtSql import QSqlDatabase, QSqlQuery
from PyQt6.QtWidgets import (QApplication, QLineEdit, QPushButton, QHeaderView,
                             QTableWidget, QTableWidgetItem, QComboBox, QFileDialog, QInputDialog, QMenu,
                             QGraphicsScene, QGraphicsView, QMessageBox, QWidget, QStackedWidget, QUndoView, QDialog,
                             )
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

from components.db import DatabaseManager
from components.dialogs import DualInputDialog
from components.shading_drawer import ShadingDrawer
from components.table_knbk import KNBK_Table
from config import path2
from components.comands import PasteCommand


class MplCanvas(FigureCanvas):
    def __init__(self, parent=None, width=4.5, height=1.5, dpi=100):
        self.fig = Figure(figsize=(width, height), dpi=dpi)
        self.axes = self.fig.add_subplot(111)
        super().__init__(self.fig)
        self.setParent(parent)
        self.setFixedSize(int(width * dpi), int(height * dpi))
class Tables(QWidget):
    def __init__(self, parent=None):
        super(Tables, self).__init__(parent)
        uic.loadUi('tables.ui', self)
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
        self.stackedWidget.setCurrentIndex(0)
        self.stackedWidget.insertWidget(4, KNBK_Table(parent=self))

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
        self.btn_load_profile: QPushButton = self.findChild(QPushButton, 'pushButton_load_profile')

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
        self.btn_load_profile.clicked.connect(self.open_file)

        self.tbl_profile.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.tbl_stratigraphy.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

        header = self.tbl_casing_strings.horizontalHeader()

        for column in range(self.tbl_casing_strings.columnCount() - 1):
            header.setSectionResizeMode(column, QHeaderView.ResizeMode.ResizeToContents)

        header.setSectionResizeMode(self.tbl_casing_strings.columnCount() - 1, QHeaderView.ResizeMode.Stretch)

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
        self.merge_columns_3(0, 8, 11)
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
        if current_column_count != len(headers):
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
            selected_data = self.calculate_coords(data)
            self.plot_graph(selected_data)

            if self.is_row_complete(row, [0, 1, 2], self.tbl_profile):
                '''
                if self.tbl_profile.columnCount() == 4:
                    self.tbl_profile.insertColumn(4)
                    self.tbl_profile.insertColumn(5)
                    self.tbl_profile.insertColumn(6)
                '''

                if data.shape[1] < 7:
                    start_column = 3
                else:
                    start_column = 5
                self.calculate_additional_columns(selected_data, row, start_column)


    def calculate_additional_columns(self, selected_data, row, start_column=5):
        delta_z = selected_data[row][2]
        delta_y = selected_data[row][1]
        delta_x = selected_data[row][0]
        self.vertical_depth_and_vertical_deviation_on_item_changed(row, delta_z, start_column)
        L_current = self.extract_number(self.tbl_profile.item(row, 0).text())
        L_prev = self.extract_number(self.tbl_profile.item(row - 1 if row > 0 else row, 0).text())
        delta_L = L_current - L_prev
        sum_proection = (delta_y ** 2 + delta_x ** 2) ** 0.5
        self.tbl_profile.setItem(row, start_column + 1, QTableWidgetItem(str(round(sum_proection, 2))))
        if delta_L != 0:
            self.tbl_profile.setItem(row, start_column + 2, QTableWidgetItem(str(round(sum_proection / delta_L, 2))))


    def vertical_depth_and_vertical_deviation_on_item_changed(self, row, delta_z, delta_x, delta_y, start_column):
        delta_z_old = self.extract_number(self.tbl_profile.item(row, start_column).text())
        delta_delta = delta_z - delta_z_old
        self.tbl_profile.setItem(row, start_column, QTableWidgetItem(str(round(delta_z, 2))))
        for row_i in range(row + 1, self.tbl_profile.rowCount()):
            old_value = self.extract_number(self.tbl_profile.item(row_i, start_column).text())
            self.tbl_profile.item(row_i, start_column).setText(str(round(old_value + delta_delta, 2)))
        delta_x_old = self.extract_number(self.tbl_profile.item(row, start_column).text())
        delta_y_old = self.extract_number(self.tbl_profile.item(row, start_column).text())
        delta_delta = delta_z - delta_z_old
        self.tbl_profile.setItem(row, start_column, QTableWidgetItem(str(round(delta_z, 2))))


    def process_excel_data(self, data):
        try:
            '''
            if self.has_headers(data):
                data.columns = data.iloc[0]
                data = data[1:].reset_index(drop=True)
            '''
            if data.shape[1] < 3:
                raise ValueError("The file must have at least 3 columns for calculations.")

            if data.shape[1] < 4:
                vertical_depths = self.calculate_vertical_depth(data)
                data['Глубина по верт(м)'] = vertical_depths

            vertical_deviation = self.calculate_vertical_deviation(data)
            data['Отход от верт(м)'] = vertical_deviation
            intensity_curvature = self.calculate_intensity_curvature(data)
            data['Интенсивность искревления'] = intensity_curvature

            self.populate_table(data)
            selected_data = self.calculate_coords(data.to_numpy())
            self.plot_graph(selected_data)
        except Exception as e:
            logging.error(f"Error processing Excel data: {e}")


    '''
    def has_headers(self, df):
        first_row = df.iloc[0]
        return all(isinstance(val, str) for val in first_row) and len(set(first_row)) == len(first_row)
    '''


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


    def calculate_vertical_deviation(self, df):
        vertical_deviation = [0.0]

        for i in range(1, len(df)):
            delta_L = df.iloc[i, 0] - df.iloc[i - 1, 0]

            current_zenith_angle = np.radians(df.iloc[i, 1])
            delta_vertical_deviation = delta_L * np.sin(current_zenith_angle)
            '''
            azimuth_angle = np.radians(df.iloc[i, 2])
    
            # Normalize azimuth angle to avoid large changes due to wrapping
    
            # Calculate deviations in x and y directions
            delta_x = delta_L * np.sin(zenith_angle) * np.cos(azimuth_angle)
            delta_y = delta_L * np.sin(zenith_angle) * np.sin(azimuth_angle)
    
            # Calculate vertical deviation (assuming only horizontal deviations are needed)
            delta_vertical_deviation = np.sqrt(delta_x ** 2 + delta_y ** 2)
            '''
            vertical_deviation.append(vertical_deviation[-1] + delta_vertical_deviation)

        return vertical_deviation


    def calculate_vertical_depth(self, df):
        vertical_depth = [0.0]
        for i in range(1, len(df)):
            delta_L = df.iloc[i, 0] - df.iloc[i - 1, 0]
            current_zenith_angle = np.radians(df.iloc[i, 1])
            delta_z = delta_L * np.cos(current_zenith_angle)
            vertical_depth.append(vertical_depth[-1] + delta_z)
        return vertical_depth


    def calculate_intensity_curvature(self, df):
        intensity_curvature = [0.0]
        for i in range(1, len(df)):
            delta_L = df.iloc[i, 0] - df.iloc[i - 1, 0]
            delta_deviation = df.iloc[i, 4] - df.iloc[i - 1, 4]
            value_intensity_curvature = delta_deviation / delta_L
            intensity_curvature.append(value_intensity_curvature)
        return intensity_curvature


    def populate_table(self, df):
        self.tbl_profile.blockSignals(True)
        num_rows, num_cols = df.shape
        self.tbl_profile.setRowCount(num_rows)
        # self.tbl_profile.setColumnCount(num_cols)
        '''
        headers = ["Глубина по стволу (м)", "Зенитный угол (град)", "Азимут (град)", "Азимут маг(град)",
                   "Азимут дир(град)", "Глубина по верт(м)"]
        '''
        if len(df.columns) == 6:
            text = ""
            header_item = QTableWidgetItem(text)
            self.tbl_profile.setHorizontalHeaderItem(2, header_item)
            header = ComboHeader(self)
            self.tbl_profile.setHorizontalHeader(header)
            header.valueEntered.connect(self.handle_value_entered)
            text = "Глубина по верт(м)"
            header_item = QTableWidgetItem(text)
            self.tbl_profile.setHorizontalHeaderItem(3, header_item)
            self.tbl_profile.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
            text = "Отклонение от верт(м)"
            header_item = QTableWidgetItem(text)
            self.tbl_profile.setHorizontalHeaderItem(4, header_item)
            self.tbl_profile.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
            text = "Интенсивность искривления"
            header_item = QTableWidgetItem(text)
            self.tbl_profile.setHorizontalHeaderItem(5, header_item)
            self.tbl_profile.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

        elif len(df.columns) == 8:
            self.tbl_profile.insertColumn(6)
            self.tbl_profile.insertColumn(7)
            text = "Отклонение от верт(м)"
            header_item = QTableWidgetItem(text)
            self.tbl_profile.setHorizontalHeaderItem(6, header_item)
            self.tbl_profile.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
            text = "Интенсивность искривления"
            header_item = QTableWidgetItem(text)
            self.tbl_profile.setHorizontalHeaderItem(7, header_item)
            self.tbl_profile.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        '''
        if self.has_headers(df):
            headers = df.columns.astype(str).tolist()
            logging.info(f"Headers from file: {headers}")
            self.tbl_profile.setHorizontalHeaderLabels(headers)
        '''
        for index, row in df.iterrows():
            for col_index, value in enumerate(row):
                item = QTableWidgetItem(str(round(value, 2)))
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.tbl_profile.setItem(index, col_index, item)
        self.tbl_profile.blockSignals(False)
        self.tbl_profile.horizontalHeader().setVisible(True)


    def calculate_coords(self, data: np.array):
        current_zenith_angle = np.radians(0)
        current_azimuth_angle = np.radians(0)
        selected_data = np.zeros((1, 3))
        for i in range(1, len(data)):
            delta_L = data[i, 0] - data[i - 1, 0]
            delta_zenith_angle = np.radians(data[i, 1] - data[i - 1, 1])
            delta_azimuth_angle = np.radians(data[i, 2] - data[i - 1, 2])

            next_zenith_angle = current_zenith_angle + delta_zenith_angle
            next_azimuth_angle = current_azimuth_angle + delta_azimuth_angle

            delta_x = delta_L * np.sin(next_zenith_angle) * np.cos(next_azimuth_angle)
            delta_y = delta_L * np.sin(next_zenith_angle) * np.sin(next_azimuth_angle)
            delta_z = delta_L * np.cos(next_zenith_angle)

            selected_data = np.vstack((selected_data, np.array([delta_x, delta_y, delta_z])))

            current_zenith_angle = next_zenith_angle
            current_azimuth_angle = next_azimuth_angle
        selected_data = np.cumsum(selected_data, axis=0)
        return selected_data


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
        row = 0
        for row in range(row_count):
            item = self.tbl_casing_strings.item(row, 1)
            if item is None or item.text() == "":
                break
            end = self.extract_number(
                self.tbl_casing_strings.item(row, 1).text())
            if row > 0 and end > last_end:
                total_height = self.extract_number(
                    self.tbl_casing_strings.item(row - 1, 1).text()) + 2 * vertical_padding
            else:
                total_height = last_end
            last_end = end

        for row in range(row_count):
            if not self.is_row_complete(row, [1, 2, 3, 4], self.tbl_casing_strings):
                break
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
        if row != 0:
            ShadingDrawer(lengths, ends, diameter_offsets, first_diameter_hole, x_offset, scene).draw_curve()
            ShadingDrawer(lengths, ends, diameter_offsets, first_diameter_hole, x_offset, scene,
                          reverse=True).draw_curve()
        self.graphicsView_casing_strings.setScene(scene)
        self.graphicsView_casing_strings.fitInView(scene.sceneRect(), Qt.AspectRatioMode.KeepAspectRatio)
        self.graphicsView_casing_strings.update()


    def form_fluids(self):
        row_count_casing_strings = self.tbl_casing_strings.rowCount()

        grouped_rows = {}

        for row in range(row_count_casing_strings):
            try:
                fourth_column_item = self.tbl_casing_strings.item(row, 4)
                if fourth_column_item:
                    fourth_value = fourth_column_item.text().strip().lower()

                    first_column_item = self.tbl_casing_strings.item(row, 1)
                    first_value = self.extract_number(first_column_item.text()) if first_column_item else None

                    if first_value is not None:
                        if fourth_value not in grouped_rows:
                            grouped_rows[fourth_value] = []

                        grouped_rows[fourth_value].append((row, first_value))
            except Exception as e:
                print(f"Error processing row {row}: {e}")

        sorted_groups = sorted(grouped_rows.items(), key=lambda x: x[1][0][0])

        self.tbl_drilling_fluids.setRowCount(len(sorted_groups))

        drilling_fluids_row = 0
        previous_depth_to = None

        # Перебираем каждую группу
        for group_index, (group_key, rows) in enumerate(sorted_groups):
            rows.sort(key=lambda x: x[1],
                      reverse=True)

            max_first_value = max(x[1] for x in rows)
            sum_second_values = sum(self.extract_number(self.tbl_casing_strings.item(row, 2).text()) for row, _ in rows)

            if group_index == 0:
                depth_from = max_first_value - sum_second_values
                depth_to = max_first_value

            else:
                depth_from = previous_depth_to

                depth_to = max_first_value

            drilling_fluids_item_from = QTableWidgetItem(str(depth_from))
            drilling_fluids_item_from.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.tbl_drilling_fluids.setItem(drilling_fluids_row, 1, drilling_fluids_item_from)

            drilling_fluids_item_to = QTableWidgetItem(str(depth_to))
            drilling_fluids_item_to.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.tbl_drilling_fluids.setItem(drilling_fluids_row, 2, drilling_fluids_item_to)

            previous_depth_to = depth_to

            drilling_fluids_row += 1


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

        unique_values = set()

        existing_pages = set()
        for i in range(self.stackedWidget.count()):
            existing_pages.add(self.stackedWidget.widget(i).objectName())

        for i in range(number_of_pages):
            item = self.tbl_casing_strings.item(i, 4)
            if item:
                value = self.extract_number(item.text())
            else:
                value = "unknown"

            if i == 0:
                text = f"КНБК - {value} мм"
                if text in existing_pages:
                    continue
                else:
                    unique_values.add(value)

            elif value not in unique_values and f"КНБК - {value} мм" not in existing_pages:
                unique_values.add(value)
                text = f"КНБК - {value} мм"
                self.add_page(text, value)

                new_page = self.stackedWidget.widget(self.stackedWidget.count() - 1)
                new_page.setObjectName(text)

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
            initial_page.set_label(text_0)
            initial_page.sort_key = sort_key_0

    def add_page(self, text, sort_key=None):
        if self.stackedWidget is None:
            return

        num_pages = self.stackedWidget.count()
        insert_index = num_pages - 1

        new_page = KNBK_Table(sort_key=sort_key, parent=self)
        new_page.set_label(text)
        new_page.sort_key = sort_key
        self.stackedWidget.insertWidget(insert_index, new_page)
        self.stackedWidget.setCurrentWidget(new_page)

    def add_page_2(self):
        if self.stackedWidget is None:
            return
        num_pages = self.stackedWidget.count()
        insert_index = num_pages - 1
        print(f"Current number of pages: {num_pages}, Inserting at index: {insert_index}")
        new_page = KNBK_Table(parent=self)
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

        if ok and stratigraphy_name.strip():
            database = QSqlDatabase.addDatabase("QSQLITE")
            database.setDatabaseName(path2 + "stratigraphy.db")
            if not database.open():
                QMessageBox.critical(self, "Ошибка", "Не удалось открыть файл базы данных.")
                return

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

    def print_report(self):
        file_path, _ = QFileDialog.getSaveFileName(self, "Сохранить отчет", "", "Excel Files (*.xlsx)")
        if file_path:
            if os.path.exists(file_path):
                os.remove(file_path)
            resources_dir = Path(file_path).parent / "resources"
            if resources_dir.exists():
                shutil.rmtree(resources_dir)

            workbook = xlsxwriter.Workbook(file_path)

            self.export_page(workbook, self.tbl_profile, "Профиль", resources_dir, self.graphicsView_profile)
            self.export_page(workbook, self.tbl_stratigraphy, "Стратиграфия", resources_dir)
            self.export_page(workbook, self.tbl_pressure, "Давления", resources_dir, self.graphicsView_pressure,
                             self.graphicsView_gradient_pressure)
            self.export_page(workbook, self.tbl_casing_strings, "Обсадные колонны", resources_dir, self.graphicsView_casing_strings)
            self.export_page(workbook, self.tbl_drilling_fluids, "Буровые растворы", resources_dir)
            self.export_page(workbook, self.stackedWidget.widget(4).tbl_KNBK, "КНБК", resources_dir)

            workbook.close()
            QMessageBox.information(self, "Успех", "Отчет успешно сохранен.")

    def export_page(self, workbook, table, sheet_name, resources_dir, *graphics_views):
        worksheet = workbook.add_worksheet(sheet_name)

        center_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 2})

        # Экспорт горизонтальных заголовков
        for col in range(table.columnCount()):
            header_item = table.horizontalHeaderItem(col)
            if header_item:
                worksheet.write(0, col + 1, header_item.text(), center_format)

        # # Экспорт вертикальных заголовков
        # for row in range(table.rowCount()):
        #     header_item = table.verticalHeaderItem(row)
        #     if header_item:
        #         worksheet.write(row + 1, 0, header_item.text(), center_format)

        # Экспорт данных таблицы
        merged_cells = []

        for row in range(table.rowCount()):
            col = 0
            while col < table.columnCount():
                cell_widget = table.cellWidget(row, col)
                if isinstance(cell_widget, QComboBox):
                    value = cell_widget.currentText()
                else:
                    item = table.item(row, col)
                    value = item.text() if item else ""

                # Проверка на объединенные ячейки в col
                col_span = table.columnSpan(row, col)
                if col_span != 1:
                    merged_cells.append((row + 1, col + 1, row, col + col_span, value))
                    col += col_span
                else:
                    worksheet.write(row + 1, col + 1, value, center_format)
                    col += 1

        # Объединение ячеек
        for start_row, start_col, end_row, end_col, value in merged_cells:
            worksheet.merge_range(start_row, start_col, end_row, end_col, value, center_format)

        worksheet.autofit()

        resources_dir.mkdir(parents=True, exist_ok=True)
        # Экспорт графиков
        for i, view in enumerate(graphics_views):
            filename = f"resources/{sheet_name}_{i}.png"
            self.get_image_from_graphics_view(view, filename)
            worksheet.insert_image(0, table.columnCount() + 3 + i * 10, filename)

    def get_image_from_graphics_view(self, graphics_view, filename: str):
        pixmap = graphics_view.grab()
        pixmap.save(filename, "PNG")

    def disable_editing_for_rows(self):
        for row in [0, 1]:
            for column in range(self.tbl_pressure.columnCount()):
                item = self.tbl_pressure.item(row, column)
                if item:
                    item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)

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


class ComboHeader(QHeaderView):
    valueEntered = pyqtSignal(str)
    def __init__(self, parent=None):
        super(ComboHeader, self).__init__(Qt.Orientation.Horizontal, parent)
        self.setStretchLastSection(True)
        self.combobox = QComboBox(self)
        self.combobox.addItems(["Азимут (град)", "Азимут маг(град)", "Азимут дир(град)"])
        self.combobox.setStyleSheet("QComboBox { text-align: center; }")
        for i in range(self.combobox.count()):
            self.combobox.setItemData(i, Qt.AlignmentFlag.AlignCenter, Qt.ItemDataRole.TextAlignmentRole)
        self.setSectionsClickable(True)
        self.combobox.currentIndexChanged.connect(self.on_combobox_header_changed)

    def on_combobox_header_changed(self, index):
        if index == 1:
            text, ok = QInputDialog.getText(self, "Магнитный угол", "Введите значение:")
            if ok and text:
                self.valueEntered.emit(text)
        elif index == 2:
            dialog = DualInputDialog()
            if dialog.exec() == QDialog.DialogCode.Accepted:  # Если нажата кнопка ОК
                first_value, second_value = dialog.get_inputs()
                print(f"Первое значение: {first_value}, Второе значение: {second_value}")
    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self.combobox:
            index = 2
            x = self.sectionViewportPosition(index)
            w = self.sectionSize(index)
            self.combobox.setGeometry(x, 0, w, self.height())
