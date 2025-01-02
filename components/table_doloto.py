import csv
import sys
from collections import OrderedDict

from PyQt6 import uic
from PyQt6.QtCore import Qt, QStringListModel, pyqtSignal
from PyQt6.QtGui import QAction, QUndoStack, QPixmap
from PyQt6.QtWidgets import QWidget, QScrollArea, QVBoxLayout, QUndoView, QLabel, QTableWidget, QPushButton, \
    QTableWidgetItem, QComboBox, QListView, QLineEdit, QMenu, QStyledItemDelegate, QApplication, QMainWindow, \
    QHeaderView, QSpinBox

from components.adaptive_table import AdaptiveTable
from components.dialogs import CsvTableDialog
from config import IMAGE_PATH, CSV_PATH, BASE_DIR
from components.comands import UpdateTableCommand


class CenteredItemDelegate(QStyledItemDelegate):
    def initStyleOption(self, option, index):
        super().initStyleOption(option, index)
        option.displayAlignment = Qt.AlignmentFlag.AlignCenter


class Doloto_Table(QWidget):
    add_page = pyqtSignal()
    delete_page = pyqtSignal()

    def __init__(self, sort_key=None, parent=None):
        super().__init__(parent)
        uic.loadUi(BASE_DIR / 'table_doloto.ui', self, package='components')
        self.setup_ui()
        self.set_spin_box()

    def setup_ui(self):
        self.tbl_nozzle: QTableWidget = self.findChild(QTableWidget, 'table_nozzle')
        self.label: QLabel = self.findChild(QLabel, 'label')
        self.label.setVisible(False)

    def set_label(self, text):
        self.label.setVisible(True)
        self.label.setText(f"Параметры расчета для {text}")

    def set_spin_box(self):
        row_count = self.tbl_nozzle.rowCount()
        for row in range(row_count):
            spin_box = QSpinBox()
            spin_box.setMinimum(0)
            spin_box.setMaximum(100)
            self.tbl_nozzle.setCellWidget(row, 1, spin_box)


class TestDolotoTable(QMainWindow):
    """ТЕСТОВЫЙ КЛАСС, ЧТОБЫ ЗАПУСКАЛАСЬ KNBK_Table"""

    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        ex = Doloto_Table(parent=self)
        self.setCentralWidget(ex)


# Тестирование KNBK
if __name__ == '__main__':
    app = QApplication(sys.argv)
    widget = TestDolotoTable()
    widget.show()
    sys.exit(app.exec())
