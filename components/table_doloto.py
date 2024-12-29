import csv
import sys
from collections import OrderedDict

from PyQt6 import uic
from PyQt6.QtCore import Qt, QStringListModel, pyqtSignal
from PyQt6.QtGui import QAction, QUndoStack, QPixmap
from PyQt6.QtWidgets import QWidget, QScrollArea, QVBoxLayout, QUndoView, QLabel, QTableWidget, QPushButton, \
    QTableWidgetItem, QComboBox, QListView, QLineEdit, QMenu, QStyledItemDelegate, QApplication, QMainWindow, \
    QHeaderView

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
        uic.loadUi(BASE_DIR / 'table.ui', self, package='components')

        self.setup_ui()
        self.labels = OrderedDict()



class TestKNBKTable(QMainWindow):
    """ТЕСТОВЫЙ КЛАСС, ЧТОБЫ ЗАПУСКАЛАСЬ KNBK_Table"""
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        ex = KNBK_Table(parent=self)
        self.setCentralWidget(ex)


# Тестирование KNBK
if __name__ == '__main__':
    app = QApplication(sys.argv)
    widget = TestKNBKTable()
    widget.show()
    sys.exit(app.exec())