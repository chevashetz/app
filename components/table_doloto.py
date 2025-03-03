import sys

from PyQt6 import uic
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtWidgets import QWidget, QLabel, QTableWidget, QPushButton, \
    QStyledItemDelegate, QApplication, QMainWindow, \
    QSpinBox, QTableWidgetItem

from components.adaptive_table import AdaptiveTable
from config import BASE_DIR


class CenteredItemDelegate(QStyledItemDelegate):
    def initStyleOption(self, option, index):
        super().initStyleOption(option, index)
        option.displayAlignment = Qt.AlignmentFlag.AlignCenter


class Doloto_Table(QWidget):
    add_page = pyqtSignal()
    delete_page = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        uic.loadUi(BASE_DIR / 'table_doloto.ui', self, package='components')
        self.setup_ui()
        self.set_spin_box()

    def setup_ui(self):
        self.tbl_nozzle: AdaptiveTable = self.findChild(AdaptiveTable, 'table_nozzle')
        self.tbl_vzd: AdaptiveTable = self.findChild(AdaptiveTable, 'table_vzd')
        self.tbl_nozzle.calculate_min_column_widths_by_header()
        self.tbl_vzd.calculate_min_column_widths_by_header()
        self.btn_clear_spinbox: QPushButton = self.findChild(QPushButton, 'pushButton_clear_nozzle')
        self.label: QLabel = self.findChild(QLabel, 'label')

        self.btn_clear_spinbox.clicked.connect(self.clear_spinbox)
        self.label.setVisible(False)

        #self.merge_columns(0, 1, 2, text="Интервал")

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
        for row in range(row_count-1):
            spin_box = QSpinBox()
            spin_box.setMinimum(0)
            spin_box.setMaximum(100)
            self.tbl_nozzle.setCellWidget(row, 3, spin_box)

    def clear_spinbox(self):
        row_count = self.tbl_nozzle.rowCount()
        for row in range(row_count):
            spin_box = self.tbl_nozzle.cellWidget(row, 1)
            if isinstance(spin_box, QSpinBox):
                spin_box.setValue(0)

    def merge_columns(self, row, start_col, end_col, text):
        merged_item = QTableWidgetItem(text.strip())
        merged_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        self.tbl_vzd.setItem(row, start_col, merged_item)

        self.tbl_vzd.setSpan(row, start_col, 1, end_col - start_col + 1)

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
