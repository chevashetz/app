import sys
from datetime import date

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFontMetrics
from PyQt6.QtWidgets import QWidget, QVBoxLayout, QTableWidget, QTableWidgetItem, QApplication


class AdaptiveTable(QTableWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.min_column_widths = []
        self.setup_table()

    def setup_table(self):
        header = self.horizontalHeader()

        self.horizontalScrollBar().valueChanged.connect(self.adjust_columns)

    def adjust_columns(self):
        width = self.viewport().width()
        total = sum(self.min_column_widths)


        for column in range(self.columnCount()):
            column_width = self.min_column_widths[column] if column < len(self.min_column_widths) else 100
            fraction = column_width / total
            self.setColumnWidth(column, int((width if width > total else total) * fraction))

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.adjust_columns()

    def calculate_min_column_widths_by_header(self):
        font_metrics = QFontMetrics(self.font())
        self.min_column_widths = []
        for col in range(self.columnCount()):
            header_width = font_metrics.horizontalAdvance(self.horizontalHeaderItem(col).text())
            self.min_column_widths.append(header_width + 20)
            # self.setColumnWidth(col, header_width + 20)
        # self.setMinimumWidth(sum(self.min_column_widths))

    def calculate_min_column_widths(self, row):
        font_metrics = QFontMetrics(self.font())
        self.min_column_widths = []
        for col in range(self.columnCount()):
            item = self.item(row, col)
            header_width = font_metrics.horizontalAdvance(sorted(item.text().split('\n'), key=len, reverse=True)[0]) if item else 10
            self.min_column_widths.append(header_width + 20)

    def set_min_column_widths(self, index, min_width):
        self.min_column_widths[index] = min_width

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        # Create table
        self.table = AdaptiveTable()
        self.setup_table()

        layout.addWidget(self.table)
        self.setLayout(layout)

        self.setGeometry(300, 300, 400, 400)
        self.setWindowTitle('Simple Adaptive Table Example')

    def setup_table(self):
        # Define headers and data
        headers = ["Name", "Age", "City", "Registration Date", "Balance"]
        data = [
            ["John Doe", 30, "New York", date(2022, 5, 15), 1500.75],
            ["Jane Smith", 28, "Los Angeles", date(2023, 1, 10), 2750.50],
            ["Bob Johnson", 35, "Chicago", date(2021, 11, 3), 500.25],
            ["Alice Brown", 22, "Houston", date(2023, 8, 22), 3000.00],
            ["Charlie Davis", 40, "San Francisco", date(2022, 3, 7), 1250.60]
        ]

        # Set up table structure
        self.table.setColumnCount(len(headers))
        self.table.setRowCount(len(data))
        self.table.setHorizontalHeaderLabels(headers)
        self.table.calculate_min_column_widths_by_header()

        # Populate table with data
        for row, row_data in enumerate(data):
            for col, value in enumerate(row_data):
                item = QTableWidgetItem(str(value))
                if isinstance(value, (int, float)):
                    item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                self.table.setItem(row, col, item)




if __name__ == '__main__':
    app = QApplication(sys.argv)
    mainWin = MainWindow()
    mainWin.show()
    sys.exit(app.exec())