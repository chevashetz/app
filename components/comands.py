from PyQt6.QtCore import Qt
from PyQt6.QtGui import QUndoCommand, QTextDocument
from PyQt6.QtWidgets import QTableWidgetItem


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
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)  # Центрация текста
                self.tableWidget.setItem(row, col, item)
        # Удаление добавленных строк
        for _ in range(self.new_rows_needed):
            self.tableWidget.removeRow(self.tableWidget.rowCount() - 1)

    def redo(self):
        rows = self.text_data
        self.old_data = []

        total_cells_needed = sum(len(row) for row in rows)
        current_cells_available = ((self.tableWidget.rowCount() - self.start_row) * self.tableWidget.columnCount()
                                   - self.start_col)
        self.new_rows_needed = max(0,
                                   (total_cells_needed - current_cells_available + self.tableWidget.columnCount() - 1)
                                   // self.tableWidget.columnCount())

        for _ in range(self.new_rows_needed):
            self.tableWidget.insertRow(self.tableWidget.rowCount())

        doc = QTextDocument()

        current_row = self.start_row
        for row_data in rows:
            columns = row_data
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

                item = self.tableWidget.item(current_row, current_col)
                old_row_data[current_col] = item.text() if item else ""

                item = QTableWidgetItem()
                doc.setHtml(value)
                item.setData(Qt.ItemDataRole.DisplayRole, doc.toPlainText())
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)  # Центрация текста
                self.tableWidget.setItem(current_row, current_col, item)
                current_col += 1

            self.old_data.append((current_row, old_row_data))
            current_row += 1


class UpdateTableCommand(QUndoCommand):
    def __init__(self, knbk_table_instance, old_data, new_data, description="загрузку КНБК"):
        super().__init__(description)
        self.knbk_table_instance = knbk_table_instance
        self.old_data = old_data
        self.new_data = new_data

    def undo(self):
        self.knbk_table_instance.update_table_widget(self.old_data)

    def redo(self):
        self.knbk_table_instance.update_table_widget(self.new_data)

    def update_table_widget(self, data):
        self.knbk_table_instance.clear_images()

        self.knbk_table_instance.tbl_KNBK.clearContents()
        self.knbk_table_instance.tbl_KNBK.setRowCount(len(data))

        for row_index, row_data in enumerate(data):
            for col_index, value in enumerate(row_data):
                if value is None:
                    value = ""
                item = QTableWidgetItem(value)
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.knbk_table_instance.tbl_KNBK.setItem(row_index, col_index, item)
            self.knbk_table_instance.file_key = self.knbk_table_instance.tbl_KNBK.item(row_index, 0).text()
            self.knbk_table_instance.add_image()

        # Восстановление начального состояния
        self.knbk_table_instance.restore_initial_state()
        self.knbk_table_instance.set_label(f"КНБК - {self.knbk_table_instance.tbl_KNBK.item(0, 4).text()} мм")
