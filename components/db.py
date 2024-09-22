from PyQt6.QtCore import Qt
from PyQt6.QtSql import QSqlDatabase, QSqlTableModel
from PyQt6.QtWidgets import QMessageBox, QDialog, QVBoxLayout, QTableWidget, QTableWidgetItem, QHeaderView, QHBoxLayout, \
    QPushButton


class DatabaseManager:
    def __init__(self, db_path=None):
        self.database = QSqlDatabase.addDatabase("QSQLITE")
        if db_path:
            self.database.setDatabaseName(db_path)
            self.open_database()

    def open_database(self):
        if not self.database.open():
            QMessageBox.critical(None, "Ошибка", "Не удалось открыть файл базы данных.")
            return False
        return True

    def get_tables(self):
        return self.database.tables()

    def load_table(self, table_name):
        model = QSqlTableModel()
        model.setTable(table_name)
        model.select()

        if model.lastError().isValid():
            QMessageBox.critical(None, "Ошибка", f"Не удалось загрузить данные: {model.lastError().text()}")
            return None
        return model

    def close_database(self):
        self.database.close()

    custom_headers = ["Название", "Индексация", "Верх", "Низ", "Коэффициент ", "Плотность"]

    def preview_table(self, table_name, custom_headers=None, parent=None):
        model = self.load_table(table_name)
        if not model:
            return None

        dialog = QDialog(parent)
        dialog.setWindowTitle(f"Предварительный просмотр - {table_name}")
        dialog.setGeometry(30, 150, 1410, 420)

        layout = QVBoxLayout()

        table_widget = QTableWidget()
        table_widget.setRowCount(model.rowCount())
        table_widget.setColumnCount(model.columnCount() - 1)

        if custom_headers:
            headers = custom_headers
        else:
            headers = [model.headerData(i, Qt.Orientation.Horizontal) for i in
                       range(1, model.columnCount())]

        table_widget.setHorizontalHeaderLabels(headers)

        for row in range(model.rowCount()):
            for col in range(1, model.columnCount()):
                data = model.data(model.index(row, col))
                item = QTableWidgetItem(str(data))
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                table_widget.setItem(row, col - 1, item)

        header = table_widget.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

        layout.addWidget(table_widget)

        button_layout = QHBoxLayout()

        insert_button = QPushButton("Вставить в таблицу")
        return_button = QPushButton("Вернуться к выбору таблицы")

        insert_button.setFixedSize(685, 40)
        return_button.setFixedSize(685, 40)

        button_layout.addWidget(insert_button)
        button_layout.addWidget(return_button)
        layout.addLayout(button_layout)

        dialog.setLayout(layout)

        # Связываем кнопки с действиями
        insert_button.clicked.connect(dialog.accept)
        return_button.clicked.connect(dialog.reject)

        if dialog.exec() == QDialog.DialogCode.Accepted:
            return model
        return None
