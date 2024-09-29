import logging
import sys
from pathlib import Path

from PyQt6 import uic
from PyQt6.QtGui import (QAction)
from PyQt6.QtWidgets import (QApplication, QMainWindow, QDialog, QInputDialog, QMenu, QDockWidget, QTreeWidget,
                             QTreeWidgetItem, QFileDialog,
                             )

from components.dialogs import WellDialog, CustDialog
from components.tables import Tables

logging.basicConfig(level=logging.ERROR, format='%(asctime)s - %(levelname)s - %(message)s')


class MainWindow(QMainWindow):
    def __init__(self, *args, **kwargs):
        super(MainWindow, self).__init__(*args, **kwargs)
        uic.loadUi('app.ui', self)
        self.setup_ui()
        self.wellbores = {}
        # Проверка загрузки файла .ui
        print("app.ui loaded successfully")

    def setup_ui(self):

        self.tree_widget: QTreeWidget = self.findChild(QTreeWidget, 'treeWidget')
        self.tree_widget.setHeaderLabels(["Наименование"])

        self.project = QTreeWidgetItem(self.tree_widget, ["Проект"])
        self.fields = QTreeWidgetItem(self.project, ["Месторождения"])
        self.new_field = QTreeWidgetItem(self.fields, ["+"])

        self.tree_widget.itemClicked.connect(self.on_item_clicked_tree)

        # menubar = self.menuBar()
        # file_menu = menubar.addMenu('Файл')

        self.open_file_action = self.findChild(QAction, 'open_file_action')
        self.open_file_action.triggered.connect(self.open_file_all)

        self.save_file_action = self.findChild(QAction, 'save_file_action')
        # self.save_file_action.triggered.connect(self.save_file)
        self.save_file_action.setDisabled(True)

        self.print_action: QAction = self.findChild(QAction, 'print_action')
        # self.print_action.triggered.connect(self.tables.print_report)
        self.print_action.setDisabled(True)

        self.create_action = self.findChild(QAction, 'create_action')
        self.create_action.triggered.connect(self.create)

        self.view_menu = self.findChild(QMenu, 'view_menu')
        self.dockWidget = self.findChild(QDockWidget, 'project_dockWidget')
        self.toggle_dock_act = self.dockWidget.toggleViewAction()
        self.view_menu.addAction(self.toggle_dock_act)

    def open_file_all(self):
        file_path = QFileDialog.getExistingDirectory(self, "Загрузить проект", "")
        if file_path:
            file_path = Path(file_path)
            self.create_tables(file_path.name)
            self.tables.load_all(file_path)

    def create_tables(self, name):
        self.tables = Tables(self)
        self.wellbores[name] = self.tables
        self.setCentralWidget(self.tables)

        self.save_file_action.setDisabled(False)
        self.save_file_action.triggered.connect(lambda _ : self.tables.save_all(name))

        self.print_action.setDisabled(False)
        self.print_action.triggered.connect(self.tables.print_report)

    def create(self):

        wellbore_name, ok = QInputDialog.getText(self, "Проект", "Введите название:")

        if ok and wellbore_name.strip():
            self.create_tables(wellbore_name.strip())

    def handle_value_entered(self, value):
        print(f"Получено значение: {value}")

    def on_item_clicked_tree(self, item, column):
        # Проверяем, по какому элементу кликнули
        if item == self.new_field:
            self.create_new_field()
        elif item.text(0) == "+" and item.parent().text(0) == "Кусты":
            # Если клик на "+" для создания кустов
            parent = item.parent()
            self.create_new_cust(parent)
        elif item.text(0) == "+" and item.parent().text(0) == "Скважины":
            # Если клик на "+" для создания скважин
            parent = item.parent()
            self.create_new_well(parent)

    def create_new_field(self):
        # Используем стандартное диалоговое окно для ввода текста
        field_value, ok = QInputDialog.getText(self, "Добавить месторождение", "Введите название месторождения:")

        if ok and field_value:  # Если пользователь нажал OK и ввел значение
            # Создаем новый элемент "Месторождение"
            new_field_item = QTreeWidgetItem([field_value])

            # Создаем вложенный элемент "Кусты"
            new_custs_item = QTreeWidgetItem(new_field_item, ["Кусты"])
            # Добавляем знак "+" для добавления новых кустов
            new_cust_plus_item = QTreeWidgetItem(new_custs_item, ["+"])

            # Вставляем новый элемент перед "self.new_field"
            self.fields.insertChild(self.fields.indexOfChild(self.new_field), new_field_item)
            self.fields.setExpanded(True)

    def create_new_cust(self, parent_item):
        dialog = CustDialog()
        if dialog.exec() == QDialog.DialogCode.Accepted:
            cust_value = dialog.get_inputs()

            if cust_value:
                # Создаем новый куст
                new_cust_item = QTreeWidgetItem([cust_value])

                # Создаем вложенный элемент "Скважины"
                new_wells_item = QTreeWidgetItem(new_cust_item, ["Скважины"])
                # Добавляем знак "+" для добавления новых скважин
                new_well_plus_item = QTreeWidgetItem(new_wells_item, ["+"])

                # Вставляем перед элементом "+"
                plus_item = parent_item.child(parent_item.childCount() - 1)  # Это элемент с "+"
                parent_item.insertChild(parent_item.indexOfChild(plus_item), new_cust_item)
                parent_item.setExpanded(True)

    def create_new_well(self, parent_item):
        dialog = WellDialog()
        if dialog.exec() == QDialog.DialogCode.Accepted:
            well_value = dialog.get_inputs()

            if well_value:
                # Создаем новую скважину
                new_well_item = QTreeWidgetItem([well_value])

                # Вставляем перед элементом "+"
                plus_item = parent_item.child(parent_item.childCount() - 1)  # Это элемент с "+"
                parent_item.insertChild(parent_item.indexOfChild(plus_item), new_well_item)
                parent_item.setExpanded(True)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    mainWindow = MainWindow()
    mainWindow.show()
    sys.exit(app.exec())
