import logging
import sys
from pathlib import Path

from PyQt6 import uic
from PyQt6.QtCore import Qt
from PyQt6.QtGui import (QAction)
from PyQt6.QtWidgets import (QApplication, QMainWindow, QDialog, QInputDialog, QMenu, QDockWidget, QTreeWidget,
                             QTreeWidgetItem, QFileDialog, QTabWidget, )

from components.dialogs import WellDialog, CustDialog, WellboreDialog
from components.tables import Tables

logging.basicConfig(level=logging.ERROR, format='%(asctime)s - %(levelname)s - %(message)s')


class MainWindow(QMainWindow):
    def __init__(self, *args, **kwargs):
        super(MainWindow, self).__init__(*args, **kwargs)
        uic.loadUi('app.ui', self)

        self.setup_ui()
        self.setup_actions()

        self.wellbores = {}
        self.undo_stack = []
        self.redo_stack = []
        # Проверка загрузки файла .ui
        print("app.ui loaded successfully")

    def setup_ui(self):

        self.tree_widget: QTreeWidget = self.findChild(QTreeWidget, 'treeWidget')
        self.tab_widget: QTabWidget = self.findChild(QTabWidget, 'tabWidget')

        self.tree_widget.setHeaderLabels(["Наименование"])

        self.project = QTreeWidgetItem(self.tree_widget, ["Проект"])
        self.fields = QTreeWidgetItem(self.project, ["Месторождения"])
        self.new_field = QTreeWidgetItem(self.fields, ["+"])

        self.tree_widget.itemClicked.connect(self.on_item_clicked_tree)

        self.view_menu = self.findChild(QMenu, 'view_menu')
        self.dockWidget = self.findChild(QDockWidget, 'project_dockWidget')
        self.toggle_dock_act = self.dockWidget.toggleViewAction()
        self.view_menu.addAction(self.toggle_dock_act)
        # Установка контекстного меню
        self.tree_widget.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.tree_widget.customContextMenuRequested.connect(self.open_context_menu)

    def setup_actions(self):
        self.open_file_action = self.findChild(QAction, 'open_file_action')
        self.open_file_action.triggered.connect(self.open_file_all)
        self.save_file_action = self.findChild(QAction, 'save_file_action')
        self.save_file_action.setDisabled(True)
        self.print_action: QAction = self.findChild(QAction, 'print_action')
        self.print_action.setDisabled(True)
        self.create_action = self.findChild(QAction, 'create_action')
        self.create_action.triggered.connect(self.create_project)

    def open_context_menu(self, position):
        item = self.tree_widget.itemAt(position)
        if item:
            # Создание контекстного меню
            menu = QMenu()

            # Создаем действия Undo и Redo
            undo_action = QAction("Отменить последнее действие", self)
            undo_action.triggered.connect(self.undo_last_action)
            undo_action.setEnabled(bool(self.undo_stack))  # Активируем, только если есть что отменять

            redo_action = QAction("Повторить действие", self)
            redo_action.triggered.connect(self.redo_last_action)
            redo_action.setEnabled(bool(self.redo_stack))  # Активируем, только если есть что повторять

            # Добавляем действия Undo и Redo в меню
            menu.addAction(undo_action)
            menu.addAction(redo_action)

            # Добавление действий "Редактировать" и "Удалить"
            edit_action = QAction("Редактировать", self)
            delete_action = QAction("Удалить", self)

            # Пример действий
            edit_action.triggered.connect(lambda: self.edit_item(item))
            delete_action.triggered.connect(lambda: self.delete_item(item))

            # Добавляем действия в меню
            menu.addAction(edit_action)
            menu.addAction(delete_action)

            # Отображаем меню
            menu.exec(self.tree_widget.viewport().mapToGlobal(position))

    def edit_item(self, item):
        # Названия элементов, которые нельзя редактировать
        non_editable_items = ["Проект", "Месторождения", "Кусты"]

        # Проверяем, можно ли редактировать элемент
        if item.text(0) in non_editable_items:
            print(f"Нельзя редактировать элемент: {item.text(0)}")
        else:
            print(f"Редактирование элемента: {item.text(0)}")

            # Открываем диалоговое окно для редактирования текста элемента
            new_text, ok = QInputDialog.getText(self, "Редактирование элемента",
                                                "Введите новое имя:", text=item.text(0))

            if ok and new_text:
                item.setText(0, new_text)
                print(f"Элемент изменен на: {new_text}")

    def undo_last_action(self):
        if self.undo_stack:
            last_action = self.undo_stack.pop()
            action_type, item, parent = last_action

            if action_type == 'delete':
                # Если элемент был верхнеуровневым, добавляем его обратно в QTreeWidget
                if isinstance(parent, QTreeWidget):
                    parent.addTopLevelItem(item)
                    print(f"Отменено удаление верхнеуровневого элемента: {item.text(0)}")
                else:
                    # Если элемент был дочерним, добавляем его обратно к родительскому элементу
                    parent.addChild(item)
                    print(f"Отменено удаление дочернего элемента: {item.text(0)}")

                self.redo_stack.append(last_action)  # Добавляем действие в стек redo

    def redo_last_action(self):
        if self.redo_stack:
            last_action = self.redo_stack.pop()
            action_type, item, parent = last_action

            if action_type == 'delete':
                # Если элемент был верхнеуровневым
                if isinstance(parent, QTreeWidget):
                    index = parent.indexOfTopLevelItem(item)
                    if index != -1:
                        parent.takeTopLevelItem(index)
                        print(f"Действие повторено: Удаление верхнеуровневого элемента {item.text(0)}")
                else:
                    # Если элемент был дочерним
                    parent.removeChild(item)
                    print(f"Действие повторено: Удаление дочернего элемента {item.text(0)}")

                self.undo_stack.append(last_action)  # Возвращаем действие в undo-стек

    def delete_item(self, item):
        parent = item.parent()
        if parent is None:
            # Если элемент верхнеуровневый
            index = self.tree_widget.indexOfTopLevelItem(item)
            if index != -1:
                self.undo_stack.append(('delete', item, self.tree_widget))
                self.tree_widget.takeTopLevelItem(index)
                print(f"Элемент удален: {item.text(0)}")
                self.redo_stack.clear()  # Очищаем стек redo при новом действии
            else:
                print(f"Не удалось удалить элемент: {item.text(0)}")
        else:
            # Если элемент является дочерним элементом
            parent.removeChild(item)
            self.undo_stack.append(('delete', item, parent))
            print(f"Элемент удален: {item.text(0)}")
            self.redo_stack.clear()  # Очищаем стек redo при новом действии

    def on_item_clicked_tree(self, item, column):
        # Обработка клика на элементе дерева
        print(f"Клик на элементе: {item.text(column)}")

    def open_file_all(self):
        # Пример функции для открытия файла
        print("Файл открыт")

    def open_file_all(self):
        file_path = QFileDialog.getExistingDirectory(self, "Загрузить проект", "")
        if file_path:
            file_path = Path(file_path)
            self.create_tables(file_path.name)
            self.tables.load_all(file_path)

    def create_tables(self, name):
        self.tables = Tables(self)
        self.wellbores[name] = self.tables
        self.tab_widget.addTab(self.tables, name)
        self.set_current_tab()

        self.save_file_action.setDisabled(False)
        self.save_file_action.triggered.connect(lambda _: self.tables.save_all(name))

        self.print_action.setDisabled(False)
        self.print_action.triggered.connect(self.tables.print_report)

    def create_project(self):

        wellbore_name, ok = QInputDialog.getText(self, "Проект", "Введите название:")

        if ok and wellbore_name.strip():
            self.create_tables(wellbore_name.strip())

    def handle_value_entered(self, value):
        print(f"Получено значение: {value}")

    def on_item_clicked_tree(self, item, column):
        # Проверяем, по какому элементу кликнули
        if item.parent() is None:  # Корневой элемент "Проект"
            return
        elif item == self.new_field:
            self.create_new_item("Месторождение", self.create_new_field)
        elif item.text(0) == "+":
            parent_text = item.parent().text(0)
            if parent_text == "Кусты":
                self.create_new_item("Куст", self.create_new_cust, item.parent())
            elif parent_text == "Скважины":
                self.create_new_item("Скважина", self.create_new_well, item.parent())
            elif parent_text == "Стволы":
                self.create_new_item("Ствол", self.create_new_wellbore, item.parent())
        elif item.parent().text(0) == "Стволы":
            name = item.text(0)
            self.tables = self.wellbores[name]
            self.set_current_tab()

    def create_new_item(self, dialog_title, create_function, parent_item=None):
        dialog_class = {
            "Месторождение": QInputDialog,
            "Куст": CustDialog,
            "Скважина": WellDialog,
            "Ствол": WellboreDialog
        }.get(dialog_title)

        if dialog_class == QInputDialog:
            value, ok = QInputDialog.getText(self, f"Добавить {dialog_title}", f"Введите название {dialog_title}:")
        else:
            dialog = dialog_class()
            ok = dialog.exec() == QDialog.DialogCode.Accepted
            value = dialog.get_inputs() if ok else None

        if ok and value:
            create_function(parent_item, value)

    def create_new_field(self, parent_item, value):
        # Создаем новый элемент "Месторождение"
        new_field_item = self.add_tree_item(self.fields, value, "Кусты", self.new_field)
        self.expand_items(new_field_item)

    def create_new_cust(self, parent_item, value):
        # Создаем новый куст
        new_cust_item = self.add_tree_item(parent_item, value, "Скважины")
        self.expand_items(new_cust_item)

    def create_new_well(self, parent_item, value):
        # Создаем новую скважину
        new_well_item = self.add_tree_item(parent_item, value, "Стволы")
        self.expand_items(new_well_item)

    def create_new_wellbore(self, parent_item, value):
        # Создаем новый ствол
        new_wellbore_item = self.add_tree_item(parent_item, value)
        self.expand_items(new_wellbore_item)
        self.create_tables(value)

    def add_tree_item(self, parent_item, text, child_text=None, insert_before=None):
        # Создаем новый элемент дерева
        new_item = QTreeWidgetItem([text])

        # Если есть вложенный элемент, добавляем его
        if child_text:
            child_item = QTreeWidgetItem(new_item, [child_text])
            QTreeWidgetItem(child_item, ["+"])  # Знак "+" для добавления нового элемента

        # Вставляем элемент
        if insert_before:
            parent_item.insertChild(parent_item.indexOfChild(insert_before), new_item)
        else:
            plus_item = parent_item.child(parent_item.childCount() - 1)  # Это элемент с "+"
            parent_item.insertChild(parent_item.indexOfChild(plus_item), new_item)

        return new_item

    def expand_items(self, item):
        """ Рекурсивно раскрываем все вложенные элементы дерева """
        item.setExpanded(True)
        for i in range(item.childCount()):
            item.child(i).setExpanded(True)

    def set_current_tab(self):
        self.tab_widget.setCurrentIndex(self.tab_widget.indexOf(self.tables))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    mainWindow = MainWindow()
    mainWindow.show()
    sys.exit(app.exec())