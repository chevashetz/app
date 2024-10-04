import logging
import sys
from pathlib import Path

from PyQt6 import uic
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QAction
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QInputDialog, QMenu, QDockWidget, QTreeWidget,
    QTreeWidgetItem, QFileDialog, QTabWidget
)

from components.tables import Tables
from config import path5

logging.basicConfig(level=logging.ERROR, format='%(asctime)s - %(levelname)s - %(message)s')


class MainWindow(QMainWindow):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        uic.loadUi('app.ui', self)

        self.setup_ui()
        self.setup_actions()

        self.wellbores: dict[str, Tables] = {}
        self.undo_stack = []
        self.redo_stack = []

        # Start loading from "Месторождения"
        #self.load_project_structure(fields_path, self.fields_item)

        print("app.ui loaded successfully")

    def setup_ui(self):
        self.tree_widget: QTreeWidget = self.findChild(QTreeWidget, 'treeWidget')
        self.tab_widget: QTabWidget = self.findChild(QTabWidget, 'tabWidget')

        self.tree_widget.setHeaderLabels(["Наименование"])

        self.project_item = QTreeWidgetItem(self.tree_widget, ["Проект"])
        self.fields_item = QTreeWidgetItem(self.project_item, ["Месторождения"])
        self.new_field_item = QTreeWidgetItem(self.fields_item, ["+"])

        self.tree_widget.itemClicked.connect(self.on_item_clicked_tree)

        self.view_menu = self.findChild(QMenu, 'view_menu')
        self.dock_widget = self.findChild(QDockWidget, 'project_dockWidget')
        self.toggle_dock_act = self.dock_widget.toggleViewAction()
        self.view_menu.addAction(self.toggle_dock_act)

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
            menu = QMenu()
            # Undo and Redo actions
            undo_action = QAction("Отменить последнее действие", self)
            undo_action.triggered.connect(self.undo_last_action)
            undo_action.setEnabled(bool(self.undo_stack))

            redo_action = QAction("Повторить действие", self)
            redo_action.triggered.connect(self.redo_last_action)
            redo_action.setEnabled(bool(self.redo_stack))

            menu.addAction(undo_action)
            menu.addAction(redo_action)

            # Edit and Delete actions
            edit_action = QAction("Редактировать", self)
            edit_action.triggered.connect(lambda: self.edit_item(item))

            delete_action = QAction("Удалить", self)
            delete_action.triggered.connect(lambda: self.delete_item(item))

            menu.addAction(edit_action)
            menu.addAction(delete_action)

            menu.exec(self.tree_widget.viewport().mapToGlobal(position))

    def edit_item(self, item):
        non_editable_items = {"Проект", "Месторождения", "Кусты", "Скважины", "Стволы", "+"}

        if item.text(0) in non_editable_items:
            print(f"Нельзя редактировать элемент: {item.text(0)}")
            return

        new_text, ok = QInputDialog.getText(self, "Редактирование элемента",
                                            "Введите новое имя:", text=item.text(0))

        if ok and new_text.strip():
            item.setText(0, new_text.strip())
            print(f"Элемент изменен на: {new_text.strip()}")

    def undo_last_action(self):
        if self.undo_stack:
            last_action = self.undo_stack.pop()
            action_type, item, parent = last_action

            if action_type == 'delete':
                if isinstance(parent, QTreeWidget):
                    parent.addTopLevelItem(item)
                else:
                    parent.addChild(item)
                print(f"Отменено удаление элемента: {item.text(0)}")
                self.redo_stack.append(last_action)

    def redo_last_action(self):
        if self.redo_stack:
            last_action = self.redo_stack.pop()
            action_type, item, parent = last_action

            if action_type == 'delete':
                if isinstance(parent, QTreeWidget):
                    index = parent.indexOfTopLevelItem(item)
                    if index != -1:
                        parent.takeTopLevelItem(index)
                else:
                    parent.removeChild(item)
                print(f"Действие повторено: Удаление элемента {item.text(0)}")
                self.undo_stack.append(last_action)

    def delete_item(self, item):
        parent = item.parent()
        if parent is None:
            index = self.tree_widget.indexOfTopLevelItem(item)
            if index != -1:
                self.undo_stack.append(('delete', item, self.tree_widget))
                self.tree_widget.takeTopLevelItem(index)
                self.redo_stack.clear()
                print(f"Элемент удален: {item.text(0)}")
        else:
            self.undo_stack.append(('delete', item, parent))
            parent.removeChild(item)
            self.redo_stack.clear()
            print(f"Элемент удален: {item.text(0)}")

    def on_item_clicked_tree(self, item, column):
        if item.parent() is None:
            return
        elif item.text(0) == "+":
            parent_item = item.parent()
            parent_text = parent_item.text(0)
            item_mapping = {
                "Месторождения": ("Месторождение", "Кусты", QInputDialog),
                "Кусты": ("Куст", "Скважины", QInputDialog),
                "Скважины": ("Скважину", "Стволы", QInputDialog),
                "Стволы": ("Ствол", None, QInputDialog)
            }

            if parent_text in item_mapping:
                item_type_name, child_name, dialog = item_mapping[parent_text]
                item_name, ok = dialog.getText(self, f"Добавить {item_type_name}",
                                                     f"Введите название {item_type_name.lower()}:")
                if ok and item_name.strip():
                    self.create_new_item(parent_item, item_name.strip(), child_name)
        elif item.text(0) in {"Месторождения", "Кусты", "Скважины", "Стволы"}:
            # Expand or collapse the branch
            item.setExpanded(not item.isExpanded())
        elif item.parent().text(0) == "Стволы":
            # Load data for the selected borehole
            name = item.text(0)
            self.tables = self.wellbores.get(name)
            if self.tables:
                self.set_current_tab()
            else:
                self.create_tables(name)
        else:
            # Handle clicks on other items if necessary
            pass

    def create_new_item(self, parent_item, name, child_text=None):
        """Generic method to create a new item and associated folder."""
        new_item = QTreeWidgetItem([name])

        if child_text:
            child_item = QTreeWidgetItem(new_item, [child_text])
            self.add_plus_button(child_item)
            # Expand the child item
            child_item.setExpanded(True)
        # Do not add '+' if there is no child category (i.e., at the 'Ствол' level)

        plus_item = None
        for i in range(parent_item.childCount()):
            if parent_item.child(i).text(0) == "+":
                plus_item = parent_item.child(i)
                break

        if plus_item:
            parent_item.insertChild(parent_item.indexOfChild(plus_item), new_item)
        else:
            parent_item.addChild(new_item)

        self.expand_items(new_item)

        if parent_item.text(0) == "Стволы":
            self.create_tables(name)

    def create_project_folders(self, item, path: Path):
        """Creates folders corresponding to the item in the project structure."""
        if item:
            text = item.text(0).strip()
            path = path / text
            for i in range(item.childCount()):
                child = item.child(i)
                child_text = child.text(0).strip()
                if child_text and child_text != '+':
                    if text == 'Стволы':
                        self.wellbores[child_text].save_all(child_text, path)
                    else:
                        self.create_project_folders(child, path)


    def add_plus_button(self, parent_item):
        plus_item = QTreeWidgetItem(["+"])
        parent_item.addChild(plus_item)

    def expand_items(self, item):
        """Recursively expands all child items."""
        item.setExpanded(True)
        for i in range(item.childCount()):
            child = item.child(i)
            self.expand_items(child)

    def set_current_tab(self):
        """Sets the current tab to self.tables."""
        self.tab_widget.setCurrentWidget(self.tables)

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
        self.save_file_action.triggered.connect(lambda _: self.save_project(name))

        self.print_action.setDisabled(False)
        self.print_action.triggered.connect(self.tables.print_report)

    def create_project(self):
        project_name, ok = QInputDialog.getText(self, "Проект", "Введите название:")
        if ok and project_name.strip():
            self.create_new_item(self.tree_widget, project_name.strip(), "Кусты")

    def save_project(self, name):
        file_path = QFileDialog.getExistingDirectory(self, "Сохранить проект", "")
        if file_path:
            file_path = Path(file_path)
            self.create_project_folders(self.project_item, file_path)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec())
