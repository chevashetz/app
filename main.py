import logging
import sys
from pathlib import Path

from PyQt6 import uic
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QAction
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QInputDialog, QMenu, QDockWidget, QTreeWidget,
    QTreeWidgetItem, QFileDialog, QTabWidget, QDialog
)

from components.dialogs import WellboreDialog, CustDialog, WellDialog, FieldDialog
from components.tables import Tables

logging.basicConfig(level=logging.ERROR, format='%(asctime)s - %(levelname)s - %(message)s')


class MainWindow(QMainWindow):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        uic.loadUi('app.ui', self,  package='components')

        self.setup_ui()
        self.setup_actions()

        self.wellbores: dict[str, Tables] = {}
        self.undo_stack = []
        self.redo_stack = []
        self.tables = None

        # Start loading from "Месторождения"
        # self.load_project_structure(fields_path, self.fields_item)

        print("app.ui loaded successfully")

    def setup_ui(self):
        self.tree_widget: QTreeWidget = self.findChild(QTreeWidget, 'treeWidget')
        self.tab_widget: QTabWidget = self.findChild(QTabWidget, 'tabWidget')

        self.tree_widget.setHeaderLabels(["Наименование"])

        self.project_item = QTreeWidgetItem(self.tree_widget, ["Проект"])
        self.default_project_item = QTreeWidgetItem(self.tree_widget, ["Проекты по-умолчанию"])
        QTreeWidgetItem(self.default_project_item, ["+"])
        self.fields_item = QTreeWidgetItem(self.project_item, ["Месторождения"])
        QTreeWidgetItem(self.fields_item, ["+"])

        self.expand_items(self.project_item)
        self.expand_items(self.default_project_item)

        self.tree_widget.itemClicked.connect(self.on_item_clicked_tree)

        self.view_menu = self.findChild(QMenu, 'view_menu')
        self.dock_widget = self.findChild(QDockWidget, 'project_dockWidget')
        self.toggle_dock_act = self.dock_widget.toggleViewAction()
        self.view_menu.addAction(self.toggle_dock_act)

        self.tree_widget.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.tree_widget.customContextMenuRequested.connect(self.open_context_menu)
        self.tree_widget.setDragEnabled(True)
        self.tree_widget.setAcceptDrops(True)
        self.tree_widget.setDragDropMode(QTreeWidget.DragDropMode.InternalMove)
        #self.tl_br = self.findChild(QTool)

        self.showMaximized()

    def setup_actions(self):
        self.open_file_action = self.findChild(QAction, 'open_file_action')
        self.open_file_action.triggered.connect(self.open_file_all)
        self.save_file_action = self.findChild(QAction, 'save_file_action')
        self.save_file_action.setDisabled(True)
        self.print_action: QAction = self.findChild(QAction, 'print_action')
        self.print_action.setDisabled(True)
        self.create_action = self.findChild(QAction, 'create_action')
        self.create_action.triggered.connect(self.create_project)
        self.run_action = self.findChild(QAction, 'run_action')
        self.run_action.triggered.connect(self.run_project)


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
    def run_project(self):
        print(123)

    def on_item_clicked_tree(self, item: QTreeWidgetItem, column):
        if item.parent() is None:
            return
        elif item.text(0) == "+":
            parent_item = item.parent()
            parent_text = parent_item.text(0)
            item_mapping = {
                "Месторождения": ("Месторождение", "Кусты", FieldDialog),
                "Кусты": ("Куст", "Скважины", CustDialog),
                "Скважины": ("Скважину", "Стволы", WellDialog),
                "Стволы": ("Ствол", None, WellboreDialog),
                "Проекты по-умолчанию": ("Проект по-умолчанию", None, FieldDialog)
            }

            if item.text(0) == "+" and parent_text in item_mapping:
                item_type_name, child_name, dialog_class = item_mapping[parent_text]

                # Прямое создание диалога
                dialog = dialog_class(self)
                dialog.setWindowTitle(f"Добавить {item_type_name}")

                if dialog.exec() == QDialog.DialogCode.Accepted:
                    item_name = dialog.getText().strip()
                    if item_name:
                        insert_index = parent_item.indexOfChild(item)
                        if parent_item.text(0) == "Стволы" or parent_item.text(0) == "Проекты по-умолчанию":
                            self.create_table_item(parent_item, item_name, insert_index)
                        else:
                            self.create_new_folder(parent_item, item_name, child_name,
                                               insert_index=parent_item.indexOfChild(item))

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
            # Handle clicks on other items if necessary
            pass

    def create_new_folder(self, parent_item, name, child_text=None, insert_index=0):
        """Generic method to create a new item and associated folder."""
        new_item = QTreeWidgetItem([name])
        parent_item.insertChild(insert_index, new_item)

        # Do not add '+' if there is no child category (i.e., at the 'Ствол' level)
        if child_text:
            child_item = QTreeWidgetItem(new_item, [child_text])
            self.add_plus_button(child_item)
            self.expand_items(new_item)

    def create_table_item(self, parent_item, name: str, insert_index=0):
        # Контроль повторяющихся названий
        i = 0
        new_name = name
        while True:
            if new_name in self.wellbores:
                i += 1
                new_name = f'{name}({i})'
            else:
                break

        new_item = QTreeWidgetItem([new_name])
        parent_item.insertChild(insert_index, new_item)
        self.create_tables(new_name, parent_item)


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

    from pathlib import Path

    def load_project_folders(self, parent_item, path: Path):
        # Маркеры, после которых нужно добавлять "+"
        markers = ["Месторождения", "Кусты", "Скважины", "Стволы"]

        # Список для хранения последних путей
        last_paths = []
        last_parent_items = []

        # Флаг для определения наличия поддиректорий
        has_subfolders = False

        for subpath in path.iterdir():
            if subpath.is_dir():
                # Если мы находим поддиректорию, флаг становится True
                has_subfolders = True

                # Создаем элемент для папки
                folder_item = QTreeWidgetItem([subpath.name])
                parent_item.addChild(folder_item)

                last_paths, last_parent_items = self.load_project_folders(folder_item, subpath)
                # Рекурсивно собираем подпапки
                # last_paths.extend(last_paths)
                # last_parent_items.extend(last_parent_items)

        # Если папка не содержит поддиректорий, возвращаем её полный путь
        if not has_subfolders:
            last_paths.append(str(path))
            last_parent_items.append(parent_item)

        # Добавляем "+" если текст элемента совпадает с маркером
        if parent_item.text(0) in markers:
            plus_item = QTreeWidgetItem(["+"])
            parent_item.addChild(plus_item)

        # Разворачиваем элементы
        self.expand_items(parent_item)

        # Возвращаем список последних путей
        return (last_paths, last_parent_items)

    def open_file_all(self):
        file_path = QFileDialog.getExistingDirectory(self, "Открыть проект", "")
        if file_path:
            file_path = Path(file_path)
            if file_path.name == 'Проект':
                # Очищаем дерево перед загрузкой нового проекта
                self.tree_widget.clear()
                # Создаем корневой элемент для проекта
                project_item = QTreeWidgetItem([file_path.name])
                self.tree_widget.addTopLevelItem(project_item)
                # Загружаем структуру папок в дерево и получаем последние пути
                last_paths, last_parent_items = self.load_project_folders(project_item, file_path)
                # Загружаем таблицы для всех последних папок
                for path, last_parent_item in zip(last_paths, last_parent_items):
                    table_name = Path(path).name  # Используем имя папки как имя таблицы
                    self.create_tables(table_name, last_parent_item)  # Создаем таблицу для каждого пути
                    self.tables.load_all(Path(path))  # Загружаем данные из каждого последнего пути
            elif file_path.name == 'Проект по-умолчанию':
                path = next(file_path.iterdir())
                self.create_tables(path.name, self.default_project_item)
                self.tables.load_all(Path(path))


    def create_tables(self, name, parent):
        self.tables = Tables(self)
        self.wellbores[name] = self.tables
        self.tab_widget.addTab(self.tables, name)
        self.set_current_tab()

        self.save_file_action.setDisabled(False)
        self.save_file_action.triggered.connect(lambda _: self.save_project(name, parent))

        self.print_action.setDisabled(False)
        self.print_action.triggered.connect(self.tables.print_report)

    def create_project(self):
        project_name, ok = QInputDialog.getText(self, "Проект", "Введите название:")
        if ok and project_name.strip():
            self.create_table_item(self.default_project_item, project_name.strip(), self.default_project_item.childCount() - 1)
            self.default_project_item.setExpanded(True)

    def save_project(self, name, parent: QTreeWidgetItem):
        file_path = QFileDialog.getExistingDirectory(self, "Сохранить проект", "")
        if file_path:
            file_path = Path(file_path)
            if parent.text(0) == 'Ствол':
                self.create_project_folders(self.project_item, file_path)
            else:
                file_path = file_path / "Проект по-умолчанию"
                file_path.mkdir(parents=True, exist_ok=True)
                self.wellbores[name].save_all(name, file_path)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec())
