from __future__ import annotations

import sys
from dataclasses import dataclass
from enum import Enum
from pathlib import Path

from PyQt6 import QtCore
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QIcon, QAction, QKeySequence, QShortcut
from PyQt6.QtWidgets import QApplication, QMainWindow, QTreeWidget, QTreeWidgetItem, QMenu, QDialog, QInputDialog, \
    QMessageBox

from components.dialogs import FieldDialog, CustDialog, WellDialog, WellboreDialog
from config import ICONS_PATH


@dataclass
class ItemType:
    name: str = ""
    icon_path: str = ""


class ItemTypes(Enum):
    project = ItemType("Проект", str(ICONS_PATH / "folder"))
    fields = ItemType("Месторождения", str(ICONS_PATH / "folder"))
    field = ItemType("Месторождение", str(ICONS_PATH / "soil_layers"))
    custs = ItemType("Кусты", str(ICONS_PATH / "folder"))
    cust = ItemType("Куст", str(ICONS_PATH / "cust"))
    wells = ItemType("Скважины", str(ICONS_PATH / "folder"))
    well = ItemType("Скважина", str(ICONS_PATH / "well"))
    wellbores = ItemType("Стволы", str(ICONS_PATH / "wellbore"))
    tables = ItemType("Ствол", str(ICONS_PATH / "table"))

    graph = ItemType(icon_path=str(ICONS_PATH / "graph"))
    plus = ItemType(icon_path=str(ICONS_PATH / "plus"))

    def is_tables(self):
        return self == ItemTypes.tables

    def is_wellbores(self):
        return self == ItemTypes.wellbores

    def is_editable(self):
        return self not in {ItemTypes.fields, ItemTypes.custs,
                            ItemTypes.wells, ItemTypes.wellbores, ItemTypes.plus}

    def is_deletable(self):
        return self not in {ItemTypes.project, ItemTypes.fields, ItemTypes.custs,
                            ItemTypes.wells, ItemTypes.wellbores, ItemTypes.plus}

    def is_savable(self):
        return self not in {ItemTypes.plus}

    @staticmethod
    def dialog_types():
        return ItemTypes.fields, ItemTypes.custs, ItemTypes.wells, ItemTypes.wellbores

    @staticmethod
    def user_create_types():
        return ItemTypes.field, ItemTypes.cust, ItemTypes.well, ItemTypes.tables

    @staticmethod
    def item_by_name(name):
        for i in ItemTypes:
            if i.name == name:
                return i

    def next_value(self) -> ItemTypes:
        item_types = list(ItemTypes)
        return item_types[item_types.index(self) + 1]

    def prev_value(self) -> ItemTypes:
        item_types = list(ItemTypes)
        return item_types[item_types.index(self) - 1]


class ProjectItem(QTreeWidgetItem):
    def __init__(self, name="", parent=None, item_type: ItemTypes = ItemTypes.plus):
        if name == "":  # Если имя не передано в конструктор
            name = item_type.value.name  # Взять его из типа

        super().__init__(parent, [name])
        self.item_type: ItemTypes = item_type
        self.tab: "TablesTab" = None
        icon = QIcon(self.item_type.value.icon_path)
        self.setIcon(0, icon)


class ProjectTree(QTreeWidget):
    table_created = pyqtSignal(ProjectItem, str)
    table_deleted = pyqtSignal(ProjectItem)
    table_renamed = pyqtSignal(ProjectItem, str)
    table_clicked = pyqtSignal(ProjectItem)

    def __init__(self, parent):
        super().__init__(parent)
        self.undo_stack = []
        self.redo_stack = []
        self.tables = {}

        self.dialogs = {type: dialog for type, dialog in
                        zip(ItemTypes.dialog_types(), (FieldDialog, CustDialog, WellDialog, WellboreDialog,))}

        self.setup_tree()
        self.setup_actions()

    def setup_tree(self):
        self.setColumnCount(1)
        self.setHeaderLabels(["Структура проекта"])
        self.setDragEnabled(True)
        self.setAcceptDrops(True)
        self.setDropIndicatorShown(True)

        # Создаем корневой элемент
        self.root = self.invisibleRootItem()

        # Создаем элемент "Project"
        self.project = ProjectItem("Новый проект", parent=self.root, item_type=ItemTypes.project)
        ProjectItem(parent=self.root)
        self.fields = ProjectItem(parent=self.project, item_type=ItemTypes.fields)
        ProjectItem(parent=self.fields)

        self.expand_items(self.project)

        # Добавляем контекстное меню
        self.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.customContextMenuRequested.connect(self.show_context_menu)

        self.itemClicked.connect(self.on_item_clicked_tree)

        # Нажатие enter клавиши
        shortcut = QShortcut(QtCore.Qt.Key.Key_Return, self,
                             context=QtCore.Qt.ShortcutContext.WidgetShortcut)
        # noinspection PyTypeChecker
        shortcut.activated.connect(lambda: self.on_item_clicked_tree(item, 0) if (item := self.currentItem()) else None)


    def setup_actions(self):
        self.undo_action = QAction("Отменить последнее действие", self)
        self.undo_action.triggered.connect(lambda: self.undo_redo_last_action(self.undo_stack, self.redo_stack))
        self.undo_action.setShortcut(QKeySequence.StandardKey.Undo)
        self.undo_action.setEnabled(False)

        self.redo_action = QAction("Повторить действие", self)
        self.redo_action.triggered.connect(lambda: self.undo_redo_last_action(self.redo_stack, self.undo_stack))
        self.redo_action.setShortcut(QKeySequence.StandardKey.Redo)
        self.redo_action.setEnabled(False)

        self.addActions([self.undo_action, self.redo_action])

    def expand_items(self, item):
        """Recursively expands all child items."""
        item.setExpanded(True)
        for i in range(item.childCount()):
            child = item.child(i)
            self.expand_items(child)

    def dragMoveEvent(self, event):
        # noinspection PyTypeChecker
        item: ProjectItem = self.itemAt(event.position().toPoint())
        # noinspection PyTypeChecker
        dragged_item: ProjectItem = self.currentItem()

        if dragged_item and dragged_item.item_type.is_tables():
            if item is None or item == self.root or (item and item.item_type.is_wellbores()):
                event.accept()
                return

        event.ignore()

    def dropEvent(self, event):
        # noinspection PyTypeChecker
        item: ProjectItem = self.itemAt(event.position().toPoint())
        # noinspection PyTypeChecker
        dragged_item: ProjectItem = self.currentItem()

        if dragged_item and dragged_item.item_type.is_tables():
            if item is None or item == self.root or item.item_type.is_wellbores():
                parent = item or self.root
                old_parent = dragged_item.parent() or self.root
                if parent is not old_parent:
                    move_params = {'parent': old_parent}
                    # Если в новой папке такое имя таблицы существует - переименовать
                    neighbors = self.get_all_children(parent)
                    if (name := dragged_item.text(0)) in [item.text(0) for item in neighbors]:
                        dragged_item.setText(0, self.get_unique_name(name, dragged_item.item_type, neighbors))
                        move_params['name'] = name
                    old_parent.removeChild(dragged_item)
                    parent.insertChild(parent.childCount() - 1, dragged_item)
                    self.append_undo_stack(('move', dragged_item, move_params))
            event.accept()
        else:
            event.ignore()

    def show_context_menu(self, position):
        menu = QMenu()
        # noinspection PyTypeChecker
        item: ProjectItem = self.itemAt(position)
        if item:
            # Edit and Delete actions
            edit_action = QAction("Редактировать", self)
            edit_action.triggered.connect(lambda: self.edit_item(item))
            edit_action.setEnabled(item.item_type.is_editable())

            delete_action = QAction("Удалить", self)
            delete_action.triggered.connect(lambda: self.delete_item(item))
            delete_action.setEnabled(item.item_type.is_deletable())

            menu.addAction(edit_action)
            menu.addAction(delete_action)
        else:
            menu.addAction(self.undo_action)
            menu.addAction(self.redo_action)

        menu.exec(self.mapToGlobal(position))

    def update_actions(self):
        self.undo_action.setEnabled(bool(self.undo_stack))
        self.redo_action.setEnabled(bool(self.redo_stack))

    def on_item_clicked_tree(self, item: ProjectItem, column):
        if item.item_type == ItemTypes.plus:
            parent_item: ProjectItem = item.parent()

            if parent_item is None:
                self.create_fast_project()  # Создать Быстрый проект
            elif (parent_type := parent_item.item_type) in self.dialogs:
                dialog_class = self.dialogs[parent_type]
                child_type = parent_type.next_value()

                # Прямое создание диалога
                dialog = dialog_class(self)
                dialog.setWindowTitle(f"Добавить {child_type.name}")

                if dialog.exec() == QDialog.DialogCode.Accepted:
                    item_name = dialog.getText().strip()
                    if item_name:
                        item_name = self.get_unique_name(item_name, child_type, self.get_all_children(parent_item))
                        new_item = ProjectItem(item_name, item_type=child_type)
                        parent_item.insertChild(parent_item.childCount() - 1, new_item)
                        if child_type == ItemTypes.tables:
                            self.table_created.emit(new_item, item_name)
                        else:
                            folder_type = child_type.next_value()
                            folder = ProjectItem(parent=new_item, item_type=folder_type)
                            plus = ProjectItem(parent=folder)  # plus sign inside
                            self.setCurrentItem(plus)

                            self.expand_items(new_item)
                        self.append_undo_stack(('create', new_item, {'parent': parent_item}))
        elif item.item_type is ItemTypes.tables:
            self.table_clicked.emit(item)

    def create_fast_project(self):
        project_name, ok = QInputDialog.getText(self, "Проект", "Введите название:")
        if ok and project_name.strip():
            self.create_fast_project_item(project_name)

    def create_fast_project_item(self, project_name):
        project_name = self.get_unique_name(project_name.strip(), ItemTypes.tables,
                                            self.get_all_children(self.root))
        new_item = ProjectItem(project_name, item_type=ItemTypes.tables)
        self.root.insertChild(self.root.childCount() - 1, new_item)
        self.append_undo_stack(('create', new_item, {'parent': self.root}))
        self.table_created.emit(new_item, project_name)
        return new_item

    def get_unique_name(self, name: str, item_type: ItemTypes, children: list[ProjectItem]) -> str:
        # Контроль повторяющихся названий
        i = 0
        new_name = name
        item_names = [item.text(0) for item in children if item.item_type == item_type]
        while True:
            if new_name in item_names:
                i += 1
                new_name = f'{name}({i})'
            else:
                break
        return new_name

    @staticmethod
    def get_all_children(parent):
        return [parent.child(child_idx) for child_idx in range(parent.childCount())]

    @staticmethod
    def find_leaves(item: ProjectItem):
        """
        Рекурсивно находит и возвращает список всех листьев для данного QTreeWidgetItem.

        :param item: QTreeWidgetItem, для которого ищутся листья.
        :return: Список объектов QTreeWidgetItem, которые являются листьями.
        """
        leaves = []
        if item.childCount() == 0 and item.item_type is ItemTypes.tables:
            # Если у узла нет потомков, он является листом
            leaves.append(item)
        else:
            # Если есть потомки, рекурсивно ищем листья у каждого
            for i in range(item.childCount()):
                child = item.child(i)
                leaves.extend(ProjectTree.find_leaves(child))
        return leaves

    def edit_item(self, item: ProjectItem):
        old_text = item.text(0)
        new_text, ok = QInputDialog.getText(self, "Редактирование элемента",
                                            "Введите новое имя:", text=old_text)
        new_text = new_text.strip()
        if ok and new_text and new_text != old_text:
            item.setText(0, self.get_unique_name(new_text.strip(), item.item_type,
                                                 self.get_all_children(item.parent() or self.root)))
            if item.item_type == ItemTypes.tables:
                self.table_renamed.emit(item, new_text)
            self.append_undo_stack(('rename', item, {'text': old_text}))

    def delete_item(self, item):
        child_tables = ProjectTree.find_leaves(item)
        parent = (item.parent() or self.root)
        parent.removeChild(item)
        for table in child_tables:
            self.table_deleted.emit(table)
        self.append_undo_stack(('delete', item, {'parent': parent}))

    def append_undo_stack(self, last_action):
        self.undo_stack.append(last_action)
        self.undo_action.setEnabled(True)
        # self.redo_stack.clear()

    def undo_redo_last_action(self, out_stack: list, in_stack: list):
        if not out_stack:
            return
        last_action: tuple[str, ProjectItem, dict] = out_stack.pop()
        action_type, item, params = last_action

        if action_type == 'create':
            parent: ProjectItem = params.get('parent')
            parent.removeChild(item)
            if item.item_type == ItemTypes.tables:
                self.table_deleted.emit(item)
            last_action = ('delete', item, {'parent': parent})
        elif action_type == 'rename':
            old_text: str = params.get('text')
            new_text: str = item.text(0)
            item.setText(0, old_text)
            if item.item_type == ItemTypes.tables:
                self.table_renamed.emit(item, old_text)
            last_action[2]['text'] = new_text
        elif action_type == 'move':
            old_parent = item.parent() or self.root
            new_parent = params.get('parent')
            old_parent.removeChild(item)
            new_parent.insertChild(new_parent.childCount() - 1, item)
            last_action[2]['parent'] = old_parent
            if name := params.get('name'):  # Если при переносе было переименованние
                params['name'] = item.text(0)
                item.setText(0, name)
        elif action_type == 'delete':
            parent: ProjectItem = params.get('parent')
            parent.insertChild(parent.childCount() - 1, item)
            if item.item_type == ItemTypes.tables:
                self.table_created.emit(item, item.text(0))
            self.expand_items(item)
            last_action = ('create', item, {'parent': parent})

        in_stack.append(last_action)

        self.undo_action.setEnabled(bool(self.undo_stack))
        self.redo_action.setEnabled(bool(self.redo_stack))

    def load_project_folders(self, parent_item, path: Path):
        if parent_item:
            for path in path.iterdir():
                child_type = parent_item.item_type.next_value()
                item_name = path.name
                new_item = ProjectItem(item_name, parent=parent_item, item_type=child_type)
                if child_type == ItemTypes.tables:
                    self.table_created.emit(new_item, item_name)
                    if check_correct_files(path):
                        new_item.tab.tables.load_all(path)
                else:
                    self.load_project_folders(new_item, path)
                    if not child_type.is_editable():
                        ProjectItem(parent=new_item)  # plus sign inside



    def load_project(self, file_path: Path):
        if (file_path / 'Месторождения').exists():
            # Очищаем дерево перед загрузкой нового проекта
            self.root.removeChild(self.project)
            self.append_undo_stack(('remove', self.project, {'parent': self.root}))
            # Создаем корневой элемент для проекта
            self.project = ProjectItem(file_path.name, item_type=ItemTypes.project)
            self.root.insertChild(0, self.project)

            # Загружаем структуру папок в дерево и получаем последние пути
            self.load_project_folders(self.project, file_path)
            self.expand_items(self.project)

            self.append_undo_stack(('create', self.fields, {'parent': self.project}))
        else:
            if check_correct_files(file_path):
                new_item = self.create_fast_project_item(file_path.name)
                new_item.tab.tables.load_all(file_path)
            else:
                dlg = QMessageBox(self)
                dlg.setWindowTitle("Проект не найден!")
                dlg.setText("В этой папке проекта не найдено")
                dlg.exec()


def check_correct_files(file_path: Path):
    list_files = set(list(file_path.iterdir()))
    if {file.name for file in list_files} <= {'Буровые растворы.csv', 'Давления.csv', 'КНБК_1.csv',
                                              'Обсадные колонны.csv', 'Профиль.csv', 'Стратиграфия.csv'}:
        return True
    return False


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Project Tree")
        self.tree_widget = ProjectTree(self)
        self.setCentralWidget(self.tree_widget)
        self.setGeometry(400, 200, 400, 600)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
