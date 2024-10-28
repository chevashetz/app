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
from openpyxl.chart.trendline import Trendline

from components.dialogs import WellboreDialog, CustDialog, WellDialog, FieldDialog
from components.project_tree import ProjectTree
from components.tables import Tables

logging.basicConfig(level=logging.ERROR, format='%(asctime)s - %(levelname)s - %(message)s')


class MainWindow(QMainWindow):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        uic.loadUi('app.ui', self,  package='components')

        self.setup_ui()
        self.setup_actions()

        self.wellbores: dict[str, Tables] = {}
        self.tables = None

        print("app.ui loaded successfully")

    def setup_ui(self):
        self.tree_widget: ProjectTree = self.findChild(ProjectTree, 'treeWidget')
        self.tree_widget.table_created.connect(self.create_tables)

        self.tab_widget: QTabWidget = self.findChild(QTabWidget, 'tabWidget')

        self.view_menu = self.findChild(QMenu, 'view_menu')
        self.dock_widget = self.findChild(QDockWidget, 'project_dockWidget')

        self.showMaximized()

    def setup_actions(self):
        self.open_file_action = self.findChild(QAction, 'open_file_action')
        self.open_file_action.triggered.connect(self.open_file_all)
        self.save_file_action = self.findChild(QAction, 'save_file_action')
        self.save_file_action.setEnabled(False)
        self.print_action: QAction = self.findChild(QAction, 'print_action')
        self.print_action.setEnabled(False)
        self.create_action = self.findChild(QAction, 'create_action')
        self.create_action.triggered.connect(self.tree_widget.create_fast_project)
        self.run_action = self.findChild(QAction, 'run_action')
        self.run_action.triggered.connect(self.run_project)

        self.toggle_dock_act = self.dock_widget.toggleViewAction()
        self.view_menu.addAction(self.toggle_dock_act)

    def run_project(self):
        print(123)

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

    def set_current_tab(self):
        """Sets the current tab to self.tables."""
        self.tab_widget.setCurrentWidget(self.tables)

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
        return last_paths, last_parent_items

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


    def create_tables(self, name):
        self.tables = Tables(self)
        self.wellbores[name] = self.tables
        self.tab_widget.addTab(self.tables, name)
        self.set_current_tab()

        self.save_file_action.setEnabled(True)
        self.save_file_action.triggered.connect(lambda _: self.save_project(name, None))

        self.print_action.setEnabled(True)
        self.print_action.triggered.connect(self.tables.print_report)

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
