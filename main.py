import logging
import sys
from pathlib import Path

from PyQt6 import uic
from PyQt6.QtGui import QAction
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QMenu, QDockWidget, QTreeWidgetItem, QFileDialog, QTabWidget, QWidget, QDialog,
    QMessageBox, QVBoxLayout
)

from components.project_tree import ProjectTree, ProjectItem
from components.tables import Tables

logging.basicConfig(level=logging.ERROR, format='%(asctime)s - %(levelname)s - %(message)s')

class TablesTab(QWidget):
    """Таб с таблицей"""
    def __init__(self, project_item: ProjectItem, name: str, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout()
        self.setLayout(layout)
        self.tables = Tables(self)
        layout.addWidget(self.tables)
        self.name = name
        self.project_item: ProjectItem = project_item
        self.is_saved = False

        self.project_item.tab = self

    def save_tables(self, file_path: Path):
        parent = self.project_item.parent()
        if parent is None:
            # Сохранить как быстрый проект
            self.tables.save_all(self.name, file_path)
        else:
            # Сохранить как полноценный проект
            while parent.parent() is not None: # получаем самую верхнеуровневую папку проекта
                parent = parent.parent()
            self.create_project_folders(parent, file_path)

    def create_project_folders(self, item: ProjectItem, path: Path):
        """Creates folders corresponding to the item in the project structure."""
        if item:
            text = item.text(0).strip()
            path = path / text
            child_count = item.childCount()
            for i in range(child_count):
                child: ProjectItem = item.child(i)
                if child.item_type.is_savable():
                    if child.item_type.is_tables():
                        self.tables.save_all(self.name, path)
                    else:
                        self.create_project_folders(child, path)
            path.mkdir(parents=True, exist_ok=True)  # создать папку


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
        self.tree_widget.table_renamed.connect(self.rename_tables)
        self.tree_widget.table_deleted.connect(self.delete_tables)

        self.tab_widget: QTabWidget = self.findChild(QTabWidget, 'tabWidget')

        self.view_menu = self.findChild(QMenu, 'view_menu')
        self.dock_widget = self.findChild(QDockWidget, 'project_dockWidget')

        self.showMaximized()

    def setup_actions(self):
        self.open_file_action = self.findChild(QAction, 'open_file_action')
        self.open_file_action.triggered.connect(self.open_file)
        self.save_file_action = self.findChild(QAction, 'save_file_action')
        self.save_file_action.setEnabled(False)
        self.save_file_action.triggered.connect(self.save_project)

        self.print_action: QAction = self.findChild(QAction, 'print_action')
        self.print_action.setEnabled(False)
        self.print_action.triggered.connect(self.print_report)
        self.create_action = self.findChild(QAction, 'create_action')
        self.create_action.triggered.connect(self.tree_widget.create_fast_project)
        self.run_action = self.findChild(QAction, 'run_action')
        self.run_action.triggered.connect(self.run_project)

        self.toggle_dock_act = self.dock_widget.toggleViewAction()
        self.view_menu.addAction(self.toggle_dock_act)

    def run_project(self):
        print("Запуск расчетов...")

    def open_file(self):
        file_path = QFileDialog.getExistingDirectory(self, "Открыть проект", "")
        if file_path:
            file_path = Path(file_path)
            self.tree_widget.load_project(file_path)


    def create_tables(self, project_item, name):
        tables_tab = TablesTab(project_item, name, self)
        self.tab_widget.addTab(tables_tab, name)
        self.tab_widget.setCurrentWidget(tables_tab)

        self.save_file_action.setEnabled(True)
        self.print_action.setEnabled(True)

    def rename_tables(self, project_item: ProjectItem, name):
        tab: TablesTab = project_item.tab
        tab.name = name
        index = self.tab_widget.indexOf(tab)
        self.tab_widget.setTabText(index, name)

    def delete_tables(self, project_item: ProjectItem):
        tab: TablesTab = project_item.tab
        index = self.tab_widget.indexOf(tab)
        self.tab_widget.removeTab(index)

    def save_project(self):
        file_path = QFileDialog.getExistingDirectory(self, "Сохранить проект", "")
        if file_path:
            file_path = Path(file_path)
            current_tab: TablesTab = self.tab_widget.currentWidget()

            current_tab.save_tables(file_path)

    def print_report(self):
        current_tab: TablesTab = self.tab_widget.currentWidget()
        current_tab.tables.print_report()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec())
