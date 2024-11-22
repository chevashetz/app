from pathlib import Path

from PyQt6.QtWidgets import QWidget, QVBoxLayout

from components.project_tree import ProjectItem
from components.results import Results
from components.tables import Tables


class ResultsTab(QWidget):
    """Таб с результатами"""

    def __init__(self, project_item: ProjectItem, name: str):
        super().__init__()
        layout = QVBoxLayout()
        self.setLayout(layout)
        self.content = Results(self)
        layout.addWidget(self.content)
        self.name = name
        self.project_item: ProjectItem = project_item
        self.is_saved = False

        self.project_item.tab = self

    def add_new_data(self, data: dict):
        """Перезаписать и добавить новую дату"""
        print(data)


class TablesTab(QWidget):
    """Таб с таблицей"""

    def __init__(self, project_item: ProjectItem, name: str):
        super().__init__()
        layout = QVBoxLayout()
        self.setLayout(layout)
        self.content = Tables(self)
        layout.addWidget(self.content)
        self.name = name
        self.project_item: ProjectItem = project_item
        self.project_item.tab = self

        self.is_saved = False  # Сохранены изменения
        self.processing = False  # На вкладке ведутся расчеты?

    def save_tables(self, file_path: Path):
        parent = self.project_item.parent()
        if parent is None:
            # Сохранить как быстрый проект
            self.content.save_all(self.name, file_path)
        else:
            # Сохранить как полноценный проект
            while parent.parent() is not None:  # получаем самую верхнеуровневую папку проекта
                parent = parent.parent()
            self.create_project_folders(parent, file_path)

    def get_results_tab(self) -> None | ResultsTab:
        """Получить результат у таблицы, вернуть None если его нету"""
        child: ProjectItem | None = self.project_item.child(0)
        return child and child.tab  # вернет None или ResultsTab


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
                        self.content.save_all(self.name, path)
                    else:
                        self.create_project_folders(child, path)
            path.mkdir(parents=True, exist_ok=True)  # создать папку
