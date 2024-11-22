import json
import logging
import sys
from pathlib import Path

from PyQt6 import uic
from PyQt6.QtCore import QUrl
from PyQt6.QtGui import QAction
from PyQt6.QtNetwork import QNetworkRequest, QNetworkReply, QNetworkAccessManager
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QMenu, QDockWidget, QFileDialog, QTabWidget, QMessageBox
)

from components.project_tree import ProjectTree, ProjectItem
from components.tables import Tables
from components.tabs import TablesTab, ResultsTab

logging.basicConfig(level=logging.ERROR, format='%(asctime)s - %(levelname)s - %(message)s')


class MainWindow(QMainWindow):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        uic.loadUi('app.ui', self,  package='components')

        self.setup_ui()
        self.setup_actions()

        self.wellbores: dict[str, Tables] = {}
        self.tables = None

        # Создаем менеджер сетевых запросов
        self.network_manager = QNetworkAccessManager()
        self.network_manager.finished.connect(self.handle_response)

        self.reply_to_tab = {}  # Словарь для хранения связи reply -> current_tab

        print("app.ui loaded successfully")

    def setup_ui(self):
        self.tree_widget: ProjectTree = self.findChild(ProjectTree, 'treeWidget')
        self.tree_widget.table_created.connect(self.create_tables)
        self.tree_widget.table_renamed.connect(self.rename_tables)
        self.tree_widget.table_deleted.connect(self.delete_tables)
        self.tree_widget.table_clicked.connect(self.set_tables_tab)

        self.tab_widget: QTabWidget = self.findChild(QTabWidget, 'tabWidget')
        self.tab_widget.tabCloseRequested.connect(lambda index: self.tab_widget.removeTab(index))

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
        self.run_action.triggered.connect(self.start_response)

        self.toggle_dock_act = self.dock_widget.toggleViewAction()
        self.view_menu.addAction(self.toggle_dock_act)

    def start_response(self):
        current_tab: TablesTab = self.tab_widget.currentWidget()

        """Начать сетевой запрос"""
        self.run_action.setEnabled(False)

        # Создаем URL и request
        url = QUrl("http://localhost:8000/calculate/")
        request = QNetworkRequest(url)
        request.setHeader(
            QNetworkRequest.KnownHeaders.ContentTypeHeader,
            "application/json"
        )

        if current_tab is None:
            return

        results_tab = None  # изначально результатов нет
        # TODO: цикл
        # Подготавливаем данные для отправки
        for i in range(1):
            data = {
                "name": "",
                "depth_from": current_tab.content.tbl_drilling_fluids.item(i, 0).text(),
                "depth_to": 0,
                "transport_model": "Bingam",
                "density": 0,
                "viscosity": 0,
                "dns": 0
            }

            # Конвертируем данные в JSON и затем в QByteArray
            json_data = json.dumps(data).encode('utf-8')

            # Отправляем POST запрос
            reply = self.network_manager.post(request, json_data)

            # Связываем reply с current_tab
            self.reply_to_tab[reply] = current_tab


    def handle_response(self, reply: QNetworkReply):
        """Обработка ответа от сервера"""
        try:
            current_tab: TablesTab | None = self.reply_to_tab.pop(reply, None)  # Извлекаем current_tab
            if current_tab is None:
                QMessageBox.critical(self, "Результат", "Проект был не найден или удален")
                return

            if reply.error() == QNetworkReply.NetworkError.NoError:
                # Читаем данные из ответа
                data = reply.readAll().data().decode('utf-8')
                # Парсим JSON
                result = json.loads(data)
                # Показываем результат
                self.create_results(result, current_tab)
            else:
                error_message = f"Ошибка запроса: {reply.errorString()}"
                QMessageBox.critical(self, "Ошибка", error_message)

        except json.JSONDecodeError as e:
            QMessageBox.critical(
                self,
                "Ошибка",
                f"Ошибка парсинга JSON: {str(e)}"
            )
        except Exception as e:
            QMessageBox.critical(
                self,
                "Ошибка",
                f"Неизвестная ошибка: {str(e)}"
            )
        finally:
            self.run_action.setEnabled(True)
            reply.deleteLater()  # Очищаем память

    def open_file(self):
        file_path = QFileDialog.getExistingDirectory(self, "Открыть проект", "")
        if file_path:
            file_path = Path(file_path)
            self.tree_widget.load_project(file_path)

    def create_tables(self, project_item, name):
        tables_tab = TablesTab(project_item, name)
        self.tab_widget.addTab(tables_tab, name)
        self.tab_widget.setCurrentWidget(tables_tab)

        self.save_file_action.setEnabled(True)
        self.print_action.setEnabled(True)

    def create_results(self, data: dict, current_tab: TablesTab):
        name = f"{current_tab.name}_результаты"

        results_tab: ResultsTab | None = current_tab.get_results_tab()

        if results_tab is None:  # когда еще не создали вкладку под результаты
            result_project_item = self.tree_widget.create_results(current_tab.project_item, name)
            results_tab = ResultsTab(result_project_item, current_tab.name)
            self.tab_widget.addTab(results_tab, name)
            self.current_processing_results[name] = results_tab

        results_tab.add_new_data(data)



    def rename_tables(self, project_item: ProjectItem, name):
        tab: TablesTab = project_item.tab
        tab.name = name
        index = self.tab_widget.indexOf(tab)
        self.tab_widget.setTabText(index, name)

    def delete_tables(self, project_item: ProjectItem):
        tab: TablesTab = project_item.tab
        index = self.tab_widget.indexOf(tab)
        self.tab_widget.removeTab(index)

    def set_tables_tab(self, project_item: ProjectItem):
        tab = project_item.tab
        index = self.tab_widget.indexOf(tab)
        if index == -1:
            self.tab_widget.addTab(tab, tab.name)
        self.tab_widget.setCurrentWidget(tab)

    def save_project(self):
        file_path = QFileDialog.getExistingDirectory(self, "Сохранить проект", "")
        if file_path:
            file_path = Path(file_path)
            current_tab: TablesTab = self.tab_widget.currentWidget()

            current_tab.save_tables(file_path)

    def print_report(self):
        current_tab: TablesTab = self.tab_widget.currentWidget()
        current_tab.content.print_report()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec())
