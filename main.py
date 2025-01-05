import json
import logging
import sys
from pathlib import Path

from PyQt6 import uic
from PyQt6.QtCore import QUrl
from PyQt6.QtGui import QAction, QIcon
from PyQt6.QtNetwork import QNetworkRequest, QNetworkReply, QNetworkAccessManager, QAbstractSocket
from PyQt6.QtWebSockets import QWebSocket
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QMenu, QDockWidget, QFileDialog, QTabWidget, QMessageBox, QStatusBar, QProgressBar,
    QLabel
)

from components.project_tree import ProjectTree, ProjectItem, get_name_for_results
from components.tables import Tables
from components.results import Results
from components.tabs import TablesTab, ResultsTab
from config import WS_HOST, WS_PORT, WS_PORT, WS_ENDPOINT
from utils import extract_number, countFilledRows

logging.basicConfig(level=logging.ERROR, format='%(asctime)s - %(levelname)s - %(message)s')


class MainWindow(QMainWindow):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        uic.loadUi('app.ui', self, package='components')

        self.setup_ui()
        self.setup_actions()

        self.wellbores: dict[str, Tables] = {}
        self.tables = None

        self.ws = QWebSocket()
        self.ws.textMessageReceived.connect(self.handle_response)

        self.tab_widget.currentChanged.connect(self.update_run_state)
        self.tab_widget.tabCloseRequested.connect(self.update_run_state)

        self.task_id_to_tab = {}  # Словарь для хранения связи task_id -> current_tab

        self.update_run_state()

    def setup_ui(self):
        self.tree_widget: ProjectTree = self.findChild(ProjectTree, 'treeWidget')
        self.tree_widget.table_created.connect(self.create_tables)
        self.tree_widget.table_renamed.connect(self.rename_tables)
        self.tree_widget.table_deleted.connect(self.delete_tables)
        self.tree_widget.table_clicked.connect(self.set_tab)

        self.tab_widget: QTabWidget = self.findChild(QTabWidget, 'tabWidget')
        self.tab_widget.tabCloseRequested.connect(lambda index: self.tab_widget.removeTab(index))

        self.view_menu = self.findChild(QMenu, 'view_menu')
        self.dock_widget = self.findChild(QDockWidget, 'project_dockWidget')

        self.status_bar: QStatusBar = self.findChild(QStatusBar, "statusbar")
        self.setStatusBar(self.status_bar)
        self.result_progress = QProgressBar()
        self.result_progress_label = QLabel()

        self.status_bar.addWidget(self.result_progress_label)
        self.status_bar.addWidget(self.result_progress)

        self.result_progress.setVisible(False)
        self.result_progress_label.setVisible(False)
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
        self.run_action: QAction = self.findChild(QAction, 'run_action')
        self.run_action.triggered.connect(self.start_response)

        self.stop_action: QAction = self.findChild(QAction, 'stop_action')
        self.stop_action.triggered.connect(lambda _: print("Напиши функцию для меня"))

        self.toggle_dock_act = self.dock_widget.toggleViewAction()
        self.view_menu.addAction(self.toggle_dock_act)

    def update_run_state(self):
        current_tab = self.tab_widget.currentWidget()
        enabled = current_tab is not None and hasattr(current_tab, 'processing') and not current_tab.processing
        self.run_action.setEnabled(enabled)
        enabled = current_tab is not None and hasattr(current_tab, 'processing') and current_tab.processing
        self.stop_action.setEnabled(enabled)

        if current_tab is not None:
            if current_tab.total_tasks != 0:

                self.result_progress.setValue(int(100 * current_tab.executed_tasks / current_tab.total_tasks))
                self.result_progress_label.setText(
                    f"Выполнение задач... ({current_tab.executed_tasks}/{current_tab.total_tasks})")

                self.result_progress.setVisible(True)
                self.result_progress_label.setVisible(True)
            else:
                self.result_progress.setVisible(False)
                self.result_progress_label.setVisible(False)


    def start_response(self):
        current_tab: TablesTab = self.tab_widget.currentWidget()
        current_tab.processing = True

        fluids = current_tab.content.tbl_drilling_fluids
        tasks = countFilledRows(fluids)

        if tasks == 0:
            QMessageBox.critical(self, "Ошибка", "Нет данных")
            return

        current_tab.total_tasks = tasks

        self.update_run_state()

        """Начать сетевой запрос"""
        self.ws.open(QUrl(f"ws://{WS_HOST}:{WS_PORT}{WS_ENDPOINT}"))

        # Подготавливаем данные для отправки
        for i in range(fluids.rowCount()):
            try:
                payload = {
                    "name": fluids.item(i, 0).text(),
                    "depth_from": extract_number(fluids.item(i, 1).text()),
                    "depth_to": extract_number(fluids.item(i, 2).text()),
                    "transport_model": "Bingam",
                    "density": extract_number(fluids.item(i, 3).text()),
                    "viscosity": extract_number(fluids.item(i, 4).text()),
                    "dns": extract_number(fluids.item(i, 5).text())
                }
                message = {
                    "action": "create",
                    "payload": payload,
                    "tab_name": "fluids"
                }
                self.ws.sendTextMessage(json.dumps(message))


            except ValueError:  # строка недозаполнена
                continue


    def getTabByName(self, tab_name):
        for i in range(self.tab_widget.count()):
            if self.tab_widget.tabText(i) == tab_name:
                return self.tab_widget.widget(i)
        return None

    def handle_response(self, message: str):
        """Обработка ответа от сервера"""
        try:
            request_tab: TablesTab | None = self.task_id_to_tab.pop(reply, None)  # Извлекаем current_tab
            if request_tab is None:
                QMessageBox.critical(self, "Результат", "Проект был не найден или удален")
                return

            request_tab.executed_tasks += 1

            if reply.error() == QNetworkReply.NetworkError.NoError:
                # Читаем данные из ответа
                data = reply.readAll().data().decode('utf-8')
                # Парсим JSON
                result = json.loads(data)
                # Показываем результат
                self.create_results(result, request_tab)
            else:
                error_message = f"Ошибка запроса: {reply.errorString()}"
                QMessageBox.critical(self, "Ошибка", error_message)

        except json.JSONDecodeError as e:
            QMessageBox.critical(
                self,
                "Ошибка",
                f"Ошибка парсинга JSON: {str(e)}"
            )
        # except Exception as e:
        #     QMessageBox.critical(
        #         self,
        #         "Ошибка",
        #         f"Неизвестная ошибка: {str(e)}"
        #     )
        finally:
            if request_tab.name not in self.task_id_to_tab:
                request_tab.processing = False

            if request_tab is self.tab_widget.currentWidget():
                self.update_run_state()

            if request_tab.name not in self.task_id_to_tab:
                result_tab = request_tab.get_results_tab()
                request_tab.total_tasks = 0
                request_tab.executed_tasks = 0
                if result_tab is not None:
                    result_tab.total_tasks = 0
                    result_tab.executed_tasks = 0

    def open_file(self):
        file_path = QFileDialog.getExistingDirectory(self, "Открыть проект", "")
        if file_path:
            file_path = Path(file_path)
            self.tree_widget.load_project(file_path)

    def create_tables(self, project_item, name):
        tables_tab = TablesTab(project_item, name)
        self.tab_widget.addTab(tables_tab, name)
        self.tab_widget.setCurrentWidget(tables_tab)

        self.update_run_state()

        self.save_file_action.setEnabled(True)
        self.print_action.setEnabled(True)

    def create_results(self, data: dict, current_tab: TablesTab):
        name = get_name_for_results(current_tab.name)

        results_tab: ResultsTab | None = current_tab.get_results_tab()

        if results_tab is None:  # когда еще не создали вкладку под результаты
            result_project_item = self.tree_widget.create_results(current_tab.project_item, name)
            results_tab = ResultsTab(result_project_item, current_tab.name)
            results_tab.total_tasks = current_tab.total_tasks
            self.tab_widget.addTab(results_tab, name)

        results_tab.add_new_data(data)
        results_tab.executed_tasks = current_tab.executed_tasks

    def rename_tables(self, project_item: ProjectItem, name):
        tab: TablesTab = project_item.tab
        tab.name = name
        index = self.tab_widget.indexOf(tab)
        self.tab_widget.setTabText(index, name)

        if result_item := project_item.child(0):
            result_tab: ResultsTab = result_item.tab
            result_tab_name = get_name_for_results(name)
            result_tab.name = result_tab_name
            index = self.tab_widget.indexOf(result_tab)
            self.tab_widget.setTabText(index, result_tab_name)

    def delete_tables(self, project_item: ProjectItem):
        tab: TablesTab = project_item.tab
        index = self.tab_widget.indexOf(tab)
        self.tab_widget.removeTab(index)

        if result_item := project_item.child(0):
            result_tab: ResultsTab = result_item.tab
            index = self.tab_widget.indexOf(result_tab)
            self.tab_widget.removeTab(index)

    def set_tab(self, project_item: ProjectItem):
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


