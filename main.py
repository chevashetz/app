import json
import logging
import sys
import uuid
from enum import Enum
from pathlib import Path

import numpy as np
from PyQt6 import uic
from PyQt6.QtCore import QUrl, QObject, QTimer, QByteArray, pyqtSignal
from PyQt6.QtGui import QAction
from PyQt6.QtWebSockets import QWebSocket
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QMenu, QDockWidget, QFileDialog, QTabWidget, QMessageBox, QStatusBar, QProgressBar,
    QLabel
)
from components.project_tree import ProjectTree, ProjectItem, get_name_for_results
from components.results_main import Results
from components.results_graphics import Results_graphics
from components.tables import Tables
from components.tabs import TablesTab, ResultsTab
from config import RESULTS
from utils import extract_number, countFilledRows

logging.basicConfig(level=logging.ERROR, format='%(asctime)s - %(levelname)s - %(message)s')


class TaskStatus(str, Enum):
    CREATED = "created"
    CANCELED = "canceled"
    ERROR = "error"
    COMPLETED = "completed"


class MainWindow(QMainWindow):
    file_received_signal = pyqtSignal(str)
    task_completed_signal = pyqtSignal(float, float)

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        uic.loadUi('app.ui', self, package='components')

        self.setup_ui()
        self.setup_actions()
        self.results_widget = Results(self)
        self.results_graphics = Results_graphics(self)
        # У проекта свой вебсокет
        self.ws = QWebSocket()
        self.ws.textMessageReceived.connect(self.handle_response)
        self.ws.binaryMessageReceived.connect(self.file_response)
        self.ws.connected.connect(self.start_response)

        self.tab_widget.currentChanged.connect(self.update_run_state)
        self.tab_widget.tabCloseRequested.connect(self.update_run_state)

        self.project_id_to_project_item = {}
        self.update_run_state()
        #self.file_received_signal.connect(self.results_graphics.on_file_received)
        #self.task_completed_signal.connect(results_tab.content.add_page_for_task)

        self.project_id_to_data = {}

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
        self.run_action.triggered.connect(self.connect_server)

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

    def connect_server(self):
        # Открываем соединение
        self.ws.open(QUrl("ws://localhost:8000/ws"))

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

        project_id = str(uuid.uuid4())
        self.project_id_to_project_item[project_id] = current_tab.project_item

        # Подготовим список сообщений
        for i in range(fluids.rowCount()):
            try:
                payload = {
                    "name": "str",  # fluids.item(i, 0).text(),
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
                    "project_id": project_id
                }
                json_msg = json.dumps(message)
                self.ws.sendTextMessage(json_msg)
            except ValueError:
                continue

    def handle_response(self, message: str):
        """Обработка ответа от сервера через WebSocket"""
        try:
            data = json.loads(message)
            status = data.get('status')
            result = data.get('result')
            self.last_input_data = data.get('data')
            project_id = data.get('project_id')
            self.last_project_item = project_id
            if status == TaskStatus.CREATED:
                return

            project_item: ProjectItem | None = self.project_id_to_project_item.get(project_id)
            if project_item is None:
                QMessageBox.critical(self, "Результат", "Проект был не найден или удален")
                return
            tab: TablesTab = project_item.tab

            if status == TaskStatus.COMPLETED:
                tab.executed_tasks += 1
                self.create_results(data, tab)
                depth_from = self.last_input_data.get('depth_from')
                depth_to = self.last_input_data.get('depth_to')
                self.task_completed_signal.emit(depth_from, depth_to)

            elif status == TaskStatus.ERROR:
                QMessageBox.critical(self, "Ошибка", f"Проект {tab.name}, ошибка: {result}")
            elif status == TaskStatus.CANCELED:
                QMessageBox.information(self, "Отмена", f"Task {project_id} был отменен.")

            if tab.executed_tasks == tab.total_tasks:
                tab.processing = False
                self.project_id_to_project_item.pop(project_id)

            if tab is self.tab_widget.currentWidget():
                self.update_run_state()

            if tab.executed_tasks == tab.total_tasks:
                result_tab = tab.get_results_tab()
                tab.total_tasks = 0
                tab.executed_tasks = 0
                if result_tab is not None:
                    result_tab.total_tasks = 0
                    result_tab.executed_tasks = 0

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

    def file_response(self, result_file: QByteArray):
        """Обработка бинарного сообщения от сервера."""
        try:
            depth_from = self.last_input_data["depth_from"]
            depth_to = self.last_input_data["depth_to"]
            file_path = RESULTS / f"{self.last_project_item}{depth_from}{depth_to}.txt"  # или .bin, если данные бинарные
            with open(file_path, "wb") as f:
                f.write(result_file)
            logging.info(f"Binary file was received and saved as '{file_path}'")

            with open(file_path, 'r', encoding='utf-8') as file:
                data = file.read()

            self.file_received_signal.emit(data)

        except Exception as e:
            logging.error(f"Failed to handle binary data: {str(e)}")
            QMessageBox.critical(self, "Ошибка", f"Не удалось обработать бинарные данные: {str(e)}")

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
            self.task_completed_signal.connect(results_tab.content.add_page_for_task)

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

    def open_file(self):
        file_path = QFileDialog.getExistingDirectory(self, "Открыть проект", "")
        if file_path:
            file_path = Path(file_path)
            self.tree_widget.load_project(file_path)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec())
