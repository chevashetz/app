import io
import logging
import sys

import numpy as np
from PyQt6 import uic
from PyQt6.QtWidgets import QWidget, QPushButton, QRadioButton, QStackedWidget, QButtonGroup, \
    QGraphicsView, QMainWindow, QApplication, QMessageBox, QGraphicsScene, QSizePolicy
from matplotlib.backends.backend_template import FigureCanvas
from matplotlib.figure import Figure

from config import BASE_DIR


class Results(QWidget):

    def __init__(self, parent=None):
        super(Results, self).__init__(parent)
        uic.loadUi(BASE_DIR / 'results.ui', self, package='components')
        self.setup_ui()

    def setup_ui(self):
        self.button_group: QButtonGroup = self.findChild(QButtonGroup, 'buttonGroup')  # объединили кнопку в группу
        self.btn_go_to_next_page: QPushButton = self.findChild(QPushButton, 'pushButton_next_page')
        self.btn_go_to_previous_page: QPushButton = self.findChild(QPushButton, 'pushButton_previous_page')
        self.radio_btn_1: QRadioButton = self.findChild(QRadioButton, 'radioButton_page_1')
        self.radio_btn_2: QRadioButton = self.findChild(QRadioButton, 'radioButton_page_2')
        self.stackedWidget: QStackedWidget = self.findChild(QStackedWidget, 'stackedWidget')
        self.stackedWidget2: QStackedWidget = self.findChild(QStackedWidget, 'stackedWidget_2')
        self.graphicsView_resalt: QGraphicsView = self.findChild(QGraphicsView, 'graphicsView_3')
        self.graphicsView_resalt.setScene(QGraphicsScene())

        self.button_group.buttonClicked.connect(
            lambda btn: self.stackedWidget.setCurrentIndex(self.button_group.buttons().index(btn)))
        self.radio_btn_1.setChecked(True)
        self.stackedWidget.setCurrentIndex(0)
        self.stackedWidget2.currentChanged.connect(self.on_current_index_changed)

        self.init_graphics_views()

    def on_current_index_changed(self, index):
        total_pages = self.get_page_count()

        if index == 0:
            self.btn_go_to_previous_page.setVisible(False)
        else:
            self.btn_go_to_previous_page.setVisible(True)

        if index == total_pages - 1:
            self.btn_go_to_next_page.setVisible(False)
        else:
            self.btn_go_to_next_page.setVisible(True)

    def on_file_received(self, data: str):
        try:
            buf = io.StringIO(data)

            # Считаем весь текст как массив чисел.
            # По умолчанию split по любым пробельным символам (пробел, табуляция).
            arr = np.loadtxt(buf)

            # Теперь arr — двумерный numpy-массив shape = (число_строк, число_столбцов)
            # Получаем 1-й и 2-й столбцы (индексация с 0)
            col1 = arr[:, 0]
            col2 = arr[:, 1]
            #self.plot_graph(col1, col2)
        except Exception as e:
            logging.error(f"Failed to process in-memory data: {str(e)}")
            QMessageBox.critical(self, "Ошибка", f"Не удалось обработать данные в памяти: {str(e)}")

    def init_graphics_views(self):
        self.canvas_pressure = FigureCanvas(Figure(figsize=(7, 1.8)))
        self.canvas_gradient = FigureCanvas(Figure(figsize=(3.5, 1.8)))

        self.canvas_pressure.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.canvas_gradient.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)

        self.set_canvas(self.canvas_pressure, "Давления", "Глубина по вертикали, м")
        self.set_canvas(self.canvas_gradient, "Градиент давления", "Градиент давления, кгс/см2/м")

        self.graphics_layout.addWidget(self.canvas_pressure, stretch=23)
        self.graphics_layout.addWidget(self.canvas_gradient, stretch=9)

    def set_canvas(self, canvas: FigureCanvas, title: str, x_label: str):
        axes = canvas.figure.add_subplot(111)
        axes.set_title(title, fontsize=10)
        axes.set_xlabel(x_label, fontsize=8)
        axes.set_ylabel("Глубина по вертикали, м", fontsize=8)
        axes.tick_params(axis='x', labelsize=8)
        axes.invert_yaxis()
        axes.tick_params(axis='y', labelsize=8)
        axes.grid(True)
        canvas.figure.tight_layout()
    '''
    def plot_graph(self, col1, col2):
        try:
            # Создаем новую фигуру для графика
            figure = Figure()
            canvas = FigureCanvas(figure)
            ax = figure.add_subplot(111)

            # Строим график
            ax.plot(col1, col2, marker='o', linestyle='-', label="Data")
            ax.set_title("График данных")
            ax.set_xlabel("X-axis (1-й столбец)")
            ax.set_ylabel("Y-axis (2-й столбец)")
            ax.legend()
            ax.grid()

            # Удаляем предыдущий график из GraphicsView, если он есть
            for child in self.graphicsView.children():
                if isinstance(child, FigureCanvas):
                    child.setParent(None)

            # Встраиваем новый график в QGraphicsView
            scene = QGraphicsScene(self.graphicsView)
            proxy_widget = scene.addWidget(canvas)
            self.graphicsView.setScene(scene)

        except Exception as e:
            logging.error(f"Ошибка построения графика: {str(e)}")
            QMessageBox.critical(self, "Ошибка", f"Не удалось построить график: {str(e)}")
    '''

class ResultsTest(QMainWindow):
    """ТЕСТОВЫЙ КЛАСС, ЧТОБЫ ЗАПУСКАЛАСЬ Result_Tables"""

    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        ex = Results()
        self.setCentralWidget(ex)


# Тестирование tables
if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = ResultsTest()
    ex.show()
    sys.exit(app.exec())
