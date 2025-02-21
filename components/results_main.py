import io
import logging
import sys

import numpy as np
from PyQt6 import uic
from PyQt6.QtWidgets import QWidget, QPushButton, QRadioButton, QStackedWidget, QButtonGroup, \
    QGraphicsView, QMainWindow, QApplication, QMessageBox, QGraphicsScene, QSizePolicy, QFrame, QVBoxLayout, QLabel
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
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
        self.stackedWidget_2: QStackedWidget = self.findChild(QStackedWidget, 'stackedWidget_2')
        self.lbl_title_1: QLabel = self.findChild(QLabel, 'label_title_1')

        self.button_group.buttonClicked.connect(
            lambda btn: self.stackedWidget.setCurrentIndex(self.button_group.buttons().index(btn)))
        self.radio_btn_1.setChecked(True)
        self.stackedWidget.setCurrentIndex(0)
        self.stackedWidget_2.currentChanged.connect(self.on_current_index_changed)

        self.btn_go_to_next_page.clicked.connect(self.go_to_next_page)
        self.btn_go_to_previous_page.clicked.connect(self.go_to_previous_page)
        self.on_current_index_changed(0)

    def get_page_count(self):
        return self.stackedWidget.count()

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

    def go_to_next_page(self):
        current_index = self.stackedWidget_2.currentIndex()
        self.stackedWidget_2.setCurrentIndex(current_index + 1)

    def go_to_previous_page(self):
        current_index = self.stackedWidget_2.currentIndex()
        if current_index > 0:
            self.stackedWidget_2.setCurrentIndex(current_index - 1)

    def on_file_received(self, data: str):
        try:
            '''
            buf = io.StringIO(data)

            # Считаем весь текст как массив чисел.
            # По умолчанию split по любым пробельным символам (пробел, табуляция).
            arr = np.loadtxt(buf)

            # Теперь arr — двумерный numpy-массив shape = (число_строк, число_столбцов)
            # Получаем 1-й и 2-й столбцы (индексация с 0)
            col1 = arr[:, 0]
            col2 = arr[:, 1]
            #self.plot_graph(col1, col2)
            '''
        except Exception as e:
            logging.error(f"Failed to process in-memory data: {str(e)}")
            QMessageBox.critical(self, "Ошибка", f"Не удалось обработать данные в памяти: {str(e)}")
    '''
    def init_graphics_views(self):
        self.canvas_U_1 = FigureCanvas(Figure(figsize=(3, 2.5)))
        self.canvas_U_2 = FigureCanvas(Figure(figsize=(3, 2.5)))

        self.canvas_U_1.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.canvas_U_2.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

        self.set_canvas(self.canvas_U_1, "в трубном пространстве", ", м")
        self.set_canvas(self.canvas_U_2, "за ВЗД", ", м")

        self.graphics_layout_1.addWidget(self.canvas_U_1, stretch=1)
        self.graphics_layout_1.addWidget(self.canvas_U_2, stretch=1)

        #self.canvas_U_1.figure.suptitle('Главный заголовок для обоих графиков', fontsize=12)

    def set_canvas(self, canvas: FigureCanvas, title: str, x_label: str):
        axes = canvas.figure.add_subplot(111)
        axes.set_title(title, fontsize=10)
        axes.set_xlabel(x_label, fontsize=8)
        axes.set_ylabel("U, м/с", fontsize=8)
        axes.tick_params(axis='x', labelsize=8)
        axes.invert_yaxis()
        axes.tick_params(axis='y', labelsize=8)
        axes.grid(True)
        axes.invert_yaxis()
        axes.set_ylim(bottom=0)
        canvas.figure.tight_layout()
        canvas.axes = axes
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
