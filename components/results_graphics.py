import logging
import sys

from PyQt6 import uic
from PyQt6.QtWidgets import QWidget, QMainWindow, QApplication, QMessageBox, QSizePolicy, QFrame, QVBoxLayout, \
    QStackedWidget

from config import BASE_DIR
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure


class Results_graphics(QWidget):

    def __init__(self, parent=None):
        super().__init__(parent)
        uic.loadUi(BASE_DIR / 'results_graphics.ui', self, package='components')
        self.setup_ui()

    def setup_ui(self):
        self.stackedWidget: QStackedWidget = self.findChild(QStackedWidget, 'stackedWidget')
        self.stackedWidget.setCurrentIndex(0)

        self.graphics_1: QFrame = self.findChild(QFrame, 'graphics_1')
        self.graphics_layout_1: QVBoxLayout = self.graphics_1.layout()
        self.init_graphics_views()

    def on_file_received(self, data: str):
        pass
        '''
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
        '''

    def init_graphics_views(self):
        self.canvas_U_1 = FigureCanvas(Figure(figsize=(4, 4)))
        self.canvas_U_2 = FigureCanvas(Figure(figsize=(4, 4)))

        self.canvas_U_1.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.canvas_U_2.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

        self.set_canvas(self.canvas_U_1, "в трубном пространстве", ", м")
        self.set_canvas(self.canvas_U_2, "за ВЗД", ", м")

        self.graphics_layout_1.addWidget(self.canvas_U_1, stretch=1)
        self.graphics_layout_1.addWidget(self.canvas_U_2, stretch=1)

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

class ResultsTest(QMainWindow):
    """ТЕСТОВЫЙ КЛАСС, ЧТОБЫ ЗАПУСКАЛАСЬ Result_graphics_Tables"""

    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        ex = Results_graphics()
        self.setCentralWidget(ex)

    # Тестирование tables
if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = ResultsTest()
    ex.show()
    sys.exit(app.exec())
