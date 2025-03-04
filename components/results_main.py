import sys

from PyQt6 import uic
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import QWidget, QPushButton, QRadioButton, QStackedWidget, QButtonGroup, \
    QMainWindow, QApplication, QSizePolicy, QFrame, QVBoxLayout, QLabel
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

from components.results_graphics import Results_graphics
from config import BASE_DIR


class Results(QWidget):

    def __init__(self, parent=None):
        super(Results, self).__init__(parent)
        uic.loadUi(BASE_DIR / 'results.ui', self, package='components')
        self.setup_ui()
        self.first_page_used = False
        self.current_result_index = 0


    def setup_ui(self):
        self.stackedWidget: QStackedWidget = self.findChild(QStackedWidget, 'stackedWidget')
        self.stackedWidget.setCurrentIndex(0)

        self.stackedWidget_2: QStackedWidget = self.findChild(QStackedWidget, 'stackedWidget_2')
        self.stackedWidget_2.setCurrentIndex(0)

        self.graphics_widget_1 = Results_graphics(self)

        page_widget_1 = self.findChild(QWidget, "page_interval_1")
        page_layout_1 = page_widget_1.layout()
        page_layout_1.addWidget(self.graphics_widget_1)

        self.button_group: QButtonGroup = self.findChild(QButtonGroup, 'buttonGroup')  # объединили кнопку в группу
        self.btn_go_to_next_page: QPushButton = self.findChild(QPushButton, 'pushButton_next_page')
        self.btn_go_to_previous_page: QPushButton = self.findChild(QPushButton, 'pushButton_previous_page')
        self.radio_btn_1: QRadioButton = self.findChild(QRadioButton, 'radioButton_page_1')
        self.radio_btn_2: QRadioButton = self.findChild(QRadioButton, 'radioButton_page_2')

        self.lbl_title_1: QLabel = self.findChild(QLabel, 'label_title')

        self.button_group.buttonClicked.connect(
            lambda btn: self.stackedWidget.setCurrentIndex(self.button_group.buttons().index(btn)))
        self.radio_btn_1.setChecked(True)

        self.btn_go_to_next_page.clicked.connect(self.go_to_next_page)
        self.btn_go_to_previous_page.clicked.connect(self.go_to_previous_page)

        self.stackedWidget_2.currentChanged.connect(self.on_current_index_changed)
        self.graphics: QFrame = self.findChild(QFrame, 'graphics')
        self.graphics_layout: QVBoxLayout = self.graphics.layout()
        self.init_graphics_views()
        self.btn_go_to_previous_page.setVisible(False)
        self.btn_go_to_previous_page.setVisible(False)

    def add_page_for_task(self, depth_from: float, depth_to: float):
        if not self.first_page_used:
            self.lbl_title_1.setText(f"Расчет для {depth_from} - {depth_to}")
            self.first_page_used = True

        else:
            new_page = Results_graphics(parent=self)
            #new_page.set_label(depth_from, depth_to)
            self.stackedWidget.addWidget(new_page)

            #self.stackedWidget.setCurrentWidget(new_page)

            # Создаем новую страницу динамически
            #page_widget = QWidget()
            #page_layout = QVBoxLayout(page_widget)

            #lbl_title = QLabel(f"Расчет для {depth_from} - {depth_to}")
            ##page_layout.addWidget(lbl_title)

            #graphics_widget = Results_graphics(self)
            #page_layout.addWidget(graphics_widget)
            #num_pages = self.stackedWidget.count()
            #insert_index = num_pages - 1

            #self.stackedWidget_2.addWidget(page_widget)
            #self.stackedWidget_2.insertWidget(insert_index + 1, page_widget)
            #self.stackedWidget_2.setCurrentIndex(self.stackedWidget_2.count() - 1)


    def get_page_count(self):
        return self.stackedWidget_2.count()

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

    def init_graphics_views(self):
        self.canvas_ECD_1 = FigureCanvas(Figure(figsize=(4, 4)))
        self.canvas_ECD_2 = FigureCanvas(Figure(figsize=(4, 4)))

        self.canvas_ECD_1.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.canvas_ECD_2.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

        self.set_canvas(self.canvas_ECD_1, "ЭЦП в трубном пространстве", "Эквивалентная плотность, кг/м^3")
        self.set_canvas(self.canvas_ECD_2, "ЭЦП на забое и в кольцевом пространстве",
                        "Эквивалентная плотность, кг/м^3")

        self.graphics_layout.addWidget(self.canvas_ECD_1, stretch=1)
        self.graphics_layout.addWidget(self.canvas_ECD_2, stretch=1)

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
        # Сохраняем axes в объекте canvas для дальнейшего использования
        canvas.axes = axes


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
