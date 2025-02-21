import sys

from PyQt6 import uic
from PyQt6.QtWidgets import QWidget, QMainWindow, QApplication

from config import BASE_DIR


class Results_graphics(QWidget):

    def __init__(self, parent=None):
        super(Results_graphics, self).__init__(parent)
        uic.loadUi(BASE_DIR / 'results_graphics.ui', self, package='components')
        self.setup_ui()

    def setup_ui(self):
        pass

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
