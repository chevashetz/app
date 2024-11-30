import sys
from PyQt6 import uic
from PyQt6.QtWidgets import QApplication, QWidget, QMainWindow, QPushButton, QRadioButton, QGroupBox, QLabel, \
    QStackedWidget
from config import BASE_DIR

class Results(QWidget):
    def __init__(self, parent=None):
        super(Results, self).__init__(parent)
        uic.loadUi(BASE_DIR / 'results.ui', self, package='components')

    def setup_ui(self):
        self.group_box: QGroupBox = self.findChild(QGroupBox, 'groupBox')
        self.btn_go_to_next_page: QPushButton = self.findChild(QPushButton, 'pushButton_next_page')
        self.btn_go_to_previous_page: QPushButton = self.findChild(QPushButton, 'pushButton_previous_page')
        self.radio_btn_1: QRadioButton = self.group_box.findChild(QRadioButton, 'radioButton_page_1')
        self.radio_btn_2: QRadioButton = self.group_box.findChild(QRadioButton, 'radioButton_page_2')
        self.stackedWidget: QStackedWidget = self.findChild(QStackedWidget, 'stackedWidget')
        self.stackedWidget2: QStackedWidget = self.findChild(QStackedWidget, 'stackedWidget_2')

        self.radio_btn_1.setChecked(True)
        self.stackedWidget.setCurrentIndex(0)
        self.stackedWidget2.currentChanged.connect(self.on_current_index_changed)

        self.radio_btn_1.clicked.connect(self.switch_page_0)
        self.radio_btn_2.clicked.connect(self.switch_page_1)

    def switch_page_0(self):
        self.stackedWidget.setCurrentIndex(0)

    def switch_page_1(self):
        self.stackedWidget.setCurrentIndex(1)

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
