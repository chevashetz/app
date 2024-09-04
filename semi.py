'''

        #self.add_button_to_page(2, 'New Button', (1400, 400), (400, 400))
        #self.add_button_with_form_layout(1, 'Button with FormLayout', (200, 100))

    def add_button_with_form_layout(self, page_index, button_text, size):
        page = self.stackedWidget.widget(page_index)
        button = QPushButton(button_text)
        button.setMinimumSize(size[0], size[1])
        button.setMaximumSize(size[0], size[1])
        button.setStyleSheet("background-color: lightblue; font-size: 16px;")

        layout = page.layout()
        if not layout or not isinstance(layout, QFormLayout):
            layout = QFormLayout(page)
            page.setLayout(layout)
        layout.addRow(button)
        button.setFixedSize(25, 25)
        button.move(1400, 75)
'''