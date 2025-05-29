import sys
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton,
    QFileDialog, QGridLayout, QMessageBox
)

from main import do_research
from handle_xlsx import parse_xlsx, create_result_table, save_xlsx

class XlsxProcessorGUI(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        # Widgets
        input_label = QLabel('Input XLSX:')
        self.input_path = QLineEdit()
        input_browse = QPushButton('Browse...')
        input_browse.clicked.connect(self.browse_input)

        output_label = QLabel('Output XLSX:')
        self.output_path = QLineEdit()
        output_browse = QPushButton('Browse...')
        output_browse.clicked.connect(self.browse_output)

        gid_label = QLabel('Group ID:')
        self.group_id = QLineEdit()

        group_size_label = QLabel('Group size:')
        self.group_size = QLineEdit()

        run_button = QPushButton('Run')
        run_button.clicked.connect(self.run_processing)

        # Layout
        layout = QGridLayout()
        layout.setContentsMargins(10, 10, 10, 10)  # Отступы вокруг layout
        layout.setSpacing(10)  # Расстояние между виджетами

        # Строка 0: Входной файл
        layout.addWidget(input_label, 0, 0)
        layout.addWidget(self.input_path, 0, 1)
        layout.addWidget(input_browse, 0, 2)

        # Строка 1: Выходной файл
        layout.addWidget(output_label, 1, 0)
        layout.addWidget(self.output_path, 1, 1)
        layout.addWidget(output_browse, 1, 2)

        # Строка 2: ID группы
        layout.addWidget(gid_label, 2, 0)
        layout.addWidget(self.group_id, 2, 1, 1, 2)  # Занимает 1 строку и 2 колонки

        # Строка 3: Размер группы
        layout.addWidget(group_size_label, 3, 0)
        layout.addWidget(self.group_size, 3, 1, 1, 2)  # Исправлено: было 2,1 теперь 3,1

        # Строка 4: Кнопка запуска (занимает всю ширину)
        layout.addWidget(run_button, 4, 0, 1, 3)  # Перенесено на строку 4

        self.setLayout(layout)
        self.setWindowTitle('XLSX Processor')
        self.resize(500, 150)

    def browse_input(self):
        file_name, _ = QFileDialog.getOpenFileName(self, 'Select Input File', '', 'Excel Files (*.xlsx)')
        if file_name:
            self.input_path.setText(file_name)

    def browse_output(self):
        file_name, _ = QFileDialog.getSaveFileName(self, 'Select Output File', '', 'Excel Files (*.xlsx)')
        if file_name:
            if not file_name.lower().endswith('.xlsx'):
                file_name += '.xlsx'
            self.output_path.setText(file_name)

    def run_processing(self):
        input_file = self.input_path.text().strip()
        output_file = self.output_path.text().strip()
        gid = self.group_id.text().strip()
        gsize = self.group_size.text().strip()

        if not input_file or not output_file or not gid or not gsize:
            QMessageBox.warning(self, 'Missing Data', 'Please specify all fields.')
            return

        try:
            group_size = int(gsize)
            if group_size <= 0:
                raise ValueError("Group size must be positive")
        except ValueError:
            QMessageBox.warning(self, 'Invalid Data', 'Group size must be a positive integer.')
            return

        try:
            do_research(input_file, output_file, gid, group_size)
            QMessageBox.information(self, 'Success', 'Processing completed successfully.')
        except Exception as e:
            QMessageBox.critical(self, 'Error', f'An error occurred:\n{str(e)}')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = XlsxProcessorGUI()
    window.show()
    sys.exit(app.exec_())
