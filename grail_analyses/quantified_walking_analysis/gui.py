import sys

from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QFileDialog, QTextEdit, QVBoxLayout, QHBoxLayout, \
    QWidget, QLabel

from .file_io import process_files


def select_files(files_to_fill: list[str, ...], selected_files):
    files, _ = QFileDialog.getOpenFileNames(None, "QFileDialog.getOpenFileNames()", "", "CSV files (*.csv)")
    if files:
        selected_files.append(";".join(files))
        for i in range(len(files)):
            files_to_fill.append(files[i])


def cancel(window):  # Cancel button
    window.close()


def confirm(window, files: list[str, ...]):  # Confirm button
    window.close()
    process_files(files=files)


def gui():
    files = []
    app = QApplication(sys.argv)

    window = QMainWindow()
    window.setWindowTitle("File selectors")
    window.setGeometry(400, 400, 500, 300)

    layout = QVBoxLayout()

    selected_files_label = QLabel()
    selected_files_label.setText("Selected Files:")
    layout.addWidget(selected_files_label)

    selected_files = QTextEdit()
    selected_files.setReadOnly(True)
    layout.addWidget(selected_files)

    buttonBrowse = QPushButton("Browse")
    buttonBrowse.clicked.connect(lambda: select_files(files, selected_files))
    layout.addWidget(buttonBrowse)

    buttonLayout = QHBoxLayout()

    buttonConfirm = QPushButton("Confirm")
    buttonConfirm.clicked.connect(lambda: confirm(window, files))
    buttonLayout.addWidget(buttonConfirm)

    buttonCancel = QPushButton("Cancel")
    buttonCancel.clicked.connect(lambda: cancel(window))
    buttonLayout.addWidget(buttonCancel)

    layout.addLayout(buttonLayout)

    widget = QWidget()
    widget.setLayout(layout)
    window.setCentralWidget(widget)

    window.show()
    sys.exit(app.exec_())
