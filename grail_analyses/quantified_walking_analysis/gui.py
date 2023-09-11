from PyQt5.QtWidgets import (
    QMainWindow,
    QPushButton,
    QFileDialog,
    QTextEdit,
    QVBoxLayout,
    QHBoxLayout,
    QWidget,
    QLabel,
    QMessageBox,
)

from .file_io import process_files


class QwaAnalysis(QMainWindow):
    def __init__(self, *args, **kwargs):
        self.files: list[str, ...] = []

        # Declaring the meta aspect of the window
        super(QwaAnalysis, self).__init__(*args, **kwargs)
        self.setWindowTitle("File selectors")
        self.setGeometry(400, 400, 500, 300)

        self.main_layout = QVBoxLayout()
        widget = QWidget(self)
        widget.setLayout(self.main_layout)
        self.setCentralWidget(widget)

        # Declaring the UI
        selected_files_label = QLabel()
        selected_files_label.setText("Selected Files:")
        self.main_layout.addWidget(selected_files_label)

        self.selected_files_q_text = QTextEdit()
        self.selected_files_q_text.setReadOnly(True)
        self.main_layout.addWidget(self.selected_files_q_text)

        button_browse = QPushButton("Browse")
        button_browse.clicked.connect(self._select_files)
        self.main_layout.addWidget(button_browse)

        button_layout = QHBoxLayout()
        button_confirm = QPushButton("Confirm")
        button_confirm.clicked.connect(self._confirm)
        button_layout.addWidget(button_confirm)

        button_cancel = QPushButton("Cancel")
        button_cancel.clicked.connect(self._cancel)
        button_layout.addWidget(button_cancel)

        self.main_layout.addLayout(button_layout)

    def _select_files(self):
        files, _ = QFileDialog.getOpenFileNames(None, "QFileDialog.getOpenFileNames()", "", "CSV files (*.csv)")
        if not files:
            return

        self.selected_files_q_text.append(";".join(files))
        for i in range(len(files)):
            self.files.append(files[i])

    def _cancel(self):
        self.close()

    def _confirm(self):
        try:
            process_files(files=self.files)
        except IOError:
            QMessageBox.warning(
                self, "Erreur", "Le fichier d'exportation n'est pas accessible. S'il est ouvert, veuillez le fermer."
            )
            return
        except (Exception,):
            QMessageBox.warning(self, "Erreur", "Erreur inconnue dans le code")
            self.id_edit.clear()

        QMessageBox.warning(self, "Succès", "L'exportation a réussi")
        self.close()
