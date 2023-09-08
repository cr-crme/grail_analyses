# -*- coding: utf-8 -*-
"""
Created on Wed Jul 26 11:32:50 2023

@author: Florence
"""

import sys

from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QPushButton, QLineEdit, QGridLayout, QMessageBox

from .file_io import export_results


def _export_results_callback(window, id_edit):
    try:
        export_results(id_edit.text())
    except RuntimeError:
        QMessageBox.warning(window, "Erreur", "Entrer un identifiant valide.")
        id_edit.clear()
    except IOError:
        QMessageBox.warning(window, "Erreur",
                            "Le fichier d'exportation n'est pas accessible. S'il est ouvert, veuillez le fermer.")
        id_edit.clear()
        return
    window.close()


def gui():
    app = QApplication(sys.argv)

    window = QWidget()
    window.setWindowTitle('Recherche de patients')
    window.setGeometry(500, 500, 300, 20)

    layout = QGridLayout()

    id_label = QLabel('ID patient', window)
    layout.addWidget(id_label, 0, 0)

    id_edit = QLineEdit(window)
    layout.addWidget(id_edit, 0, 1)

    search_button = QPushButton('Terminer', window)
    search_button.clicked.connect(lambda: _export_results_callback(window, id_edit))
    layout.addWidget(search_button, 4, 0, 1, 2)
    id_edit.returnPressed.connect(search_button.click)

    window.setLayout(layout)
    window.show()

    sys.exit(app.exec_())
