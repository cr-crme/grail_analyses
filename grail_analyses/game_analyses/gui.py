# -*- coding: utf-8 -*-
"""
Created on Wed Jul 26 11:32:50 2023

@author: Florence
"""

from PyQt5.QtWidgets import QMainWindow, QWidget, QLabel, QPushButton, QLineEdit, QGridLayout, QMessageBox

from .file_io import export_results


class GameAnalysis(QMainWindow):
    def __init__(self, *args, **kwargs):
        # Declaring the meta aspect of the window
        super(GameAnalysis, self).__init__(*args, **kwargs)
        self.setWindowTitle("Recherche de patients")
        self.setGeometry(500, 500, 300, 20)

        self.main_layout = QGridLayout()
        widget = QWidget(self)
        widget.setLayout(self.main_layout)
        self.setCentralWidget(widget)

        # Declaring the UI
        id_label = QLabel("ID patient")
        self.main_layout.addWidget(id_label, 0, 0)

        self.id_edit = QLineEdit()
        self.main_layout.addWidget(self.id_edit, 0, 1)

        search_button = QPushButton("Terminer")
        search_button.clicked.connect(self._export_results_callback)
        self.main_layout.addWidget(search_button, 4, 0, 1, 2)

    def _export_results_callback(self):
        try:
            export_results(self.id_edit.text())
        except RuntimeError:
            QMessageBox.warning(self, "Erreur", "Entrer un identifiant valide.")
            self.id_edit.clear()
        except IOError:
            QMessageBox.warning(
                self, "Erreur", "Le fichier d'exportation n'est pas accessible. S'il est ouvert, veuillez le fermer."
            )
            self.id_edit.clear()
            return
        except (Exception,):
            QMessageBox.warning(self, "Erreur", "Erreur inconnue dans le code")
            self.id_edit.clear()

        QMessageBox.warning(self, "Succès", "L'exportation a réussi")
        self.close()
