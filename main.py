import sys

from grail_analyses import QwaAnalysis, GameAnalysis
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget


def run_analysis(main_window: QMainWindow, analyse_gui: QMainWindow):
    analyse_gui.show()
    main_window.hide()


def main():
    app = QApplication(sys.argv)
    qwa_analysis = QwaAnalysis()
    game_analysis = GameAnalysis()

    window = QMainWindow()
    window.setWindowTitle("Sélection de l'analyse")
    window.setGeometry(400, 400, 500, 150)

    main_layout = QVBoxLayout()
    widget = QWidget()
    widget.setLayout(main_layout)
    window.setCentralWidget(widget)

    aqm_button = QPushButton("AQM")
    aqm_button.clicked.connect(lambda: run_analysis(window, qwa_analysis))
    main_layout.addWidget(aqm_button)

    game_button = QPushButton("Analyse des jeux")
    game_button.clicked.connect(lambda: run_analysis(window, game_analysis))
    main_layout.addWidget(game_button)

    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
    