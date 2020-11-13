from PyQt5.QtWidgets import QApplication
from src.main_window import MainWindow
import sys

if __name__ == '__main__':
    app = QApplication(sys.argv)
    screen = MainWindow()
    screen.showFullScreen()
    sys.exit(app.exec_())