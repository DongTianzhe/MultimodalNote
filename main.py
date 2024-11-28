from PySide6.QtWidgets import QApplication
import sys
import UI

if __name__ == "__main__":
    app = QApplication(sys.argv)

    window = UI.MainWindow()
    window.show()

    app.exec()