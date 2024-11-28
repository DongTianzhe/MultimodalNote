from PySide6.QtGui import Qt
from PySide6.QtWidgets import QToolButton, QWidget, QGridLayout, QLabel, QApplication, QPushButton
import sys

class CollapseButton(QWidget):
    def __init__(self, text = '', parent = None):
        super().__init__(parent)
        self.text = text

        # Toggle button
        self.toggleButton = QToolButton()
        
        self.toggleButton.setStyleSheet('QToolButton {border: none;}')
        self.toggleButton.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon)
        self.toggleButton.setArrowType(Qt.ArrowType.RightArrow)
        self.toggleButton.setText(str(text))
        self.toggleButton.setCheckable(True)
        self.toggleButton.setChecked(False)
        self.toggleButton.clicked.connect(self.buttonCLicked)

        # content
        self.content = QWidget()
        self.content.setVisible(False)
        self.contentLayout = QGridLayout()
        # self.label = QLabel('label')
        # self.contentLayout.addWidget(self.label, 0, 0)
        self.content.setLayout(self.contentLayout)
        
        # 0 is button, 1 is content
        self.totalLayout = QGridLayout()
        self.totalLayout.setAlignment(Qt.AlignmentFlag.AlignTop)
        self.totalLayout.addWidget(self.toggleButton, 0, 0, Qt.AlignmentFlag.AlignLeft)
        self.totalLayout.addWidget(self.content, 1, 0)

        self.setLayout(self.totalLayout)
    
    def buttonCLicked(self, checked):
        if checked:
            self.arrowType = Qt.ArrowType.DownArrow
            self.content.setVisible(True)
        else:
            self.arrowType = Qt.ArrowType.RightArrow
            self.content.setVisible(False)
        self.toggleButton.setArrowType(self.arrowType)
    
    def addWidget(self, widget, verticalPos, horizontalPos, verticalLength=1, horizontalLength=1):
        self.contentLayout.addWidget(widget, verticalPos, horizontalPos, verticalLength, horizontalLength)

if __name__ == '__main__':
    app = QApplication(sys.argv)

    window = CollapseButton('button')
    window.resize(800, 600)
    label = QLabel('Label')
    button = QPushButton('Button')
    window.addWidget(label, 0, 0)
    window.addWidget(button, 1, 0)
    window.show()

    app.exec()
