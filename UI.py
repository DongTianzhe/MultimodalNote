from PySide6.QtCore import QSize, Qt, QRect
from PySide6.QtWidgets import (
    QMainWindow, QPushButton, QGridLayout, QWidget, QLabel, QDockWidget, QCheckBox, QTableWidget,
    QHeaderView, QAbstractItemView, QTableWidgetItem, QTextEdit, QScrollArea, QSpinBox, QFileDialog,
    QMessageBox)
from PySide6.QtGui import QAction, QPixmap, QImage
from collapseButton import *
from const import *
from action import *
import xlwings as xw
import os
import time

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle('MultimodalNote')
        self.resize(400, 400)

        # self.button = QPushButton('button')
        # self.button.clicked.connect(self.buttonClicked)
        # self.secondCB = collapseButton.CollapseButton('Second', self)

        self.currentAction = 0
        self.actionList: list[Action] = [Action()]

        self.errorMessage = QMessageBox()

        # Action
        saveAction = QAction('&Save', self)
        saveAction.setShortcut('Ctrl+S')
        saveAction.setCheckable(False)
        saveAction.triggered.connect(self.saveFile)

        openAction = QAction('&Open', self)
        openAction.setShortcut('Ctrl+O')
        openAction.setCheckable(True)
        openAction.triggered.connect(self.openFile)

        self.showSelectionPanelAction = QAction('&Selection Panel', self)
        self.showSelectionPanelAction.setShortcut('Ctrl+D')
        self.showSelectionPanelAction.setCheckable(True)
        self.showSelectionPanelAction.triggered.connect(self.showSelectionPanel)

        self.addActionAtBottomAction = QAction('&Add action at buttom', self)
        self.addActionAtBottomAction.setShortcut('Ctrl+B')
        self.addActionAtBottomAction.setChecked(False)
        self.addActionAtBottomAction.triggered.connect(self.addActionAtBottom)

        # Dock widget
        self.dock  = QDockWidget('Selection Panel', self)
        self.dock.setVisible(False)
        # self.dock.setAllowedAreas(Qt.DockWidgetArea.LeftDockWidgetArea | Qt.DockWidgetArea.RightDockWidgetArea)
        self.dockLayout = QGridLayout()
        self.dockLayout.setAlignment(Qt.AlignmentFlag.AlignTop)

        # Add CollapseButton to Dock widget
        # Current
        self.currentCollapseButton = CollapseButton('Current', self)
        self.currentSceneLabel = QLabel('Scene')
        self.currentCollapseButton.addWidget(self.currentSceneLabel, 0, 0)
        self.currentSceneTextEdit = QTextEdit()
        self.currentSceneTextEdit.textChanged.connect(self.currentSceneTextEditChanged)
        self.currentSceneTextEdit.setFixedSize(50, 25)
        self.currentCollapseButton.addWidget(self.currentSceneTextEdit, 0, 1)
        self.currentActionLabel = QLabel('Action')
        self.currentCollapseButton.addWidget(self.currentActionLabel, 1, 0)
        self.currentActionTextEdit = QTextEdit()
        self.currentActionTextEdit.textChanged.connect(self.currentActionTextEditChanged)
        self.currentActionTextEdit.setFixedSize(50, 25)
        self.currentCollapseButton.addWidget(self.currentActionTextEdit, 1, 1)
        self.dockLayout.addWidget(self.currentCollapseButton)

        # Picture
        self.pictureCollapseButton = CollapseButton('Picture', self)
        self.pictureUploadButton = QPushButton('Upload Picture')
        self.pictureUploadButton.setFixedSize(150, 25)
        self.pictureUploadButton.clicked.connect(self.uploadPicture)
        self.pictureCollapseButton.addWidget(self.pictureUploadButton, 0, 0)
        self.dockLayout.addWidget(self.pictureCollapseButton)

        # Speech
        self.speechCollapseButton = CollapseButton('Speech', self)
        self.speechTextEdit = QTextEdit()
        self.speechTextEdit.textChanged.connect(self.speechTextEditChanged)
        self.speechCollapseButton.addWidget(self.speechTextEdit, 0, 0)
        self.dockLayout.addWidget(self.speechCollapseButton)
        
        # Music
        self.musicCollapseButton = CollapseButton('Music', self)
        self.musicScorePicture = CollapseButton('Score(Picture)', self)
        self.musicScorePictureUploadButton = QPushButton('Upload music score', self)
        self.musicScorePictureUploadButton.clicked.connect(self.uploadMusicScore)
        self.musicScorePicture.addWidget(self.musicScorePictureUploadButton, 0, 0)
        self.musicCollapseButton.addWidget(self.musicScorePicture, 0, 0)

        self.musicScoreText = CollapseButton('Score(Text)', self)
        self.copyToLilyPondFormButton = QPushButton('Copy to LilyPond form')
        self.copyToLilyPondFormButton.clicked.connect(self.actionList[self.currentAction].getLilyPondscript)
        self.copyToLilyPondFormButton.setFixedSize(150, 25)
        self.musicScoreText.addWidget(self.copyToLilyPondFormButton, 0, 0)
        self.scoreTextLabel = QLabel('Score text:')
        self.musicScoreText.addWidget(self.scoreTextLabel, 1, 0)
        self.scoreTextEdit = QTextEdit()
        self.scoreTextEdit.textChanged.connect(self.scoreTextEditChanged)
        self.musicScoreText.addWidget(self.scoreTextEdit, 2, 0)
        self.musicKeyLabel = QLabel('Key (Major): ')
        self.musicScoreText.addWidget(self.musicKeyLabel, 3, 0)
        self.musicKeyTextEdit = QTextEdit()
        self.musicKeyTextEdit.textChanged.connect(self.musicKeyTextEditChanged)
        self.musicKeyTextEdit.setFixedSize(50, 25)
        self.musicScoreText.addWidget(self.musicKeyTextEdit, 3, 1)
        self.musicTimeLabel = QLabel('Time: ')
        self.musicScoreText.addWidget(self.musicTimeLabel, 4, 0)
        self.musicScoreTimeNumeratorTextEdit = QTextEdit()
        self.musicScoreTimeNumeratorTextEdit.textChanged.connect(self.musicScoreTimeNumeratorTextEditChanged)
        self.musicScoreTimeNumeratorTextEdit.setFixedSize(50, 25)
        self.musicScoreText.addWidget(self.musicScoreTimeNumeratorTextEdit, 4, 1)
        self.musicScoreSlashLabel = QLabel('/')
        self.musicScoreText.addWidget(self.musicScoreSlashLabel, 4, 2)
        self.musicScoreTimeDenominatorTextEdit = QTextEdit()
        self.musicScoreTimeDenominatorTextEdit.textChanged.connect(self.musicScoreTimeDenominatorTextEditChanged)
        self.musicScoreTimeDenominatorTextEdit.setFixedSize(50, 25)
        self.musicScoreText.addWidget(self.musicScoreTimeDenominatorTextEdit, 4, 3)
        self.musicScoreTempoLabel = QLabel('Tempo (BPM): ')
        self.musicScoreText.addWidget(self.musicScoreTempoLabel, 5, 0)
        self.musicScoreTempoTextEdit = QTextEdit()
        self.musicScoreTempoTextEdit.textChanged.connect(self.musicScoreTempoTextEditChanged)
        self.musicScoreTempoTextEdit.setFixedSize(50, 25)
        self.musicScoreText.addWidget(self.musicScoreTempoTextEdit, 5, 1)

        self.musicScoreInstrumentsCollapseButton = CollapseButton('Instruments', self)
        self.musicScoreInstrumentVocalLabel = QLabel('Vocal: ')
        self.musicScoreInstrumentsCollapseButton.addWidget(self.musicScoreInstrumentVocalLabel, 0, 0)
        self.musicScoreInstrumentVocalSpinBox = QSpinBox()
        self.musicScoreInstrumentVocalSpinBox.textChanged.connect(self.musicScoreInstrumentVocalSpinBoxChanged)
        self.musicScoreInstrumentVocalSpinBox.setSingleStep(1)
        self.musicScoreInstrumentsCollapseButton.addWidget(self.musicScoreInstrumentVocalSpinBox, 0, 1)
        self.musicScoreInstrumentGuitarLabel = QLabel('Guitar: ')
        self.musicScoreInstrumentsCollapseButton.addWidget(self.musicScoreInstrumentGuitarLabel, 1, 0)
        self.musicScoreInstrumentGuitarSpinBox = QSpinBox()
        self.musicScoreInstrumentGuitarSpinBox.textChanged.connect(self.musicScoreInstrumentGuitarSpinBoxChanged)
        self.musicScoreInstrumentsCollapseButton.addWidget(self.musicScoreInstrumentGuitarSpinBox, 1, 1)
        self.musicScoreInstrumentGuitarSpinBox.setSingleStep(1)
        self.musicScoreINstrumentBassLabel = QLabel('Bass: ')
        self.musicScoreInstrumentsCollapseButton.addWidget(self.musicScoreINstrumentBassLabel, 2, 0)
        self.musicScoreInstrumentBassSpinBox = QSpinBox()
        self.musicScoreInstrumentBassSpinBox.textChanged.connect(self.musicScoreInstrumentBassSpinBoxChanged)
        self.musicScoreInstrumentBassSpinBox.setSingleStep(1)
        self.musicScoreInstrumentsCollapseButton.addWidget(self.musicScoreInstrumentBassSpinBox, 2, 1)
        self.musicScoreInstrumentDrumLabel = QLabel('Drum: ')
        self.musicScoreInstrumentsCollapseButton.addWidget(self.musicScoreInstrumentDrumLabel, 3, 0)
        self.musicScoreInstrumentDrumSpinBox = QSpinBox()
        self.musicScoreInstrumentDrumSpinBox.textChanged.connect(self.musicScoreInstrumentDrumSpinBoxChanged)
        self.musicScoreInstrumentDrumSpinBox.setSingleStep(1)
        self.musicScoreInstrumentsCollapseButton.addWidget(self.musicScoreInstrumentDrumSpinBox, 3, 1)
        self.musicScoreInstrumentTrebleClefLabel = QLabel('Treble clef: ')
        self.musicScoreInstrumentsCollapseButton.addWidget(self.musicScoreInstrumentTrebleClefLabel, 4, 0)
        self.musicScoreInstrumentTrebleClefSpinBox = QSpinBox()
        self.musicScoreInstrumentTrebleClefSpinBox.textChanged.connect(self.musicScoreInstrumentTrebleClefSpinBoxChanged)
        self.musicScoreInstrumentTrebleClefSpinBox.setSingleStep(1)
        self.musicScoreInstrumentsCollapseButton.addWidget(self.musicScoreInstrumentTrebleClefSpinBox, 4, 1)
        self.musicScoreInstrumentAltoClefLabel = QLabel('Alto clef: ')
        self.musicScoreInstrumentsCollapseButton.addWidget(self.musicScoreInstrumentAltoClefLabel, 5, 0)
        self.musicScoreInstrumentAltoClefSpinBox = QSpinBox()
        self.musicScoreInstrumentAltoClefSpinBox.textChanged.connect(self.musicScoreInstrumentAltoClefSpinBoxChanged)
        self.musicScoreInstrumentAltoClefSpinBox.setSingleStep(1)
        self.musicScoreInstrumentsCollapseButton.addWidget(self.musicScoreInstrumentAltoClefSpinBox, 5, 1)
        self.musicScoreInstrumentBassClefLabel = QLabel('Bass clef: ')
        self.musicScoreInstrumentsCollapseButton.addWidget(self.musicScoreInstrumentBassClefLabel, 6, 0)
        self.musicScoreInstrumentBassClefSpinBox = QSpinBox()
        self.musicScoreInstrumentBassClefSpinBox.textChanged.connect(self.musicScoreInstrumentBassClefSpinBoxChanged)
        self.musicScoreInstrumentBassClefSpinBox.setSingleStep(1)
        self.musicScoreInstrumentsCollapseButton.addWidget(self.musicScoreInstrumentBassClefSpinBox, 6, 1)

        self.musicScoreText.addWidget(self.musicScoreInstrumentsCollapseButton, 6, 0)

        self.musicCollapseButton.addWidget(self.musicScoreText, 1, 0)
        self.dockLayout.addWidget(self.musicCollapseButton)

        # Dramatic action
        self.dramaticActionCollapseButton = CollapseButton('Dramatic action', self)
        self.gestureCollapseButton = CollapseButton('Gesture', self)
        self.gestureTextEdit = QTextEdit()
        self.gestureTextEdit.textChanged.connect(self.gestureTextEditChanged)
        self.gestureCollapseButton.addWidget(self.gestureTextEdit, 0, 0)
        self.dramaticActionCollapseButton.addWidget(self.gestureCollapseButton, 0, 0)

        self.facialExpressionCollapseButton = CollapseButton('Facial express', self)
        self.facialExpressionTextEdit = QTextEdit()
        self.facialExpressionTextEdit.textChanged.connect(self.facialExpressionTextEditChanged)
        self.facialExpressionCollapseButton.addWidget(self.facialExpressionTextEdit, 0, 0)
        self.dramaticActionCollapseButton.addWidget(self.facialExpressionCollapseButton, 1, 0)

        self.movementCollapseButton = CollapseButton('Movement', self)
        self.movementTextEdit = QTextEdit()
        self.movementTextEdit.textChanged.connect(self.movementTextEditChanged)
        self.movementCollapseButton.addWidget(self.movementTextEdit, 0, 0)
        self.dramaticActionCollapseButton.addWidget(self.movementCollapseButton, 2, 0)

        # self.CollapseButton = collapseButton.CollapseButton('Text', self)
        # self.CollapseButton.addWidget(QLabel('label'), 0, 0)
        # self.CollapseButton.addWidget(QCheckBox(), 0, 1)
        # self.CollapseButton.addWidget(self.button, 1, 0)
        # self.secondCB.addWidget(QLabel('Second label'), 0, 0)
        # self.CollapseButton.addWidget(self.secondCB, 2, 0)
        # self.dockLayout.addWidget(self.CollapseButton)
        self.dockLayout.addWidget(self.dramaticActionCollapseButton)

        # Filming
        self.filmingCollapseButton = CollapseButton('Filming', self)
        self.shotCollapseButton = CollapseButton('Shot', self)
        self.shotTextEdit = QTextEdit()
        self.shotTextEdit.textChanged.connect(self.shotTextEditChanged)
        self.shotCollapseButton.addWidget(self.shotTextEdit, 0, 0)
        self.filmingCollapseButton.addWidget(self.shotCollapseButton, 0, 0)
        self.focalLensCollapseButton = CollapseButton('Focal Lens')
        self.focalLensTextEdit = QTextEdit()
        self.focalLensTextEdit.textChanged.connect(self.focalLensTextEditChanged)
        self.focalLensCollapseButton.addWidget(self.focalLensTextEdit, 0, 0)
        self.filmingCollapseButton.addWidget(self.focalLensCollapseButton, 1, 0)
        self.cameraMovementCollapseButton = CollapseButton('Camera Movement', self)
        self.cameraMovementTextEdit = QTextEdit()
        self.cameraMovementTextEdit.textChanged.connect(self.cameraMovementTextEditChanged)
        self.cameraMovementCollapseButton.addWidget(self.cameraMovementTextEdit, 0, 0)
        self.filmingCollapseButton.addWidget(self.cameraMovementCollapseButton, 2, 0)
        self.dockLayout.addWidget(self.filmingCollapseButton)

        # Editing
        self.editingCollapseButton = CollapseButton('Editing', self)
        self.transitionCollapseButton = CollapseButton('Transition', self)
        self.editingCollapseButton.addWidget(self.transitionCollapseButton, 0, 0)
        self.transitionTextEdit = QTextEdit()
        self.transitionTextEdit.textChanged.connect(self.transitionTextEditChanged)
        self.transitionCollapseButton.addWidget(self.transitionTextEdit, 0, 0)
        self.specialEffectCollapseButton = CollapseButton('Special effect')
        self.editingCollapseButton.addWidget(self.specialEffectCollapseButton, 1, 0)
        self.specialEffectTextEdit = QTextEdit()
        self.specialEffectTextEdit.textChanged.connect(self.specialEffectTextEditChanged)
        self.specialEffectCollapseButton.addWidget(self.specialEffectTextEdit, 0, 0)
        self.dockLayout.addWidget(self.editingCollapseButton)

        # Dock layout
        dockWidget = QWidget()
        dockWidget.setLayout(self.dockLayout)
        self.scrollArea = QScrollArea(self)
        self.scrollArea.setWidgetResizable(True)
        self.scrollArea.setWidget(dockWidget)
        self.dock.setWidget(self.scrollArea)
        self.addDockWidget(Qt.DockWidgetArea.LeftDockWidgetArea, self.dock)

        # Table
        self.table = QTableWidget(1, 8)
        # self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        # self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.setColumnWidth(0, 50)
        self.table.setColumnWidth(1, 50)
        self.table.setHorizontalHeaderLabels(headerList)
        # self.table.setItem(1, 1, QTableWidgetItem('123'))
        self.table.cellClicked.connect(self.rowSelect)
        
        # Layout
        self.layout = QGridLayout()

        self.layout.addWidget(self.table, 0, 0)

        widget = QWidget()
        widget.setLayout(self.layout)
        self.setCentralWidget(widget)

        # Menu
        menu = self.menuBar()

        fileMenu = menu.addMenu('&File')
        fileMenu.addAction(saveAction)
        fileMenu.addAction(openAction)

        viewMenu = menu.addMenu('&View')
        viewMenu.addAction(self.showSelectionPanelAction)
        viewMenu.addAction(self.addActionAtBottomAction)

        self.panelInitialise()
    
    def uploadPicture(self):
        filePath, _ = QFileDialog.getOpenFileName(self,'Select Picture', '', 'Images (*.png *.jpg *.jpeg *.bmp *.gif)')
        self.actionList[self.currentAction].picturePath = filePath
        self.showPicture(filePath, 2)
    
    def uploadMusicScore(self):
        filePath, _ = QFileDialog.getOpenFileName(self, 'Select Music Score', '', 'Images (*.png *.jpg *.jpeg *.bmp *.gif *.svg)')
        self.actionList[self.currentAction].musicScorePath = filePath
        self.showMusicScore(filePath, 4)
    
    def showMusicScore(self, path, column):
        try:
            pixmap = QPixmap(path)

            # rect = QRect(0, 0, pixmap.width(), 192)
            # croppedPixmap = pixmap.copy(rect)

            # croppedPixmap = pixmap.scaled(307, 192, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)

            item = QLabel()
            item.setScaledContents(True)
            item.setPixmap(pixmap)
            item.setAlignment(Qt.AlignmentFlag.AlignCenter)
            
            self.table.setRowHeight(self.currentAction, 192)
            self.table.setColumnWidth(column, 307)

            self.table.setCellWidget(self.currentAction, column, item)
        except BaseException as e:
            self.errorMessage.warning(self, 'Error', f'showPicture:\n{str(e)}')

    def showPicture(self, path, column):
        try:
            pixmap = QPixmap(path)
            pixmap = pixmap.scaled(307, 192, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)

            item = QLabel()
            item.setScaledContents(True)
            item.setPixmap(pixmap)
            item.setAlignment(Qt.AlignmentFlag.AlignCenter)
            
            self.table.setRowHeight(self.currentAction, 192)
            self.table.setColumnWidth(column, 307)

            self.table.setCellWidget(self.currentAction, column, item)
        except BaseException as e:
            self.errorMessage.warning(self, 'Error', f'showPicture:\n{str(e)}')
    
    def currentSceneTextEditChanged(self):
        self.actionList[self.currentAction].scene = self.currentSceneTextEdit.toPlainText()
        self.table.setItem(self.currentAction, 0, QTableWidgetItem(str(self.actionList[self.currentAction].scene)))
    
    def currentActionTextEditChanged(self):
        self.actionList[self.currentAction].action = self.currentActionTextEdit.toPlainText()
        self.table.setItem(self.currentAction, 1, QTableWidgetItem(str(self.actionList[self.currentAction].action)))
    
    def speechTextEditChanged(self):
        self.actionList[self.currentAction].speechText = self.speechTextEdit.toPlainText()
        self.table.setItem(self.currentAction, 3, QTableWidgetItem(self.actionList[self.currentAction].speechText))
    
    def scoreTextEditChanged(self):
        self.actionList[self.currentAction].scoreText = self.scoreTextEdit.toPlainText()
    
    def musicKeyTextEditChanged(self):
        self.actionList[self.currentAction].key != self.musicKeyTextEdit.toPlainText()
    
    def musicScoreTimeNumeratorTextEditChanged(self):
        self.actionList[self.currentAction].timeNumeratorText = self.musicScoreTimeNumeratorTextEdit.toPlainText()
    
    def musicScoreTimeDenominatorTextEditChanged(self):
        self.actionList[self.currentAction].timeDenominator = self.musicScoreTimeDenominatorTextEdit.toPlainText()
    
    def musicScoreTempoTextEditChanged(self):
        self.actionList[self.currentAction].tempo = self.musicScoreTempoTextEdit.toPlainText()
    
    def musicScoreInstrumentVocalSpinBoxChanged(self):
        self.actionList[self.currentAction].musicScoreInstrumentsCount['Vocal'] = self.musicScoreInstrumentVocalSpinBox.value()
    
    def musicScoreInstrumentGuitarSpinBoxChanged(self):
        self.actionList[self.currentAction].musicScoreInstrumentsCount['Guitar'] = self.musicScoreInstrumentGuitarSpinBox.value()
    
    def musicScoreInstrumentBassSpinBoxChanged(self):
        self.actionList[self.currentAction].musicScoreInstrumentsCount['Bass'] = self.musicScoreInstrumentBassSpinBox.value()
    
    def musicScoreInstrumentDrumSpinBoxChanged(self):
        self.actionList[self.currentAction].musicScoreInstrumentsCount['Drum'] = self.musicScoreInstrumentDrumSpinBox.value()
    
    def musicScoreInstrumentTrebleClefSpinBoxChanged(self):
        self.actionList[self.currentAction].musicScoreInstrumentsCount['Treble clef'] = self.musicScoreInstrumentTrebleClefSpinBox.value()
    
    def musicScoreInstrumentAltoClefSpinBoxChanged(self):
        self.actionList[self.currentAction].musicScoreInstrumentsCount['Alto clef'] = self.musicScoreInstrumentAltoClefSpinBox.value()
    
    def musicScoreInstrumentBassClefSpinBoxChanged(self):
        self.actionList[self.currentAction].musicScoreInstrumentsCount['Bass clef'] = self.musicScoreInstrumentBassClefSpinBox.value()
    
    def gestureTextEditChanged(self):
        self.actionList[self.currentAction].gestureText = self.gestureTextEdit.toPlainText()
        self.table.setItem(self.currentAction, 5, QTableWidgetItem(self.actionList[self.currentAction].getDramaticActionText()))
    
    def facialExpressionTextEditChanged(self):
        self.actionList[self.currentAction].facialExpressionText = self.facialExpressionTextEdit.toPlainText()
        self.table.setItem(self.currentAction, 5, QTableWidgetItem(self.actionList[self.currentAction].getDramaticActionText()))
    
    def movementTextEditChanged(self):
        self.actionList[self.currentAction].movementText = self.movementTextEdit.toPlainText()
        self.table.setItem(self.currentAction, 5, QTableWidgetItem(self.actionList[self.currentAction].getDramaticActionText()))
    
    def shotTextEditChanged(self):
        self.actionList[self.currentAction].shotText = self.shotTextEdit.toPlainText()
        self.table.setItem(self.currentAction, 6, QTableWidgetItem(self.actionList[self.currentAction].getFilmingText()))

    
    def focalLensTextEditChanged(self):
        self.actionList[self.currentAction].focalLensText = self.focalLensTextEdit.toPlainText()
        self.table.setItem(self.currentAction, 6, QTableWidgetItem(self.actionList[self.currentAction].getFilmingText()))

    
    def cameraMovementTextEditChanged(self):
        self.actionList[self.currentAction].cameraMovementText = self.cameraMovementTextEdit.toPlainText()
        self.table.setItem(self.currentAction, 6, QTableWidgetItem(self.actionList[self.currentAction].getFilmingText()))

    
    def transitionTextEditChanged(self):
        self.actionList[self.currentAction].transitionText = self.transitionTextEdit.toPlainText()
        self.table.setItem(self.currentAction, 7, QTableWidgetItem(self.actionList[self.currentAction].getEditingText()))
    
    def specialEffectTextEditChanged(self):
        self.actionList[self.currentAction].specialEffectText = self.specialEffectTextEdit.toPlainText()
        self.table.setItem(self.currentAction, 7, QTableWidgetItem(self.actionList[self.currentAction].getEditingText()))
    
    def panelInitialise(self):
        # self.table.setItem(self.currentAction, 0, QTableWidgetItem(self.actionList[self.currentAction].scene))
        self.currentSceneTextEdit.setText(self.actionList[self.currentAction].scene)
        # self.table.setItem(self.currentAction, 1, QTableWidgetItem(str(self.actionList[self.currentAction].action)))
        self.currentActionTextEdit.setText(self.actionList[self.currentAction].action)

        if self.actionList[self.currentAction].picturePath != '' and self.actionList[self.currentAction].picturePath != None:
            self.showPicture(self.actionList[self.currentAction].picturePath, 2)

        # self.table.setItem(self.currentAction, 3, QTableWidgetItem(str(self.actionList[self.currentAction].speechText)))
        self.speechTextEdit.setText(self.actionList[self.currentAction].speechText)

        if self.actionList[self.currentAction].musicScorePath != '' and self.actionList[self.currentAction].musicScorePath != None:
            self.showMusicScore(self.actionList[self.currentAction].musicScorePath, 4)
        self.scoreTextEdit.setText(self.actionList[self.currentAction].scoreText)
        self.musicKeyTextEdit.setText(self.actionList[self.currentAction].key)
        self.musicScoreTimeNumeratorTextEdit.setText(self.actionList[self.currentAction].timeNumeratorText)
        self.musicScoreTimeDenominatorTextEdit.setText(self.actionList[self.currentAction].timeDenominator)
        self.musicScoreTempoTextEdit.setText(self.actionList[self.currentAction].tempo)
        self.musicScoreInstrumentVocalSpinBox.setValue(self.actionList[self.currentAction].musicScoreInstrumentsCount['Vocal'])
        self.musicScoreInstrumentGuitarSpinBox.setValue(self.actionList[self.currentAction].musicScoreInstrumentsCount['Guitar'])
        self.musicScoreInstrumentBassSpinBox.setValue(self.actionList[self.currentAction].musicScoreInstrumentsCount['Bass'])
        self.musicScoreInstrumentDrumSpinBox.setValue(self.actionList[self.currentAction].musicScoreInstrumentsCount['Drum'])
        self.musicScoreInstrumentTrebleClefSpinBox.setValue(self.actionList[self.currentAction].musicScoreInstrumentsCount['Treble clef'])
        self.musicScoreInstrumentAltoClefSpinBox.setValue(self.actionList[self.currentAction].musicScoreInstrumentsCount['Alto clef'])
        self.musicScoreInstrumentBassClefSpinBox.setValue(self.actionList[self.currentAction].musicScoreInstrumentsCount['Bass clef'])

        self.table.setItem(self.currentAction, 5, QTableWidgetItem(self.actionList[self.currentAction].getDramaticActionText()))
        self.gestureTextEdit.setText(self.actionList[self.currentAction].gestureText)
        self.facialExpressionTextEdit.setText(self.actionList[self.currentAction].facialExpressionText)
        self.movementTextEdit.setText(self.actionList[self.currentAction].movementText)

        self.table.setItem(self.currentAction, 6, QTableWidgetItem(self.actionList[self.currentAction].getFilmingText()))
        self.shotTextEdit.setText(self.actionList[self.currentAction].shotText)
        self.focalLensTextEdit.setText(self.actionList[self.currentAction].focalLensText)
        self.cameraMovementTextEdit.setText(self.actionList[self.currentAction].cameraMovementText)

        self.table.setItem(self.currentAction, 7, QTableWidgetItem(self.actionList[self.currentAction].getEditingText()))
        self.transitionTextEdit.setText(self.actionList[self.currentAction].transitionText)
        self.specialEffectTextEdit.setText(self.actionList[self.currentAction].specialEffectText)

    
    def buttonClicked(self):
        button = self.sender()
        if isinstance(button, QPushButton):
            print(button.text())
    
    def rowSelect(self, row, column):
        self.currentAction = row
        self.panelInitialise()
    
    def showSelectionPanel(self):
        self.dock.setVisible(not self.dock.isVisible())
    
    def addActionAtBottom(self):
        self.table.insertRow(len(self.actionList))
        newAction = Action()
        newAction.action = str(int(self.actionList[-1].action) + 1)
        self.actionList.append(newAction)
    
    def saveFile(self):
        filePath, _ = QFileDialog.getSaveFileName(self, 'Save file', '', 'Excel file (*.xlsx)')
        print(filePath)

        if filePath:
            try:
                if not os.path.exists(filePath):
                    with xw.App(visible=False) as app:
                        book = app.books.add()
                        sheet = book.sheets.add()
                        sheet.name = 'Note'
                        sheet.range('a:a').api.NumberFormat = '@'
                        sheet.range('b:b').api.NumberFormat = '@'
                        sheet.range('A1').value = headerList
                        book.save(filePath)

                with xw.App(visible=False) as app:
                    book = app.books.open(filePath)
                    sheet = book.sheets['Note']
                    for i in range(len(self.actionList)):
                        sheet[f'A{i + 2}'].value = self.actionList[i].scene
                        sheet[f'B{i + 2}'].value = self.actionList[i].action
                        sheet[f'C{i + 2}'].value = self.actionList[i].picturePath
                        sheet[f'D{i + 2}'].value = self.actionList[i].speechText
                        sheet[f'E{i + 2}'].value = self.actionList[i].musicScorePath
                        sheet[f'F{i + 2}'].value = f'{self.actionList[i].gestureText}\n{self.actionList[i].facialExpressionText}\n{self.actionList[i].movementText}'
                        sheet[f'G{i + 2}'].value = f'{self.actionList[i].shotText}\n{self.actionList[i].focalLensText}\n{self.actionList[i].cameraMovementText}'
                        sheet[f'H{i + 2}'].value = f'{self.actionList[i].transitionText}\n{self.actionList[i].specialEffectText}'
                    book.save(filePath)

            except BaseException as e:
                self.errorMessage.warning(self, 'Error', f'saveFile:\n{str(e)}')

    def openFile(self):
        filePath, _ = QFileDialog.getOpenFileName(self, 'Open file', '', 'Excel file (*.xlsx)')
        with xw.App(visible=False) as app:
            book = app.books.open(filePath)
            sheet = book.sheets['Note']
            i = 2
            self.actionList = []
            self.table.clearContents()
            while sheet.range(i, 1).value != None:
                self.actionList.append(Action())
                self.actionList[-1].scene = str(sheet.range(i, 1).value)
                self.actionList[-1].action = str(sheet.range(i, 2).value)
                self.actionList[-1].picturePath = sheet.range(i, 3).value
                self.actionList[-1].speechText = sheet.range(i, 4).value
                self.actionList[-1].musicScorePath = sheet.range(i, 5).value
                self.actionList[-1].gestureText, self.actionList[-1].facialExpressionText, self.actionList[-1].movementText = sheet.range(i, 6).value.split('\n')
                self.actionList[-1].shotText, self.actionList[-1].focalLensText, self.actionList[-1].cameraMovementText = sheet.range(i, 7).value.split('\n')
                self.actionList[-1].transitionText, self.actionList[-1].specialEffectText = sheet.range(i, 8).value.split('\n')
                if i != 2:
                    self.table.insertRow(i - 2)
                self.currentAction = i - 2
                self.panelInitialise()
                i += 1
