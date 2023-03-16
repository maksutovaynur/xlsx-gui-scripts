import sys
from os import path as P
from typing import Iterable
from PyQt6.QtWidgets import QApplication, QMainWindow, QFileDialog, QPushButton, \
    QVBoxLayout, QHBoxLayout, QWidget, QErrorMessage, QLabel, QStatusBar
import pandas as pd

def V(*widgets):
    l = QVBoxLayout()
    for w in widgets:
        l.addWidget(w)
    w = QWidget()
    w.setLayout(l)
    return w

def H(*widgets):
    l = QHBoxLayout()
    for w in widgets:
        l.addWidget(w)
    w = QWidget()
    w.setLayout(l)
    return w

def B(title, callback):
    b = QPushButton(title)
    b.clicked.connect(callback)
    return b


class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
        self.err = QErrorMessage()
        self.setWindowTitle("Python File processor")
        self.open_button = B('Open files', self.open_files)
        self.process_button = B('Process', self.process_files)
        self.label = QLabel('')
        self.status = QStatusBar(self)
        self.setCentralWidget(V(H(self.open_button, self.process_button), self.label, self.status))
        self.process_button.setEnabled(False)
        self.hasProcessedData = False

    def open_files(self):
        self.open_button.setEnabled(False)
        self.status.showMessage('')
        fileNames, _ = QFileDialog.getOpenFileNames(self, "Open Excel", None, "Excel Files (*.xlsx)")

        if not fileNames:
            self.err.showMessage("No files were selected")
        else:
            fileNames = [P.abspath(f) for f in fileNames]
            for f in fileNames:
                if not P.isfile(f):
                    self.err.showMessage(f"File {f} does not exist")
            fileNames = [f for f in fileNames if P.isfile(f)]

        # Comment next 4 lines if more than one file is needed:
        if len(fileNames) != 1:
            fileNames = []
            self.err.showMessage('Only one file is allowed here')

        self.filesSelected = fileNames
        self.label.setText(', '.join(self.filesSelected))

        self.process_button.setEnabled(bool(fileNames))
        self.open_button.setEnabled(True)

    def process_files(self):
        self.process_button.setEnabled(False)
        self.open_button.setEnabled(False)
        self.status.showMessage('Processing ...')
        dfs = [pd.read_excel(f) for f in self.filesSelected]
        out_dfs = self.do_what_you_want(*dfs)
        for df in out_dfs:
            fileName, _ = QFileDialog.getSaveFileName(None, "Save Excel", None, "Excel Files (*.xlsx)")
            df.to_excel(fileName)
        self.label.setText('')
        self.status.showMessage('Done')
        self.open_button.setEnabled(True)



    def do_what_you_want(self, *dfs: pd.DataFrame) -> Iterable[pd.DataFrame]:
        # replace this with everything you want

        # Now it returns the first dataframe given. 
        # So, Total result of the program is just copying excel file to another excel file
        return [dfs[0]]


app = QApplication(sys.argv)

window = MainWindow()
window.show()

# Запускаем цикл событий.
app.exec()