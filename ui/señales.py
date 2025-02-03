from PyQt6.QtCore import QObject, pyqtSignal

class Signals(QObject):
    progress = pyqtSignal(int)
    status = pyqtSignal(str)
    finished = pyqtSignal(bool)
    error = pyqtSignal(str)