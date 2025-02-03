from PyQt6.QtCore import QThread

class ProcessingWorker(QThread):
    def __init__(self, controlador, archivos, fecha_inicial, fecha_final):
        super().__init__()
        self.controlador = controlador
        self.archivos = archivos
        self.fecha_inicial = fecha_inicial
        self.fecha_final = fecha_final

    def run(self):
        self.controlador.iniciar_procesamiento(
            self.archivos,
            self.fecha_inicial,
            self.fecha_final
        )