from pathlib import Path

from PyQt6.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox
from PyQt6.QtCore import QStringListModel
from ui.interfaz_usuario import Ui_MainWindow
from ui.controlador import Controlador
import sys

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.controlador = Controlador()
        
        # Conectar señales
        self.ui.btn_select_files.clicked.connect(self.select_files)
        self.ui.btn_process.clicked.connect(self.process_data)
        self.ui.btn_save.clicked.connect(self.save_results)
        self.ui.btn_delete.clicked.connect(self.delete_files)
        
        # Inicializar modelo para la lista
        self.file_list_model = []
        self.list_model = QStringListModel()
        self.ui.listView.setModel(self.list_model)
        
    def select_files(self):
        """Permite seleccionar los archivos de entrada"""
        try:
            files, _ = QFileDialog.getOpenFileNames(
                self,
                "Seleccionar archivos",
                "",
                "Excel Files (*.xlsx *.xls)"
            )
            if files:
                new_files = [f for f in files if f not in self.file_list_model]
                self.file_list_model.extend(new_files)
                self.list_model.setStringList([Path(f).name for f in self.file_list_model])
                self.ui.status_label.setText(f"Archivos seleccionados: {len(self.file_list_model)}")
        except Exception as e:
            self.show_error("Error al seleccionar archivos", str(e))
            
    def delete_files(self):
        """Elimina los archivos seleccionados de la lista"""
        try:
            selected_indexes = self.ui.listView.selectedIndexes()
            if not selected_indexes:
                raise Exception("No se han seleccionado archivos para eliminar")
            
            selected_files = [self.file_list_model[index.row()] for index in selected_indexes]
            self.file_list_model = [f for f in self.file_list_model if f not in selected_files]
            self.list_model.setStringList(self.file_list_model)
            self.list_model.setStringList([Path(f).name for f in self.file_list_model])
        except Exception as e:
            self.show_error("Error al eliminar archivos", str(e))

    def show_error(self, title, message):
        """Muestra un diálogo de error"""
        QMessageBox.critical(self, title, message)
    def save_results(self):
        """Guarda los resultados procesados"""
        try:
            
            pass
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudieron guardar los resultados: {e}")

    def process_data(self):
        """Procesa los datos según las fechas seleccionadas"""
        try:
            if not self.file_list_model:
                raise Exception("No se han seleccionado archivos")
                
            fecha_inicial = self.ui.date_start.date().toPyDate()
            fecha_final = self.ui.date_end.date().toPyDate()
            
            self.ui.progressBar.setValue(50)
            self.ui.label_status_process.setText("Procesando...")
            
            resultado = self.controlador.iniciar_procesamiento(
                self.file_list_model,
                fecha_inicial,
                fecha_final
            )
            # Aquí puedes agregar el código para manejar el resultado del procesamiento
            self.ui.progressBar.setValue(100)
            self.ui.label_status_process.setText("Procesamiento completado")
        except Exception as e:
            self.show_error("Error al procesar datos", str(e))

# Código adicional para iniciar la aplicación
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())