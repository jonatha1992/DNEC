from pathlib import Path
from turtle import pd
from PyQt6.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox
from PyQt6.QtCore import QStringListModel
from ui.interfaz_usuario import Ui_MainWindow
from ui.controlador import Controlador
import sys
from PyQt6.QtCore import QAbstractTableModel, Qt
from PyQt6.QtWidgets import QMainWindow, QFileDialog, QMessageBox
import pandas as pd
from pathlib import Path
class FileTableModel(QAbstractTableModel):
    def __init__(self, files=None):
        super(FileTableModel, self).__init__()
        self._files = files or []

    def data(self, index, role):
        if role == Qt.ItemDataRole.DisplayRole:
            file = self._files[index.row()]
            if index.column() == 0:
                return Path(file).name
            elif index.column() == 1:
                df = pd.read_excel(file)
                return df['DENUNCIAFECHA'].min().strftime('%Y-%m-%d')
            elif index.column() == 2:
                df = pd.read_excel(file)
                return df['DENUNCIAFECHA'].max().strftime('%Y-%m-%d')

    def rowCount(self, index):
        return len(self._files)

    def columnCount(self, index):
        return 3

    def headerData(self, section, orientation, role):
        if role == Qt.ItemDataRole.DisplayRole:
            if orientation == Qt.Orientation.Horizontal:
                if section == 0:
                    return "Archivo"
                elif section == 1:
                    return "Fecha Inicial"
                elif section == 2:
                    return "Fecha Final"



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
        self.table_model = FileTableModel(self.file_list_model)
        self.ui.tableView.setModel(self.table_model)
        
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
                self.table_model.layoutChanged.emit()
                self.ui.status_label.setText(f"Archivos seleccionados: {len(self.file_list_model)}")
        except Exception as e:
            self.show_error("Error al seleccionar archivos", str(e))
            
    def delete_files(self):
        """Elimina los archivos seleccionados de la tabla"""
        try:
            selected_indexes = self.ui.tableView.selectionModel().selectedRows()
            if not selected_indexes:
                raise Exception("No se han seleccionado archivos para eliminar")
            
            selected_files = [self.file_list_model[index.row()] for index in selected_indexes]
            self.file_list_model = [f for f in self.file_list_model if f not in selected_files]
            self.table_model.layoutChanged.emit()
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
            
            self.ui.progressBar.setVisible(True)
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