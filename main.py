import warnings
import pandas as pd
from pathlib import Path
from PyQt6.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox
from PyQt6.QtCore import QAbstractTableModel, Qt, QThread
from ui.interfaz_usuario import Ui_MainWindow
from ui.controlador import Controlador
import sys

# Suppress specific openpyxl warning about Data Validation
warnings.filterwarnings('ignore', category=UserWarning, 
                       module='openpyxl.worksheet._reader')

class FileTableModel(QAbstractTableModel):
    def __init__(self, files=None):
        super(FileTableModel, self).__init__()
        self._files = files or []

    def data(self, index, role):
        if role == Qt.ItemDataRole.DisplayRole:
            file = self._files[index.row()]
            if index.column() == 0:
                return Path(file).name
            
            skip_rows = 1 if 'OPER' in Path(file).name.upper() else 0
            df = pd.read_excel(file, skiprows=skip_rows)
            date_column = 'DENUNCIAFECHA' if 'DENUNCIAFECHA' in df.columns else 'FECHA'
            
            if index.column() in [1, 2]:
                if date_column in df.columns:
                    date_value = df[date_column].min() if index.column() == 1 else df[date_column].max()
                    return date_value.strftime('%d-%m-%Y')
                return "S/D"
            elif index.column() == 3:
                return len(df)

    def rowCount(self, index):
        return len(self._files)

    def columnCount(self, index):
        return 4

    def headerData(self, section, orientation, role):
        if role == Qt.ItemDataRole.DisplayRole:
            if orientation == Qt.Orientation.Horizontal:
                if section == 0:
                    return "Archivo"
                elif section == 1:
                    return "Fecha Inicial"
                elif section == 2:
                    return "Fecha Final"
                elif section == 3:
                    return "Filas"



class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.controlador = Controlador()
        
        # Inicializar lista de archivos y modelo
        self.file_list_model = []
        self.table_model = FileTableModel(self.file_list_model)
        self.loaded_files = set()  # Para controlar archivos ya cargados
        
        # Configurar TableView
        self.ui.tableView.setModel(self.table_model)
        self.ui.tableView.horizontalHeader().setStretchLastSection(True)  # type: ignore
        
        # Conectar señales
        self.ui.btn_select_files.clicked.connect(self.select_files)
        self.ui.btn_process.clicked.connect(self.process_data)
        self.ui.btn_save.clicked.connect(self.save_results)
        self.ui.btn_delete.clicked.connect(self.delete_files)
        
        self.ui.progressBar.setValue(0)
        self.ui.progressBar.setVisible(False)

    def select_files(self):
        try:
            self.ui.progressBar.setVisible(True)
            self.ui.progressBar.setValue(0)
            self.ui.label_status_process.setText("Seleccionando archivos...")
            
            files, _ = QFileDialog.getOpenFileNames(
                self,
                "Seleccionar archivos",
                "",
                "Excel Files (*.xlsx *.xls)"
            )
            
            if files:
                # Filtrar archivos que ya están cargados
                new_files = []
                duplicates = []
                for file in files:
                    if file not in self.loaded_files:
                        new_files.append(file)
                    else:
                        duplicates.append(Path(file).name)
                
                if duplicates:
                    self.ui.label_status_process.setText(f"Archivos duplicados: {', '.join(duplicates)}")
                    if not new_files:
                        self.ui.progressBar.setVisible(False)
                        return
                
                total_files = len(new_files)
                for i, file in enumerate(new_files, 1):
                    progress = int((i / total_files) * 90)
                    self.ui.progressBar.setValue(progress)
                    self.ui.label_status_process.setText(f"Cargando archivo {i} de {total_files}...")
                    
                    # Agregar archivo a la lista y al set de control
                    self.file_list_model.append(file)
                    self.loaded_files.add(file)
                
                # Actualizar tabla y contador
                self.table_model = FileTableModel(self.file_list_model)
                self.ui.tableView.setModel(self.table_model)
                self.table_model.layoutChanged.emit()
                
                # Actualizar label con contador total
                self.actualizar_contador()
                
                # Finalizar progreso
                self.ui.progressBar.setValue(100)
                self.ui.label_status_process.setText(
                    f"Se agregaron {len(new_files)} archivos nuevos" + 
                    (f" ({len(duplicates)} duplicados ignorados)" if duplicates else "")
                )
            
            self.ui.progressBar.setVisible(False)
                
        except Exception as e:
            self.ui.progressBar.setVisible(False)
            self.show_error("Error al seleccionar archivos", str(e))

    def delete_files(self):
        try:
            selected_indexes = self.ui.tableView.selectionModel().selectedIndexes() # type: ignore
            if not selected_indexes:
                raise Exception("No se han seleccionado archivos para eliminar")
            
            # Obtener archivos seleccionados
            selected_files = [self.file_list_model[index.row()] for index in selected_indexes]
            
            # Eliminar de ambas estructuras
            self.file_list_model = [f for f in self.file_list_model if f not in selected_files]
            self.loaded_files = self.loaded_files - set(selected_files)
            
            # Actualizar tabla
            self.ui.tableView.clearSelection()
            self.table_model = FileTableModel(self.file_list_model)
            self.ui.tableView.setModel(self.table_model)
            self.table_model.layoutChanged.emit()
            
            # Actualizar contador
            self.actualizar_contador()
            
        except Exception as e:
            self.show_error("Error al eliminar archivos", str(e))

    def actualizar_contador(self):
        """Actualiza el contador de archivos en el label"""
        total = len(self.file_list_model)
        if total == 0:
            self.ui.status_label.setText("No hay archivos cargados")
        else:
            tipos = len({Path(f).stem.split('_')[0].lower() for f in self.file_list_model})
            self.ui.status_label.setText(f"Archivos cargados: {total} ({tipos} tipos)")

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
        try:
            if not self.file_list_model:
                raise Exception("No se han seleccionado archivos")
                
            fecha_inicial = self.ui.date_start.date().toPyDate()
            fecha_final = self.ui.date_end.date().toPyDate()
            
            self.ui.progressBar.setVisible(True)
            self.ui.progressBar.setValue(0)
            self.ui.label_status_process.setText("Iniciando...")
            
            # Deshabilitar controles
            self.ui.btn_process.setEnabled(False)
            self.ui.btn_select_files.setEnabled(False)
            
            # Conectar señales
            self.controlador.progress.connect(self.ui.progressBar.setValue)
            self.controlador.status.connect(self.ui.label_status_process.setText)
            self.controlador.error.connect(self.on_process_error)
            
            # Iniciar procesamiento sin hilo separado
            self.controlador.iniciar_procesamiento(self.file_list_model, fecha_inicial, fecha_final ,self.ui.chk_verificar.isChecked()) 
            
            self.ui.label_status_process.setText("Finalizado...")
            self.ui.progressBar.setVisible(False)
            self.ui.btn_process.setEnabled(True)
            self.ui.btn_select_files.setEnabled(True)
        except Exception as e:
            self.show_error("Error al procesar datos", str(e))

        
    def on_process_error(self, error_msg):
        self.show_error("Error al procesar datos", error_msg)
        self.ui.btn_process.setEnabled(True)
        self.ui.btn_select_files.setEnabled(True)

# Código adicional para iniciar la aplicación
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())