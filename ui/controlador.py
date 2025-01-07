import os
from datetime import datetime
from consolidado_semanal import procesar_consolidado_semanal

class Controlador:
    def __init__(self):
        self.archivos_seleccionados = []
        self.fecha_inicial = None
        self.fecha_final = None
        self.actualizar_base = True
        self.verificar_base = True
        
    def validar_fechas(self, fecha_inicial, fecha_final):
        """Valida que las fechas sean correctas"""
        try:
            if fecha_inicial > fecha_final:
                raise ValueError("La fecha inicial no puede ser mayor que la fecha final")
            return True
        except Exception as e:
            raise Exception(f"Error en las fechas: {str(e)}")
            
    def validar_archivos(self, archivos):
        """Valida que los archivos necesarios estén presentes"""
        tipos_requeridos = ['procedimiento', 'arma', 'divisa', 'narcotrafico', 
                          'objetos', 'persona', 'vehiculo', 'OPER']
        
        archivos_encontrados = {tipo: False for tipo in tipos_requeridos}
        
        for archivo in archivos:
            nombre = archivo.lower()
            for tipo in tipos_requeridos:
                if tipo in nombre:
                    archivos_encontrados[tipo] = True
                    
        faltantes = [tipo for tipo, encontrado in archivos_encontrados.items() 
                    if not encontrado]
        
        if faltantes:
            raise Exception(f"Faltan archivos requeridos: {', '.join(faltantes)}")
        
        return True
    
    def iniciar_procesamiento(self, archivos, fecha_inicial, fecha_final, 
                            actualizar_base=True, verificar_base=True):
        """Inicia el proceso de consolidación"""
        try:
            # Validaciones
            self.validar_fechas(fecha_inicial, fecha_final)
            self.validar_archivos(archivos)
            
            # Guardar estado
            self.archivos_seleccionados = archivos
            self.fecha_inicial = fecha_inicial
            self.fecha_final = fecha_final
            self.actualizar_base = actualizar_base
            self.verificar_base = verificar_base
            
            # Procesar consolidado
            resultado = procesar_consolidado_semanal(archivos)
            
            if resultado:
                return {
                    'estado': 'éxito',
                    'mensaje': 'Procesamiento completado correctamente',
                    'fecha_proceso': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                }
            
        except Exception as e:
            raise Exception(f"Error en el procesamiento: {str(e)}")
