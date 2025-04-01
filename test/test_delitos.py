import pytest
import pandas as pd
import os
from Funciones import cargar_delitos_codigos_desde_excel, procesar_tipo_delito_codigo
from Parametros import DELITOS_CODIGOS

# Fixture para cargar el diccionario antes de las pruebas
@pytest.fixture(scope="session")
def load_delitos_codigos():
    """Fixture que carga el diccionario DELITOS_CODIGOS una vez por sesión de pruebas"""
    # Obtener ruta absoluta del directorio actual (test/)
    current_dir = os.path.dirname(os.path.abspath(__file__))
    # Subir un nivel para llegar a la raíz del proyecto
    project_root = os.path.dirname(current_dir)
    # Construir ruta al archivo Matriz.xlsx
    ruta_matriz = os.path.join(project_root, "models", "Matriz.xlsx")
    
    print(f"\nBuscando archivo en: {ruta_matriz}")
    
    if os.path.exists(ruta_matriz):
        print(f"Archivo encontrado en {ruta_matriz}")
        try:
            delitos_codigos = cargar_delitos_codigos_desde_excel(ruta_matriz)
            print(f"Se cargaron {len(delitos_codigos)} códigos de delitos.")
            if len(delitos_codigos) == 0:
                print("ADVERTENCIA: El diccionario está vacío después de cargar el archivo.")
            return delitos_codigos
        except Exception as e:
            print(f"Error al cargar el archivo: {e}")
            raise
    else:
        print(f"No se encontró el archivo en {ruta_matriz}")
        print(f"Contenido del directorio models/:")
        models_dir = os.path.join(project_root, "models")
        if os.path.exists(models_dir):
            print(os.listdir(models_dir))
        else:
            print("El directorio models/ no existe")
        raise FileNotFoundError(f"No se encontró el archivo {ruta_matriz}")

def test_diccionario_no_vacio(load_delitos_codigos):
    """Prueba que el diccionario se haya cargado correctamente"""
    assert len(load_delitos_codigos) > 0, "El diccionario DELITOS_CODIGOS está vacío"

def test_procesar_delito_completo(load_delitos_codigos):
    """Prueba el procesamiento de un delito con las tres clasificaciones"""
    row = pd.Series({
        'CLASIFICACION_NIVEL_1': 'DELITOS CONTRA LA PROPIEDAD',
        'CLASIFICACION_NIVEL_2': 'HURTO',
        'CLASIFICACION_NIVEL_3': 'CALIFICADO'
    })
    resultado = procesar_tipo_delito_codigo(row)
    assert resultado is not None, "El resultado no debería ser None"
    print(f"Resultado para delito completo: {resultado}")

def test_procesar_delito_sin_nivel3(load_delitos_codigos):
    """Prueba el procesamiento de un delito sin clasificación nivel 3"""
    row = pd.Series({
        'CLASIFICACION_NIVEL_1': 'DELITOS CONTRA LA PROPIEDAD',
        'CLASIFICACION_NIVEL_2': 'HURTO',
        'CLASIFICACION_NIVEL_3': ''
    })
    resultado = procesar_tipo_delito_codigo(row)
    assert resultado is not None, "El resultado no debería ser None"
    print(f"Resultado para delito sin nivel 3: {resultado}")

def test_procesar_delito_no_existente(load_delitos_codigos):
    """Prueba el procesamiento de un delito que no existe en el diccionario"""
    row = pd.Series({
        'CLASIFICACION_NIVEL_1': 'DELITO INEXISTENTE',
        'CLASIFICACION_NIVEL_2': 'PRUEBA',
        'CLASIFICACION_NIVEL_3': 'TEST'
    })
    resultado = procesar_tipo_delito_codigo(row)
    assert resultado == 'DELITO INEXISTENTE - PRUEBA - TEST', \
           "Debería devolver la clave original para delitos no existentes"
    print(f"Resultado para delito no existente: {resultado}")