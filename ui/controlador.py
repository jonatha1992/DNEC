import os
from datetime import datetime
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment
from PyQt6.QtCore import QObject, pyqtSignal
from Funciones import *
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter


PATH_BASE = 'db/base_informada.xlsx'
PATH_TEMPLATE = 'models/modelo_informe.xlsx'
PATH_OUTPUT = 'informes/'

class Controlador(QObject):
    # Señales para comunicar progreso
    progress = pyqtSignal(int, str)
    error = pyqtSignal(str)
    completed = pyqtSignal(dict)

    def __init__(self):
        super().__init__()
        self.archivos = {}
        self.fecha_inicial = None
        self.fecha_final = None
        self.verificar_base = True
        self.CONTADOR = self.inicializar_contador()
        
    def inicializar_contador(self):
        """Inicializa el diccionario contador"""
        return {
            "OPERATIVOS_BASE":0,
            "PROCEDIMIENTOS_BASE":0,
            "FECHA_MENOR_BASE":datetime,
            "FECHA_MAYOR_BASE":datetime,
            'PROCEDIMIENTOS_NUEVOS': 0,
            'ORDEN_SERVICIOS_NUEVOS': 0,
            'BAJADA_ORDEN_SERVICIOS': 0,
            'BAJADA_PROCEDIMIENTOS': 0,
            'BAJADA_ARMAS': 0,
            'BAJADA_DIVISAS': 0,
            'BAJADA_NARCOTRAFICO': 0,
            'BAJADA_OBJETOS': 0,
            'BAJADA_PERSONAS': 0,
            'BAJADA_VEHICULOS': 0,
            'BAJADA_TRATA': 0,
            'GEOG_FINAL': 0,
            'VICTIMAS_FINAL': 0,
            'DETENIDOS_FINAL': 0,
            'ARMAS_FINAL': 0,
            'DIVISAS_FINAL': 0,
            'NARCOTRAFICO_FINAL': 0,
            'OBJETOS_FINAL': 0,
            'VEHICULOS_FINAL': 0,
            'TRATA_FINAL': 0,
        }

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
                          'objetos', 'persona', 'vehiculo', 'oper']
        
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

    def iniciar_procesamiento(self, archivos, fecha_inicial, fecha_final, verificar_base=True):
        """Inicia el proceso de consolidación"""
        try:
            # Validaciones iniciales
            self.validar_fechas(fecha_inicial, fecha_final)
            self.validar_archivos(archivos)
            
            # Guardar parámetros
            self.archivos = self.clasificar_archivos(archivos)
            self.fecha_inicial = fecha_inicial
            self.fecha_final = fecha_final
            self.verificar_base = verificar_base

            # 1. Procesar procedimientos (20%)
            self.progress.emit(10, "Procesando procedimientos...")
            df_procedimientos = self.procesar_procedimientos()
            
            self.progress.emit(20, "Procesando personas...")
            df_detenidos , df_otros_delitos = self.procesar_personas(df_procedimientos)
            
            # 3. Procesar incautaciones (50%)
            self.progress.emit(30, "Procesando incautaciones...")
            df_incautaciones = self.procesar_incautaciones(df_procedimientos)
            
            # 4. Procesar operaciones (70%)
            self.progress.emit(40, "Procesando operaciones...")
            df_operaciones, df_controlados, df_afectados, df_codigos = self.procesar_operaciones()
            
            self.progress.emit(50, "Procesando trata...")
            df_trata = self.procesar_trata(df_procedimientos)
            
            self.progress.emit(60, "Consolidando datos...")
            dataframes = self.consolidar_datos( df_procedimientos, df_operaciones, df_incautaciones, df_detenidos, df_otros_delitos, df_trata , df_afectados, df_controlados, df_codigos)
            
            self.progress.emit(80, "Generando informe...")
            self.copiar_formato_template(dataframes )
            

            # Completado
            self.progress.emit(100, "Procesamiento completado")
            self.completed.emit({
                'estado': 'éxito',
                'mensaje': 'Procesamiento completado correctamente',
                'fecha_proceso': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'estadisticas': self.CONTADOR
            })
            
            return True
            
        except Exception as e:
            self.error.emit(f"Error en el procesamiento: {str(e)}")
            raise

    def clasificar_archivos(self, archivos):
        """Clasifica los archivos según su tipo"""
        clasificados = {
            'procedimiento': None,
            'persona':None,
            'arma': None,
            'divisas': None,
            'narcotrafico': None,
            'objetos': None,
            'vehiculos': None,
            'operaciones': None,
            'trata': None
        }
        
        for archivo in archivos:
            nombre = archivo.lower()
            for tipo in clasificados.keys():
                if tipo in nombre:
                    clasificados[tipo] = archivo
                    break
                    
        return clasificados

    def procesar_procedimientos(self):
        """Procesa los archivos de procedimientos"""
        self.progress.emit(12, "Leyendo archivo de procedimientos...")
        
        df = pd.read_excel(self.archivos['procedimiento'])
        
        self.progress.emit(15, "Procesando procedimientos...")
        df_procedimientos = pd.DataFrame()
        df_procedimientos['ID_OPERATIVO'] = df.apply(procesar_causa_judicial, axis=1)
        df_procedimientos['FUERZA_INTERVINIENTE'] = "PSA"
        df_procedimientos['ID_PROCEDIMIENTO'] = df.apply(generar_uid_sigpol, axis=1)
        df_procedimientos['CAUSAJUDICIALNUMERO'] = df['CAUSAJUDICIALNUMERO'].copy()
        df_procedimientos['UNIDAD_INTERVINIENTE'] = df['UOSP'].fillna(df['URSA'])
        df_procedimientos['DESCRIPCIÓN'] = df.apply(procesar_descripcion, axis=1)
        df_procedimientos['TIPO_INTERVENCION'] = df.apply(procesar_tipo_procedimiento, axis=1)
        df_procedimientos['PROVINCIA'] = df.apply(procesar_provincia, axis=1)
        df_procedimientos['DEPARTAMENTO O PARTIDO'] = df.apply(procesar_municipio, axis=1)
        df_procedimientos['LOCALIDAD'] = "-"
        df_procedimientos['DIRECCION'] = df.apply(procesar_direccion, axis=1)
        df_procedimientos[['LATITUD', 'LONGITUD']] = df.apply(procesar_geog, axis=1, result_type='expand')
        df_procedimientos['FECHA'] = pd.to_datetime(df['DENUNCIAFECHA'], errors='coerce').dt.date
        df_procedimientos['HORA'] = pd.to_datetime(df['DENUNCIAFECHA'], errors='coerce').dt.strftime('%H:%M')
        df_procedimientos['ZONA_SEGURIDAD_FRONTERAS'] = "-"
        df_procedimientos['PASO_FRONTERIZO'] = "-"
        df_procedimientos['OTRAS AGENCIAS INTERVINIENTES'] = "-"
        df_procedimientos['Observaciones - Detalles'] = "-"
        df_procedimientos['LATITUD'] = df_procedimientos['LATITUD'].astype(str).str.replace(',', '.')
        df_procedimientos['LONGITUD'] = df_procedimientos['LONGITUD'].astype(str).str.replace(',', '.')

        df_procedimientos_completado = df_procedimientos[['FUERZA_INTERVINIENTE', 'ID_OPERATIVO', 'ID_PROCEDIMIENTO',
                                            'UNIDAD_INTERVINIENTE', 'DESCRIPCIÓN', 'TIPO_INTERVENCION',
                                            'PROVINCIA', 'DEPARTAMENTO O PARTIDO', 'LOCALIDAD', 'DIRECCION',
                                            'ZONA_SEGURIDAD_FRONTERAS', 'PASO_FRONTERIZO', 'LATITUD', 'LONGITUD',
                                            'FECHA', 'HORA', 'OTRAS AGENCIAS INTERVINIENTES', 'Observaciones - Detalles']]
        self.CONTADOR['BAJADA_PROCEDIMIENTOS'] = len(df)
        
        
        return df_procedimientos_completado

    def procesar_personas(self, df_procedimientos):
        """Procesa los archivos de personas"""
        self.progress.emit(32, "Procesando personas y detenidos...")
        df = pd.read_excel(self.archivos['persona'])
        
        if df.empty:
            return pd.DataFrame()
            
        self.CONTADOR['BAJADA_PERSONAS'] = len(df)
        
        # Generar UID y procesar datos
        df['UOSP'] = df['UOSP'].fillna(df['URSA'])
        df['ID_PROCEDIMIENTO'] = df.apply(generar_uid_sigpol, axis=1)
        # Paso 3: Coincidencia de UID entre excel_bajada_personas y df_geog_final
        uid_coincidentes = df[df['ID_PROCEDIMIENTO'].isin(df_procedimientos['ID_PROCEDIMIENTO'])]

        # Crear DataFrame de detenidos y aprendidos
        df_detenidos_aprendidos = pd.DataFrame()
        df_detenidos_aprendidos['ID_PROCEDIMIENTO'] = uid_coincidentes['ID_PROCEDIMIENTO']
        df_detenidos_aprendidos['EDAD'] = uid_coincidentes.apply(procesar_edad, axis=1)
        df_detenidos_aprendidos['SEXO'] = uid_coincidentes.apply(procesar_sexo, axis=1)
        df_detenidos_aprendidos['GENERO'] = uid_coincidentes.apply(procesar_genero, axis=1)
        df_detenidos_aprendidos['NACIONALIDAD'] = uid_coincidentes.apply(procesar_nacionalidad, axis=1)
        df_detenidos_aprendidos['SITUACION_PROCESAL'] = uid_coincidentes.apply(procesar_situacion_judicial, axis=1)
        df_detenidos_aprendidos['DELITO_IMPUTADO'] = uid_coincidentes.apply(procesar_tipo_delito, axis=1)
        df_detenidos_aprendidos['JUZGADO_INTERVINIENTE'] = uid_coincidentes.apply(procesar_juzgado, axis=1)
        df_detenidos_aprendidos['CARATULA_CAUSA'] = uid_coincidentes.apply(procesar_caratula, axis=1)
        df_detenidos_aprendidos['NUM_CAUSA'] = uid_coincidentes['ID_PROCEDIMIENTO']


        df_detenidos_aprendidos = df_detenidos_aprendidos[df_detenidos_aprendidos['SITUACION_PROCESAL'] != 'NO INFORMAR']
        df_detenidos_aprendidos = df_detenidos_aprendidos[df_detenidos_aprendidos['ID_PROCEDIMIENTO'].isin(df_procedimientos['ID_PROCEDIMIENTO'])]
        print(df_detenidos_aprendidos.nunique())
        
        
    
        if not df.empty:
            
            df_otros_delitos = df_detenidos_aprendidos[df_detenidos_aprendidos['SITUACION_PROCESAL'] == 'VICTIMA']
            # Filtro los que son victima de detenidos y aprendidos
            df_detenidos_aprendidos_completado = df_detenidos_aprendidos[df_detenidos_aprendidos['SITUACION_PROCESAL'] != 'VICTIMA']

            # Cambiar el nombre de algunas columnas en el DataFrame de víctimas
            df_otros_delitos_completado = df_otros_delitos.rename(columns={
                'DELITO_IMPUTADO': 'TIPO_OTRO_DELITO',
                'GENERO': 'GENERO_VICTIMA',
                'EDAD': 'EDAD_VICTIMA',
            }) # type: ignore
            df_otros_delitos_completado["OBSERVACIONES"] ="-"
            
            self.CONTADOR['DETENIDOS_FINAL'] = len(df_detenidos_aprendidos_completado['ID_PROCEDIMIENTO'])
            self.CONTADOR['VICTIMAS_FINAL'] = len(df_otros_delitos_completado['ID_PROCEDIMIENTO'])
            
            print(df_otros_delitos_completado.nunique())
                
        return df_detenidos_aprendidos_completado, df_otros_delitos_completado

    def procesar_incautaciones(self, df_procedimientos):
        """Procesa todas las incautaciones"""
        self.progress.emit(52, "Procesando armas...")
        df_armas = self.procesar_armas(df_procedimientos)
        
        self.progress.emit(54, "Procesando divisas...")
        df_divisas = self.procesar_divisas(df_procedimientos)
        
        self.progress.emit(56, "Procesando objetos...")
        df_objetos = self.procesar_objetos(df_procedimientos)
        
        self.progress.emit(58, "Procesando vehículos...")
        df_vehiculos = self.procesar_vehiculos(df_procedimientos)
        
        self.progress.emit(60, "Procesando narcotráfico...")
        df_narcotrafico = self.procesar_narcotrafico(df_procedimientos)
        
        df_incautaaciones = pd.concat([df_armas, df_divisas, df_objetos, df_vehiculos, df_narcotrafico])

        
        df_incautaaciones = df_incautaaciones[df_incautaaciones['ID_PROCEDIMIENTO'].isin(df_procedimientos['ID_PROCEDIMIENTO'])]
        return df_incautaaciones 

    def procesar_armas(self, df_procedimientos):
        """Procesa incautaciones de armas"""
        df = pd.read_excel(self.archivos['arma'])
        if df.empty:
            return pd.DataFrame()
            
        self.CONTADOR['BAJADA_ARMAS'] = len(df)

        df = df[df["TIPO_ESTADO_OBJETO"] == "SECUESTRADO"]
        df['UOSP'] = df['UOSP'].fillna(df['URSA'])
        df['TIPO_CAUSA_INTERNA'] = df.apply( procesar_tipo_causa_interna, axis=1)
        df['ID_PROCEDIMIENTO'] = df.apply(generar_uid_sigpol, axis=1)


        df_arma = pd.DataFrame()
        df_arma['ID_PROCEDIMIENTO'] = df['ID_PROCEDIMIENTO']
        df_arma['TIPO_INCAUTACION'] = "ARMAS"
        df_arma['TIPO'] = df['TIPO_ARMA']
        df_arma['SUBTIPO'] = "-"
        df_arma['CANTIDAD'] = df.apply(procesar_cantidad_arma, axis=1)
        df_arma['MEDIDAS'] = "UNIDADES"
        df_arma['AFORO'] = "-"
        df_arma['OBSERVACIONES'] = df.apply(procesar_observaciones_arma, axis=1)

        df_arma_completado = df_arma[df_arma['ID_PROCEDIMIENTO'].isin(df_procedimientos['ID_PROCEDIMIENTO'])]

        self.CONTADOR['ARMAS_FINAL'] = len(df_arma['ID_PROCEDIMIENTO'])
        print(df_arma_completado.nunique())

        return df_arma_completado

    def procesar_divisas(self, df_procedimientos):
        """Procesa incautaciones de divisas"""
        df = pd.read_excel(self.archivos['divisas'])
        if df.empty:
            return pd.DataFrame()
            
        self.CONTADOR['BAJADA_DIVISAS'] = len(df)

        df = df[df["TIPO_ESTADO_OBJETO"] == "SECUESTRADO"]
        df['UOSP'] = df['UOSP'].fillna(df['URSA'])
        df['TIPO_CAUSA_INTERNA'] = df.apply(procesar_tipo_causa_interna, axis=1)
        df['ID_PROCEDIMIENTO'] = df.apply(generar_uid_sigpol, axis=1)

        df_divisa = pd.DataFrame()
        df_divisa['ID_PROCEDIMIENTO'] = df['ID_PROCEDIMIENTO']
        df_divisa['TIPO_INCAUTACION'] = "DIVISAS"
        df_divisa['TIPO'] = df['TIPO_MONEDA'] 
        df_divisa['SUBTIPO'] = "-"
        df_divisa['CANTIDAD'] = df['CANTIDAD']
        df_divisa['MEDIDAS'] = df['TIPO_MONEDA']
        df_divisa['AFORO'] = df['VALOR_MONEDA'].fillna('-')
        df_divisa['OBSERVACIONES'] = "-"

        df_divisa_completado = df_divisa[df_divisa['ID_PROCEDIMIENTO'].isin(df_procedimientos['ID_PROCEDIMIENTO'])]

        self.CONTADOR['DIVISAS_FINAL'] = len(df_divisa_completado['ID_PROCEDIMIENTO'])
        print(df_divisa_completado.nunique())

        return df_divisa_completado
        

    def procesar_objetos(self, df_procedimientos):
        """Procesa incautaciones de objetos"""
        df = pd.read_excel(self.archivos['objetos'])
        if df.empty:
            return pd.DataFrame()
            
        self.CONTADOR['BAJADA_OBJETOS'] = len(df)


        df = df[df["TIPO_ESTADO_OBJETO"] == "SECUESTRADO"]
        df['UOSP'] = df['UOSP'].fillna(df['URSA'])
        df['TIPO_CAUSA_INTERNA'] = df.apply(procesar_tipo_causa_interna, axis=1)
        df['ID_PROCEDIMIENTO'] = df.apply(generar_uid_sigpol, axis=1)

        df_objetos = pd.DataFrame()
        df_objetos['ID_PROCEDIMIENTO'] = df['ID_PROCEDIMIENTO']
        df_objetos['TIPO_INCAUTACION'] = "OBJETOS"
        df_objetos['TIPO'] = df['OBJETO_SECUESTRADO'] 
        df_objetos['SUBTIPO'] = "-"
        df_objetos['CANTIDAD'] = df['CANTIDAD']
        df_objetos['MEDIDAS'] = df['MEDIDA']
        df_objetos['AFORO'] = df['VALOR_MONEDA'].fillna('-')
        df_objetos['OBSERVACIONES'] = "-"

        df_objetos_completado = df_objetos[df_objetos['ID_PROCEDIMIENTO'].isin(df_procedimientos['ID_PROCEDIMIENTO'])]

        self.CONTADOR['OBJETOS_FINAL'] = len(df_objetos_completado['ID_PROCEDIMIENTO'])
        print(df_objetos_completado.nunique())

        return df_objetos_completado

    def procesar_vehiculos(self, df_procedimientos):
        """Procesa incautaciones de vehículos"""
        df = pd.read_excel(self.archivos['vehiculos'])
        if df.empty:
            return pd.DataFrame()
            
        self.CONTADOR['BAJADA_VEHICULOS'] = len(df)

        df = df[df["TIPO_ESTADO_OBJETO"] == "SECUESTRADO"]
        df['UOSP'] = df['UOSP'].fillna(df['URSA'])
        df['TIPO_CAUSA_INTERNA'] = df.apply(procesar_tipo_causa_interna, axis=1)
        df['ID_PROCEDIMIENTO'] = df.apply(generar_uid_sigpol, axis=1)

        df_vehiculos = pd.DataFrame()
        df_vehiculos['ID_PROCEDIMIENTO'] = df['ID_PROCEDIMIENTO']
        df_vehiculos['TIPO_INCAUTACION'] = "VEHICULOS"
        df_vehiculos['TIPO'] = df['TIPO_VEHICULO']
        df_vehiculos['SUBTIPO'] = df['MARCA_VEHICULO']
        df_vehiculos['CANTIDAD'] = 1
        df_vehiculos['MEDIDAS'] = "UNIDADES"
        df_vehiculos['AFORO'] = "-"
        df_vehiculos['OBSERVACIONES'] = df.apply(observaciones_vehiculo, axis=1)

        df_vehiculos_completado = df_vehiculos[df_vehiculos['ID_PROCEDIMIENTO'].isin(df_procedimientos['ID_PROCEDIMIENTO'])]

        self.CONTADOR['VEHICULOS_FINAL'] = len(df_vehiculos_completado['ID_PROCEDIMIENTO'])

        return df_vehiculos_completado

    def procesar_narcotrafico(self, df_procedimientos):
        """Procesa incautaciones de narcotráfico"""
        df = pd.read_excel(self.archivos['narcotrafico'])
        if df.empty:
            return pd.DataFrame()
            
        self.CONTADOR['BAJADA_NARCOTRAFICO'] = len(df)

        df['UOSP'] = df['UOSP'].fillna(df['URSA'])
        df['TIPO_CAUSA_INTERNA'] = df.apply(procesar_tipo_causa_interna, axis=1)
        df['ID_PROCEDIMIENTO'] = df.apply(generar_uid_sigpol, axis=1)

        df_narcotrafico = pd.DataFrame()
        df_narcotrafico['ID_PROCEDIMIENTO'] = df['ID_PROCEDIMIENTO']
        df_narcotrafico['TIPO_INCAUTACION'] = "ESTUPEFACIENTE"
        df_narcotrafico['TIPO'] = df.apply(clasificar_tipo_sustancia, axis=1)
        df_narcotrafico['SUBTIPO'] = "-"
        df_narcotrafico[['CANTIDAD', 'MEDIDAS']] = df.apply(clasificar_medida, axis=1, result_type='expand')
        df_narcotrafico['AFORO'] = "-"
        df_narcotrafico['OBSERVACIONES'] = df.apply(observaciones_sustancia, axis=1)

        df_narcotrafico_completado = df_narcotrafico[df_narcotrafico['ID_PROCEDIMIENTO'].isin(df_procedimientos['ID_PROCEDIMIENTO'])]
        self.CONTADOR['NARCOTRAFICO_FINAL'] = len(df_narcotrafico_completado['ID_PROCEDIMIENTO'])

        return df_narcotrafico_completado

    def procesar_operaciones(self):
        """Procesa las operaciones y órdenes de servicio"""
        self.progress.emit(72, "Procesando operaciones...")
        df = pd.read_excel(self.archivos['operaciones'], sheet_name="ORDEN_SERVICIOS", skiprows=1)
        
        self.CONTADOR['BAJADA_ORDEN_SERVICIOS'] = len(df)
        
        # Procesar operaciones principales
        df_operaciones = self.procesar_operaciones_principales(df)
        
        # Procesar datos relacionados
        df_controlados = self.procesar_controlados(df)
        df_afectados = self.procesar_afectados(df)
        df_codigos = self.procesar_codigos_operativos(df)
        
        return df_operaciones, df_controlados, df_afectados, df_codigos

    def procesar_operaciones_principales(self, df):
        """Procesa las operaciones principales"""
        if df.empty:
            return pd.DataFrame()
        
        # Procesar operaciones
        df_operaciones = pd.DataFrame()
        df_operaciones["ID_PROCEDIMIENTO"] = df["ID_PROCEDIMIENTO"]
        df_operaciones["FUERZA_INTERVINIENTE"] = "PSA"
        df_operaciones["ID_OPERATIVO"] = df["ID_OPERATIVO"]
        df_operaciones["UNIDAD_INTERVINIENTE"] = df["UNIDAD_INTERVINIENTE"]
        df_operaciones["DESCRIPCIÓN"] = df["DESCRIPCIÓN"]
        df_operaciones["TIPO_INTERVENCION"] = df["TIPO_INTERVENCION"]
        df_operaciones["PROVINCIA"] = df["PROVINCIA"].str.replace("_", " ", regex=False)
        df_operaciones["DEPARTAMENTO O PARTIDO"] = df["DEPARTAMENTO O PARTIDO"].str.upper()
        df_operaciones["LOCALIDAD"] = df["LOCALIDAD"]
        df_operaciones["DIRECCION"] = df["DIRECCION"]
        df_operaciones['FECHA'] = pd.to_datetime(df["FECHA"]).dt.date
        df_operaciones['HORA'] = df['HORA'].apply(lambda x: x.strftime('%H:%M') if pd.notnull(x) else None)
        df_operaciones["ZONA_SEGURIDAD_FRONTERAS"] = "-"
        df_operaciones["PASO_FRONTERIZO"] = "-"
        df_operaciones['OTRAS AGENCIAS INTERVINIENTES'] = df["OTRAS AGENCIAS INTERVINIENTES"]
        df_operaciones['Observaciones - Detalles'] = "PATRULLAJE DINAMICO"
        df_operaciones[['LATITUD', 'LONGITUD']] = df.apply(procesar_geog_oper, axis=1, result_type='expand')

        return df_operaciones

    def procesar_controlados(self, df):
        """Procesa los datos de vehículos y personas controladas"""
        if df.empty:
            return pd.DataFrame()
            
        df_controlados = pd.DataFrame()
        df_controlados["ID_PROCEDIMIENTO"] = df["ID_PROCEDIMIENTO"]
        df_controlados["FUERZA_INTERVINIENTE"] = "PSA"
        df_controlados["ID_OPERATIVO"] = df["ID_OPERATIVO"]
        df_controlados["UNIDAD_INTERVINIENTE"] = df["UNIDAD_INTERVINIENTE"]
        df_controlados["VEHICULOS_CONTROLADOS"] = df["VEHICULOS_CONTROLADOS"]
        df_controlados["PERSONAS_CONTROLADAS"] = df["PERSONAS_CONTROLADAS"]
        df_controlados["AVERIGUACIONES"] = df["CANT_AVERIGUACIONES_SECUESTRO"]
        df_controlados["ANTECEDENTES"] = df["CANT_SOLICITUDES_ANTECEDENTES"]
        
        return df_controlados

    def procesar_afectados(self, df):
        """Procesa los datos del personal y elementos afectados"""
        if df.empty:
            return pd.DataFrame()
            
        df_afectados = pd.DataFrame()
        df_afectados["ID_PROCEDIMIENTO"] = df["ID_PROCEDIMIENTO"]
        df_afectados["FUERZA_INTERVINIENTE"] = "PSA"
        df_afectados["ID_OPERATIVO"] = df["ID_OPERATIVO"]
        df_afectados["UNIDAD_INTERVINIENTE"] = df["UNIDAD_INTERVINIENTE"]
        df_afectados["CANT_EFECTIVOS"] = df["CANT_EFECTIVOS"]
        df_afectados["CANT_VEHICULOS"] = df["CANT_AUTOS_CAMIONETAS"]
        df_afectados["CANT_MOTOS"] = df["CANT_MOTOS"]
        df_afectados["CANT_SCANNERS"] = df["CANT_SCANNERS"]
        df_afectados["CANT_CANES"] = df["CANT_CANES"]
        df_afectados["OTROS_ELEMENTOS"] = df.apply(self.procesar_otros_elementos, axis=1)
        
        return df_afectados

    def procesar_otros_elementos(self, row):
        """Procesa otros elementos afectados"""
        elementos = []
        if row["CANT_EMBARCACIONES"] > 0:
            elementos.append(f"EMBARCACIONES: {row['CANT_EMBARCACIONES']}")
        if row["CANT_CABALLOS"] > 0:
            elementos.append(f"CABALLOS: {row['CANT_CABALLOS']}")
        if row["CANT_MORPHRAPID"] > 0:
            elementos.append(f"MORPHRAPID: {row['CANT_MORPHRAPID']}")
        if row["CANT_LPR"] > 0:
            elementos.append(f"LPR: {row['CANT_LPR']}")
        
        return " - ".join(elementos) if elementos else "-"

    def procesar_codigos_operativos(self, df):
        """Procesa los códigos operativos"""
        if df.empty:
            return pd.DataFrame()
            
        df_codigos = pd.DataFrame()
        df_codigos["ID_PROCEDIMIENTO"] = df["ID_PROCEDIMIENTO"]
        df_codigos["CODIGO_OPERATIVO"] = df["CODIGO_OPERATIVO"]
        df_codigos["DESCRIPCION"] = df_codigos["CODIGO_OPERATIVO"].map(CODIGOS_OPERATIVOS)
        
        return df_codigos

    def procesar_trata(self, df_procedimientos):
        """Procesa casos de trata de personas"""
        self.progress.emit(82, "Procesando casos de trata...")
        df = pd.read_excel(self.archivos['trata'])
        
        if df.empty:
            return pd.DataFrame()
            
        self.CONTADOR['BAJADA_TRATA'] = len(df)

        df['UOSP'] = df['UOSP'].fillna(df['URSA'])
        df['ID_PROCEDIMIENTO'] = df.apply(generar_uid_sigpol, axis=1)
        
        df_trata = pd.DataFrame()
        df_trata['ID_PROCEDIMIENTO'] = df['ID_PROCEDIMIENTO']
        df_trata['PROPOSITO_EXPLOTACION'] = df['MODALIDAD_EXPL'] 
        df_trata['OBSERVACIONES'] = "-"
        df_trata_final = df_trata[df_trata['ID_PROCEDIMIENTO'].isin(df_procedimientos['ID_PROCEDIMIENTO'])]
        self.CONTADOR['TRATA_FINAL'] = len(df_trata_final['ID_PROCEDIMIENTO'])
        print(df_trata_final.nunique())

        return df_trata_final

    def consolidar_datos(self, df_procedimientos, df_operaciones, df_incautaciones, df_detenidos, df_otros_delitos, df_trata, df_afectados, df_controlados, df_codigos):
        """Consolida todos los datos procesados"""
        # Unir datos de procedimientos y operaciones
        df_geog_final = pd.concat([df_procedimientos, df_operaciones])
        self.CONTADOR["GEOG_FINAL"] = len(df_geog_final)

        # Realizar los merge necesarios
        if self.CONTADOR['PROCEDIMIENTOS_NUEVOS'] > 0:
            df_incautados_final = pd.merge(df_geog_final, df_incautaciones, on='ID_PROCEDIMIENTO', how='left')
            df_detenidos_final = pd.merge(df_geog_final, df_detenidos, on='ID_PROCEDIMIENTO', how='left')
            df_otros_delitos_final = pd.merge(df_geog_final, df_otros_delitos, on='ID_PROCEDIMIENTO', how='right')
            df_trata_final = pd.merge(df_geog_final, df_trata, on='ID_PROCEDIMIENTO', how='right')

        if self.CONTADOR['ORDEN_SERVICIOS_NUEVOS'] > 0:
            df_afectados_final = pd.merge(df_geog_final, df_afectados, on='ID_PROCEDIMIENTO', how='left')
            df_controlados_final = pd.merge(df_geog_final, df_controlados, on='ID_PROCEDIMIENTO', how='left')
            df_codigos_final = pd.merge(df_geog_final, df_codigos, on='ID_PROCEDIMIENTO', how='left')

        
        
        dataframes = {
                'GEOG. PROCEDIMIENTO': df_geog_final.reset_index(drop=True).fillna("").replace("", "-") if 'df_geog_final' in locals() else pd.DataFrame(),
                'VEHI. Y PERSO. CONTROLADAS': df_controlados_final.reset_index(drop=True).fillna("").replace("", "-") if 'df_controlados_final' in locals() else pd.DataFrame(),
                'PERSONAL Y ELEMENTOS AFECTADOS': df_afectados_final.reset_index(drop=True).fillna("").replace("", "-") if 'df_afectados_final' in locals() else pd.DataFrame(),
                'INCAUTACIONES': df_incautados_final.reset_index(drop=True).fillna("").replace("", "-") if 'df_incautados_final' in locals() else pd.DataFrame(),
                'DETENIDOS Y APREHENDIDOS': df_detenidos_final.reset_index(drop=True).fillna("").replace("", "-")if 'df_detenidos_final' in locals() else pd.DataFrame(),
                'OTROS DELITOS': df_otros_delitos_final.reset_index(drop=True).fillna("").replace("", "-") if 'df_otros_delitos_final' in locals() else pd.DataFrame(),
                'TRATA O TRAFIC PERSONAS': df_trata_final.reset_index(drop=True).fillna("").replace("", "-") if 'df_trata_final' in locals() else pd.DataFrame(),
                'CODIGO OPERATIVO': df_codigos_final.reset_index(drop=True).fillna("").replace("", "-") if 'df_codigos_final' in locals() else pd.DataFrame(),
            }

        return dataframes
    
    
    
    def copiar_formato_template(self, dataframes):
        # 1. Cargar template
        wb_template = load_workbook(PATH_TEMPLATE)
        
        # 2. Definir estilos base
        header_style = {
            'fill': PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid'),
            'font': Font(bold=True, color='FFFFFF'),
            'border': Border(
                left=Side(style='thin'),
                right=Side(style='thin'), 
                top=Side(style='thin'),
                bottom=Side(style='thin')
            ),
            'alignment': Alignment(horizontal='center', vertical='center')
        }
        
        data_style = {
            'border': Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'), 
                bottom=Side(style='thin')
            ),
            'alignment': Alignment(horizontal='center', vertical='center')
        }

        # 3. Procesar cada hoja
        for sheet_name, df in dataframes.items():
            if sheet_name in wb_template.sheetnames:
                sheet = wb_template[sheet_name]
                template_sheet = wb_template[sheet_name]
                
                # Copiar formatos de columnas
                for column in range(1, template_sheet.max_column + 1):
                    letter = get_column_letter(column)
                    sheet.column_dimensions[letter].width = template_sheet.column_dimensions[letter].width
                
                # Escribir datos con formato
                for i, row in df.iterrows():
                    for j, value in enumerate(row, 1):
                        cell = sheet.cell(row=i+4, column=j+1, value=value)
                        
                        # Aplicar estilo según posición
                        if i == 0:
                            for k, v in header_style.items():
                                setattr(cell, k, v)
                        else:
                            for k, v in data_style.items():
                                setattr(cell, k, v)
                                
                # Preservar fórmulas existentes         
                for row in template_sheet.rows:
                    for cell in row:
                        if cell.value and str(cell.value).startswith('='):
                            sheet[cell.coordinate].value = cell.value

        # 4. Guardar resultado
        wb_template.save(PATH_OUTPUT)
    

    def actualizar_base_datos(self):
        """Actualiza la base de datos con los nuevos registros"""
        try:
            fecha_actual = datetime.now().strftime('%Y%m%d')

            # Leer el archivo de base de datos
            try:
                wb_base = load_workbook(PATH_BASE)
            except:
                wb_base = Workbook()
                wb_base.remove(wb_base.active)  # Eliminar hoja por defecto
                
                # Crear hojas necesarias
                wb_base.create_sheet("PROCEDIMIENTOS")
                wb_base.create_sheet("DETENIDOS_APREHENDIDOS")
                wb_base.create_sheet("OTROS DELITOS") 
                wb_base.create_sheet("INCAUTACIONES")
                wb_base.create_sheet("TRATA")
                wb_base.save("BASE_DATOS.xlsx")

            # Actualizar cada hoja con los nuevos datos
            hojas = {
                "PROCEDIMIENTOS": self.df_procedimientos,
                "DETENIDOS_APREHENDIDOS": self.df_detenidos,
                "OTROS DELITOS": self.df_otros_delitos,
                "INCAUTACIONES": self.df_incautaciones,
                "TRATA": self.df_trata if hasattr(self, 'df_trata') else None
            }

            for nombre_hoja, df in hojas.items():
                if df is not None and not df.empty:
                    # Convertir DataFrame a filas
                    filas = [df.columns.tolist()] + df.values.tolist()
                    
                    # Limpiar hoja actual
                    hoja = wb_base[nombre_hoja]
                    hoja.delete_rows(1, hoja.max_row)
                    
                    # Escribir nuevos datos
                    for fila in filas:
                        hoja.append(fila)

                    # Ajustar ancho de columnas
                    for columna in hoja.columns:
                        max_length = 0
                        for celda in columna:
                            try:
                                if len(str(celda.value)) > max_length:
                                    max_length = len(str(celda.value))
                            except:
                                pass
                        adjusted_width = (max_length + 2)
                        hoja.column_dimensions[columna[0].column_letter].width = adjusted_width

            # Guardar cambios
            wb_base.save("BASE_DATOS.xlsx")
            
            # Hacer backup
            wb_base.save(f"BASE_DATOS_{fecha_actual}.xlsx")

            self.progress.emit(98, "Base de datos actualizada correctamente")
            return True

        except Exception as e:
            raise Exception(f"Error actualizando base de datos: {str(e)}")

