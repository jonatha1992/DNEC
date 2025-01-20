import os
from datetime import datetime
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment
from PyQt6.QtCore import QObject, pyqtSignal

from Funciones import *

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
        self.actualizar_base = True
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

    def iniciar_procesamiento(self, archivos, fecha_inicial, fecha_final, 
                            actualizar_base=True, verificar_base=True):
        """Inicia el proceso de consolidación"""
        try:
            # Validaciones iniciales
            self.validar_fechas(fecha_inicial, fecha_final)
            self.validar_archivos(archivos)
            
            # Guardar parámetros
            self.archivos = self.clasificar_archivos(archivos)
            self.fecha_inicial = fecha_inicial
            self.fecha_final = fecha_final
            self.actualizar_base = actualizar_base
            self.verificar_base = verificar_base

            # 1. Procesar procedimientos (20%)
            self.progress.emit(10, "Procesando procedimientos...")
            df_procedimientos = self.procesar_procedimientos()
            
            # 2. Procesar personas (40%)
            self.progress.emit(30, "Procesando personas...")
            df_personas = self.procesar_personas(df_procedimientos)
            
            # 3. Procesar incautaciones (60%)
            self.progress.emit(50, "Procesando incautaciones...")
            df_incautaciones = self.procesar_incautaciones(df_procedimientos)
            
            # 4. Procesar operaciones (80%)
            self.progress.emit(70, "Procesando operaciones...")
            df_operaciones = self.procesar_operaciones()
            
            # 5. Generar informe final (90%)
            self.progress.emit(85, "Generando informe...")
            self.generar_informe_consolidado()
            
            # 6. Actualizar base de datos si corresponde (95%)
            if self.actualizar_base:
                self.progress.emit(95, "Actualizando base de datos...")
                self.actualizar_base_datos()

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
            'procedimientos': None,
            'personas': None,
            'armas': None,
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
        
        df = pd.read_excel(self.archivos['procedimientos'])
        
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
        df = pd.read_excel(self.archivos['personas'])
        
        if df.empty:
            return pd.DataFrame()
            
        self.CONTADOR['BAJADA_PERSONAS'] = len(df)
        
        # Generar UID y procesar datos
        df['UOSP'] = df['UOSP'].fillna(df['URSA'])
        df['ID_PROCEDIMIENTO'] = df.apply(generar_uid_sigpol, axis=1)
        # Paso 3: Coincidencia de UID entre excel_bajada_personas y df_geog_final
        uid_coincidentes = df[df['ID_PROCEDIMIENTO'].isin(df['ID_PROCEDIMIENTO'])]

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
        
        
        
        if self.CONTADOR['PROCEDIMIENTOS_NUEVOS'] != 0:
    
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
        
        # Unir todas las incautaciones
        return pd.concat([df_armas, df_divisas, df_objetos, df_vehiculos, df_narcotrafico])

    def procesar_armas(self, df_procedimientos):
        """Procesa incautaciones de armas"""
        df = pd.read_excel(self.archivos['armas'])
        if df.empty:
            return pd.DataFrame()
            
        self.CONTADOR['BAJADA_ARMAS'] = len(df)
        if self.CONTADOR['PROCEDIMIENTOS_NUEVOS'] != 0:

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
        if self.CONTADOR['PROCEDIMIENTOS_NUEVOS'] != 0:

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

        if self.CONTADOR['PROCEDIMIENTOS_NUEVOS'] != 0:

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

        if self.CONTADOR['PROCEDIMIENTOS_NUEVOS'] != 0:

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
        pass

    def procesar_narcotrafico(self, df_procedimientos):
        """Procesa incautaciones de narcotráfico"""
        # ...similar a procesar_armas...
        pass

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

    def procesar_trata(self, df_procedimientos):
        """Procesa casos de trata de personas"""
        self.progress.emit(82, "Procesando casos de trata...")
        df = pd.read_excel(self.archivos['trata'])
        
        if df.empty:
            return pd.DataFrame()
            
        self.CONTADOR['BAJADA_TRATA'] = len(df)

        if self.CONTADOR['PROCEDIMIENTOS_NUEVOS'] != 0:
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

    def generar_informe_consolidado(self):
        """Genera el informe final consolidado"""
        self.progress.emit(92, "Generando informe consolidado...")

        try:
            # Crear archivo Excel
            wb = Workbook()
<<<<<<< HEAD
            wb.create_sheet("PROCEDIMIENTOS")
            hoja_proc = wb["PROCEDIMIENTOS"]
            
=======
            
            # Hoja de Procedimientos
            wb.create_sheet("PROCEDIMIENTOS")
            hoja_proc = wb["PROCEDIMIENTOS"]
>>>>>>> parent of 5f289a8 (eliminar archivos de interfaz de usuario innecesarios)
            for row in dataframe_to_rows(self.df_procedimientos, index=False, header=True):
                hoja_proc.append(row)

            # Hoja de Detenidos/Aprehendidos
            wb.create_sheet("DETENIDOS_APREHENDIDOS") 
            hoja_det = wb["DETENIDOS_APREHENDIDOS"]
            for row in dataframe_to_rows(self.df_detenidos, index=False, header=True):
                hoja_det.append(row)

            # Hoja de Otros Delitos
            wb.create_sheet("OTROS DELITOS")
            hoja_otros = wb["OTROS DELITOS"]
            for row in dataframe_to_rows(self.df_otros_delitos, index=False, header=True):
                hoja_otros.append(row)

            # Hoja de Incautaciones
            wb.create_sheet("INCAUTACIONES")
            hoja_inc = wb["INCAUTACIONES"]
            for row in dataframe_to_rows(self.df_incautaciones, index=False, header=True):
                hoja_inc.append(row)
                
            # Hoja de Trata
            if self.df_trata is not None and not self.df_trata.empty:
                wb.create_sheet("TRATA")
                hoja_trata = wb["TRATA"]
                for row in dataframe_to_rows(self.df_trata, index=False, header=True):
                    hoja_trata.append(row)

            # Eliminar hoja por defecto
            wb.remove(wb['Sheet'])

            # Ajustar ancho de columnas
            for hoja in wb.sheetnames:
                for column in wb[hoja].columns:
                    max_length = 0
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = (max_length + 2)
                    wb[hoja].column_dimensions[column[0].column_letter].width = adjusted_width

            # Guardar archivo
            fecha_actual = datetime.now().strftime('%Y%m%d')
            nombre_archivo = f'Consolidado_{fecha_actual}.xlsx'
            wb.save(nombre_archivo)
            
            return nombre_archivo

        except Exception as e:
            raise Exception(f"Error generando informe consolidado: {str(e)}")

    def actualizar_base_datos(self):
        """Actualiza la base de datos con los nuevos registros"""
        try:
            fecha_actual = datetime.now().strftime('%Y%m%d')

            # Leer el archivo de base de datos
            try:
                wb_base = load_workbook("BASE_DATOS.xlsx")
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

