import glob
import os
import pandas as pd
import re 
import datetime
from Parametros import *
from datetime import datetime 


def obtener_ruta_bajada(nombre_clave):
    """
    Busca la ruta de un archivo relacionado con la clave proporcionada que puede tener variaciones en el nombre.
    
    Args:
        nombre_clave (str): Nombre base del archivo a buscar (sin extensión).

    Returns:
        str: Ruta del primer archivo encontrado que coincide con el patrón o una cadena vacía si no se encuentra.
    """
    base = "bajadas"
    # Patrón para buscar archivos similares (puede incluir variaciones)
    patron_busqueda = os.path.join(base, f"*{nombre_clave}*.xls*")

    # Buscar archivos que coincidan con el patrón
    archivos_encontrados = glob.glob(patron_busqueda)
    
    if not archivos_encontrados:
        print(f"No se encontró ningún archivo con el nombre '{nombre_clave}' en la carpeta '{base}'.")
        return ""

    archivo_encontrado = archivos_encontrados[0]  # Tomar el primer archivo encontrado
    print(f"Archivo encontrado: {archivo_encontrado}")
    
    return archivo_encontrado


def generar_uid_sigpol(row):
        tipo = str(row['TIPO_CAUSA_INTERNA'])
        numero_parte = str(row['NUMERO_PARTE'])
        uosp = str(row['UOSP'])
        anio = str(row['ANIO_PARTE'])
        return tipo + "-"+ numero_parte  + "-" + uosp + "-"+ anio

def procesar_descripcion(row):
    tipo = row['TIPO_PROCEDIMIENTO']
    if tipo == "DENUNCIA":
        return "DENUNCIA POLICIAL"
    elif tipo == "CONTROL PREVENTIVO":
        return f"CONTROL PREVENTIVO - {procesar_lugar(row)}"
    elif tipo == "ORDEN DE ALLANAMIENTO":
        return "ORDEN DE ALLANAMIENTO"
    elif tipo == "ORDEN DE ALLANAMIENTO / DETENCIÓN":
        return "ORDEN DE ALLANAMIENTO"
    else:
        return "OTRO MANDATO JUDICIAL"

def procesar_tipo_procedimiento(row):
    tipo = row['TIPO_PROCEDIMIENTO']
    if pd.isna(tipo):
        return ""
    elif tipo == "DENUNCIA"  or tipo == "CONTROL PREVENTIVO" :
        return "ORDEN POLICIAL"
    else:
        return "ORDEN JUDICIAL"

def procesar_provincia(row):
    provincia = row['PROVINCIA']
    jurisdiccion = row['JURISDICCION']
    if pd.isna(provincia):
        return PROVINCIAS.get(jurisdiccion, provincia)
    else:
        return PROVINCIAS.get(provincia, provincia)
        
def procesar_municipio(row):
    unidad = row['UOSP']
    if pd.isna(unidad):
        return ""
    return UNIDADES_MUNICIPIOS.get(unidad, unidad)

def procesar_lugar(row):
    lugar = row['LUGAR_CATALOGADO_NIVEL_1']
    if pd.isna(lugar):
        return "-"
    return  LUGARES_CATALOGADOS[lugar]





def procesar_geog(row):
    latitud = row['GEOREFERENCIA_Y']
    longitud = row['GEOREFERENCIA_X']
    
    if pd.isna(latitud) or pd.isna(longitud):
        unidad = row['UOSP']
        if unidad in GEOS_UNIDADES:
            latitud = GEOS_UNIDADES[unidad]['LATITUD']
            longitud = GEOS_UNIDADES[unidad]['LONGITUD']
            return [latitud,longitud]
        else:
            return ["-","-"]
    else:
        return [latitud,longitud]

def procesar_direccion(row):
    """
    Asigna valores y maneja el caso en que sean None o NaN para luego construir la
    dirección basada en las condiciones.
    Si el lugar es "FUERA DE JURISDICCION" y la ciudad es "ROSARIO", se devuelve "ROSARIO".
    Si el lugar es "FUERA DE JURISDICCION", se construye la dirección con los valores
    de calle, numero, ciudad y partido. Si todos los valores son NaN, devuelve una cadena vacía.
    Si el lugar no es "FUERA DE JURISDICCION", se devuelve "-".
    """
    # Asignamos valores y verificamos si son NaN o None, si es así los dejamos como ""
    lugar = row['LUGAR_CATALOGADO_NIVEL_1']
    ciudad = str(row['CIUDAD']) if not pd.isna(row['CIUDAD']) else ""
    calle = str(row['CALLE']) if not pd.isna(row['CALLE']) else ""
    numero = str(row['NUMERO']) if not pd.isna(row['NUMERO']) else ""
    partido = str(row['PARTIDO']) if not pd.isna(row['PARTIDO']) else ""

    # Si después de la conversión el valor es 'nan', también lo dejamos como ""
    ciudad = "" if ciudad.lower() == 'nan' else ciudad
    calle = "" if calle.lower() == 'nan' else calle
    numero = "" if numero.lower() == 'nan' else numero
    partido = "" if partido.lower() == 'nan' else partido

    # Lógica para construir la dirección basada en las condiciones
    if lugar == "FUERA DE JURISDICCION" and ciudad == "ROSARIO":
        direccion = "ROSARIO"
    elif lugar == "FUERA DE JURISDICCION" and not (calle == "" and numero == "" and ciudad == "" and partido == ""):
        direccion = f"{calle} {numero} {ciudad} {partido}".strip()
    else:
        direccion = "-"

    return direccion


def controlar_estado (row):
    ursa = row['URSA']
    unidad = row['UOSP']
    estado = row['ESTADO_PARTE']
    if pd.isna(unidad)  and ursa == 'RG4' and estado == 'NO DISPONIBLE ESTADISTICA':
        return  "DISPONIBLE ESTADISTICA"
    else:
        return estado

def leer_excel_a_df(worksheet):
    # Leer títulos desde la fila 2
    titulos = [worksheet.cell(row=2, column=col).value for col in range(1, worksheet.max_column + 1)]
    
    # Leer datos desde la fila 3
    data = []
    for row in worksheet.iter_rows(min_row=3, min_col=1, max_col=worksheet.max_column, values_only=True):
        data.append(row)
    
    # Ajustar el número de títulos para que coincidan con las columnas de datos
    if len(titulos) < worksheet.max_column:
        print(f"Advertencia: hay menos títulos ({len(titulos)}) que columnas en los datos ({worksheet.max_column}).")
        titulos.extend([f"Columna_{i}" for i in range(len(titulos) + 1, worksheet.max_column + 1)])
    elif len(titulos) > worksheet.max_column:
        print(f"Advertencia: hay más títulos ({len(titulos)}) que columnas en los datos ({worksheet.max_column}).")
        titulos = titulos[:worksheet.max_column]
    
    # Crear el DataFrame
    df = pd.DataFrame(data, columns=titulos)
    return df


def procesar_causa_judicial(row):
    # procesar los valores de las columnas, asegurándose de que no sean `None` o `NaN`
    causa = str(row.get('CAUSAJUDICIALNUMERO', '')).strip()
    tipo = str(row.get('TIPO_CAUSA_INTERNA', '')).strip()
    causa_int = str(row.get('CAUSA_INTERNA_NUMERO', '')).strip()


    # Validar si el campo 'causa' tiene un valor no deseado
    if causa in ["", "S/D", "A/S", "N/C", "None","nan", "--", "SN°","S/N","-", "---"]:
        # Retornar el tipo y el número de causa interna, asegurando que no haya guiones redundantes
        resultado = f"{tipo}-{causa_int}".replace("--", "-").strip("-")
        return resultado

    # Definir los prefijos que deben ser eliminados
    prefijos = ["NRO", "N°", "EXPTE", "EXPEDIENTE", "EXPT", "N ", '.', ':', '"', '_', '´', "`","°",""]
    
    # Crear una expresión regular para eliminar los prefijos
    
    causa = causa.strip()
    for prefijo in prefijos:
        causa = causa.replace(prefijo, "").strip()

    # Normalizar la causa, añadiendo guion entre letras y números si es necesario
    causa = re.sub(r'([A-Za-z]+)\s*(\d+)', r'\1-\2', causa)

    # Eliminar guiones dobles o triples si existen
    causa = re.sub(r'-{2,}', '-', causa)

    return causa

def filtrar_procedimientos_generales (ruta_archivo):
    cantidad_partes_inical = 0
    cantidad_partes = 0
    cantidad_partes_duplicados = 0
    cantidad_partes_no_disponible = 0
    
    df = pd.read_excel(ruta_archivo)
    
    
    
    df["UID"] = df.apply(generar_uid_sigpol,axis=1)
    cantidad_partes_inical = df['UID'].count()
    cantidad_partes = df['UID'].count()

    df.drop_duplicates(subset='UID', keep='first', inplace=True)
    
    
    
    cantidad_partes_duplicados = cantidad_partes - df['UID'].count()
    cantidad_partes = df['UID'].count()


    df['TIPO_CAUSA_INTERNA'] =  df.apply(procesar_tipo_causa_interna ,axis=1)
    
    df['ESTADO_PARTE'] = df.apply(controlar_estado ,axis=1)
    df['UOSP'] = df['UOSP'].fillna(df['URSA'])

    
    df = df[df['ESTADO_PARTE'] != "NO DISPONIBLE ESTADISTICA"]

    
    df["UID"] = df.apply(generar_uid_sigpol,axis=1)
    
    cantidad_partes_no_disponible = cantidad_partes - df['UID'].count()
    cantidad_partes = df['UID'].count()
    
    print(f"Estadistica de Partes\n")
    print(f"Total de Partes: {cantidad_partes}"  )
    print(f"Cantidad Duplicado: {cantidad_partes_duplicados}" )
    print(f"Cantidad No diponible: {cantidad_partes_no_disponible}" )
    print(f"Cantidad de Partes final: {cantidad_partes}" )
    
    return df

### funciones para el procesamiento de datos de OPERACIONES


def procesar_geog_oper(row):
    latitud = str(row['LATITUD'])
    longitud = str(row['LONGITUD'])
    
    if latitud == "N/C" or longitud == "N/C" or latitud == "-" or longitud == "NO CORRESPONDE" or longitud == "S/D" or latitud == "CONTROLES ALEATORIOS Y DINAMICOS":
        return ["-","-"]
    
    # Verificar si ya tiene punto decimal
    if '.' in latitud:
        latitud = latitud
    else:
        latitud = latitud[:3] + '.' + latitud[3:]
        
    if '.' in longitud:
        longitud = longitud
    else:
        longitud = longitud[:3] + '.' + longitud[3:]
    
    if pd.isna(latitud) or pd.isna(longitud):
        unidad = row['UOSP']
        if unidad in GEOS_UNIDADES:
            latitud = GEOS_UNIDADES[unidad]['LATITUD']
            longitud = GEOS_UNIDADES[unidad]['LONGITUD']
            return [latitud,longitud]
        else:
            return ["-","-"]
        
    elif latitud == "N/C" or longitud == "N/C" or latitud == "-" or longitud == "NO CORRESPONDE" or longitud == "S/D":
        return ["-","-"]
    else:
        return [latitud,longitud]
    
    
def procesar_unidad (row):
    unidad = row['UNIDAD_INTERVINIENTE']
    unidad ="UR1" if unidad in "DROPA I" else unidad
    return  unidad

def colocar_guion_espacio(texto):
    # Limpieza inicial del texto: eliminar caracteres innecesarios y normalizar guiones
    print("=== Inicio ===")
    print(f"Texto original: {texto}")
    
    caracteres_no_deseados = ['N°', ' N', '.', ':', '-', '"', '_', '´',"`"]

    # Eliminamos caracteres listados arriba en una sola expresión regular
    texto = re.sub(r'({})'.format('|'.join(map(re.escape, caracteres_no_deseados))), '', texto)
    
    # Reemplazar múltiples espacios por un solo guion para normalizar los separadores
    texto = re.sub(r'\s+', '-', texto)
    
    print(f"Texto después de la limpieza inicial: {texto}")

    # Inicializar componentes vacíos
    prefijo = ""
    numero = ""
    ubicacion = ""
    year = ""
    suffix = "-(1)"  # Valor por defecto del sufijo

    # Extraer prefijo si está en la lista de prefijos conocidos
    for p in PREFIJOS:
        if texto.upper().startswith(p):
            prefijo = p
            texto = texto[len(p):].strip()  # Recortar el prefijo del texto
            break
    print(f"Prefijo encontrado: {prefijo}")


    # Limpiar posibles espacios adicionales antes de buscar la ubicación
    texto = texto.strip()

    # Extraer ubicación si está en la lista de ubicaciones conocidas
    for u in UBICACIONES:
        if texto.upper().startswith(u):
            ubicacion = u
            texto = texto[len(u):].strip()  # Recortar la ubicación del texto
            break

    # Mejorar el chequeo para la ubicación, verificando si existe la ubicación en cualquier parte
    if not ubicacion:
        for u in UBICACIONES:
            if u in texto.upper():
                ubicacion = u
                texto = texto.replace(u, "", 1).strip()
                break
    print(f"Ubicación encontrada: {ubicacion}")

    # Extraer número principal
    match_numero = re.search(r"(\d+)", texto)
    if match_numero:
        numero = match_numero.group(0).zfill(4)
        texto = texto[len(match_numero.group(0)):].strip()  # Recortar el número del texto
    print(f"Número principal encontrado: {numero}")
    
    # Extraer año en formato de 4 dígitos
    match_year = re.search(r"(\d{4})", texto)
    if match_year:
        year = match_year.group(0)
        texto = texto[len(year):].strip()  # Recortar el año del texto
    print(f"Año encontrado: {year}")

    # Extraer sufijo al final (número entre paréntesis)
    match_suffix = re.search(r'\((\d+)\)$', texto)
    if match_suffix:
        suffix = f"-({match_suffix.group(1)})"
    print(f"Sufijo encontrado: {suffix}")

    # Formatear y retornar el texto resultante
    resultado = f"{prefijo}-{numero}-{ubicacion}/{year}{suffix}"
    print(f"Resultado formateado: {resultado}")
    print("=== Fin del proceso ===")
    return resultado

def formatear_contador(texto):
    print(f"Texto original: {texto}")
    texto_procesado = re.sub(r'-+\(\d+\)$', '', texto)
    print(f"Texto procesado: {texto_procesado}")
    return texto_procesado

def colocar_contador (df_operaciones, base):
    conteo_base_datos = base['ID_OPERATIVO'].value_counts()
    conteo_acumulado  = conteo_base_datos.to_dict()
    df_ordenes_no_informadas = pd.DataFrame()
    for index, row in df_operaciones.iterrows():
        id_operativa = row['ID_OPERATIVO']
        
        # Verificar cuántas veces ha aparecido el ID_operativa en total hasta ahora (base + nuevos)
        if id_operativa in conteo_acumulado:
            conteo_acumulado[id_operativa] += 1
        else:
            conteo_acumulado[id_operativa] = 1
        
        nuevo_id_procedimiento = f"{id_operativa}-({conteo_acumulado[id_operativa]})"
        
        df_ordenes_no_informadas.at[index, 'ID_PROCEDIMIENTO'] = nuevo_id_procedimiento
        
    return df_ordenes_no_informadas

def generar_uid_operaciones(row):
    texto = str(row['ID_PROCEDIMIENTO'])
    prefijo = ""
    for p in PREFIJOS:
        if texto.upper().startswith(p):
            prefijo = p
            texto = texto[len(p):].strip()  # Recortar el prefijo del texto
            break
    unidad = str(row['UNIDAD_INTERVINIENTE'])
    fecha_completa = str(row['FECHA'])
    fecha = fecha_completa.split()[0]  # Tomar solo la parte antes del espacio (la fecha)
    hora = str(row['HORA']).replace(":","-")
    conjunto = prefijo + "-" + unidad + "-" + fecha + "-" + hora
    
    return conjunto

### funciones para el procesamiento de datos de PERSONAS

def procesar_edad(row):
    fecha_nacimiento = row['FECHA_NACIMIENTO']
    if pd.isna(fecha_nacimiento):
        return "-"
    denuncia_fecha = row['DENUNCIAFECHA']
    edad = (denuncia_fecha.year - fecha_nacimiento.year) - ((denuncia_fecha.month, denuncia_fecha.day) < (fecha_nacimiento.month, fecha_nacimiento.day))
    
    return edad
def procesar_sexo(row):
    sexo = row['SEXO']
    if sexo == 'F':
        return 'FEMENINO'
    else:
        return 'MASCULINO'
def procesar_genero(row):
    sexo = row['SEXO']
    if sexo == 'F':
        return 'MUJER'
    else:
        return 'VARON'
def procesar_nacionalidad(row):
    nacionalidad = row['NACIONALIDAD1']
    if pd.isna(nacionalidad):
        return "-"
    return NACIONALIADADES.get(nacionalidad, nacionalidad)
    
def procesar_situacion_judicial(row):
    tipo_persona = row['TIPO_PERSONA']
    situacion = row['SITUACION_JUDICIAL']
    union = tipo_persona + " - " + situacion
    return SITUACIONES_JUDICIALES.get(union, union)

def procesar_tipo_delito(row):
    clasificion_1 = row['CLASIFICACION_NIVEL_1']
    clasificion_2 = row['CLASIFICACION_NIVEL_2']
    union = clasificion_1 + " - " + clasificion_2
    return DELITOS.get(union, union)

def procesar_caratula(row):
    caratula_judicial = row['CARATULAJUDICIAL']
    caratula_interna = row['CARATULAINTERNA']
    
    if pd.isna(caratula_judicial) or caratula_judicial == "S/D" or caratula_judicial == "A/S" or caratula_judicial == "N/C":
        return caratula_interna
    return caratula_judicial
    
def procesar_juzgado(row):
    juzgado = row['JUZGADO']
    fiscalia = row['FISCALIA']
    if pd.isna(juzgado) or juzgado == "S/D" or juzgado == "N/C":
        return fiscalia
    return juzgado
### funciones genericas

def procesar_tipo_causa_interna(row):
    tipo = row['TIPO_CAUSA_INTERNA']
    if tipo == "ACTUACIÓN JUDICIAL" or tipo == "ACTUACIONES JUDICIALES":
        return "AJ"
    elif tipo == "RESTRICCIÓN A LA LIBERTAD":
        return "RL"
    else:
        return tipo

### funciones de armas 


def procesar_cantidad_arma(row):
    cantidad = row['CANTIDAD']
    if pd.isna(cantidad):  # Verifica si es NaN
        cantidad = 1.0     # Asigna 1.0 en caso de NaN
    else:
        cantidad = float(cantidad)  # Convierte a float sin intentar cambiar a int
    return cantidad



def procesar_observaciones_arma(row):
    marca = str(row['MARCA']).replace("-", "")
    calibre = str(row['CALIBRE'])
    observaciones = "-"
    if (marca != "nan" or marca == "") and calibre != "nan":
        observaciones = f"MARCA: {marca} - CALIBRE: {calibre}"
    elif  marca == "nan" and calibre != "nan":
        observaciones = f"CALIBRE: {calibre}"
    else:
        observaciones = "-"
    return observaciones
    


### funciones de objetos 
def clasificar_tipo_objeto(row):
    clasificacion_nivel_2, tipo_objeto, cantidad = row["CLASIFICACION_NIVEL_2"], row["TIPO_OBJETO"], row["CANTIDAD"]
    if clasificacion_nivel_2 == "CONTRABANDO":
        pre_tipo = f"{clasificacion_nivel_2} - {tipo_objeto}"
    elif tipo_objeto == "OTRO":
        pre_tipo = f"{tipo_objeto} - {clasificacion_nivel_2}"
    else:
        pre_tipo = tipo_objeto
    return pre_tipo if cantidad else F"{pre_tipo} - VACIO"



### funciones de vehiculos 
def clasificar_tipo_vehiculo(row):
    dominio = row["VEHICULO_DOMINIO"].replace(" ", "")
    modelo = str(row["VEHICULO_MODELO"]).upper()  # Convertir a mayúsculas para evitar problemas de coincidencia
    print(f"Dominio: {dominio}, Modelo: {modelo}")
    # Definir patrones para cada tipo de vehículo
    patrones = {
        "AUTO": r'^[A-Z]{3}\d{3}$|^[A-Z]{2}\d{3}[A-Z]{2}$',  # Formatos ABC123 o AB123CD
        "MOTO": r'^[A-Z]{2}\d{3}$',                          # Formato AB123
        "CUATRICICLO": r'^[A-Z]{2}\d{3}$',                   # Similar a las motos
        "AVION": r'^(LV|LQ)-[A-Z]{3}$'                       # Formato LV-XXX o LQ-XXX
    }
    
    camionetas = [
        "RANGER", "HILUX", "D-MAX", "ALASKAN", "TROOPER",
        "PICK UP", "AMAROK", "FRONTIER", "F-100", "RAM",
        "TRANSIT", "MASTER", "SPRINTER", "BOXER", "DUCATO", "SW4"
    ]
    
    
    # Verificar si el modelo es una camioneta
    palabras_modelo = modelo.split()
    for palabra in palabras_modelo:
        if palabra in camionetas:
            print(f"Palabra detectada en camionetas: {palabra} - Clasificación detectada: CAMIONETA")
            return "CAMIONETA"
    
    # Intentar clasificar el dominio según los patrones
    
    for tipo, patron in patrones.items():
        if re.match(patron, dominio):
            return tipo
        
        
    return "VERIFIQUE"


def observaciones_vehiculo(row):
    dominio = row["VEHICULO_DOMINIO"]
    marca = row["VEHICULO_MARCA"]
    modelo = row["VEHICULO_MODELO"]


    return f"{marca} - {modelo} - {dominio} "



### funciones de narcotrafico 
def  clasificar_tipo_sustancia(row):
    tipo_sustancia = row['TIPO_ESTUPEFACIENTE']
    return TIPO_SUSTANCIA.get(tipo_sustancia, tipo_sustancia)


def clasificar_medida(row):
    """
    Clasifica la medida de una sustancia según el tipo de sustancia
    
    Parameters
    ----------
    row : pandas Series
        Fila de un DataFrame que contiene las columnas 'UNIDADES', 'PESO' y 'TIPO_ESTUPEFACIENTE'
    
    Returns
    -------
    list
        Lista de dos elementos, donde el primer elemento es la cantidad y el segundo es la unidad de medida.
        Si no se puede clasificar, devuelve [0, "VERIFIQUE"]
    """
    
    unidad = row['UNIDADES']
    peso = row['PESO']
    tipo_sustancia = row['TIPO_ESTUPEFACIENTE']
    tipo_sustancia = TIPO_SUSTANCIA.get(tipo_sustancia, tipo_sustancia)
    
    if tipo_sustancia == "COCAINA":
        if pd.isna(peso):
            return [unidad, "UNIDADES"]
        return [peso, "GRAMOS"]
    elif tipo_sustancia == "MARIHUANA":
        if pd.isna(peso):
            return [unidad, "UNIDADES"]
        return [peso, "GRAMOS"]
    elif pd.isna(unidad) and pd.isna(peso):
        return [0, "VERIFIQUE"]
    elif pd.isna(unidad) and not pd.isna(peso):
        return [unidad, "UNIDADES"]
    else:
        return [peso, "GRAMOS"]


def observaciones_sustancia(row):
    tipo_sustancia = row["TIPO_ESTUPEFACIENTE"]
    tipo_sustancia2 = TIPO_SUSTANCIA.get(tipo_sustancia, tipo_sustancia)
    
    if tipo_sustancia2 == "OTROS":
        return f"{tipo_sustancia}" 
    else:
        return  "-"