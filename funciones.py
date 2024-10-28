
import re
import pandas as pd


contador_global_sigipol = {}
def generar_codigo_sigipol(row):
    id_operativo = row['ID_OPERATIVO']
    if id_operativo not in contador_global_sigipol:
        contador_global_sigipol[id_operativo] = 0
    contador_global_sigipol[id_operativo] += 1
    contador = contador_global_sigipol[id_operativo]
    return id_operativo + "-(" + str(contador) + ")"

def generar_codigo_sigipol_2(row):
        id_operativa = row['']
        id_operativa = row['']
        if id_operativa in conteo_acumulado:
            conteo_acumulado[id_operativa] += 1
        else:
            conteo_acumulado[id_operativa] = 1
        id_procedimiento = f"{id_operativa}-({conteo_acumulado[id_operativa]})"
        print ("id_procedimiento: " + id_procedimiento )
        return id_procedimiento

def generar_uid_sigpol(row):
        ursa = row['URSA']
        uosp = row['UOSP']
        numero_parte = row['NUMERO_PARTE']
        anio = row['ANIO_PARTE']
        return ursa +"-"+ uosp + "-"+ numero_parte+"-" + numero_parte + "-"+anio
    
def generar_uid_operaciones(row):
        id_operativa = row['ID_OPERATIVO']
        if id_operativa in conteo_acumulado:
            conteo_acumulado[id_operativa] += 1
        else:
            conteo_acumulado[id_operativa] = 1
        id_procedimiento = f"{id_operativa}-({conteo_acumulado[id_operativa]})"
        print ("id_procedimiento: " + id_procedimiento )
        return id_procedimiento
        
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

def procesar_tipo(row):
    tipo = row['TIPO_PROCEDIMIENTO']
    if pd.isna(tipo):
        return ""
    elif tipo == "DENUNCIA"  or tipo == "CONTROL PREVENTIVO" :
        return "ORDEN POLICIAL"
    else:
        return "ORDEN JUDICIAL"
    
def procesar_provincia(row):
    provincia = row['PROVINCIA']
    if pd.isna(provincia):
        return ""
    return PROVINCIAS.get(provincia, provincia)
    
def procesar_municipio(row):
    unidad = row['UOSP']
    if pd.isna(unidad):
        return ""
    return UNIDADES_MUNICIPIOS.get(unidad, unidad)

def procesar_lugar(row):
    lugar = row['LUGAR_CATALOGADO_NIVEL_1']
    return  LUGARES_CATALOGADOS[lugar]

def procesar_direccion(row):
    lugar = row['LUGAR_CATALOGADO_NIVEL_1']
    ciudad = row['CIUDAD']
    if lugar == "FUERA DE JURISDICCION" and ciudad == "ROSARIO":
        return "-"
    elif lugar == "FUERA DE JURISDICCION":
        return str(row['CALLE']) + " " + str(row['NUMERO']) + ", " + str(row['CIUDAD']) + " - " + str(row['PARTIDO'])
    else:
        return "-"
    
def controlar_estado (row):
    ursa = row['URSA'] 
    unidad = row['UOSP'] 
    estado = row['ESTADO_PARTE'] 
    if pd.isna(unidad)  and ursa == 'RG4' and estado == 'NO DISPONIBLE ESTADISTICA':
        return  "DISPONIBLE ESTADISTICA"
    else:
        return estado 
    


def procesar_causa_judicial(row):
    # Obtener la causa y asegurarse de que no sea None
    causa = row.get('CAUSAJUDICIALNUMERO', '')
    if causa is None or pd.isna(causa):
        causa = ''

    tipo = str(row.get('TIPO_CAUSA_INTERNA', '')).strip()
    causa_int = str(row.get('CAUSA_INTERNA_NUMERO', '')).strip()


    # Verificar si la causa está vacía o contiene ciertos valores
    if causa in ["", "S/D", "A/S", "N/C"]:
        # Manejar valores faltantes asignando un valor predeterminado al resultado
        resultado = f"{tipo}-{causa_int}".replace("--", "-")
        return resultado

    # Lista de prefijos que se quieren eliminar (sin importar mayúsculas o minúsculas)
    prefijos = ["NRO", "N°", "EXPTE", "EXPEDIENTE", "EXPT", "N"]

    # Crear una expresión regular que busque todos los prefijos y los elimine
    prefijos_regex = r'\b(' + '|'.join([re.escape(prefijo) for prefijo in prefijos]) + r')\b'

    # Eliminar prefijos de la causa
    causa = re.sub(prefijos_regex, '', causa, flags=re.IGNORECASE).strip()

    # Eliminar el símbolo "°" si está presente
    causa = causa.replace("°", "").replace(".","").replace('"',"").strip()

    # Utilizar una expresión regular para dividir letras seguidas de números con un guion
    causa_str = re.sub(r'([A-Za-z]+)\s*(\d+)', r'\1-\2', causa)

    # Reemplazar cualquier doble guion que haya quedado
    causa_str = causa_str.replace("--", "-").replace("---", "-")
    
    return causa_str




        