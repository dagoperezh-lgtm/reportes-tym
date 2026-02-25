# TYM PLATAFORMA - VERSION: 2.2.19-PCVCI-RECOVERY
# OBJETIVO: RESTAURAR EXTENSIÓN COMPLETA Y ELIMINAR SÍNTESIS AUTOMÁTICA
# LINEAS DE CODIGO: 785
# ESTADO: MODELO FUNCIONAL EXTENDIDO - SELLO DE INGENIERÍA FINAL

import streamlit as st
import pandas as pd
import numpy as np
import re
import io
import zipfile
import unicodedata
import matplotlib.pyplot as plt
from datetime import time, datetime, timedelta
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

# *****************************************************************************
# --- 1. CONFIGURACIÓN DE PÁGINA (BLINDADO - NO TOCAR) ---
# *****************************************************************************

st.set_page_config(
    page_title="Plataforma TYM 2026 - V2.2.19", 
    page_icon="🏆", 
    layout="wide"
)

st.title("🏆 Gestión de Reportes y Estadísticas - Club TYM")

# *****************************************************************************
# --- 2. UTILIDADES DE PROCESAMIENTO Y TIEMPO (BLINDADO - NO SINTETIZAR) ---
# *****************************************************************************

def clean_string(text):
    """
    Normaliza nombres para asegurar coincidencias entre Strava y Excel.
    Elimina tildes, espacios extra y convierte a mayúsculas.
    """
    if text is None or pd.isna(text):
        return ""
    
    # Proceso de normalización de caracteres paso a paso
    nombre_limpio_temp = str(text).strip()
    
    nombre_limpio_temp = nombre_limpio_temp.upper()
    
    # Uso de NFKD para descomponer caracteres con tildes
    info_normalizada = unicodedata.normalize('NFKD', nombre_limpio_temp)
    
    resultado_final_nombre = ""
    
    for caracter_indiv in info_normalizada:
        # Filtro para ignorar los caracteres de combinación (tildes)
        if not unicodedata.combining(caracter_indiv):
            resultado_final_nombre = resultado_final_nombre + caracter_indiv
            
    return resultado_final_nombre

def to_mins(valor_entrada_tiempo):
    """
    Convierte cualquier formato de tiempo a minutos totales de forma explícita.
    Maneja decimales de Excel, objetos datetime, strings HH:MM y formato Strava.
    """
    if pd.isna(valor_entrada_tiempo):
        return 0
    
    string_valor = str(valor_entrada_tiempo).strip()
    
    # Listado exhaustivo de casos nulos detectados en la operativa real
    lista_casos_nulos = ['--:--', '0', '', '00:00:00', '0:00:00', '00:00', '0.0', 'NC', '0:00']
    
    if string_valor in lista_casos_nulos:
        return 0
        
    try:
        # 🛡️ REGLA ARITMÉTICA: Si el valor es numérico (fracción de día de Excel)
        if isinstance(valor_entrada_tiempo, (float, int)):
            # Excel almacena 1 día completo como 1.0. 
            # Multiplicamos por 1440 para obtener la cifra real de minutos.
            minutos_finales_calculados = int(round(valor_entrada_tiempo * 1440))
            return minutos_finales_calculados
        
        # Si el dato es un objeto de tiempo nativo de Python
        if isinstance(valor_entrada_tiempo, (time, datetime)):
            minutos_finales_calculados = (valor_entrada_tiempo.hour * 60) + valor_entrada_tiempo.minute
            return minutos_finales_calculados
            
        # Si el string representa un número decimal puro
        try:
            conversion_float = float(string_valor)
            minutos_finales_calculados = int(round(conversion_float * 1440))
            return minutos_finales_calculados
        except ValueError:
            # No es numérico, continuamos con la lógica de parsing de texto
            pass
            
        # Formato de hora estándar con separador de dos puntos (HH:MM)
        if ':' in string_valor:
            bloques_tiempo = string_valor.split(':')
            if len(bloques_tiempo) >= 2:
                horas_bloque = int(bloques_tiempo[0])
                
                minutos_raw_bloque = bloques_tiempo[1]
                # Se eliminan segundos o microsegundos si existen
                minutos_clean_bloque = int(minutos_raw_bloque.split('.')[0])
                
                total_minutos_bloque = (horas_bloque * 60) + minutos_clean_bloque
                return total_minutos_bloque
        
        # Formato nativo de Strava (ejemplo: 11h 6min)
        busqueda_horas = re.search(r'(\d+)h', string_valor)
        busqueda_minutos = re.search(r'(\d+)min', string_valor)
        
        h_resultado = 0
        if busqueda_horas:
            h_resultado = int(busqueda_horas.group(1))
            
        m_resultado = 0
        if busqueda_minutos:
            m_resultado = int(busqueda_minutos.group(1))
            
        resultado_total_minutos = (h_resultado * 60) + m_resultado
        return resultado_total_minutos
        
    except Exception:
        # Fallback de seguridad para evitar que la aplicación se detenga
        return 0

def to_excel_time_value(dato_entrada_original):
    """
    Transforma la entrada en la fracción decimal exacta que requiere el motor de Excel.
    Este paso es vital para que las celdas sean sumables y promediables.
    """
    minutos_para_excel = to_mins(dato_entrada_original)
    
    # 24 horas equivalen a 1440 minutos totales
    valor_decimal_excel = minutos_para_excel / 1440.0
    
    return valor_decimal_excel

def to_hhmmss_display(minutos_totales_input):
    """
    Formato de texto HH:MM:00 exclusivo para la estética del reporte Word.
    """
    valor_horas_v = int(minutos_totales_input // 60)
    valor_minutos_v = int(minutos_totales_input % 60)
    
    # Generación de la cadena de texto con formato de reloj
    string_formato_reloj = f"{valor_horas_v:02d}:{valor_minutos_v:02d}:00"
    
    return string_formato_reloj

# *****************************************************************************
# --- 3. MOTOR DE COMENTARIOS TÉCNICOS (PROTEGIDO - BLOQUEADO) ---
# *****************************************************************************

def generar_comentario(datos_de_fila, nombre_categoria, rank_posicion):
    """
    Genera el análisis cualitativo extenso para los podios del reporte Word.
    Este bloque debe ser extenso y descriptivo para mantener el nivel profesional del informe.
    """
    identidad_atleta = datos_de_fila['Deportista']
    
    # Análisis para categorías de Volumen y Clasificación General
    if nombre_categoria == 'Completos' or nombre_categoria == 'General':
        if rank_posicion == 1:
            texto_comentario = f"Dominio absoluto de {identidad_atleta}. Se consolida en la cima del club con un volumen total envidiable. Su capacidad para sostener sesiones de alta intensidad demuestra una preparación de élite y una disciplina inquebrantable."
            return texto_comentario
        
        if rank_posicion == 2:
            texto_comentario = f"Una semana brillante para {identidad_atleta}. Se queda con la plata manteniendo una presión constante sobre el líder. Su solidez fue el motor que lo mantuvo en la parte más alta de la tabla."
            return texto_comentario
            
        if rank_posicion == 3:
            texto_comentario = f"{identidad_atleta} cierra el podio con un rendimiento muy equilibrado. Demuestra que para estar en el grupo de avanzada no se puede regalar nada, sumando minutos de calidad en las tres disciplinas."
            return texto_comentario
            
        return f"Desempeño consistente de {identidad_atleta} en la zona alta de la tabla clasificatoria del club TYM."
    
    # Análisis para el equilibrio de carga técnica (CV)
    if nombre_categoria == 'CV':
        valor_cv_fila = datos_de_fila.get('CV', 0)
        texto_comentario = f"¡El reloj suizo del club! {identidad_atleta} logra una simetría casi perfecta ({valor_cv_fila}), demostrando una planificación milimétrica de sus cargas y un control total del entrenamiento en todas las áreas técnicas."
        return texto_comentario
    
    # Análisis por disciplina individual técnica
    tiempo_especifico_txt = datos_de_fila.get(nombre_categoria, "00:00:00")
    
    if nombre_categoria == 'Natación':
        texto_comentario = f"Fuerza pura en el agua. {identidad_atleta} registra un tiempo de {tiempo_especifico_txt}, liderando el podio de la disciplina con una técnica depurada y un gran volumen acumulado en la pileta."
        return texto_comentario
        
    if nombre_categoria == 'Bicicleta':
        texto_comentario = f"Potencia pura sobre ruedas. {identidad_atleta} devoró la ruta con un tiempo de {tiempo_especifico_txt}. Demuestra ser el gran motor del equipo en la carretera con promedios que intimidan a sus rivales."
        return texto_comentario
        
    if nombre_categoria == 'Trote':
        texto_comentario = f"Resistencia inalcanzable. {identidad_atleta} domina el asfalto con un tiempo de {tiempo_especifico_txt} y una fase de carrera soberbia, cerrando una semana de alta calidad técnica en la zancada."
        return texto_comentario
    
    return "Desempeño técnico destacado durante la jornada de entrenamiento semanal del equipo TYM."

# *****************************************************************************
# --- 4. PARSERS DE ENTRADA (BLINDADO - NO SINTETIZAR) ---
# *****************************************************************************

def parse_raw_data(bloque_input_strava):
    """
    Procesa el bloque de texto copiado de Strava (Tiempo Total).
    No utiliza síntesis; cada paso de extracción es explícito y visible.
    """
    lista_de_registros_atleta = []
    valor_rank_contador = 1
    
    # Limpieza de caracteres de control web (espacios de no ruptura)
    bloque_input_strava = bloque_input_strava.replace('\xa0', ' ')
    lineas_encontradas = bloque_input_strava.strip().split('\n')
    
    for fila_texto in lineas_encontradas:
        if not fila_texto:
            continue
            
        if 'Deportista' in fila_texto:
            continue
            
        try:
            # Expresión regular para detectar tiempos con formato h y min
            patron_tiempos = r'(\d+h\s*\d*min|\d+h|\d+min|--:--)'
            tiempos_en_linea = re.findall(patron_tiempos, fila_texto)
            
            # 🛡️ CORRECCIÓN SINTAXIS AUDITADA:
            if not tiempos_en_linea:
                continue
                
            # El Tiempo Total es siempre el primer elemento detectado
            string_del_total = tiempos_en_linea[0]
            ubicacion_del_tiempo = fila_texto.find(string_del_total)
            
            # El nombre del deportista precede a la cifra de tiempo
            segmento_del_nombre = fila_texto[:ubicacion_del_tiempo].strip()
            
            # Limpieza del número de ranking si está presente en el copiado (ej: "1 Rodrigo")
            nombre_limpio_final = re.sub(r'^\d+\s*', '', segmento_del_nombre).strip()
            
            # Conversión de los bloques de tiempo a minutos enteros
            minutos_volumen_total = to_mins(string_del_total)
            
            minutos_nat = 0
            if len(tiempos_en_linea) > 1:
                minutos_nat = to_mins(tiempos_en_linea[1])
                
            minutos_bici = 0
            if len(tiempos_en_linea) > 2:
                minutos_bici = to_mins(tiempos_en_linea[2])
                
            minutos_trote = 0
            if len(tiempos_en_linea) > 3:
                minutos_trote = to_mins(tiempos_en_linea[3])
                
            # Cálculo del Coeficiente de Variación (CV)
            lista_tiempos_cv = [minutos_nat, minutos_bici, minutos_trote]
            
            if 0 in lista_tiempos_cv:
                valor_cv_final = "NC"
            else:
                calculo_std = np.std(lista_tiempos_cv)
                calculo_mean = np.mean(lista_tiempos_cv)
                valor_cv_final = round(calculo_std / calculo_mean, 4)
            
            # Extracción del conteo de actividades (dato tras el tiempo total)
            segmento_final_linea = fila_texto[ubicacion_del_tiempo + len(string_total := string_del_total):]
            match_de_actividades = re.search(r'\d+', segmento_final_linea)
            
            numero_de_actividades = 0
            if match_de_actividades:
                numero_de_actividades = int(match_de_actividades.group())
            
            # Construcción del registro detallado por cada deportista
            diccionario_de_atleta = {
                '#': valor_rank_contador,
                'Deportista': nombre_limpio_final,
                'Tiempo Total': to_hhmmss_display(minutos_volumen_total),
                'Actividades': numero_de_actividades,
                'Natación': to_hhmmss_display(minutos_nat),
                'Bicicleta': to_hhmmss_display(minutos_bici),
                'Trote': to_hhmmss_display(minutos_trote),
                'CV': valor_cv_final,
                'T_Mins': minutos_volumen_total,
                'N_Mins': minutos_nat,
                'B_Mins': minutos_bici,
                'R_Mins': minutos_trote
            }
            
            lista_de_registros_atleta.append(diccionario_de_atleta)
            valor_rank_contador = valor_rank_contador + 1
            
        except Exception:
            # Omisión de líneas corruptas o sin formato válido
            continue
            
    # Retorno estructurado para procesamiento masivo en hojas de cálculo
    df_resultado_parsing = pd.DataFrame(lista_de_registros_atleta)
    
    return df_resultado_parsing

def parse_ocr_data(texto_ocr_crudo):
    """
    Parsea la tabla de traducción OCR (Distancia y Salida Larga).
    Filtra mandatoriamente los encabezados técnicos de la tabla procesada.
    """
    lista_podio_distancia = []
    lista_podio_larga = []
    
    # Listado de términos prohibidos en las celdas de datos (encabezados)
    filtro_nombres_tabla = ["Nombre", "Distancia", "Actividad", "Tiempo", "Km", "total", "Clasificación"]
    
    lineas_de_la_entrada_ocr = texto_ocr_crudo.strip().split('\n')
    
    for fila_ocr in lineas_de_la_entrada_ocr:
        celdas_ocr = fila_ocr.split(';')
        
        if len(celdas_ocr) >= 6:
            # Datos pertenecientes al bloque de Distancia Total
            nombre_atleta_d = celdas_ocr[2].strip()
            valor_metrica_d = celdas_ocr[3].strip()
            
            # Datos pertenecientes al bloque de Salida Larga semanal
            nombre_atleta_l = celdas_ocr[4].strip()
            valor_metrica_l = celdas_ocr[5].strip()
            
            # Validación de integridad para columna de Distancia
            es_titulo_d = False
            for term_f in filtro_nombres_tabla:
                if term_f.lower() in nombre_atleta_d.lower():
                    es_titulo_d = True
            
            if es_titulo_d == False and nombre_atleta_d != "":
                lista_podio_distancia.append({'nombre': nombre_atleta_d, 'valor': valor_metrica_d})
                
            # Validación de integridad para columna de Salida Larga
            es_titulo_l = False
            for term_f in filtro_nombres_tabla:
                if term_f.lower() in nombre_atleta_l.lower():
                    es_titulo_l = True
                    
            if es_titulo_l == False and nombre_atleta_l != "":
                lista_podio_larga.append({'nombre': nombre_atleta_l, 'valor': valor_metrica_l})
                
    # Retornar exclusivamente el Top 3 de cada categoría de honor
    return lista_podio_distancia[:3], lista_podio_larga[:3]

# *****************************************************************************
# --- 5. ACTUALIZADOR DE EXCEL (OBJETIVO: INTEGRIDAD Y ORDEN NUMÉRICO) ---
# *****************************************************************************

def crear_excel_actualizado(referencia_maestro, df_actualizacion, input_semana_n):
    """
    Genera el archivo Excel inyectando la nueva semana y eliminando columnas futuras.
    Lógica de ordenamiento numérico real implementada para evitar errores en el historial.
    """
    # Carga del libro maestro con tipos de datos protegidos (object)
    lector_maestro_excel = pd.ExcelFile(referencia_maestro)
    hojas_existentes_maestro = lector_maestro_excel.sheet_names
    label_de_la_semana_actual = f"Sem {input_semana_n.strip()}"
    
    # 🛡️ FILTRO DE INGENIERÍA: Extraer el límite numérico de la semana procesada
    match_numero_semana = re.search(r'\d+', input_semana_n)
    valor_entero_limite = int(match_numero_semana.group()) if match_numero_semana else 0
    
    # DETERMINACIÓN DEL ORDEN OPERATIVO MANDATORIO (PROTOCOLO TYM):
    
    # Bloque 1: Identificación de Hojas Técnicas de Trabajo
    hojas_de_trabajo_lista = []
    for h_name in hojas_existentes_maestro:
        if not h_name.startswith("Sem "):
            hojas_de_trabajo_lista.append(h_name)
            
    # Bloque 2: Identificación y limpieza del historial semanal
    hojas_del_historial_lista = []
    for h_name in hojas_existentes_maestro:
        if h_name.startswith("Sem "):
            # Condición crítica: No duplicar la semana cargada actualmente
            if h_name != label_de_la_semana_actual:
                hojas_del_historial_lista.append(h_name)
    
    # Bloque 3: Ordenamiento Numérico Real (Auditado)
    def extraer_numero_pestaña(texto_pestaña):
        match_p = re.search(r'\d+', str(texto_pestaña))
        if match_p:
            return int(match_p.group())
        return 0
        
    hojas_del_historial_lista.sort(key=extraer_numero_pestaña, reverse=True)
    
    # Bloque 4: Construcción del libro final de salida
    secuencia_de_hojas_final = hojas_de_trabajo_lista + [label_de_la_semana_actual] + hojas_del_historial_lista

    # Creación del stream de memoria binario
    buffer_binario_descarga = io.BytesIO()
    
    with pd.ExcelWriter(buffer_binario_descarga, engine='xlsxwriter') as motor_escritura_excel:
        libro_excel_obj = motor_escritura_excel.book
        
        # Formato de tiempo TYM para celdas operables numéricamente
        formato_hora_tym = libro_excel_obj.add_format({'num_format': '[h]:mm:ss'})
        
        for hoja_en_curso in secuencia_de_hojas_final:
            
            # ESCENARIO A: GENERACIÓN DE LA NUEVA PESTAÑA SEMANAL DETALLADA
            if hoja_en_curso == label_de_la_semana_actual:
                cols_requeridas_sem = ['#', 'Deportista', 'Tiempo Total', 'Actividades', 'Natación', 'Bicicleta', 'Trote', 'CV']
                df_hoja_semana_final = df_actualizacion[cols_requeridas_sem].copy()
                df_hoja_semana_final.rename(columns={'#': 'Clasificación'}, inplace=True)
                
                # Conversión mandatoria a valores decimales de Excel
                for columna_t_name in ['Tiempo Total', 'Natación', 'Bicicleta', 'Trote']:
                    df_hoja_semana_final[columna_t_name] = df_hoja_semana_final[columna_t_name].apply(to_excel_time_value)
                
                # Escritura de la pestaña de clasificación
                df_hoja_semana_final.to_excel(motor_escritura_excel, sheet_name=hoja_en_curso, index=False)
                
                # Formateo visual de celdas horarias
                pestaña_actual_obj = motor_escritura_excel.sheets[hoja_en_curso]
                # Índices de columnas de tiempo (C, E, F, G)
                for id_columna_v in [2, 4, 5, 6]:
                    pestaña_actual_obj.set_column(id_columna_v, id_columna_v, 12, formato_hora_tym)
            
            # ESCENARIO B: ACTUALIZACIÓN DE LAS HOJAS DE TRABAJO TÉCNICAS (CON FILTRO DE COLUMNAS)
            elif hoja_en_curso in ["Tiempo Total", "Natación", "Ciclismo", "Trote"]:
                # Lectura blindada de la hoja histórica
                df_maestro_h_tecnica = pd.read_excel(lector_maestro_excel, sheet_name=hoja_en_curso, dtype=object)
                
                # 🛡️ FILTRO DE SEGURIDAD DE COLUMNAS (PROTOCOLO V2.2.19)
                # Remoción física de columnas "Sem XX" que exceden la semana actual procesada
                lista_de_cols_desbordadas = []
                for nombre_col_check in df_maestro_h_tecnica.columns:
                    str_check_nom = str(nombre_col_check)
                    if str_check_nom.startswith("Sem "):
                        match_num_check = re.search(r'\d+', str_check_nom)
                        if match_num_check:
                            if int(match_num_check.group()) > valor_entero_limite:
                                lista_de_cols_desbordadas.append(nombre_col_check)
                                
                if len(lista_de_cols_desbordadas) > 0:
                    df_maestro_h_tecnica = df_maestro_h_tecnica.drop(columns=lista_de_cols_desbordadas)
                
                # Sello de Transcripción Aritmética para columnas de cálculo
                cols_arit_tecnicas = []
                for c_header_t in df_maestro_h_tecnica.columns:
                    c_header_t_str = str(c_header_t)
                    if c_header_t_str.startswith("Sem ") or "Promedio" in c_header_t_str or "Acumulado" in c_header_t_str:
                        cols_arit_tecnicas.append(c_header_t)
                
                for col_a_t in cols_arit_tecnicas:
                    if col_a_t != label_de_la_semana_actual:
                        df_maestro_h_tecnica[col_a_t] = df_maestro_h_tecnica[col_a_t].apply(to_excel_time_value)
                
                # Match de deportistas mediante columna identificadora
                id_col_identificadora = df_maestro_h_tecnica.columns[0]
                for col_it_t in df_maestro_h_tecnica.columns:
                    if "nombre" in str(col_it_t).lower() or "deportista" in str(col_it_t).lower():
                        id_col_identificadora = col_it_t
                        break
                
                dict_mapeo_hojas_tym = {'Tiempo Total': 'Tiempo Total', 'Natación': 'Natación', 'Ciclismo': 'Bicicleta', 'Trote': 'Trote'}
                nombre_fuente_actual_tym = dict_mapeo_hojas_tym.get(hoja_en_curso)
                
                if nombre_fuente_actual_tym:
                    df_prep_normalizacion = df_actualizacion[['Deportista', nombre_fuente_actual_tym]].copy()
                    df_prep_normalizacion['MatchKey'] = df_prep_normalizacion['Deportista'].apply(clean_string)
                    df_maestro_h_tecnica['MatchKey'] = df_maestro_h_tecnica[id_col_identificadora].astype(str).apply(clean_string)
                    dict_valores_inyectar = df_prep_normalizacion.set_index('MatchKey')[nombre_fuente_actual_tym].apply(to_excel_time_value).to_dict()
                    df_maestro_h_tecnica[label_de_la_semana_actual] = df_maestro_h_tecnica['MatchKey'].map(dict_valores_inyectar).fillna(0)
                    
                    # RECALCULO DE COLUMNA TIEMPO ACUMULADO
                    lista_sem_para_suma = [c_f_s for c_f_s in df_maestro_h_tecnica.columns if str(c_f_s).startswith("Sem ")]
                    if 'Tiempo Acumulado' in df_maestro_h_tecnica.columns:
                        df_maestro_h_tecnica['Tiempo Acumulado'] = df_maestro_h_tecnica[lista_sem_para_suma].sum(axis=1)
                    
                    df_maestro_h_tecnica = df_maestro_h_tecnica.drop(columns=['MatchKey'])
                
                # Escritura definitiva de la hoja técnica
                df_maestro_h_tecnica.to_excel(motor_escritura_excel, sheet_name=hoja_en_curso, index=False)
                
                # Formatear visualmente las columnas de cálculo matemático
                pestaña_trabajo_activa = motor_escritura_excel.sheets[hoja_en_curso]
                for id_c_f, nom_c_f in enumerate(df_maestro_h_tecnica.columns):
                    nom_c_f_str = str(nom_c_f)
                    if nom_c_f_str.startswith("Sem ") or "Promedio" in nom_c_f_str or "Acumulado" in nom_c_f_str:
                        pestaña_trabajo_activa.set_column(id_c_f, id_c_f, 13, formato_hora_tym)

            # ESCENARIO C: ACTUALIZACIÓN DE LA HOJA COEFICIENTE DE VARIACIÓN (CV)
            elif hoja_en_curso == "CV":
                df_maestro_cv_hoja = pd.read_excel(lector_maestro_excel, sheet_name=hoja_en_curso, dtype=object)
                
                # 🛡️ BLINDAJE REGEX CV: Filtrado de columnas futuras
                cols_cv_a_borrar = []
                for c_cv_iterador in df_maestro_cv_hoja.columns:
                    if str(c_cv_iterador).startswith("Sem "):
                        m_cv_check = re.search(r'\d+', str(c_cv_iterador))
                        if m_cv_check:
                            if int(m_cv_check.group()) > valor_entero_limite:
                                cols_cv_a_borrar.append(c_cv_iterador)
                
                if len(cols_cv_a_borrar) > 0:
                    df_maestro_cv_hoja = df_maestro_cv_hoja.drop(columns=cols_cv_a_borrar)
                
                id_col_nombre_en_cv = next((c_cv for c_cv in df_maestro_cv_hoja.columns if "nombre" in str(c_cv).lower() or "deportista" in str(c_cv).lower()), df_maestro_cv_hoja.columns[0])
                df_maestro_cv_hoja['MatchKey'] = df_maestro_cv_hoja[id_col_nombre_en_cv].astype(str).apply(clean_string)
                df_cv_semana_prep = df_actualizacion[['Deportista', 'CV']].copy()
                df_cv_semana_prep['MatchKey'] = df_cv_semana_prep['Deportista'].apply(clean_string)
                df_maestro_cv_hoja[label_de_la_semana_actual] = df_maestro_cv_hoja['MatchKey'].map(df_cv_semana_prep.set_index('MatchKey')['CV'].to_dict()).fillna("NC")
                df_maestro_cv_hoja.drop(columns=['MatchKey']).to_excel(motor_escritura_excel, sheet_name=hoja_en_curso, index=False)

            # ESCENARIO D: RÉPLICA DE TODAS LAS DEMÁS PESTAÑAS DEL LIBRO
            else:
                df_replica_hoja_estatica = pd.read_excel(lector_maestro_excel, sheet_name=hoja_en_curso, dtype=object)
                df_replica_hoja_estatica.to_excel(motor_escritura_excel, sheet_name=hoja_en_curso, index=False)
                
    return buffer_binario_descarga.getvalue()

# *****************************************************************************
# --- 6. GENERADOR DE REPORTE WORD PROFESIONAL (BLOQUEADO - NO SINTETIZAR) ---
# *****************************************************************************

def aplicar_formato_tym_word(objeto_parrafo_word, pt_fuente, negrita_activo=False, centrado_activo=False):
    """
    Aplica rigurosamente el estilo institucional Calibri con los tamaños 20/15/13/11.
    Este bloque está blindado contra síntesis para asegurar la imagen del club.
    """
    if not objeto_parrafo_word.runs:
        cursor_de_run = objeto_parrafo_word.add_run()
    else:
        cursor_de_run = objeto_parrafo_word.runs[0]
        
    cursor_de_run.font.name = 'Calibri'
    cursor_de_run.font.size = Pt(pt_fuente)
    cursor_de_run.bold = negrita_activo
    
    if centrado_activo:
        objeto_parrafo_word.alignment = WD_ALIGN_PARAGRAPH.CENTER

def crear_tabla_profesional_tym_word(doc_word_instancia, df_origen_datos, lista_de_cabeceras):
    """
    Genera tablas con anchos milimétricos (Protocolo TYM) para el informe profesional.
    Blindado contra saltos de línea inesperados.
    """
    instancia_de_tabla = doc_word_instancia.add_table(rows=1, cols=len(lista_de_cabeceras))
    instancia_de_tabla.style = 'Light Grid Accent 1'
    instancia_de_tabla.alignment = 1 # Centrado
    instancia_de_tabla.autofit = False
    
    # Anchos fijos por ingeniería
    anchos_tym_fijos = {'#': 0.4, 'Deportista': 2.8, 'Tiempo Total': 0.7, 'Natación': 0.7, 'Bicicleta': 0.7, 'Trote': 0.7, 'CV': 0.6}
    
    for idx_cab, txt_cab in enumerate(lista_de_cabeceras):
        celda_header_t = instancia_de_tabla.rows[0].cells[idx_cab]
        celda_header_t.text = txt_cab
        ancho_val_t = anchos_tym_fijos.get(txt_cab, 0.7)
        celda_header_t.width = Inches(ancho_val_t)
        aplicar_formato_tym_word(celda_header_t.paragraphs[0], 9, True, True)
        
    for _, fila_datos_w in df_origen_datos.iterrows():
        celdas_de_la_fila_w = instancia_de_tabla.add_row().cells
        for idx_dat, txt_cab_d in enumerate(lista_de_cabeceras):
            celdas_de_la_fila_w[idx_dat].text = str(fila_datos_w[txt_cab_d])
            ancho_dat_w = anchos_tym_fijos.get(txt_cab_d, 0.7)
            celdas_de_la_fila_w[idx_dat].width = Inches(ancho_dat_w)
            
            # Alineación Nombres Izquierda, resto Centro
            es_central = True
            if txt_cab_d == 'Deportista':
                es_central = False
                
            aplicar_formato_tym_word(celdas_de_la_fila_w[idx_dat].paragraphs[0], 9, False, es_central)
            
    doc_word_instancia.add_paragraph()

def generar_reporte_word_tym_completo(df_semanal_datos, num_sem_texto, podio_d_lista, podio_l_lista):
    """
    Construye el reporte Word íntegro bajo el modelo funcional V2.2.19.
    Restablece todas las secciones de análisis técnico de forma extensa.
    """
    documento_final_word = Document()
    
    # Título Principal
    p_main_title_word = documento_final_word.add_heading(f'Reporte Semanal Club Tym Triatlón - Semana {num_sem_texto}', 0)
    aplicar_formato_tym_word(p_main_title_word, 20, True, True)
    documento_final_word.add_paragraph()
    
    p_slogan_word_f = documento_final_word.add_paragraph('"(La semana de la simetría perfecta y el retorno del volumen)"')
    aplicar_formato_tym_word(p_slogan_word_f, 11, True, True)
    documento_final_word.add_paragraph()

    # BLOQUE 1: Resumen General
    h_resumen_word_f = documento_final_word.add_heading('🔍 Resumen General', level=2)
    aplicar_formato_tym_word(h_resumen_word_f, 15, True)
    documento_final_word.add_paragraph()
    
    df_filtrado_trias = df_semanal_datos[df_semanal_datos['CV'] != 'NC'].copy()
    txt_resumen_bloque_f = f"Total deportistas registrados: {len(df_semanal_datos)}\nTriatletas completos: {len(df_filtrado_trias)}\nHoras totales del club: {to_hhmmss_display(df_semanal_datos['T_Mins'].sum())}"
    p_info_word_f = documento_final_word.add_paragraph(txt_resumen_bloque_f); aplicar_formato_tym_word(p_info_word_f, 11)
    
    # Gráfico de Torta
    fig_w_f, ax_w_f = plt.subplots(figsize=(4,4))
    ax_w_f.pie([df_semanal_datos['N_Mins'].sum(), df_semanal_datos['B_Mins'].sum(), df_semanal_datos['R_Mins'].sum()], labels=['Natación', 'Ciclismo', 'Trote'], autopct='%1.1f%%', colors=['#1E90FF', '#32CD32', '#FF4500'])
    
    buffer_grafico_f = io.BytesIO()
    plt.savefig(buffer_grafico_f, format='png', bbox_inches='tight')
    plt.close(fig_w_f)
    
    p_graf_f = documento_final_word.add_paragraph()
    p_graf_f.alignment = 1 # Centrado
    p_graf_f.add_run().add_picture(buffer_grafico_f, width=Inches(3.5))

    # BLOQUE 2: Podios Honor
    for t_pod_f, d_pod_f, c_key_f in [('🏅 TOP 5 TRIATLETAS COMPLETOS', df_filtrado_trias.sort_values('T_Mins', ascending=False).head(5), 'Completos'), ('⚖️ TOP 5 TRIATLETAS MÁS BALANCEADOS', df_filtrado_trias.sort_values('CV', ascending=True).head(5), 'CV')]:
        h_s_f = documento_final_word.add_heading(t_pod_f, level=2); aplicar_formato_tym_word(h_s_f, 15, True); documento_final_word.add_paragraph()
        d_ren_f = d_pod_f.copy(); d_ren_f['#'] = range(1, len(d_ren_f) + 1)
        crear_tabla_profesional_tym_word(documento_final_word, d_ren_f, ['#', 'Deportista', 'Tiempo Total' if c_key_f=='Completos' else 'CV', 'Natación', 'Bicicleta', 'Trote'])
        h_analisis_f = documento_final_word.add_paragraph('Análisis del Desempeño:'); aplicar_formato_tym_word(h_analisis_f, 13, True)
        for _, fila_f_loop in d_ren_f.iterrows():
            p_n_f = documento_final_word.add_paragraph(f"{fila_f_loop['#']}. {fila_f_loop['Deportista']}"); aplicar_formato_tym_word(p_n_f, 11, True)
            documento_final_word.add_paragraph(generar_comentario(fila_f_loop, c_key_f, fila_f_loop['#']))

    # BLOQUE 3: TOP 15
    for tit_s_f, ico_f, col_m_f, col_t_f in [('TIEMPO GENERAL', '🥇', 'T_Mins', 'Tiempo Total'), ('NATACIÓN', '🏊‍♂️', 'N_Mins', 'Natación'), ('CICLISMO', '🚴', 'B_Mins', 'Bicicleta'), ('TROTE', '🏃‍♂️', 'R_Mins', 'Trote')]:
        documento_final_word.add_page_break()
        h_15_f = documento_final_word.add_heading(f'{ico_f} TOP 15 {tit_s_f}', level=1); aplicar_formato_tym_word(h_15_f, 15, True); documento_final_word.add_paragraph()
        d_15_f = df_semanal_datos[df_semanal_datos[col_m_f] > 0].sort_values(col_m_f, ascending=False).head(15).copy(); d_15_f['#'] = range(1, len(d_15_f) + 1)
        crear_tabla_profesional_tym_word(documento_final_word, d_15_f, ['#', 'Deportista', col_t_f, 'Natación', 'Bicicleta', 'Trote'] if tit_s_f == 'TIEMPO GENERAL' else ['#', 'Deportista', col_t_f, 'Tiempo Total'])
        h_podio_f = documento_final_word.add_paragraph('Análisis del Podio:'); aplicar_formato_tym_word(h_podio_f, 13, True)
        for _, f_p_f in d_15_f.head(3).iterrows():
            p_at_f = documento_final_word.add_paragraph(f"{'🥇' if f_p_f['#']==1 else '🥈' if f_p_f['#']==2 else '🥉'} {f_p_f['Deportista']}"); aplicar_formato_tym_word(p_at_f, 11, True)
            documento_final_word.add_paragraph(generar_comentario(f_p_f, 'General' if tit_s_f == 'TIEMPO GENERAL' else col_t_f, f_p_f['#']))

    # BLOQUE 4: OCR
    documento_final_word.add_page_break()
    h_d_w_f = documento_final_word.add_heading('📏 PODIO DISTANCIA TOTAL', level=1); aplicar_formato_tym_word(h_d_w_f, 15, True); documento_final_word.add_paragraph()
    for idx_d_f, item_d_f in enumerate(podio_d_lista): documento_final_word.add_paragraph(f"{idx_d_f+1}. {item_d_f['nombre']} ({item_d_f['valor']} km)")
    documento_final_word.add_paragraph(); h_l_w_f = documento_final_word.add_heading('⏱️ PODIO ACTIVIDAD MÁS LARGA', level=1); aplicar_formato_tym_word(h_l_w_f, 15, True); documento_final_word.add_paragraph()
    for idx_l_f, item_l_f in enumerate(podio_l_lista): documento_final_word.add_paragraph(f"{idx_l_f+1}. {item_l_f['nombre']} ({item_l_f['valor']})")
    
    stream_salida_word = io.BytesIO(); documento_final_word.save(stream_salida_word); stream_salida_word.seek(0); return stream_salida_word
    
# --- NUEVA FUNCIÓN: MOTOR NARRATIVO INDIVIDUAL (MODO INYECCIÓN) ---
def generar_reporte_narrativo_individual(atleta_nom, df_actual, dict_historicos, sem_n):
    """Genera reporte personal con insights narrativos basados en Colab V17.0."""
    doc_p = Document()
    match_key = clean_string(atleta_nom)
    
    # Header Personalizado
    p_h = doc_p.add_heading(f'Análisis de Rendimiento Personal: {atleta_nom}', 0)
    aplicar_formato_tym_word(p_h, 18, True, True)
    doc_p.add_paragraph(f"Semana de Entrenamiento: {sem_n}").alignment = 1
    
    # Localizar fila actual
    df_actual['MatchKey'] = df_actual['Deportista'].apply(clean_string)
    row_act = df_actual[df_actual['MatchKey'] == match_key]
    if row_act.empty: return None
    row_act = row_act.iloc[0]

    def get_benchmarks(hoja_key, match_k):
        df_h = dict_historicos.get(hoja_key)
        if df_h is None: return 0, 0
        df_h['MatchKey'] = df_h.iloc[:, 0].astype(str).apply(clean_string)
        cols_sem = [c for c in df_h.columns if str(c).startswith("Sem ")]
        # Promedio del Equipo en esta semana
        prefijo = "N" if "Natación" in hoja_key else "B" if "Ciclismo" in hoja_key else "R" if "Trote" in hoja_key else "T"
        avg_equipo = df_actual[f"{prefijo}_Mins"].mean()
        # Promedio Histórico del Atleta
        r_atleta = df_h[df_h['MatchKey'] == match_k]
        avg_hist = r_atleta[cols_sem].mean(axis=1).iloc[0] * 1440 if not r_atleta.empty else 0
        return avg_equipo, avg_hist

    # Bloques de Disciplina
    for tit_m, hoja_m, col_m in [("TIEMPO TOTAL", "Tiempo Total", "T_Mins"), ("NATACIÓN", "Natación", "N_Mins"), ("CICLISMO", "Ciclismo", "B_Mins"), ("TROTE", "Trote", "R_Mins")]:
        doc_p.add_heading(tit_m, level=1)
        val_act = row_act[col_m]
        bench_eq, bench_hi = get_benchmarks(hoja_m, match_key)
        
        p_m = doc_p.add_paragraph()
        p_m.add_run(f"Volumen actual: {to_hhmmss_display(val_act)}\n").bold = True
        
        diff_eq = val_act - bench_eq
        txt_eq = f"Vs Equipo: {to_hhmmss_display(abs(diff_eq))} {'MÁS' if diff_eq > 0 else 'MENOS'}."
        run_eq = p_m.add_run(txt_eq)
        # Verde si es positivo, rojo si es negativo
        from docx.shared import RGBColor
        run_eq.font.color.rgb = RGBColor(0, 100, 0) if diff_eq >= 0 else RGBColor(180, 0, 0)
        
        diff_hi = val_act - bench_hi
        p_m.add_run(f"\nVs Tu Media Histórica: {to_hhmmss_display(abs(diff_hi))} {'MÁS' if diff_hi > 0 else 'MENOS'}.")

    # Gráfico Personal
    fig_p, ax_p = plt.subplots(figsize=(5,3))
    ax_p.bar(['Nat', 'Bici', 'Tro'], [row_act['N_Mins'], row_act['B_Mins'], row_act['R_Mins']], color=['#1E90FF', '#32CD32', '#FF4500'])
    ax_p.set_title("Tu Distribución de Carga (Minutos)")
    buf_p = io.BytesIO(); plt.savefig(buf_p, format='png', bbox_inches='tight'); plt.close(fig_p)
    doc_p.add_paragraph().add_run().add_picture(buf_p, width=Inches(3.5))
    
    doc_p.add_paragraph("─" * 50)
    doc_p.add_paragraph("Generado por Agente TYM 2026").alignment = 2
    b_out = io.BytesIO(); doc_p.save(b_out); b_out.seek(0); return b_out
    
# *****************************************************************************
# --- 7. INTERFAZ DE USUARIO (STREMLIT) ---
# *****************************************************************************

st.sidebar.header("📁 Gestión de Datos Históricos TYM")
cargador_maestro_excel = st.sidebar.file_uploader("Cargar Excel Maestro", type=["xlsx"])
num_semana_procesar = st.text_input("Número de Semana (Ej: 08):", "08")
area_texto_strava = st.text_area("1. Datos Tiempo Total (Strava):")
area_texto_ocr = st.text_area("2. Datos OCR (Traducción):")

if st.button("🚀 PROCESAR JORNADA"):
    if cargador_maestro_excel and area_texto_strava.strip() and area_texto_ocr.strip():
        # Procesar Parsing
        df_resultados = parse_raw_data(area_texto_strava)
        d_p_dist, d_p_larg = parse_ocr_data(area_texto_ocr)
        st.success(f"¡Semana {num_semana_procesar} procesada!"); col1, col2 = st.columns(2)
        
        # Word
        col1.download_button(label="📄 REPORTE WORD", data=generar_reporte_word_tym_completo(df_resultados, num_semana_procesar, d_p_dist, d_p_larg), file_name=f"Reporte_TYM_{num_semana_procesar}.docx")
        
        # Excel
        col2.download_button(label="📊 EXCEL ACTUALIZADO", data=crear_excel_actualizado(cargador_maestro_excel, df_resultados, num_semana_procesar), file_name=f"00_Estadisticas_Actualizadas_{num_semana_procesar}.xlsx")
# --- INYECCIÓN DE INTERFAZ PARA REPORTES INDIVIDUALES ---
        st.divider()
        st.subheader("👤 Generador de Insights Individuales")
        
        # Cargamos los datos históricos necesarios para la comparativa
        h_t = pd.read_excel(maestro_uploader, sheet_name="Tiempo Total", dtype=object)
        h_n = pd.read_excel(maestro_uploader, sheet_name="Natación", dtype=object)
        h_c = pd.read_excel(maestro_uploader, sheet_name="Ciclismo", dtype=object)
        h_r = pd.read_excel(maestro_uploader, sheet_name="Trote", dtype=object)
        dict_h_ref = {"Tiempo Total": h_t, "Natación": h_n, "Ciclismo": h_c, "Trote": h_r}
        
        atletas_list = df_res['Deportista'].tolist()
        seleccionados = st.multiselect("Seleccionar Atletas para Reporte Individual:", atletas_list)
        
        if seleccionados:
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zf:
                for a_sel in seleccionados:
                    r_indiv = generar_reporte_narrativo_individual(a_sel, df_res, dict_h_ref, input_num_semana)
                    if r_indiv:
                        zf.writestr(f"Reporte_{clean_string(a_sel)}.docx", r_indiv.getvalue())
            
            st.download_button(
                label="⬇️ DESCARGAR ZIP INDIVIDUALES", 
                data=zip_buffer.getvalue(), 
                file_name=f"Individuales_Sem_{input_num_semana}.zip", 
                mime="application/zip"
            )
    else:
        st.error("Error Mandatorio: Excel y campos de texto no pueden estar vacíos.")
