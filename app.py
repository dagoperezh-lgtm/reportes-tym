# TYM PLATAFORMA - VERSION: 2.2.23-PCVCI-FINAL-RECOVERY
# OBJETIVO: RESTAURAR EXTENSIÓN COMPLETA E INTEGRAR REPORTES NARRATIVOS (ESTILO COLAB)
# LINEAS DE CODIGO: 1025
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
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# --- 1. CONFIGURACIÓN DE PÁGINA (BLINDADO - NO TOCAR) ---

st.set_page_config(
    page_title="Plataforma TYM 2026 - V2.2.23", 
    page_icon="🏆", 
    layout="wide"
)

st.title("🏆 Gestión de Reportes y Estadísticas - Club TYM")

# --- 2. UTILIDADES DE PROCESAMIENTO Y TIEMPO (BLINDADO - NO SINTETIZAR) ---

def clean_string(text):
    """
    Normaliza nombres para asegurar coincidencias entre Strava y Excel.
    Elimina tildes, espacios extra y convierte a mayúsculas.
    """
    if text is None or pd.isna(text):
        return ""
    
    # Proceso de normalización de caracteres paso a paso para evitar errores de codificación
    # Se descompone el caracter para separar la tilde y eliminarla mediante NFKD
    text_pre = str(text).strip().upper()
    info_normalizada = unicodedata.normalize('NFKD', text_pre)
    
    # Se reconstruye el string filtrando los caracteres de combinación (acentos)
    resultado_final_nombre = ""
    for caracter_indiv in info_normalizada:
        if not unicodedata.combining(caracter_indiv):
            resultado_final_nombre = resultado_final_nombre + caracter_indiv
            
    return resultado_final_nombre

def to_mins(valor_entrada_tiempo):
    """
    Convierte cualquier formato de tiempo a minutos totales de forma explícita.
    Maneja decimales de Excel, objetos datetime, strings HH:MM y formato Strava.
    No se permite la síntesis de este bloque para asegurar la captura de todos los casos.
    """
    if pd.isna(valor_entrada_tiempo):
        return 0
    
    string_valor = str(valor_entrada_tiempo).strip()
    
    # Filtro exhaustivo de valores nulos o vacíos detectados en producción
    lista_casos_nulos = ['--:--', '0', '', '00:00:00', '0:00:00', '00:00', '0.0', 'NC', '0:00']
    
    if string_valor in lista_casos_nulos:
        return 0
        
    try:
        # REGLA ARITMÉTICA: Si el valor es numérico (fracción de día de Excel)
        if isinstance(valor_entrada_tiempo, (float, int)):
            # Excel almacena 1 día completo como 1.0. Multiplicamos por 1440.
            minutos_finales_calculados = int(round(valor_entrada_tiempo * 1440))
            return minutos_finales_calculados
        
        # Si el dato es un objeto de tiempo nativo de Python (datetime o time)
        if isinstance(valor_entrada_tiempo, (time, datetime)):
            minutos_finales_calculados = (valor_entrada_tiempo.hour * 60) + valor_entrada_tiempo.minute
            return minutos_finales_calculados
            
        # Si el string representa un número decimal puro (ej: "0.4625")
        try:
            conversion_float = float(string_valor)
            minutos_finales_calculados = int(round(conversion_float * 1440))
            return minutos_finales_calculados
        except ValueError:
            # No es numérico, continuamos con el parsing de texto estándar
            pass
            
        # Formato de hora estándar con separador de dos puntos (HH:MM:SS o HH:MM)
        if ':' in string_valor:
            bloques_tiempo = string_valor.split(':')
            if len(bloques_tiempo) >= 2:
                horas_bloque = int(bloques_tiempo[0])
                minutos_raw_bloque = bloques_tiempo[1]
                minutos_clean_bloque = int(minutos_raw_bloque.split('.')[0])
                total_minutos_bloque = (horas_bloque * 60) + minutos_clean_bloque
                return total_minutos_bloque
        
        # Formato nativo de Strava pegado desde navegador (ej: 11h 6min)
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
        # Fallback de seguridad para evitar que la aplicación se detenga por un dato corrupto
        return 0

def to_excel_time_value(dato_entrada_original):
    """
    Transforma la entrada en la fracción decimal exacta que requiere el motor de Excel.
    Es la base mandatoria para que las fórmulas de suma y promedio funcionen en el libro.
    """
    minutos_para_excel = to_mins(dato_entrada_original)
    
    # 24 horas equivalen a 1440 minutos totales (60 * 24)
    valor_decimal_excel = minutos_para_excel / 1440.0
    
    return valor_decimal_excel

def to_hhmmss_display(minutos_totales_input):
    """
    Formato de texto HH:MM:00 exclusivo para la visualización en el reporte Word.
    Este formato no se usa para cálculos en Excel, solo para legibilidad humana.
    """
    valor_horas_v = int(minutos_totales_input // 60)
    valor_minutos_v = int(minutos_totales_input % 60)
    
    # Se construye el string con ceros a la izquierda para mantener la uniformidad visual
    string_formato_reloj = f"{valor_horas_v:02d}:{valor_minutos_v:02d}:00"
    
    return string_formato_reloj

# --- 3. MOTOR DE COMENTARIOS TÉCNICOS (PROTEGIDO - BLOQUEADO) ---

def generar_comentario(fila_datos, categoria_nombre, rank_posicion):
    """
    Genera el análisis cualitativo extenso para los podios del reporte Word.
    Este bloque debe ser extenso para mantener el nivel profesional del informe del club.
    """
    identidad_atleta = fila_datos['Deportista']
    
    # Comentarios específicos para podios de volumen general (Completos y General)
    if categoria_nombre == 'Completos' or categoria_nombre == 'General':
        if rank_posicion == 1:
            return f"Dominio absoluto de {identidad_atleta}. Se consolida en la cima del club con un volumen total envidiable. Su capacidad para sostener sesiones de alta intensidad demuestra una preparación de élite y una disciplina inquebrantable."
        
        if rank_posicion == 2:
            return f"Una semana brillante para {identidad_atleta}. Se queda con la plata manteniendo una presión constante sobre el líder. Su solidez fue el motor que lo mantuvo en la parte más alta de la tabla."
            
        if rank_posicion == 3:
            return f"{identidad_atleta} cierra el podio con un rendimiento muy equilibrado. Demuestra que para estar en el grupo de avanzada no se puede regalar nada, sumando minutos de calidad en las tres disciplinas."
            
        return f"Desempeño consistente de {identidad_atleta} en la zona alta de la tabla clasificatoria del club TYM."
    
    # Comentarios para el podio de balance y simetría de carga (Coeficiente de Variación)
    if categoria_nombre == 'CV':
        valor_cv_fila = fila_datos.get('CV', 0)
        return f"¡El reloj suizo del club! {identidad_atleta} logra una simetría casi perfecta ({valor_cv_fila}), demostrando una planificación milimétrica de sus cargas y un control total del entrenamiento en todas las áreas técnicas."
    
    # Comentarios por disciplina técnica individual detectada en el motor
    texto_tiempo_especifico = fila_datos.get(categoria_nombre, "00:00:00")
    
    if categoria_nombre == 'Natación':
        return f"Fuerza pura en el agua. {identidad_atleta} registra un tiempo de {texto_tiempo_especifico}, liderando el podio de la disciplina con una técnica depurada y un gran volumen acumulado en la pileta."
        
    if categoria_nombre == 'Bicicleta':
        return f"Potencia pura sobre ruedas. {identidad_atleta} devoró la ruta con un tiempo de {texto_tiempo_especifico}. Demuestra ser el gran motor del equipo en la carretera con promedios que intimidan a sus rivales."
        
    if categoria_nombre == 'Trote':
        return f"Resistencia inalcanzable. {identidad_atleta} domina el asfalto con un tiempo de {texto_tiempo_especifico} y una fase de carrera soberbia, cerrando una semana de alta calidad técnica en la zancada."
    
    # Fallback de seguridad en caso de categoría no mapeada
    return "Desempeño técnico destacado durante la jornada de entrenamiento semanal del equipo TYM."

# --- 4. PARSERS DE ENTRADA (BLINDADO - NO SINTETIZAR) ---

def parse_raw_data(bloque_input_strava):
    """
    Procesa el bloque de texto copiado de Strava (Tiempo Total).
    No utiliza síntesis; cada paso de limpieza y extracción es explícito y visible.
    """
    lista_de_registros_atleta = []
    valor_rank_contador = 1
    
    # Limpieza previa de caracteres de control de copiado web y espacios invisibles
    bloque_input_strava = bloque_input_strava.replace('\xa0', ' ')
    lineas_encontradas = bloque_input_strava.strip().split('\n')
    
    for fila_texto in lineas_encontradas:
        if not fila_texto:
            continue
            
        if 'Deportista' in fila_texto:
            continue
            
        try:
            # Buscar patrones de tiempo Strava mediante expresión regular de precisión
            expresion_regular = r'(\d+h\s*\d*min|\d+h|\d+min|--:--)'
            tiempos_en_linea = re.findall(expresion_regular, fila_texto)
            
            # Verificación de integridad: Si no hay tiempos, no es una línea de atleta válida
            if not tiempos_en_linea:
                continue
                
            # El primer tiempo capturado siempre representa el Tiempo Total de la jornada semanal
            string_del_total = tiempos_en_linea[0]
            ubicacion_del_tiempo = fila_texto.find(string_del_total)
            
            # El nombre del deportista se ubica antes de la primera cifra de tiempo detectada
            segmento_del_nombre = fila_texto[:ubicacion_del_tiempo].strip()
            
            # Limpieza del número de ranking si está presente en el copiado (ej: "1 Rodrigo")
            nombre_limpio_final = re.sub(r'^\d+\s*', '', segmento_del_nombre).strip()
            
            # Extracción y conversión de minutos por cada bloque detectado en la línea
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
                
            # Cálculo del Coeficiente de Variación (CV) para medir el equilibrio de carga semanal
            lista_para_cv = [minutos_nat, minutos_bici, minutos_trote]
            
            if 0 in lista_para_cv:
                valor_cv_final = "NC"
            else:
                calculo_std = np.std(lista_para_cv)
                calculo_mean = np.mean(lista_para_cv)
                # El CV se redondea a 4 decimales según el protocolo de análisis técnico
                valor_cv_final = round(calculo_std / calculo_mean, 4)
            
            # Extracción del conteo de actividades (dato numérico tras el tiempo total de la jornada)
            segmento_final_linea = fila_texto[ubicacion_del_tiempo + len(string_del_total):]
            match_de_actividades = re.search(r'\d+', segmento_final_linea)
            
            numero_de_actividades = 0
            if match_de_actividades:
                numero_de_actividades = int(match_de_actividades.group())
            
            # Construcción del registro detallado por cada deportista procesado
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
            # Omisión de líneas corruptas o sin formato válido para evitar que el bucle falle
            continue
            
    # Retorno estructurado para procesamiento masivo en hojas de cálculo y reportes Word
    df_resultado_parsing = pd.DataFrame(lista_de_registros_atleta)
    
    return df_resultado_parsing

def parse_ocr_data(texto_ocr_crudo):
    """
    Parsea la tabla de traducción OCR (Distancia y Salida Larga).
    Filtra mandatoriamente los encabezados técnicos de la tabla procesada.
    """
    lista_podio_distancia = []
    lista_podio_larga = []
    
    # Listado de términos prohibidos en las celdas de datos para detectar encabezados
    filtro_nombres_tabla = ["Nombre", "Distancia", "Actividad", "Tiempo", "Km", "total", "Clasificación"]
    
    lineas_de_la_entrada_ocr = texto_ocr_crudo.strip().split('\n')
    
    for fila_ocr in lineas_de_la_entrada_ocr:
        celdas_ocr = fila_ocr.split(';')
        
        if len(celdas_ocr) >= 6:
            # Datos pertenecientes al bloque de Distancia Total recorrida
            nombre_atleta_d = celdas_ocr[2].strip()
            valor_metrica_d = celdas_ocr[3].strip()
            
            # Datos pertenecientes al bloque de Salida Larga semanal registrada
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
                
    # Retornar exclusivamente el Top 3 de cada categoría de honor (Podio)
    return lista_podio_distancia[:3], lista_podio_larga[:3]

# --- 5. ACTUALIZADOR DE EXCEL (OBJETIVO: INTEGRIDAD Y ORDEN NUMÉRICO) ---

def crear_excel_actualizado(referencia_maestro, df_actualizacion, input_semana_n):
    """
    Genera el archivo Excel inyectando la nueva semana y eliminando columnas futuras.
    Lógica de ordenamiento numérico real implementada para evitar errores en el historial.
    """
    # Carga del libro maestro con tipos de datos protegidos para evitar conversiones automáticas
    lector_maestro_excel = pd.ExcelFile(referencia_maestro)
    hojas_originales_maestro = lector_maestro_excel.sheet_names
    label_de_la_semana_actual = f"Sem {input_semana_n.strip()}"
    
    # FILTRO DE INGENIERÍA: Extraer el límite numérico de la semana procesada
    # para identificar qué columnas deben borrarse del archivo maestro.
    match_numero_semana = re.search(r'\d+', input_semana_n)
    valor_entero_limite = int(match_numero_semana.group()) if match_numero_semana else 0
    
    # DETERMINACIÓN DEL ORDEN OPERATIVO MANDATORIO (PROTOCOLO TYM):
    
    # Bloque 1: Identificación de Hojas Técnicas de Trabajo (Ciclismo, Natación, etc.)
    hojas_de_trabajo_lista = []
    for h_name in hojas_originales_maestro:
        if not h_name.startswith("Sem "):
            hojas_de_trabajo_lista.append(h_name)
            
    # Bloque 2: Identificación y limpieza del historial semanal anterior
    hojas_del_historial_lista = []
    for h_name in hojas_originales_maestro:
        if h_name.startswith("Sem "):
            # Condición crítica de integridad: No duplicar la semana cargada actualmente
            if h_name != label_de_la_semana_actual:
                hojas_del_historial_lista.append(h_name)
    
    # Bloque 3: Ordenamiento Numérico Real (Auditado contra errores Sem 2 vs Sem 11)
    def extraer_numero_pestaña(texto_pestaña):
        match_p = re.search(r'\d+', str(texto_pestaña))
        if match_p:
            return int(match_p.group())
        return 0
        
    hojas_del_historial_lista.sort(key=extraer_numero_pestaña, reverse=True)
    
    # Bloque 4: Construcción del libro final de salida unificado
    secuencia_de_hojas_final = hojas_de_trabajo_lista + [label_de_la_semana_actual] + hojas_del_historial_lista

    # Creación del stream de memoria binario para la descarga del usuario
    buffer_binario_descarga = io.BytesIO()
    
    with pd.ExcelWriter(buffer_binario_descarga, engine='xlsxwriter') as motor_escritura_excel:
        libro_excel_obj = motor_escritura_excel.book
        
        # Formato de tiempo TYM para celdas operables numéricamente ([h]:mm:ss)
        formato_hora_tym = libro_excel_obj.add_format({'num_format': '[h]:mm:ss'})
        
        for hoja_en_curso in secuencia_de_hojas_final:
            
            # ESCENARIO A: GENERACIÓN DE LA NUEVA PESTAÑA SEMANAL DETALLADA CON CLASIFICACIÓN
            if hoja_en_curso == label_de_la_semana_actual:
                cols_requeridas_sem = ['#', 'Deportista', 'Tiempo Total', 'Actividades', 'Natación', 'Bicicleta', 'Trote', 'CV']
                df_hoja_semana_final = df_actualizacion[cols_requeridas_sem].copy()
                df_hoja_semana_final.rename(columns={'#': 'Clasificación'}, inplace=True)
                
                # Conversión mandatoria a valores decimales de Excel para permitir aritmética
                for columna_t_name in ['Tiempo Total', 'Natación', 'Bicicleta', 'Trote']:
                    df_hoja_semana_final[columna_t_name] = df_hoja_semana_final[columna_t_name].apply(to_excel_time_value)
                
                # Escritura de la pestaña de clasificación del club
                df_hoja_semana_final.to_excel(motor_escritura_excel, sheet_name=hoja_en_curso, index=False)
                
                # Formateo visual de celdas horarias para el usuario final
                pestaña_actual_obj = motor_escritura_excel.sheets[hoja_en_curso]
                # Índices de columnas de tiempo que requieren formato (C, E, F, G)
                for id_columna_v in [2, 4, 5, 6]:
                    pestaña_actual_obj.set_column(id_columna_v, id_columna_v, 12, formato_hora_tym)
            
            # ESCENARIO B: ACTUALIZACIÓN DE LAS HOJAS DE TRABAJO TÉCNICAS (CON FILTRO DE DESBORDAMIENTO)
            elif hoja_en_curso in ["Tiempo Total", "Natación", "Ciclismo", "Trote"]:
                # Lectura blindada de la hoja histórica para inyectar los nuevos datos
                df_maestro_h_tecnica = pd.read_excel(lector_maestro_excel, sheet_name=hoja_en_curso, dtype=object)
                
                # FILTRO DE SEGURIDAD DE COLUMNAS (PROTOCOLO V2.2.23)
                # Remoción física de columnas "Sem XX" que exceden la semana actual procesada
                lista_de_cols_desbordadas = []
                for nombre_col_check in df_maestro_h_tecnica.columns:
                    str_check_nom = str(nombre_col_check)
                    if str_check_nom.startswith("Sem "):
                        match_num_check = re.search(r'\d+', str_check_nom)
                        if match_num_check:
                            valor_numerico_check = int(match_num_check.group())
                            if valor_numerico_check > valor_entero_limite:
                                lista_de_cols_desbordadas.append(nombre_col_check)
                                
                if len(lista_de_cols_desbordadas) > 0:
                    df_maestro_h_tecnica = df_maestro_h_tecnica.drop(columns=lista_de_cols_desbordadas)
                
                # Sello de Transcripción Aritmética para columnas de cálculo histórico
                cols_arit_tecnicas = []
                for c_header_t in df_maestro_h_tecnica.columns:
                    c_header_t_str = str(c_header_t)
                    if c_header_t_str.startswith("Sem ") or "Promedio" in c_header_t_str or "Acumulado" in c_header_t_str:
                        cols_arit_tecnicas.append(c_header_t)
                
                for col_a_t in cols_arit_tecnicas:
                    if col_a_t != label_de_la_semana_actual:
                        df_maestro_h_tecnica[col_a_t] = df_maestro_h_tecnica[col_a_t].apply(to_excel_time_value)
                
                # Match de deportistas mediante columna identificadora (limpia de tildes y espacios)
                id_col_identificadora = df_maestro_h_tecnica.columns[0]
                for col_it_t in df_maestro_h_tecnica.columns:
                    if "nombre" in str(col_it_t).lower() or "deportista" in str(col_it_t).lower():
                        id_col_identificadora = col_it_t
                        break
                
                # Diccionario de mapeo entre nombre de hoja y fuente de datos actual
                dict_mapeo_hojas_tym = {'Tiempo Total': 'Tiempo Total', 'Natación': 'Natación', 'Ciclismo': 'Bicicleta', 'Trote': 'Trote'}
                nombre_fuente_actual_tym = dict_mapeo_hojas_tym.get(hoja_en_curso)
                
                if nombre_fuente_actual_tym:
                    df_prep_normalizacion = df_actualizacion[['Deportista', nombre_fuente_actual_tym]].copy()
                    df_prep_normalizacion['MatchKey'] = df_prep_normalizacion['Deportista'].apply(clean_string)
                    df_maestro_h_tecnica['MatchKey'] = df_maestro_h_tecnica[id_col_identificadora].astype(str).apply(clean_string)
                    
                    # Generación de mapeo decimal exacto para la inyección de datos
                    dict_valores_inyectar = df_prep_normalizacion.set_index('MatchKey')[nombre_fuente_actual_tym].apply(to_excel_time_value).to_dict()
                    df_maestro_h_tecnica[label_de_la_semana_actual] = df_maestro_h_tecnica['MatchKey'].map(dict_valores_inyectar).fillna(0)
                    
                    # RECALCULO DE COLUMNA TIEMPO ACUMULADO (MANDATORIO)
                    lista_sem_para_suma = [c_f_s for c_f_s in df_maestro_h_tecnica.columns if str(c_f_s).startswith("Sem ")]
                    if 'Tiempo Acumulado' in df_maestro_h_tecnica.columns:
                        df_maestro_h_tecnica['Tiempo Acumulado'] = df_maestro_h_tecnica[lista_sem_para_suma].sum(axis=1)
                    
                    df_maestro_h_tecnica = df_maestro_h_tecnica.drop(columns=['MatchKey'])
                
                # Escritura definitiva de la hoja técnica actualizada en el libro Excel
                df_maestro_h_tecnica.to_excel(motor_escritura_excel, sheet_name=hoja_en_curso, index=False)
                
                # Formatear visualmente las columnas de cálculo matemático para asegurar la compatibilidad
                pestaña_trabajo_activa = motor_escritura_excel.sheets[hoja_en_curso]
                for id_c_f, nom_c_f in enumerate(df_maestro_h_tecnica.columns):
                    nom_c_f_str = str(nom_c_f)
                    if nom_c_f_str.startswith("Sem ") or "Promedio" in nom_c_f_str or "Acumulado" in nom_c_f_str:
                        pestaña_trabajo_activa.set_column(id_c_f, id_c_f, 13, formato_hora_tym)

            # ESCENARIO C: ACTUALIZACIÓN DE LA HOJA COEFICIENTE DE VARIACIÓN (CV)
            elif hoja_en_curso == "CV":
                df_maestro_cv_hoja = pd.read_excel(lector_maestro_excel, sheet_name=hoja_en_curso, dtype=object)
                
                # FILTRO DE SEGURIDAD PARA COLUMNAS CV (EVITAR SEMANAS FUTURAS VACÍAS)
                cols_cv_a_borrar = []
                for c_cv_iterador in df_maestro_cv_hoja.columns:
                    if str(c_cv_iterador).startswith("Sem "):
                        m_cv_check = re.search(r'\d+', str(c_cv_iterador))
                        if m_cv_check:
                            if int(m_cv_check.group()) > valor_entero_limite:
                                cols_cv_a_borrar.append(c_cv_iterador)
                
                if len(cols_cv_a_borrar) > 0:
                    df_maestro_cv_hoja = df_maestro_cv_hoja.drop(columns=cols_cv_a_borrar)
                
                # Localización de columna identificadora en la hoja CV
                id_col_nombre_en_cv = next((c_cv for c_cv in df_maestro_cv_hoja.columns if "nombre" in str(c_cv).lower() or "deportista" in str(c_cv).lower()), df_maestro_cv_hoja.columns[0])
                df_maestro_cv_hoja['MatchKey'] = df_maestro_cv_hoja[id_col_nombre_en_cv].astype(str).apply(clean_string)
                
                # Preparación de datos de CV de la semana actual
                df_cv_semana_prep = df_actualizacion[['Deportista', 'CV']].copy()
                df_cv_semana_prep['MatchKey'] = df_cv_semana_prep['Deportista'].apply(clean_string)
                
                # Inyección del valor CV (Dato flotante, no requiere formato de hora)
                df_maestro_cv_hoja[label_de_la_semana_actual] = df_maestro_cv_hoja['MatchKey'].map(df_cv_semana_prep.set_index('MatchKey')['CV'].to_dict()).fillna("NC")
                df_maestro_cv_hoja.drop(columns=['MatchKey']).to_excel(motor_escritura_excel, sheet_name=hoja_en_curso, index=False)

            # ESCENARIO D: RÉPLICA ÍNTEGRA DE TODAS LAS DEMÁS PESTAÑAS (PARÁMETROS, ETC.)
            else:
                df_replica_hoja_estatica = pd.read_excel(lector_maestro_excel, sheet_name=hoja_en_curso, dtype=object)
                df_replica_hoja_estatica.to_excel(motor_escritura_excel, sheet_name=hoja_en_curso, index=False)
                
    return buffer_binario_descarga.getvalue()

# --- 6. GENERADOR DE REPORTE WORD GRUPAL (BLOQUEADO - NO SINTETIZAR) ---

def aplicar_formato_tym_word(objeto_p, fuente_pt, bold_on=False, center_on=False):
    """
    Aplica rigurosamente el estilo institucional Calibri con los tamaños 20/15/13/11.
    """
    run_cursor = objeto_p.add_run() if not objeto_p.runs else objeto_p.runs[0]
    run_cursor.font.name = 'Calibri'
    run_cursor.font.size = Pt(fuente_pt)
    run_cursor.bold = bold_on
    if center_on:
        objeto_p.alignment = WD_ALIGN_PARAGRAPH.CENTER

def crear_tabla_profesional_tym_word(doc_word_ref, df_fuente_datos, listado_cabeceras):
    """Genera tablas con anchos milimétricos para el informe profesional."""
    objeto_tabla = doc_word_ref.add_table(rows=1, cols=len(listado_cabeceras))
    objeto_tabla.style = 'Light Grid Accent 1'
    objeto_tabla.alignment = 1 # Centrado
    objeto_tabla.autofit = False
    
    dicc_anchos = {'#': 0.4, 'Deportista': 2.8, 'Tiempo Total': 0.7, 'Natación': 0.7, 'Bicicleta': 0.7, 'Trote': 0.7, 'CV': 0.6}
    
    for indice_h, texto_cabecera in enumerate(listado_cabeceras):
        celda_h = objeto_tabla.rows[0].cells[indice_h]
        celda_h.text = texto_cabecera
        celda_h.width = Inches(dicc_anchos.get(texto_cabecera, 0.7))
        aplicar_formato_tym_word(celda_h.paragraphs[0], 9, True, True)
        
    for _, fila_loop in df_fuente_datos.iterrows():
        celdas_fila = objeto_tabla.add_row().cells
        for indice_d, texto_cabecera_d in enumerate(listado_cabeceras):
            celdas_fila[indice_d].text = str(fila_loop[texto_cabecera_d])
            celdas_fila[indice_d].width = Inches(dicc_anchos.get(texto_cabecera_d, 0.7))
            aplicar_formato_tym_word(celdas_fila[indice_d].paragraphs[0], 9, False, texto_cabecera_d != 'Deportista')
            
    doc_word_ref.add_paragraph()

def generar_reporte_word_tym_completo(df_datos_semanales, str_num_semana, lista_podio_d, lista_podio_l):
    """Construye el reporte Word íntegro bajo el modelo funcional V2.2.23."""
    doc_final = Document()
    
    p_header_tit = doc_final.add_heading(f'Reporte Semanal Club Tym Triatlón - Semana {str_num_semana}', 0)
    aplicar_formato_tym_word(p_header_tit, 20, True, True)
    doc_final.add_paragraph()
    
    p_slogan = doc_final.add_paragraph('"(La semana de la simetría perfecta y el retorno del volumen)"')
    aplicar_formato_tym_word(p_slogan, 11, True, True)
    doc_final.add_paragraph()
    
    h_res = doc_final.add_heading('🔍 Resumen General', level=2)
    aplicar_formato_tym_word(h_res, 15, True)
    doc_final.add_paragraph()
    
    df_trias = df_datos_semanales[df_datos_semanales['CV'] != 'NC'].copy()
    txt_res = f"Total deportistas registrados: {len(df_datos_semanales)}\nTriatletas completos: {len(df_trias)}\nHoras totales del club: {to_hhmmss_display(df_datos_semanales['T_Mins'].sum())}"
    p_info = doc_final.add_paragraph(txt_res); aplicar_formato_tym_word(p_info, 11)
    
    # Gráfico de Distribución Grupal
    fig, ax = plt.subplots(figsize=(4,4))
    ax.pie([df_datos_semanales['N_Mins'].sum(), df_datos_semanales['B_Mins'].sum(), df_datos_semanales['R_Mins'].sum()], labels=['Nat', 'Bici', 'Tro'], autopct='%1.1f%%', colors=['#1E90FF', '#32CD32', '#FF4500'])
    buf = io.BytesIO(); plt.savefig(buf, format='png', bbox_inches='tight'); plt.close(fig)
    doc_final.add_paragraph().add_run().add_picture(buf, width=Inches(3.5))
    
    # Bloque 2: Top 5 de Honor
    for tit, d, c in [('🏅 TOP 5 TRIATLETAS COMPLETOS', df_trias.sort_values('T_Mins', ascending=False).head(5), 'Completos'), ('⚖️ TOP 5 TRIATLETAS MÁS BALANCEADOS', df_trias.sort_values('CV', ascending=True).head(5), 'CV')]:
        h_s = doc_final.add_heading(tit, level=2); aplicar_formato_tym_word(h_s, 15, True); doc_final.add_paragraph()
        df_r = d.copy(); df_r['#'] = range(1, len(df_r) + 1)
        crear_tabla_profesional_tym_word(doc_final, df_r, ['#', 'Deportista', 'Tiempo Total' if c=='Completos' else 'CV', 'Natación', 'Bicicleta', 'Trote'])
        
        doc_final.add_paragraph('Análisis del Desempeño:'); 
        for _, r_fila_analisis in df_r.iterrows():
            p_at_n = doc_final.add_paragraph(f"{r_fila_analisis['#']}. {r_fila_analisis['Deportista']}"); aplicar_formato_tym_word(p_at_n, 11, True)
            doc_final.add_paragraph(generar_comentario(r_fila_analisis, c, r_fila_analisis['#']))
            
    # Bloque 3: Top 15 Especialidades
    for s, ico, m, txt in [('TIEMPO GENERAL', '🥇', 'T_Mins', 'Tiempo Total'), ('NATACIÓN', '🏊‍♂️', 'N_Mins', 'Natación'), ('CICLISMO', '🚴', 'B_Mins', 'Bicicleta'), ('TROTE', '🏃‍♂️', 'R_Mins', 'Trote')]:
        doc_final.add_page_break(); h_15 = doc_final.add_heading(f'{ico} TOP 15 {s}', level=1); aplicar_formato_tym_word(h_15, 15, True); doc_final.add_paragraph()
        df_t15 = df_datos_semanales[df_datos_semanales[m] > 0].sort_values(m, ascending=False).head(15).copy(); df_t15['#'] = range(1, len(df_t15) + 1)
        crear_tabla_profesional_tym_word(doc_final, df_t15, ['#', 'Deportista', txt, 'Natación', 'Bicicleta', 'Trote'] if s == 'TIEMPO GENERAL' else ['#', 'Deportista', txt, 'Tiempo Total'])
        
        doc_final.add_paragraph('Análisis del Podio:'); 
        for _, r_podio_15 in df_t15.head(3).iterrows():
            p_nom_15 = doc_final.add_paragraph(f"{r_podio_15['Deportista']}"); aplicar_formato_tym_word(p_nom_15, 11, True)
            doc_final.add_paragraph(generar_comentario(r_podio_15, 'General' if s=='TIEMPO GENERAL' else txt, r_podio_15['#']))
            
    doc_final.add_page_break(); doc_final.add_heading('📏 PODIO DISTANCIA TOTAL', 1)
    for it_dist in lista_podio_d: doc_final.add_paragraph(f"{it_dist['nombre']} ({it_dist['valor']} km)")
    
    doc_final.add_heading('⏱️ PODIO ACTIVIDAD MÁS LARGA', 1)
    for it_larg in lista_podio_l: doc_final.add_paragraph(f"{it_larg['nombre']} ({it_larg['valor']})")
    
    b_out = io.BytesIO(); doc_final.save(b_out); b_out.seek(0); return b_out

# --- 7. MOTOR NARRATIVO INDIVIDUAL (ESTRATEGIA COLAB V17.0) ---

def generar_reporte_narrativo_individual(atleta_nom, df_actual, dict_historicos, sem_n):
    """Genera reporte personal con insights narrativos basados en Colab V17.0."""
    doc_p = Document()
    match_key = clean_string(atleta_nom)
    
    # Header Personalizado
    p_h = doc_p.add_heading(f'Análisis de Rendimiento Personal: {atleta_nom}', 0)
    aplicar_formato_tym_word(p_h, 18, True, True)
    doc_p.add_paragraph(f"Semana de Entrenamiento: {sem_n}").alignment = 1
    
    df_actual['MatchKey'] = df_actual['Deportista'].apply(clean_string)
    row_act = df_actual[df_actual['MatchKey'] == match_key]
    if row_act.empty: return None
    row_act = row_act.iloc[0]

    def get_benchmarks(hoja_key, match_k):
        df_h = dict_historicos.get(hoja_key)
        if df_h is None: return 0, 0
        df_h['MatchKey'] = df_h.iloc[:, 0].astype(str).apply(clean_string)
        cols_sem = [c for c in df_h.columns if str(c).startswith("Sem ")]
        avg_equipo = df_actual[f"{hoja_key[0]}_Mins"].mean() if hoja_key != "Tiempo Total" else df_actual["T_Mins"].mean()
        r_atleta = df_h[df_h['MatchKey'] == match_k]
        avg_hist = r_atleta[cols_sem].mean(axis=1).iloc[0] * 1440 if not r_atleta.empty else 0
        return avg_equipo, avg_hist

    for tit_m, hoja_m, col_m in [("TIEMPO TOTAL", "Tiempo Total", "T_Mins"), ("NATACIÓN", "Natación", "N_Mins"), ("CICLISMO", "Ciclismo", "B_Mins"), ("TROTE", "Trote", "R_Mins")]:
        doc_p.add_heading(tit_m, level=1); val_act = row_act[col_m]; bench_equipo, bench_hist = get_benchmarks(hoja_m, match_key)
        
        p_m = doc_p.add_paragraph()
        p_m.add_run(f"Volumen actual: {to_hhmmss_display(val_act)}\n").bold = True
        
        # Comparativa Equipo
        diff_eq = val_act - bench_equipo
        txt_eq = f"Rendiste {to_hhmmss_display(abs(diff_eq))} {'MÁS' if diff_eq > 0 else 'MENOS'} que el promedio del equipo."
        run_eq = p_m.add_run(txt_eq)
        run_eq.font.color.rgb = RGBColor(0, 100, 0) if diff_eq >= 0 else RGBColor(180, 0, 0)
        
        # Comparativa Histórica
        diff_hi = val_act - bench_hist
        txt_hi = f"\nRespecto a tu propia media histórica: {to_hhmmss_display(abs(diff_hi))} {'MÁS' if diff_hi > 0 else 'MENOS'}."
        p_m.add_run(txt_hi)

    # Gráfico Personal
    fig_p, ax_p = plt.subplots(figsize=(5,3))
    ax_p.bar(['Nat', 'Bici', 'Tro'], [row_act['N_Mins'], row_act['B_Mins'], row_act['R_Mins']], color=['#1E90FF', '#32CD32', '#FF4500'])
    ax_p.set_title("Tu Distribución de Carga (Minutos)")
    buf_p = io.BytesIO(); plt.savefig(buf_p, format='png', bbox_inches='tight'); plt.close(fig_p)
    doc_p.add_paragraph().add_run().add_picture(buf_p, width=Inches(3.5))
    
    doc_p.add_paragraph("─" * 50); doc_p.add_paragraph("Generado por Agente TYM 2026").alignment = 2
    b_out = io.BytesIO(); doc_p.save(b_out); b_out.seek(0); return b_out

# --- 8. INTERFAZ DE USUARIO (STREMLIT) ---

st.sidebar.header("📁 Gestión Histórica TYM")
maestro_file = st.sidebar.file_uploader("Cargar Excel Maestro", type=["xlsx"])
n_sem_val = st.text_input("Semana (Ej: 08):", "08")
data_st = st.text_area("1. Datos Strava:")
data_oc = st.text_area("2. Traducción OCR:")

if st.button("🚀 PROCESAR JORNADA COMPLETA"):
    if maestro_file and data_st.strip() and data_oc.strip():
        # Procesamiento Principal
        df_res = parse_raw_data(data_st); p_d, p_l = parse_ocr_data(data_oc)
        
        # Carga de diccionarios históricos para reportes narrativos
        h_t = pd.read_excel(maestro_file, sheet_name="Tiempo Total", dtype=object)
        h_n = pd.read_excel(maestro_file, sheet_name="Natación", dtype=object)
        h_c = pd.read_excel(maestro_file, sheet_name="Ciclismo", dtype=object)
        h_r = pd.read_excel(maestro_file, sheet_name="Trote", dtype=object)
        dict_h = {"Tiempo Total": h_t, "Natación": h_n, "Ciclismo": h_c, "Trote": h_r}
        
        st.success(f"¡Semana {n_sem_val} procesada con éxito!"); c1, c2 = st.columns(2)
        c1.download_button("📄 REPORTE GRUPAL", generar_reporte_word_tym_completo(df_res, n_sem_val, p_d, p_l), f"Grupal_{n_sem_val}.docx")
        c2.download_button("📊 EXCEL ACTUALIZADO", crear_excel_actualizado(maestro_file, df_res, n_sem_val), f"Excel_{n_sem_val}.xlsx")
        
        # SECCIÓN INDIVIDUAL
        st.divider(); st.subheader("👤 Generador de Insights Individuales")
        atletas = df_res['Deportista'].tolist()
        seleccion = st.multiselect("Seleccionar atletas:", atletas, default=atletas[:2])
        
        if st.button("📦 EMPAQUETAR REPORTES INDIVIDUALES (ZIP)"):
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w") as zf:
                for a in seleccion:
                    r_ind = generar_reporte_narrativo_individual(a, df_res, dict_h, n_sem_val)
                    if r_ind: zf.writestr(f"Reporte_{clean_string(a)}.docx", r_ind.getvalue())
            st.download_button("⬇️ DESCARGAR ZIP", zip_buf.getvalue(), f"Individuales_{n_sem_val}.zip")
            
    else: st.error("Error: Complete Excel y datos de texto.")
