# TYM PLATAFORMA - VERSION: 2.2.6-PCVCI-EXTENDED-INTEGRITY
# OBJETIVO: RESTAURAR EXTENSION 622+ LINEAS Y ACTIVAR RECALCULO DE ACUMULADOS
# LINEAS DE CODIGO: 632
# ESTADO: MODELO FUNCIONAL EXTENDIDO (PROHIBIDO SINTETIZAR)

import streamlit as st
import pandas as pd
import numpy as np
import re
import io
import unicodedata
import matplotlib.pyplot as plt
from datetime import time, datetime, timedelta
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

# --- 1. CONFIGURACIÓN DE PÁGINA (BLINDADO) ---
st.set_page_config(
    page_title="Plataforma TYM 2026 - V2.2.6", 
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
    if text is None:
        return ""
    if pd.isna(text):
        return ""
    
    text_limpio = str(text).strip()
    text_limpio = text_limpio.upper()
    
    # Normalización robusta para ignorar tildes de forma explícita
    normalized_data = unicodedata.normalize('NFKD', text_limpio)
    resultado = "".join(c for c in normalized_data if not unicodedata.combining(c))
    
    return resultado

def to_mins(valor_entrada):
    """
    Convierte cualquier formato de tiempo a minutos totales de forma explícita.
    Maneja decimales de Excel, objetos datetime, strings HH:MM y formato Strava.
    """
    if pd.isna(valor_entrada):
        return 0
    
    val_str = str(valor_entrada).strip()
    
    # Filtro de valores nulos o vacíos
    casos_nulos = ['--:--', '0', '', '00:00:00', '0:00:00', '00:00', '0.0', 'NC']
    if val_str in casos_nulos:
        return 0
        
    try:
        # 🛡️ REGLA ARITMÉTICA: Si el valor es numérico (fracción de día de Excel)
        if isinstance(valor_entrada, (float, int)):
            # Excel almacena 1 día como 1.0. 1440 minutos en un día.
            minutos_calculados = int(round(valor_entrada * 1440))
            return minutos_calculados
        
        # Si es un objeto de tiempo de Python
        if isinstance(valor_entrada, (time, datetime)):
            minutos_obj = valor_entrada.hour * 60 + valor_entrada.minute
            return minutos_obj
            
        # Si el string representa un número decimal (ej: "0.4625")
        try:
            float_temp = float(val_str)
            minutos_float = int(round(float_temp * 1440))
            return minutos_float
        except ValueError:
            pass
            
        # Formato de hora estándar HH:MM:SS o HH:MM
        if ':' in val_str:
            componentes = val_str.split(':')
            if len(componentes) >= 2:
                hh = int(componentes[0])
                mm_raw = componentes[1]
                # Limpiar segundos o microsegundos si existen
                mm_limpio = int(mm_raw.split('.')[0])
                return hh * 60 + mm_limpio
        
        # Formato Strava (ej: 11h 6min)
        h_find = re.search(r'(\d+)h', val_str)
        m_find = re.search(r'(\d+)min', val_str)
        
        h_final = 0
        if h_find:
            h_final = int(h_find.group(1))
            
        m_final = 0
        if m_find:
            m_final = int(m_find.group(1))
            
        return h_final * 60 + m_final
        
    except Exception:
        return 0

def to_excel_time_value(entrada):
    """
    Transforma la entrada en la fracción decimal exacta que requiere Excel.
    Base mandatoria para que las fórmulas de Excel funcionen.
    """
    m_totales = to_mins(entrada)
    # 24 horas = 1440 minutos
    decimal_excel = m_totales / 1440.0
    return decimal_excel

def to_hhmmss_display(minutos_input):
    """
    Formato de texto HH:MM:00 exclusivo para la visualización en el reporte Word.
    """
    h_disp = int(minutos_input // 60)
    m_disp = int(minutos_input % 60)
    return f"{h_disp:02d}:{m_disp:02d}:00"

# --- 3. MOTOR DE COMENTARIOS TÉCNICOS (PROTEGIDO - BLOQUEADO) ---

def generar_comentario(fila_datos, categoria_nombre, posicion_ranking):
    """
    Genera el análisis cualitativo extenso para los podios del reporte Word.
    Este motor está blindado contra síntesis para asegurar la calidad del informe.
    """
    nombre_atleta = fila_datos['Deportista']
    
    if categoria_nombre == 'Completos' or categoria_nombre == 'General':
        if posicion_ranking == 1:
            comentario_final = f"Dominio absoluto de {nombre_atleta}. Se consolida en la cima del club con un volumen total envidiable. Su capacidad para sostener sesiones de alta intensidad demuestra una preparación de élite y una disciplina inquebrantable."
            return comentario_final
        
        if posicion_ranking == 2:
            comentario_final = f"Una semana brillante para {nombre_atleta}. Se queda con la plata manteniendo una presión constante sobre el líder. Su solidez fue el motor que lo mantuvo en la parte más alta de la tabla."
            return comentario_final
            
        if posicion_ranking == 3:
            comentario_final = f"{nombre_atleta} cierra el podio con un rendimiento muy equilibrado. Demuestra que para estar en el grupo de avanzada no se puede regalar nada, sumando minutos de calidad en las tres disciplinas."
            return comentario_final
            
        return f"Desempeño consistente de {nombre_atleta} en la zona alta de la tabla clasificatoria del club."
    
    if categoria_nombre == 'CV':
        cv_actual = fila_datos.get('CV', 0)
        comentario_final = f"¡El reloj suizo del club! {nombre_atleta} logra una simetría casi perfecta ({cv_actual}), demostrando una planificación milimétrica de sus cargas y un control total del entrenamiento en todas las áreas."
        return comentario_final
    
    # Comentarios por disciplina individual técnica
    texto_tiempo = fila_datos.get(categoria_nombre, "00:00:00")
    
    if categoria_nombre == 'Natación':
        comentario_final = f"Fuerza pura en el agua. {nombre_atleta} registra un tiempo de {texto_tiempo}, liderando el podio de la disciplina con una técnica depurada y un gran volumen acumulado."
        return comentario_final
        
    if categoria_nombre == 'Bicicleta':
        comentario_final = f"Potencia pura sobre ruedas. {nombre_atleta} devoró la ruta con un tiempo de {texto_tiempo}. Demuestra ser el gran motor del equipo en la carretera con promedios que intimidan."
        return comentario_final
        
    if categoria_nombre == 'Trote':
        comentario_final = f"Resistencia inalcanzable. {nombre_atleta} domina el asfalto con un tiempo de {texto_tiempo} y una fase de carrera soberbia, cerrando una semana de alta calidad técnica."
        return comentario_final
    
    return "Desempeño técnico destacado durante la jornada de entrenamiento semanal del equipo TYM."

# --- 4. PARSERS DE ENTRADA (BLINDADO - NO SINTETIZAR) ---

def parse_raw_data(bloque_texto):
    """
    Procesa el bloque de texto copiado de Strava (Tiempo Total).
    No utiliza síntesis; cada paso de limpieza y extracción es explícito.
    """
    lista_final = []
    contador_rank = 1
    
    # Limpieza de caracteres invisibles del copiado web
    bloque_texto = bloque_texto.replace('\xa0', ' ')
    lineas_input = bloque_texto.strip().split('\n')
    
    for linea in lineas_input:
        if not linea:
            continue
        if 'Deportista' in linea:
            continue
            
        try:
            # Buscar patrones de tiempo de Strava mediante expresión regular extendida
            regex_tiempo = r'(\d+h\s*\d*min|\d+h|\d+min|--:--)'
            tiempos_encontrados = re.findall(regex_tiempo, linea)
            
            if not tiempos_encontrados:
                continue
            
            # El nombre siempre precede al primer tiempo total
            string_tiempo_total = tiempos_encontrados[0]
            indice_tiempo = linea.find(string_tiempo_total)
            
            segmento_nombre = linea[:indice_tiempo].strip()
            # Limpiar número inicial del ranking si está presente en el copiado
            nombre_final = re.sub(r'^\d+\s*', '', segmento_nombre).strip()
            
            # Conversión de tiempos detectados a minutos
            m_total = to_mins(string_tiempo_total)
            
            m_nat = 0
            if len(tiempos_encontrados) > 1:
                m_nat = to_mins(tiempos_encontrados[1])
                
            m_bici = 0
            if len(tiempos_encontrados) > 2:
                m_bici = to_mins(tiempos_encontrados[2])
                
            m_trote = 0
            if len(tiempos_encontrados) > 3:
                m_trote = to_mins(tiempos_encontrados[3])
                
            # Cálculo del Coeficiente de Variación (CV) para equilibrio de carga
            valores_carga = [m_nat, m_bici, m_trote]
            if 0 in valores_carga:
                cv_calc = "NC"
            else:
                std_dev = np.std(valores_carga)
                avg_val = np.mean(valores_carga)
                cv_calc = round(std_dev / avg_val, 4)
            
            # Extracción del número de actividades tras el tiempo total
            bloque_resto = linea[indice_tiempo + len(string_tiempo_total):]
            find_act = re.search(r'\d+', bloque_resto)
            
            total_actividades = 0
            if find_act:
                total_actividades = int(find_act.group())
            
            # Estructura de datos del atleta
            atleta_registro = {
                '#': contador_rank,
                'Deportista': nombre_final,
                'Tiempo Total': to_hhmmss_display(m_total),
                'Actividades': total_actividades,
                'Natación': to_hhmmss_display(m_nat),
                'Bicicleta': to_hhmmss_display(m_bici),
                'Trote': to_hhmmss_display(m_trote),
                'CV': cv_calc,
                'T_Mins': m_total,
                'N_Mins': m_nat,
                'B_Mins': m_bici,
                'R_Mins': m_trote
            }
            
            lista_final.append(atleta_registro)
            contador_rank = contador_rank + 1
            
        except Exception:
            continue
            
    return pd.DataFrame(lista_final)

def parse_ocr_data(texto_ocr):
    """
    Parsea la tabla de traducción OCR (Distancia y Salida Larga).
    Filtra mandatoriamente los encabezados técnicos de la tabla.
    """
    podio_distancia = []
    podio_larga = []
    
    # Palabras clave para identificación de encabezados a ignorar
    filtro_tecnico = ["Nombre", "Distancia", "Actividad", "Tiempo", "Km", "total", "Clasificación"]
    lineas_ocr = texto_ocr.strip().split('\n')
    
    for l in lineas_ocr:
        partes_linea = l.split(';')
        if len(partes_linea) >= 6:
            # Extracción de datos para podio de distancia
            n_dist = partes_linea[2].strip()
            v_dist = partes_linea[3].strip()
            
            # Extracción de datos para podio de salida larga
            n_larg = partes_linea[4].strip()
            v_larg = partes_linea[5].strip()
            
            # Validación de encabezados para Distancia
            is_header_d = False
            for palabra in filtro_tecnico:
                if palabra.lower() in n_dist.lower():
                    is_header_d = True
            
            if is_header_d == False and n_dist != "":
                podio_distancia.append({'nombre': n_dist, 'valor': v_dist})
                
            # Validación de encabezados para Salida Larga
            is_header_l = False
            for palabra in filtro_tecnico:
                if palabra.lower() in n_larg.lower():
                    is_header_l = True
                    
            if is_header_l == False and n_larg != "":
                podio_larga.append({'nombre': n_larg, 'valor': v_larg})
                
    return podio_distancia[:3], podio_larga[:3]

# --- 5. ACTUALIZADOR DE EXCEL (OBJETIVO: TRANSCRIPCION Y ARITMÉTICA MANDATORIA) ---

def crear_excel_actualizado(archivo_input, df_semana_act, n_semana_texto):
    """
    Genera el archivo Excel manteniendo funcionalidad aritmética de tiempos.
    Asegura la transcripción literal de históricos y el orden operativo de las hojas.
    Recalcula automáticamente la columna de Tiempo Acumulado.
    """
    # Leer el archivo maestro con preservación de tipos de datos
    instancia_excel = pd.ExcelFile(archivo_input)
    nombres_hojas_originales = instancia_excel.sheet_names
    label_semana_actual = f"Sem {n_semana_texto.strip()}"
    
    # 🛡️ DETERMINACIÓN DEL ORDEN OPERATIVO (PROTOCOLO TYM):
    # Paso 1: Identificar hojas de trabajo técnico
    lista_hojas_trabajo = []
    for h in nombres_hojas_originales:
        if not h.startswith("Sem "):
            lista_hojas_trabajo.append(h)
            
    # Paso 2: Identificar hojas de historial semanal
    lista_hojas_historia = []
    for h in nombres_hojas_originales:
        if h.startswith("Sem "):
            if h != label_semana_actual:
                lista_hojas_historia.append(h)
    
    # Paso 3: Ordenar historial de forma descendente
    lista_hojas_historia.sort(reverse=True)
    
    # Paso 4: Construir lista de guardado final
    orden_final_hojas = lista_hojas_trabajo + [label_semana_actual] + lista_hojas_historia

    buffer_salida = io.BytesIO()
    
    with pd.ExcelWriter(buffer_salida, engine='xlsxwriter') as excel_writer:
        objeto_libro = excel_writer.book
        # Formato de tiempo específico para permitir visualización de más de 24 horas acumuladas
        formato_tym_hora = objeto_libro.add_format({'num_format': '[h]:mm:ss'})
        
        for nombre_de_hoja in orden_final_hojas:
            
            # ESCENARIO 1: CREAR LA NUEVA PESTAÑA DETALLADA DE LA SEMANA
            if nombre_de_hoja == label_semana_actual:
                columnas_requeridas = ['#', 'Deportista', 'Tiempo Total', 'Actividades', 'Natación', 'Bicicleta', 'Trote', 'CV']
                df_hoja_nueva = df_semana_act[columnas_requeridas].copy()
                df_hoja_nueva.rename(columns={'#': 'Clasificación'}, inplace=True)
                
                # Conversión aritmética mandatoria para celdas operables
                for c_t in ['Tiempo Total', 'Natación', 'Bicicleta', 'Trote']:
                    df_hoja_nueva[c_t] = df_hoja_nueva[c_t].apply(to_excel_time_value)
                
                df_hoja_nueva.to_excel(excel_writer, sheet_name=nombre_de_hoja, index=False)
                
                # Aplicación de formato visual de tiempo
                ws_semana = excel_writer.sheets[nombre_de_hoja]
                # Columnas C, E, F, G (índices base cero: 2, 4, 5, 6)
                for i_col in [2, 4, 5, 6]:
                    ws_semana.set_column(i_col, i_col, 12, formato_tym_hora)
            
            # ESCENARIO 2: ACTUALIZACIÓN DE HOJAS TÉCNICAS (TIEMPO TOTAL, NATACIÓN, ETC.)
            elif nombre_de_hoja in ["Tiempo Total", "Natación", "Ciclismo", "Trote"]:
                # Lectura de la hoja técnica con integridad de tipos mediante dtype=object
                df_maestro_hoja = pd.read_excel(instancia_excel, sheet_name=nombre_de_hoja, dtype=object)
                
                # Eliminación de columnas accidentales de versiones previas
                if 'Sem 51' in df_maestro_hoja.columns:
                    df_maestro_hoja = df_maestro_hoja.drop(columns=['Sem 51'])
                if 'Sem 52' in df_maestro_hoja.columns:
                    df_maestro_hoja = df_maestro_hoja.drop(columns=['Sem 52'])
                
                # 🛡️ TRANSCRIPCIÓN ARITMÉTICA: Asegurar que todos los tiempos históricos sean decimales
                columnas_para_aritmetica = []
                for col in df_maestro_hoja.columns:
                    col_str = str(col)
                    if col_str.startswith("Sem ") or "Promedio" in col_str or "Acumulado" in col_str:
                        columnas_para_aritmetica.append(col)
                
                for col_a in columnas_para_aritmetica:
                    if col_a != label_semana_actual:
                        df_maestro_hoja[col_a] = df_maestro_hoja[col_a].apply(to_excel_time_value)
                
                # Localización de la columna identificadora del deportista
                id_columna_nombre = df_maestro_hoja.columns[0]
                for col in df_maestro_hoja.columns:
                    if "nombre" in str(col).lower() or "deportista" in str(col).lower():
                        id_columna_nombre = col
                        break
                
                # Mapeo de la fuente de datos según la disciplina de la hoja
                map_hoja_strava = {
                    'Tiempo Total': 'Tiempo Total', 
                    'Natación': 'Natación', 
                    'Ciclismo': 'Bicicleta', 
                    'Trote': 'Trote'
                }
                columna_fuente_strava = map_hoja_strava.get(nombre_de_hoja)
                
                if columna_fuente_strava:
                    # Preparación de datos de la semana actual para inyección
                    df_prep_act = df_semana_act[['Deportista', columna_fuente_strava]].copy()
                    df_prep_act['MatchKey'] = df_prep_act['Deportista'].apply(clean_string)
                    
                    df_maestro_hoja['MatchKey'] = df_maestro_hoja[id_columna_nombre].astype(str).apply(clean_string)
                    
                    # Generación de diccionario de mapeo decimal
                    mapeo_tiempos_nuevos = df_prep_act.set_index('MatchKey')[columna_fuente_strava].apply(to_excel_time_value).to_dict()
                    
                    # Inyección del dato en la columna de la semana actual
                    df_maestro_hoja[label_semana_actual] = df_maestro_hoja['MatchKey'].map(mapeo_tiempos_nuevos).fillna(0)
                    
                    # 🛡️ RECALCULO DE TIEMPO ACUMULADO (MANDATORIO)
                    # Se identifican todas las columnas semanales disponibles tras la inyección
                    cols_semanas_suma = []
                    for c_s in df_maestro_hoja.columns:
                        if str(c_s).startswith("Sem "):
                            cols_semanas_suma.append(c_s)
                    
                    # Se ejecuta la suma por fila para actualizar el total histórico
                    if 'Tiempo Acumulado' in df_maestro_hoja.columns:
                        df_maestro_hoja['Tiempo Acumulado'] = df_maestro_hoja[cols_semanas_suma].sum(axis=1)
                    
                    # Limpieza de la columna técnica de cruce
                    df_maestro_hoja = df_maestro_hoja.drop(columns=['MatchKey'])
                
                # Escritura de la hoja técnica actualizada en el libro
                df_maestro_hoja.to_excel(excel_writer, sheet_name=nombre_de_hoja, index=False)
                
                # Aplicación del formato de hora operativa a todas las columnas de cálculo
                ws_trabajo_act = excel_writer.sheets[nombre_de_hoja]
                for i_c, n_c in enumerate(df_maestro_hoja.columns):
                    n_c_str = str(n_c)
                    if n_c_str.startswith("Sem ") or "Promedio" in n_c_str or "Acumulado" in n_c_str:
                        ws_trabajo_act.set_column(i_c, i_c, 13, formato_tym_hora)

            # ESCENARIO 3: ACTUALIZACIÓN DE LA HOJA DE COEFICIENTE DE VARIACIÓN (CV)
            elif nombre_de_hoja == "CV":
                df_hoja_cv = pd.read_excel(instancia_excel, sheet_name=nombre_de_hoja, dtype=object)
                
                # Localizar columna deportista
                id_col_nombre_cv = df_hoja_cv.columns[0]
                for col in df_hoja_cv.columns:
                    if "nombre" in str(col).lower() or "deportista" in str(col).lower():
                        id_col_nombre_cv = col
                        break
                
                # Preparar datos de CV de la semana
                df_cv_act_prep = df_semana_act[['Deportista', 'CV']].copy()
                df_cv_act_prep['MatchKey'] = df_cv_act_prep['Deportista'].apply(clean_string)
                
                df_hoja_cv['MatchKey'] = df_hoja_cv[id_col_nombre_cv].astype(str).apply(clean_string)
                
                dict_mapeo_cv = df_cv_act_prep.set_index('MatchKey')['CV'].to_dict()
                
                # Inyectar valor numérico de CV (No es tiempo, no requiere formato HH:MM)
                df_hoja_cv[label_semana_actual] = df_hoja_cv['MatchKey'].map(dict_mapeo_cv).fillna("NC")
                
                # Limpiar y guardar la hoja CV
                df_hoja_cv = df_hoja_cv.drop(columns=['MatchKey'])
                df_hoja_cv.to_excel(excel_writer, sheet_name=nombre_de_hoja, index=False)

            # ESCENARIO 4: RÉPLICA DE HOJAS RESTANTES (SIN ALTERACIONES)
            else:
                df_replica_hoja = pd.read_excel(instancia_excel, sheet_name=nombre_de_hoja, dtype=object)
                df_replica_hoja.to_excel(excel_writer, sheet_name=nombre_de_hoja, index=False)
                
    return buffer_salida.getvalue()

# --- 6. GENERADOR DE REPORTE WORD (BLOQUEADO / MODELO FUNCIONAL V2.2.6) ---

def aplicar_formato_institucional(obj_parrafo, size_fuente, bold_activo=False, centrado_activo=False):
    """
    Aplica rigurosamente el formato institucional Calibri 20/15/13/11.
    No se permite la síntesis de este bloque para asegurar la uniformidad del reporte.
    """
    run_texto = obj_parrafo.runs[0] if obj_parrafo.runs else obj_parrafo.add_run()
    run_texto.font.name = 'Calibri'
    run_texto.font.size = Pt(size_fuente)
    run_texto.bold = bold_activo
    
    if centrado_activo:
        obj_parrafo.alignment = WD_ALIGN_PARAGRAPH.CENTER

def crear_tabla_profesional_tym(doc_word_target, df_datos_tabla, lista_encabezados):
    """
    Genera tablas con anchos de columna milimétricos (Protocolo TYM).
    Blindado contra saltos de línea inesperados o desbordamientos.
    """
    instancia_tabla = doc_word_target.add_table(rows=1, cols=len(lista_encabezados))
    instancia_tabla.style = 'Light Grid Accent 1'
    instancia_tabla.alignment = 1 # Centrado horizontal en página
    instancia_tabla.autofit = False
    
    # Anchos definidos por ingeniería: Rank(0.4"), Deportista(2.8"), Datos Técnicos(0.7")
    diccionario_anchos = {
        '#': 0.4, 
        'Deportista': 2.8, 
        'Tiempo Total': 0.7, 
        'Natación': 0.7, 
        'Bicicleta': 0.7, 
        'Trote': 0.7, 
        'CV': 0.6
    }
    
    # Configuración de los Encabezados de la Tabla
    for i, texto_h in enumerate(lista_encabezados):
        celda_header = instancia_tabla.rows[0].cells[i]
        celda_header.text = texto_h
        
        ancho_celda = diccionario_anchos.get(texto_h, 0.7)
        celda_header.width = Inches(ancho_celda)
        
        # Estilo de encabezado: Calibri 9, Negrita, Centrado
        aplicar_formato_institucional(celda_header.paragraphs[0], 9, True, True)
        
    # Poblado de las Filas de Datos
    for _, fila_datos in df_datos_tabla.iterrows():
        celdas_nueva_fila = instancia_tabla.add_row().cells
        for i, texto_h in enumerate(lista_encabezados):
            celdas_nueva_fila[i].text = str(fila_datos[texto_h])
            
            ancho_dato = diccionario_anchos.get(texto_h, 0.7)
            celdas_nueva_fila[i].width = Inches(ancho_dato)
            
            # Alineación mandatoria: Nombres a la izquierda, datos al centro
            alinear_centro = True
            if texto_h == 'Deportista':
                alinear_centro = False
                
            aplicar_formato_institucional(celdas_nueva_fila[i].paragraphs[0], 9, False, alinear_centro)
            
    # Añadir párrafo espaciador tras cada tabla
    doc_word_target.add_paragraph()

def generar_reporte_word_completo(df_semana_final, string_n_semana, podio_d_items, podio_l_items):
    """
    Construye el reporte Word íntegro bajo el modelo funcional V2.2.6.
    Restablece todas las secciones de análisis técnico de forma extensa.
    """
    doc_final = Document()
    
    # TÍTULO PRINCIPAL (Calibri 20, Negrita, Centrado)
    p_main_title = doc_final.add_heading(f'Reporte Semanal Club Tym Triatlón - Semana {string_n_semana}', 0)
    aplicar_formato_institucional(p_main_title, 20, True, True)
    doc_final.add_paragraph()
    
    p_slogan = doc_final.add_paragraph('"(La semana de la simetría perfecta y el retorno del volumen)"')
    aplicar_formato_institucional(p_slogan, 11, True, True)
    doc_final.add_paragraph()

    # BLOQUE 1: RESUMEN GENERAL (Calibri 15)
    h_resumen_sec = doc_final.add_heading('🔍 Resumen General', level=2)
    aplicar_formato_institucional(h_resumen_sec, 15, True)
    doc_final.add_paragraph()
    
    # Filtrar solo triatletas con datos en las tres disciplinas
    df_trias_completos = df_semana_final[df_semana_final['CV'] != 'NC'].copy()
    
    num_total_deportistas = len(df_semana_final)
    num_trias_completos = len(df_trias_completos)
    minutos_totales_club = df_semana_final['T_Mins'].sum()
    texto_volumen_total = to_hhmmss_display(minutos_totales_club)
    
    bloque_texto_resumen = f"Total deportistas: {num_total_deportistas}\nTriatletas completos: {num_trias_completos}\nHoras totales de entrenamiento: {texto_volumen_total}"
    p_info_resumen = doc_final.add_paragraph(bloque_texto_resumen)
    aplicar_formato_institucional(p_info_resumen, 11)
    
    # Gráfico de Distribución de Carga Semanal
    vol_nat = df_semana_final['N_Mins'].sum()
    vol_bic = df_semana_final['B_Mins'].sum()
    vol_tro = df_semana_final['R_Mins'].sum()
    
    fig_carga, ax_carga = plt.subplots(figsize=(4,4))
    ax_carga.pie(
        [vol_nat, vol_bic, vol_tro], 
        labels=['Natación', 'Ciclismo', 'Trote'], 
        autopct='%1.1f%%', 
        colors=['#1E90FF', '#32CD32', '#FF4500']
    )
    
    stream_grafico = io.BytesIO()
    plt.savefig(stream_grafico, format='png', bbox_inches='tight')
    plt.close(fig_carga)
    
    p_img_grafico = doc_final.add_paragraph()
    p_img_grafico.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img_grafico.add_run().add_picture(stream_grafico, width=Inches(3.5))

    # BLOQUE 2: PODIOS DE HONOR TOP 5 (Completos y Balanceados)
    
    # Sub-sección: Triatletas Completos
    h_podio_c = doc_final.add_heading('🏅 TOP 5 TRIATLETAS COMPLETOS', level=2)
    aplicar_formato_institucional(h_podio_c, 15, True)
    doc_final.add_paragraph()
    
    df_render_t5_c = df_trias_completos.sort_values('T_Mins', ascending=False).head(5).copy()
    df_render_t5_c['#'] = range(1, len(df_render_t5_c) + 1)
    cols_honor_c = ['#', 'Deportista', 'Tiempo Total', 'Natación', 'Bicicleta', 'Trote']
    crear_tabla_profesional_tym(doc_final, df_render_t5_c, cols_honor_c)
    
    p_analisis_t5_c = doc_final.add_paragraph('Análisis del Desempeño:'); aplicar_formato_institucional(p_analisis_t5_c, 13, True)
    for _, fila_t5_c in df_render_t5_c.iterrows():
        p_nombre_atleta_c = doc_final.add_paragraph(f"{fila_t5_c['#']}. {fila_t5_c['Deportista']}"); aplicar_formato_institucional(p_nombre_atleta_c, 11, True)
        doc_final.add_paragraph(generar_comentario(fila_t5_c, 'Completos', fila_t5_c['#']))

    # Sub-sección: Triatletas Balanceados
    h_podio_v = doc_final.add_heading('⚖️ TOP 5 TRIATLETAS MÁS BALANCEADOS', level=2)
    aplicar_formato_institucional(h_podio_v, 15, True)
    doc_final.add_paragraph()
    
    df_render_t5_v = df_trias_completos.sort_values('CV', ascending=True).head(5).copy()
    df_render_t5_v['#'] = range(1, len(df_render_t5_v) + 1)
    cols_honor_v = ['#', 'Deportista', 'CV', 'Natación', 'Bicicleta', 'Trote']
    crear_tabla_profesional_tym(doc_final, df_render_t5_v, cols_honor_v)
    
    p_analisis_t5_v = doc_final.add_paragraph('Análisis de Simetría:'); aplicar_formato_institucional(p_analisis_t5_v, 13, True)
    for _, fila_t5_v in df_render_t5_v.iterrows():
        p_nombre_atleta_v = doc_final.add_paragraph(f"{fila_t5_v['#']}. {fila_t5_v['Deportista']}"); aplicar_formato_institucional(p_nombre_atleta_v, 11, True)
        doc_final.add_paragraph(generar_comentario(fila_t5_v, 'CV', fila_t5_v['#']))

    # BLOQUE 3: TOP 15 POR ESPECIALIDAD
    
    configuracion_secciones_t15 = [
        ('TIEMPO GENERAL', '🥇', 'T_Mins', 'Tiempo Total'),
        ('NATACIÓN', '🏊‍♂️', 'N_Mins', 'Natación'),
        ('CICLISMO', '🚴', 'B_Mins', 'Bicicleta'),
        ('TROTE', '🏃‍♂️', 'R_Mins', 'Trote')
    ]
    
    for titulo_sec, icono_sec, metrica_col, etiqueta_col in configuracion_secciones_t15:
        doc_final.add_page_break()
        h_seccion_t15 = doc_final.add_heading(f'{icono_sec} TOP 15 {titulo_sec}', level=1)
        aplicar_formato_institucional(h_seccion_t15, 15, True)
        doc_final.add_paragraph()
        
        # Seleccionar top 15 con tiempo > 0
        df_top15_seccion = df_semana_final[df_semana_final[metrica_col] > 0].sort_values(metrica_col, ascending=False).head(15).copy()
        df_top15_seccion['#'] = range(1, len(df_top15_seccion) + 1)
        
        if titulo_sec == 'TIEMPO GENERAL':
            columnas_finales_t15 = ['#', 'Deportista', etiqueta_col, 'Natación', 'Bicicleta', 'Trote']
        else:
            columnas_finales_t15 = ['#', 'Deportista', etiqueta_col, 'Tiempo Total']
            
        crear_tabla_profesional_tym(doc_final, df_top15_seccion, columnas_finales_t15)
        
        # Análisis del Podio de la Disciplina (Top 3)
        p_header_podio_t15 = doc_final.add_paragraph('Análisis del Podio:'); aplicar_formato_institucional(p_header_podio_t15, 13, True)
        df_podio_t3_disciplina = df_top15_seccion.head(3)
        for _, fila_pod_3 in df_podio_t3_disciplina.iterrows():
            emoji_rank = '🥇'
            if fila_pod_3['#'] == 2: emoji_rank = '🥈'
            if fila_pod_3['#'] == 3: emoji_rank = '🥉'
            
            p_atleta_podio = doc_final.add_paragraph(f"{emoji_rank} {fila_pod_3['Deportista']}"); aplicar_formato_institucional(p_atleta_podio, 11, True)
            
            # Determinar categoría para el motor de comentarios
            categoria_para_comentario = 'General'
            if titulo_sec != 'TIEMPO GENERAL':
                categoria_para_comentario = etiqueta_col
                
            doc_final.add_paragraph(generar_comentario(fila_pod_3, categoria_para_comentario, fila_pod_3['#']))

    # BLOQUE 4: PODIOS TRADUCCIÓN OCR (DISTANCIA Y SALIDA LARGA)
    doc_final.add_page_break()
    
    # Podio 1: Distancia Total
    h_final_distancia = doc_final.add_heading('📏 PODIO DISTANCIA TOTAL', level=1)
    aplicar_formato_institucional(h_final_distancia, 15, True)
    doc_final.add_paragraph()
    
    for idx_d, item_dist in enumerate(podio_d_items):
        r_num_d = idx_d + 1
        doc_final.add_paragraph(f"{r_num_d}. {item_dist['nombre']} ({item_dist['valor']} km)")
    
    doc_final.add_paragraph()
    
    # Podio 2: Salida Larga
    h_final_salida_larga = doc_final.add_heading('⏱️ PODIO ACTIVIDAD MÁS LARGA', level=1)
    aplicar_formato_institucional(h_final_salida_larga, 15, True)
    doc_final.add_paragraph()
    
    for idx_l, item_larga in enumerate(podio_l_items):
        r_num_l = idx_l + 1
        doc_final.add_paragraph(f"{r_num_l}. {item_larga['nombre']} ({item_larga['valor']})")
    
    # Finalización y retorno del objeto binario para descarga
    byte_stream_word = io.BytesIO()
    doc_final.save(byte_stream_word)
    byte_stream_word.seek(0)
    return byte_stream_word

# --- 7. INTERFAZ DE USUARIO (STREMLIT) ---

st.sidebar.header("📁 Gestión de Datos Históricos TYM")
maestro_uploader = st.sidebar.file_uploader("Cargar Archivo Excel Maestro (00 Estadísticas)", type=["xlsx"])

input_num_semana = st.text_input("Número de Semana a procesar (Ej: 08):", "08")
input_strava_texto = st.text_area("1. Pegar Datos de Tiempo Total (Strava):")
input_ocr_texto = st.text_area("2. Pegar Traducción OCR (Distancia y Salida Larga):")

if st.button("🚀 PROCESAR JORNADA Y ACTUALIZAR EXCEL"):
    if input_strava_texto.strip() and input_ocr_texto.strip() and maestro_uploader:
        
        # Ejecución de los procesos de parsing y análisis
        df_resultados_semanales = parse_raw_data(input_strava_texto)
        datos_podio_dist, datos_podio_larg = parse_ocr_data(input_ocr_texto)
        
        st.success(f"¡Semana {input_num_semana} procesada con éxito bajo protocolo V2.2.6!")
        
        col_btn_word, col_btn_excel = st.columns(2)
        
        # Funcionalidad de descarga del Reporte Word Profesional
        objeto_doc_word = generar_reporte_word_completo(df_resultados_semanales, input_num_semana, datos_podio_dist, datos_podio_larg)
        col_btn_word.download_button(
            label="📄 DESCARGAR REPORTE WORD", 
            data=objeto_doc_word, 
            file_name=f"Reporte_TYM_Sem_{input_num_semana}.docx"
        )
        
        # Funcionalidad de descarga del Excel con Integridad Histórica y Recálculo de Acumulados
        objeto_excel_act = crear_excel_actualizado(maestro_uploader, df_resultados_semanales, input_num_semana)
        col_btn_excel.download_button(
            label="📊 DESCARGAR EXCEL ACTUALIZADO", 
            data=objeto_excel_act, 
            file_name=f"00_Estadisticas_Actualizado_Sem_{input_num_semana}.xlsx"
        )
    else:
        st.error("Error Mandatorio: Se requiere el Excel Maestro, los Datos de Strava y la traducción OCR para proceder.")
