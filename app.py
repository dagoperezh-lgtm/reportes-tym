# TYM PLATAFORMA - VERSION: 2.2.4-PCVCI-FIXED
# OBJETIVO: CORREGIR ERROR SINTAXIS Y ASEGURAR INTEGRIDAD ARITMETICA
# LINEAS DE CODIGO: 450
# ESTADO: MODELO FUNCIONAL EXPANDIDO (PROHIBIDO SINTETIZAR)

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
    page_title="Plataforma TYM 2026 - V2.2.4", 
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
    """
    nombre_atleta = fila_datos['Deportista']
    
    if categoria_nombre == 'Completos' or categoria_nombre == 'General':
        if posicion_ranking == 1:
            res = f"Dominio absoluto de {nombre_atleta}. Se consolida en la cima del club con un volumen total envidiable. Su capacidad para sostener sesiones de alta intensidad demuestra una preparación de élite y una disciplina inquebrantable."
            return res
        if posicion_ranking == 2:
            res = f"Una semana brillante para {nombre_atleta}. Se queda con la plata manteniendo una presión constante sobre el líder. Su solidez fue el motor que lo mantuvo en la parte más alta de la tabla."
            return res
        if posicion_ranking == 3:
            res = f"{nombre_atleta} cierra el podio con un rendimiento muy equilibrado. Demuestra que para estar en el grupo de avanzada no se puede regalar nada, sumando minutos de calidad en las tres disciplinas."
            return res
        return f"Desempeño consistente de {nombre_atleta} en la zona alta de la tabla clasificatoria."
    
    if categoria_nombre == 'CV':
        cv_actual = fila_datos.get('CV', 0)
        res = f"¡El reloj suizo del club! {nombre_atleta} logra una simetría casi perfecta ({cv_actual}), demostrando una planificación milimétrica de sus cargas y un control total del entrenamiento en todas las áreas."
        return res
    
    # Comentarios por disciplina individual
    txt_tiempo = fila_datos.get(categoria_nombre, "00:00:00")
    
    if categoria_nombre == 'Natación':
        res = f"Fuerza pura en el agua. {nombre_atleta} registra un tiempo de {txt_tiempo}, liderando el podio de la disciplina con una técnica depurada y gran volumen."
        return res
        
    if categoria_nombre == 'Bicicleta':
        res = f"Potencia pura sobre ruedas. {nombre_atleta} devoró la ruta con un tiempo de {txt_tiempo}. Demuestra ser el gran motor del equipo en la carretera con promedios que intimidan."
        return res
        
    if categoria_nombre == 'Trote':
        res = f"Resistencia inalcanzable. {nombre_atleta} domina el asfalto con un tiempo de {txt_tiempo} y una fase de carrera soberbia, cerrando una semana de alta calidad."
        return res
    
    return "Desempeño técnico destacado durante la jornada de entrenamiento semanal."

# --- 4. PARSERS DE ENTRADA (BLINDADO - NO SINTETIZAR) ---

def parse_raw_data(bloque_texto):
    """
    Procesa el bloque de texto copiado de Strava (Tiempo Total).
    No utiliza síntesis; cada paso de limpieza y extracción es explícito.
    """
    lista_final = []
    contador_rank = 1
    
    # Limpieza de caracteres invisibles
    bloque_texto = bloque_texto.replace('\xa0', ' ')
    lineas_input = bloque_texto.strip().split('\n')
    
    for linea in lineas_input:
        if not linea:
            continue
        if 'Deportista' in linea:
            continue
            
        try:
            # Buscar patrones de tiempo de Strava
            regex_tiempo = r'(\d+h\s*\d*min|\d+h|\d+min|--:--)'
            tiempos_encontrados = re.findall(regex_tiempo, linea)
            
            if not tiempos_encontrados:
                continue
            
            # El nombre está antes del tiempo total
            string_tiempo_total = tiempos_encontrados[0]
            indice_tiempo = linea.find(string_tiempo_total)
            
            segmento_nombre = linea[:indice_tiempo].strip()
            # Eliminar número de ranking si viene pegado
            nombre_final = re.sub(r'^\d+\s*', '', segmento_nombre).strip()
            
            # Convertir todos los tiempos detectados
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
                
            # Cálculo del Coeficiente de Variación (CV)
            valores_cv = [m_nat, m_bici, m_trote]
            if 0 in valores_cv:
                cv_calc = "NC"
            else:
                std_dev = np.std(valores_cv)
                avg_val = np.mean(valores_cv)
                cv_calc = round(std_dev / avg_val, 4)
            
            # Extraer número de actividades
            bloque_act = linea[indice_tiempo + len(string_tiempo_total):]
            find_act = re.search(r'\d+', bloque_act)
            
            total_actividades = 0
            if find_act:
                total_actividades = int(find_act.group())
            
            # Construcción de la fila de datos
            atleta_data = {
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
            
            lista_final.append(atleta_data)
            contador_rank = contador_rank + 1
            
        except Exception:
            continue
            
    return pd.DataFrame(lista_final)

def parse_ocr_data(texto_ocr):
    """
    Parsea la tabla de traducción OCR (Distancia y Salida Larga).
    Filtra mandatoriamente los encabezados técnicos.
    """
    podio_distancia = []
    podio_larga = []
    
    filtro_tecnico = ["Nombre", "Distancia", "Actividad", "Tiempo", "Km", "total"]
    lineas_ocr = texto_ocr.strip().split('\n')
    
    for l in lineas_ocr:
        partes_linea = l.split(';')
        if len(partes_linea) >= 6:
            # Datos de Distancia Total
            n_dist = partes_linea[2].strip()
            v_dist = partes_linea[3].strip()
            
            # Datos de Salida Larga
            n_larg = partes_linea[4].strip()
            v_larg = partes_linea[5].strip()
            
            # Verificación para Distancia
            is_header_d = False
            for palabra in filtro_tecnico:
                if palabra.lower() in n_dist.lower():
                    is_header_d = True
            
            if is_header_d == False and n_dist != "":
                podio_distancia.append({'nombre': n_dist, 'valor': v_dist})
                
            # Verificación para Salida Larga
            is_header_l = False
            for palabra in filtro_tecnico:
                if palabra.lower() in n_larg.lower():
                    is_header_l = True
                    
            if is_header_l == False and n_larg != "":
                podio_larga.append({'nombre': n_larg, 'valor': v_larg})
                
    return podio_distancia[:3], podio_larga[:3]

# --- 5. ACTUALIZADOR DE EXCEL (OBJETIVO: TRANSCRIPCION Y ARITMÉTICA MANDATORIA) ---

def crear_excel_actualizado(file_maestro, df_actualizado, txt_semana):
    """
    Genera el archivo Excel manteniendo funcionalidad aritmética de tiempos.
    Asegura la transcripción literal de históricos y el orden de hojas de trabajo.
    """
    engine_excel = pd.ExcelFile(file_maestro)
    nombres_hojas = engine_excel.sheet_names
    label_semana = f"Sem {txt_semana.strip()}"
    
    # 🛡️ ORDEN OPERATIVO MANDATORIO:
    # 1. Hojas de Trabajo Técnicas
    list_trabajo = []
    for h in nombres_hojas:
        if not h.startswith("Sem "):
            list_trabajo.append(h)
            
    # 2. Hojas de semanas anteriores
    list_historia = []
    for h in nombres_hojas:
        if h.startswith("Sem "):
            if h != label_semana:
                list_historia.append(h)
    
    list_historia.sort(reverse=True)
    
    # Lista final unificada
    hojas_finales = list_trabajo + [label_semana] + list_historia

    output_buffer = io.BytesIO()
    
    with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer_obj:
        libro_excel = writer_obj.book
        # Formato de tiempo para celdas operables
        formato_tym = libro_excel.add_format({'num_format': '[h]:mm:ss'})
        
        for name_h in hojas_finales:
            
            # CASO A: CREAR HOJA SEMANAL NUEVA
            if name_h == label_semana:
                df_export_sem = df_actualizado[['#', 'Deportista', 'Tiempo Total', 'Actividades', 'Natación', 'Bicicleta', 'Trote', 'CV']].copy()
                df_export_sem.rename(columns={'#': 'Clasificación'}, inplace=True)
                
                # Convertir a decimales para aritmética
                for col_name in ['Tiempo Total', 'Natación', 'Bicicleta', 'Trote']:
                    df_export_sem[col_name] = df_export_sem[col_name].apply(to_excel_time_value)
                
                df_export_sem.to_excel(writer_obj, sheet_name=name_h, index=False)
                
                # Formato visual
                sheet_sem = writer_obj.sheets[name_h]
                for idx_col in [2, 4, 5, 6]:
                    sheet_sem.set_column(idx_col, idx_col, 12, formato_tym)
            
            # CASO B: HOJAS DE TRABAJO TÉCNICAS
            elif name_h in ["Tiempo Total", "Natación", "Ciclismo", "Trote"]:
                df_original = pd.read_excel(engine_excel, sheet_name=name_h, dtype=object)
                
                # Limpieza de columnas históricas residuales
                if 'Sem 51' in df_original.columns:
                    df_original = df_original.drop(columns=['Sem 51'])
                if 'Sem 52' in df_original.columns:
                    df_original = df_original.drop(columns=['Sem 52'])
                
                # 🛡️ TRANSCRIPCIÓN ARITMÉTICA: Asegurar formato decimal en históricos
                cols_calculo = []
                for c in df_original.columns:
                    c_s = str(c)
                    if c_s.startswith("Sem ") or "Promedio" in c_s or "Acumulado" in c_s:
                        cols_calculo.append(c)
                
                for c_calc in cols_calculo:
                    if c_calc != label_semana:
                        df_original[c_calc] = df_original[c_calc].apply(to_excel_time_value)
                
                # Búsqueda de columna Deportista
                key_col = df_original.columns[0]
                for c in df_original.columns:
                    if "nombre" in str(c).lower() or "deportista" in str(c).lower():
                        key_col = c
                        break
                
                # Mapeo de fuente
                mapeo_src = {
                    'Tiempo Total': 'Tiempo Total', 
                    'Natación': 'Natación', 
                    'Ciclismo': 'Bicicleta', 
                    'Trote': 'Trote'
                }
                src_col = mapeo_src.get(name_h)
                
                if src_col:
                    df_update_prep = df_actualizado[['Deportista', src_col]].copy()
                    df_update_prep['MatchKey'] = df_update_prep['Deportista'].apply(clean_string)
                    
                    df_original['MatchKey'] = df_original[key_col].astype(str).apply(clean_string)
                    
                    dict_mapping = df_update_prep.set_index('MatchKey')[src_col].apply(to_excel_time_value).to_dict()
                    
                    # Inyección de nueva semana
                    df_original[label_semana] = df_original['MatchKey'].map(dict_mapping).fillna(0)
                    # Limpiar llave técnica
                    df_original = df_original.drop(columns=['MatchKey'])
                
                df_original.to_excel(writer_obj, sheet_name=name_h, index=False)
                
                sheet_work = writer_obj.sheets[name_h]
                for i_col, name_col in enumerate(df_original.columns):
                    str_n = str(name_col)
                    if str_n.startswith("Sem ") or "Promedio" in str_n or "Acumulado" in str_n:
                        sheet_work.set_column(i_col, i_col, 13, formato_tym)

            # CASO C: HOJA CV
            elif name_h == "CV":
                df_cv_org = pd.read_excel(engine_excel, sheet_name=name_h, dtype=object)
                
                # Identificar deportista
                id_col_cv = df_cv_org.columns[0]
                for c in df_cv_org.columns:
                    if "nombre" in str(c).lower() or "deportista" in str(c).lower():
                        id_col_cv = c
                        break
                
                df_cv_update = df_actualizado[['Deportista', 'CV']].copy()
                df_cv_update['MatchKey'] = df_cv_update['Deportista'].apply(clean_string)
                
                df_cv_org['MatchKey'] = df_cv_org[id_col_cv].astype(str).apply(clean_string)
                
                dict_cv = df_cv_update.set_index('MatchKey')['CV'].to_dict()
                
                df_cv_org[label_semana] = df_cv_org['MatchKey'].map(dict_cv).fillna("NC")
                df_cv_org = df_cv_org.drop(columns=['MatchKey'])
                df_cv_org.to_excel(writer_obj, sheet_name=name_h, index=False)

            # CASO D: RESTO DE HOJAS
            else:
                df_resto_hoja = pd.read_excel(engine_excel, sheet_name=name_h, dtype=object)
                df_resto_hoja.to_excel(writer_obj, sheet_name=name_h, index=False)
                
    return output_buffer.getvalue()

# --- 6. GENERADOR DE REPORTE WORD (BLOQUEADO / MODELO FUNCIONAL) ---

def aplicar_estilo_word(p_obj, t_fuente, is_bold=False, is_center=False):
    """
    Aplica el formato institucional Calibri 20/15/13/11.
    """
    run_obj = p_obj.runs[0] if p_obj.runs else p_obj.add_run()
    run_obj.font.name = 'Calibri'
    run_obj.font.size = Pt(t_fuente)
    run_obj.bold = is_bold
    if is_center:
        p_obj.alignment = WD_ALIGN_PARAGRAPH.CENTER

def crear_tabla_tym_format(doc_target, df_tab, list_headers):
    """
    Genera tablas con anchos milimétricos para el reporte profesional.
    """
    t_word = doc_target.add_table(rows=1, cols=len(list_headers))
    t_word.style = 'Light Grid Accent 1'
    t_word.alignment = 1 # Centrado
    t_word.autofit = False
    
    # Anchos Blindados
    map_w = {
        '#': 0.4, 'Deportista': 2.8, 'Tiempo Total': 0.7, 
        'Natación': 0.7, 'Bicicleta': 0.7, 'Trote': 0.7, 'CV': 0.6
    }
    
    # Encabezados
    for i, h_text in enumerate(list_headers):
        c_head = t_word.rows[0].cells[i]
        c_head.text = h_text
        
        w_val = map_w.get(h_text, 0.7)
        c_head.width = Inches(w_val)
        
        aplicar_estilo_word(c_head.paragraphs[0], 9, True, True)
        
    # Filas
    for _, r_data in df_tab.iterrows():
        c_row = t_word.add_row().cells
        for i, h_text in enumerate(list_headers):
            c_row[i].text = str(r_data[h_text])
            
            w_cell = map_w.get(h_text, 0.7)
            c_row[i].width = Inches(w_cell)
            
            # Deportista Izquierda, resto Centro
            is_c = True
            if h_text == 'Deportista':
                is_c = False
                
            aplicar_estilo_word(c_row[i].paragraphs[0], 9, False, is_c)
            
    doc_target.add_paragraph()

def generar_word_report_total(df_res_sem, val_n_sem, p_dist_l, p_larg_l):
    """
    Construye el reporte Word íntegro bajo el modelo funcional V2.2.4.
    """
    doc_final = Document()
    
    # Título (Calibri 20)
    tit_p = doc_final.add_heading(f'Reporte Semanal Club Tym Triatlón - Semana {val_n_sem}', 0)
    aplicar_estilo_word(tit_p, 20, True, True)
    doc_final.add_paragraph()
    
    cita = doc_final.add_paragraph('"(La semana de la simetría perfecta y el retorno del volumen)"')
    aplicar_estilo_word(cita, 11, True, True)
    doc_final.add_paragraph()

    # BLOQUE 1: Resumen
    h_res = doc_final.add_heading('🔍 Resumen General', level=2)
    aplicar_estilo_word(h_res, 15, True)
    doc_final.add_paragraph()
    
    df_comp_only = df_res_sem[df_res_sem['CV'] != 'NC'].copy()
    
    t_atletis = len(df_res_sem)
    t_completos = len(df_comp_only)
    t_minutos = df_res_sem['T_Mins'].sum()
    txt_horas = to_hhmmss_display(t_minutos)
    
    txt_res_bloque = f"Total deportistas: {t_atletis}\nTriatletas completos: {t_completos}\nHoras totales: {txt_horas}"
    p_info = doc_final.add_paragraph(txt_res_bloque)
    aplicar_estilo_word(p_info, 11)
    
    # Gráfico
    s_n = df_res_sem['N_Mins'].sum()
    s_b = df_res_sem['B_Mins'].sum()
    s_t = df_res_sem['R_Mins'].sum()
    
    f_res, a_res = plt.subplots(figsize=(4,4))
    a_res.pie([s_n, s_b, s_t], labels=['Natación', 'Ciclismo', 'Trote'], autopct='%1.1f%%', colors=['#1E90FF', '#32CD32', '#FF4500'])
    
    buf_img = io.BytesIO()
    plt.savefig(buf_img, format='png', bbox_inches='tight')
    plt.close(f_res)
    
    p_gr = doc_final.add_paragraph()
    p_gr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_gr.add_run().add_picture(buf_img, width=Inches(3.5))

    # BLOQUE 2: Top 5
    podios_5 = [
        ('🏅 TOP 5 TRIATLETAS COMPLETOS', df_comp_only.sort_values('T_Mins', ascending=False).head(5), 'Completos'),
        ('⚖️ TOP 5 TRIATLETAS MÁS BALANCEADOS', df_comp_only.sort_values('CV', ascending=True).head(5), 'CV')
    ]
    
    for t_pod, d_pod, c_key in podios_5:
        h_s = doc_final.add_heading(t_pod, level=2)
        aplicar_estilo_word(h_s, 15, True)
        doc_final.add_paragraph()
        
        d_ren = d_pod.copy()
        d_ren['#'] = range(1, len(d_ren) + 1)
        cols_5 = ['#', 'Deportista', 'Tiempo Total' if c_key=='Completos' else 'CV', 'Natación', 'Bicicleta', 'Trote']
        crear_tabla_tym_format(doc_final, d_ren, cols_5)
        
        h_an = doc_final.add_paragraph('Análisis del Desempeño:'); aplicar_estilo_word(h_an, 13, True)
        for _, r_fila in d_ren.iterrows():
            p_n = doc_final.add_paragraph(f"{r_fila['#']}. {r_fila['Deportista']}"); aplicar_estilo_word(p_n, 11, True)
            doc_final.add_paragraph(generar_comentario(r_fila, c_key, r_fila['#']))

    # BLOQUE 3: Top 15
    config_15 = [
        ('TIEMPO GENERAL', '🥇', 'T_Mins', 'Tiempo Total'),
        ('NATACIÓN', '🏊‍♂️', 'N_Mins', 'Natación'),
        ('CICLISMO', '🚴', 'B_Mins', 'Bicicleta'),
        ('TROTE', '🏃‍♂️', 'R_Mins', 'Trote')
    ]
    
    for n_sec, ico, c_met, c_txt in config_15:
        doc_final.add_page_break()
        h_15 = doc_final.add_heading(f'{ico} TOP 15 {n_sec}', level=1)
        aplicar_estilo_word(h_15, 15, True)
        doc_final.add_paragraph()
        
        d_15 = df_res_sem[df_res_sem[c_met] > 0].sort_values(c_met, ascending=False).head(15).copy()
        d_15['#'] = range(1, len(d_15) + 1)
        
        c_15 = ['#', 'Deportista', c_txt, 'Natación', 'Bicicleta', 'Trote'] if n_sec == 'TIEMPO GENERAL' else ['#', 'Deportista', c_txt, 'Tiempo Total']
        crear_tabla_tym_format(doc_final, d_15, c_15)
        
        h_an_p = doc_final.add_paragraph('Análisis del Podio:'); aplicar_estilo_word(h_an_p, 13, True)
        for _, r_p in d_15.head(3).iterrows():
            emo = '🥇' if r_p['#'] == 1 else '🥈' if r_p['#'] == 2 else '🥉'
            p_at = doc_final.add_paragraph(f"{emo} {r_p['Deportista']}"); aplicar_estilo_word(p_at, 11, True)
            c_an = 'General' if n_sec == 'TIEMPO GENERAL' else c_txt
            doc_final.add_paragraph(generar_comentario(r_p, c_an, r_p['#']))

    # BLOQUE 4: OCR
    doc_final.add_page_break()
    h_d = doc_final.add_heading('📏 PODIO DISTANCIA TOTAL', level=1); aplicar_estilo_word(h_d, 15, True); doc_final.add_paragraph()
    for i, it in enumerate(p_dist_l): doc_final.add_paragraph(f"{i+1}. {it['nombre']} ({it['valor']} km)")
    doc_final.add_paragraph()
    h_l = doc_final.add_heading('⏱️ PODIO ACTIVIDAD MÁS LARGA', level=1); aplicar_estilo_word(h_l, 15, True); doc_final.add_paragraph()
    for i, it in enumerate(p_larg_l): doc_final.add_paragraph(f"{i+1}. {it['nombre']} ({it['valor']})")
    
    b_out = io.BytesIO(); doc_final.save(b_out); b_out.seek(0); return b_out

# --- 7. INTERFAZ ---

st.sidebar.header("📁 Gestión de Datos Históricos TYM")
maestro_upload = st.sidebar.file_uploader("Cargar Excel Maestro", type=["xlsx"])
n_sem_input = st.text_input("Número de Semana (Ej: 08):", "08")
txt_strava_raw = st.text_area("1. Datos Tiempo Total (Strava):")
txt_ocr_raw = st.text_area("2. Datos OCR (Captura):")

if st.button("🚀 PROCESAR JORNADA Y ACTUALIZAR EXCEL"):
    if txt_strava_raw.strip() and txt_ocr_raw.strip() and maestro_upload:
        d_res = parse_raw_data(txt_strava_raw); l_dist, l_larg = parse_ocr_data(txt_ocr_raw)
        st.success(f"¡Semana {n_sem_input} procesada!"); c_w, c_e = st.columns(2)
        c_w.download_button(label="📄 REPORTE WORD", data=generar_word_report_total(d_res, n_sem_input, l_dist, l_larg), file_name=f"Reporte_TYM_Sem_{n_sem_input}.docx")
        c_e.download_button(label="📊 EXCEL ACTUALIZADO", data=crear_excel_actualizado(maestro_upload, d_res, n_sem_input), file_name=f"00_Estadisticas_Actualizado_Sem_{n_sem_input}.xlsx")
    else:
        st.error("Error: Se requiere el Excel Maestro y completar los campos de datos.")
