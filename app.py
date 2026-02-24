import streamlit as st
import pandas as pd
import numpy as np
import re
import io
import unicodedata
import matplotlib.pyplot as plt
from datetime import time
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

# --- 1. CONFIGURACIÓN DE PÁGINA (BLINDADO) ---
st.set_page_config(page_title="Plataforma TYM 2026", page_icon="🏆", layout="wide")
st.title("🏆 Gestión de Reportes y Estadísticas - Club TYM")

# --- 2. UTILIDADES DE PROCESAMIENTO Y TIEMPO (BLINDADO - NO SINTETIZAR) ---
def clean_string(text):
    """Normaliza nombres para eliminar tildes, mayúsculas y espacios extra."""
    if not text: return ""
    text = str(text).strip().upper()
    text = "".join(c for c in unicodedata.normalize('NFKD', text) if not unicodedata.combining(c))
    return text

def to_mins(t_str):
    """Convierte cadenas de tiempo (h, min, :) a minutos totales de forma explícita."""
    if pd.isna(t_str) or str(t_str).strip() in ['--:--', '0', '', '00:00:00', '0:00:00']: 
        return 0
    try:
        if isinstance(t_str, time): 
            return t_str.hour * 60 + t_str.minute
        t_str = str(t_str).strip()
        if ':' in t_str:
            parts = t_str.split(':')
            if len(parts) >= 2:
                return int(parts[0]) * 60 + int(parts[1])
        
        h_match = re.search(r'(\d+)h', t_str)
        m_match = re.search(r'(\d+)min', t_str)
        h = int(h_match.group(1)) if h_match else 0
        m = int(m_match.group(1)) if m_match else 0
        return h * 60 + m
    except: 
        return 0

def to_excel_time_value(t_str):
    """
    Transforma un tiempo en formato decimal de Excel (fracción de 24h).
    Esta es la base para que el archivo mantenga funcionalidad aritmética.
    """
    total_minutos = to_mins(t_str)
    return total_minutos / 1440.0

def to_hhmmss_display(mins):
    """Formato de texto HH:MM:00 exclusivo para el reporte Word."""
    h = int(mins // 60)
    m = int(mins % 60)
    return f"{h:02d}:{m:02d}:00"

# --- 3. MOTOR DE COMENTARIOS TÉCNICOS (PROTEGIDO - BLOQUEADO) ---
def generar_comentario(row, cat, pos):
    """Genera el análisis cualitativo para los podios del reporte Word."""
    nombre = row['Deportista']
    
    if cat in ['Completos', 'General']:
        if pos == 1: 
            return f"Dominio absoluto de {nombre}. Se consolida en la cima del club con un volumen total envidiable. Su capacidad para sostener sesiones de alta intensidad demuestra una preparación de élite y una disciplina inquebrantable."
        if pos == 2: 
            return f"Una semana brillante para {nombre}. Se queda con la plata manteniendo una presión constante sobre el líder. Su solidez fue el motor que lo mantuvo en la parte más alta de la tabla."
        if pos == 3:
            return f"{nombre} cierra el podio con un rendimiento muy equilibrado. Demuestra que para estar en el grupo de avanzada no se puede regalar nada, sumando minutos de calidad."
        return f"Desempeño consistente de {nombre} en la zona alta de la tabla."
    
    if cat == 'CV':
        cv_val = row.get('CV', 0)
        return f"¡El reloj suizo del club! {nombre} logra una simetría casi perfecta ({cv_val}), demostrando una planificación milimétrica de sus cargas y un control total del entrenamiento."
    
    tiempo_texto = row.get(cat, "")
    if cat == 'Natación': 
        return f"Fuerza pura en el agua. {nombre} registra {tiempo_texto}, liderando el podio con una técnica depurada. Sus hombros de acero dominan el volumen acuático."
    if cat == 'Bicicleta': 
        return f"Potencia pura sobre ruedas. {nombre} devoró la ruta con {tiempo_texto}. Demuestra ser el motor del equipo en la carretera con promedios que intimidan."
    if cat == 'Trote': 
        return f"Resistencia inalcanzable. {nombre} domina el asfalto con {tiempo_texto} y una fase de carrera soberbia."
    
    return "Desempeño destacado durante la jornada de entrenamiento."

# --- 4. PARSERS DE ENTRADA (BLINDADO - NO SINTETIZAR) ---
def parse_raw_data(raw_text):
    """Extrae los datos de Strava pegados por el usuario."""
    data = []
    rank = 1
    raw_text = raw_text.replace('\xa0', ' ')
    lineas = raw_text.strip().split('\n')
    
    for line in lineas:
        if not line or 'Deportista' in line: 
            continue
        try:
            # Buscar patrones de tiempo Strava (h, min)
            tiempos_encontrados = re.findall(r'(\d+h\s*\d*min|\d+h|\d+min|--:--)', line)
            if not tiempos_encontrados: 
                continue
            
            # Extraer nombre (texto antes del primer tiempo detectado)
            pos_primer_tiempo = line.find(tiempos_encontrados[0])
            nombre_sucio = line[:pos_primer_tiempo].strip()
            nombre_limpio = re.sub(r'^\d+\s*', '', nombre_sucio).strip()
            
            # Minutos por disciplina
            min_total = to_mins(tiempos_encontrados[0])
            min_nat = to_mins(tiempos_encontrados[1] if len(tiempos_encontrados) > 1 else 0)
            min_bici = to_mins(tiempos_encontrados[2] if len(tiempos_encontrados) > 2 else 0)
            min_trote = to_mins(tiempos_encontrados[3] if len(tiempos_encontrados) > 3 else 0)
            
            # Coeficiente de Variación
            lista_vals = [min_nat, min_bici, min_trote]
            if 0 in lista_vals:
                cv = "NC"
            else:
                cv = round(np.std(lista_vals) / np.mean(lista_vals), 4)
            
            # N° Actividades
            match_act = re.search(r'\d+', line[pos_primer_tiempo + len(tiempos_encontrados[0]):])
            num_act = int(match_act.group()) if match_act else 0
            
            data.append({
                '#': rank,
                'Deportista': nombre_limpio,
                'Tiempo Total': to_hhmmss_display(min_total),
                'Actividades': num_act,
                'Natación': to_hhmmss_display(min_nat),
                'Bicicleta': to_hhmmss_display(min_bici),
                'Trote': to_hhmmss_display(min_trote),
                'CV': cv,
                'T_Mins': min_total,
                'N_Mins': min_nat,
                'B_Mins': min_bici,
                'R_Mins': min_trote
            })
            rank += 1
        except:
            pass
    return pd.DataFrame(data)

def parse_ocr_data(ocr_text):
    """Parsea la tabla OCR filtrando mandatoriamente los encabezados técnicos."""
    dist_podio, larg_podio = [], []
    palabras_filtro = ["Nombre", "Distancia", "Actividad", "Tiempo", "Km", "total"]
    lineas_ocr = ocr_text.strip().split('\n')
    
    for line in lineas_ocr:
        partes = line.split(';')
        if len(partes) >= 6:
            n_d = partes[2].strip()
            v_d = partes[3].strip()
            n_l = partes[4].strip()
            v_l = partes[5].strip()
            
            # Verificación de encabezados
            es_titulo_d = any(p.lower() in n_d.lower() for p in palabras_filtro)
            es_titulo_l = any(p.lower() in n_l.lower() for p in palabras_filtro)
            
            if not es_titulo_d and n_d:
                dist_podio.append({'nombre': n_d, 'valor': v_d})
            if not es_titulo_l and n_l:
                larg_podio.append({'nombre': n_l, 'valor': v_l})
                
    return dist_podio[:3], larg_podio[:3]

# --- 5. ACTUALIZADOR DE EXCEL (OBJETIVO: ARITMÉTICA Y POSICIÓN MANDATORIA) ---
def crear_excel_actualizado(archivo_maestro, df_semana, num_sem):
    """
    Genera el archivo Excel manteniendo funcionalidad aritmética de tiempos.
    Asegura el orden de las hojas de trabajo al inicio.
    """
    xls_engine = pd.ExcelFile(archivo_maestro)
    hojas_origen = xls_engine.sheet_names
    nombre_sem_nueva = f"Sem {num_sem.strip()}"
    
    # ORDEN DE INGENIERÍA:
    # 1. Hojas de Trabajo (Sin prefijo "Sem ")
    hojas_trabajo = [h for h in hojas_origen if not h.startswith("Sem ")]
    # 2. Hojas de semanas anteriores (En orden descendente)
    hojas_historia = [h for h in hojas_origen if h.startswith("Sem ") and h != nombre_sem_nueva]
    hojas_historia.sort(reverse=True)
    
    # Secuencia final: TRABAJO -> NUEVA -> HISTORIA
    lista_final_hojas = hojas_trabajo + [nombre_sem_nueva] + hojas_historia

    output_stream = io.BytesIO()
    with pd.ExcelWriter(output_stream, engine='xlsxwriter') as writer:
        workbook = writer.book
        # Formato de Tiempo Aritmético para Excel ([h]:mm:ss)
        formato_hora = workbook.add_format({'num_format': '[h]:mm:ss'})
        
        for hoja_nombre in lista_final_hojas:
            
            # ESCENARIO 1: LA NUEVA HOJA DETALLADA DE LA SEMANA
            if hoja_nombre == nombre_sem_nueva:
                df_nueva_sem = df_semana[['#', 'Deportista', 'Tiempo Total', 'Actividades', 'Natación', 'Bicicleta', 'Trote', 'CV']].copy()
                df_nueva_sem.rename(columns={'#': 'Clasificación'}, inplace=True)
                
                # Convertir columnas de tiempo a valor decimal para que sean operables
                for col_t in ['Tiempo Total', 'Natación', 'Bicicleta', 'Trote']:
                    df_nueva_sem[col_t] = df_nueva_sem[col_t].apply(to_excel_time_value)
                
                df_nueva_sem.to_excel(writer, sheet_name=hoja_nombre, index=False)
                
                # Aplicar formato visual a columnas de tiempo
                worksheet = writer.sheets[hoja_nombre]
                for col_num in [2, 4, 5, 6]:
                    worksheet.set_column(col_num, col_num, 12, formato_hora)
            
            # ESCENARIO 2: HOJAS DE TRABAJO (TIEMPO TOTAL, NATACIÓN, ETC.)
            elif hoja_nombre in ["Tiempo Total", "Natación", "Ciclismo", "Trote"]:
                df_trabajo = pd.read_excel(xls_engine, sheet_name=hoja_nombre)
                df_trabajo = df_trabajo.drop(columns=['Sem 51', 'Sem 52'], errors='ignore')
                
                # Identificar columna del deportista
                col_nombre_id = next((c for c in df_trabajo.columns if str(c).lower() in ['nombre', 'deportista']), df_trabajo.columns[0])
                
                # Mapeo de columna de datos según la hoja
                mapeo_disciplina = {
                    'Tiempo Total': 'Tiempo Total', 'Natación': 'Natación', 
                    'Ciclismo': 'Bicicleta', 'Trote': 'Trote'
                }
                col_datos_proceso = mapeo_disciplina.get(hoja_nombre)
                
                if col_datos_proceso:
                    # Preparar cruce
                    df_sem_clean = df_semana[['Deportista', col_datos_proceso]].copy()
                    df_sem_clean['MatchID'] = df_sem_clean['Deportista'].apply(clean_string)
                    df_trabajo['MatchID'] = df_trabajo[col_nombre_id].astype(str).apply(clean_string)
                    
                    # Diccionario con valores decimales
                    mapa_valores = df_sem_clean.set_index('MatchID')[col_datos_proceso].apply(to_excel_time_value).to_dict()
                    
                    # Inyectar dato en la columna nueva
                    df_trabajo[nombre_sem_nueva] = df_trabajo['MatchID'].map(mapa_valores).fillna(0)
                    df_trabajo.drop(columns=['MatchID'], inplace=True)
                
                df_trabajo.to_excel(writer, sheet_name=hoja_nombre, index=False)
                
                # Aplicar formato de tiempo a todas las columnas Semanas y Acumulados
                worksheet = writer.sheets[hoja_nombre]
                for idx_c, nombre_c in enumerate(df_trabajo.columns):
                    if "Sem " in str(nombre_c) or "Promedio" in str(nombre_c) or "Acumulado" in str(nombre_c):
                        worksheet.set_column(idx_c, idx_c, 13, formato_hora)

            # ESCENARIO 3: HOJA CV (NO ES TIEMPO, ES VALOR NUMÉRICO O NC)
            elif hoja_nombre == "CV":
                df_cv = pd.read_excel(xls_engine, sheet_name=hoja_nombre)
                col_nombre_id = next((c for c in df_cv.columns if str(c).lower() in ['nombre', 'deportista']), df_cv.columns[0])
                
                df_sem_cv = df_semana[['Deportista', 'CV']].copy()
                df_sem_cv['MatchID'] = df_sem_cv['Deportista'].apply(clean_string)
                df_cv['MatchID'] = df_cv[col_nombre_id].astype(str).apply(clean_string)
                
                mapa_cv = df_sem_cv.set_index('MatchID')['CV'].to_dict()
                df_cv[nombre_sem_nueva] = df_cv['MatchID'].map(mapa_cv).fillna("NC")
                df_cv.drop(columns=['MatchID'], inplace=True)
                df_cv.to_excel(writer, sheet_name=hoja_nombre, index=False)

            # ESCENARIO 4: RESTO DE HOJAS (TRANSCRIPCIÓN ÍNTEGRA)
            else:
                df_resto = pd.read_excel(xls_engine, sheet_name=hoja_nombre)
                df_resto.to_excel(writer, sheet_name=hoja_nombre, index=False)
                
    return output_stream.getvalue()

# --- 6. GENERADOR DE REPORTE WORD (BLOQUEADO / MODELO FUNCIONAL) ---
def aplicar_estilo_tym(parrafo, fuente_size, es_bold=False, centrar=False):
    """Estilo Calibri 20/15/13/11 estricto."""
    run = parrafo.runs[0] if parrafo.runs else parrafo.add_run()
    run.font.name = 'Calibri'
    run.font.size = Pt(fuente_size)
    run.bold = es_bold
    if centrar:
        parrafo.alignment = WD_ALIGN_PARAGRAPH.CENTER

def crear_tabla_profesional_tym(doc, df_podio, lista_columnas):
    """Genera tablas con anchos milimétricos (Protocolo TYM)."""
    tabla = doc.add_table(rows=1, cols=len(lista_columnas))
    tabla.style = 'Light Grid Accent 1'
    tabla.alignment = 1 # Centrado horizontal
    tabla.autofit = False
    
    # Anchos Blindados
    dict_anchos = {
        '#': 0.4, 'Deportista': 2.8, 'Tiempo Total': 0.7, 
        'Natación': 0.7, 'Bicicleta': 0.7, 'Trote': 0.7, 'CV': 0.6
    }
    
    # Encabezado de Tabla
    for i, col_txt in enumerate(lista_columnas):
        celda = tabla.rows[0].cells[i]
        celda.text = col_txt
        celda.width = Inches(dict_anchos.get(col_txt, 0.7))
        aplicar_estilo_tym(celda.paragraphs[0], 9, True, True)
        
    # Filas de Datos
    for _, fila_datos in df_podio.iterrows():
        celdas = tabla.add_row().cells
        for i, col_txt in enumerate(lista_columnas):
            celdas[i].text = str(fila_datos[col_txt])
            celdas[i].width = Inches(dict_anchos.get(col_txt, 0.7))
            # Alineación: Nombres Izquierda, resto Centro
            aplicar_estilo_tym(celdas[i].paragraphs[0], 9, False, col_txt != 'Deportista')
    doc.add_paragraph()

def generar_reporte_word_completo(df_datos, sem_num, podio_d, podio_l):
    """Construye el reporte Word completo sin síntesis de código."""
    doc_word = Document()
    
    # Título Principal (20)
    tit = doc_word.add_heading(f'Reporte Semanal Club Tym Triatlón - Semana {sem_num}', 0)
    aplicar_estilo_tym(tit, 20, True, True)
    doc_word.add_paragraph()
    
    txt_intro = doc_word.add_paragraph('"(La semana de la simetría perfecta y el retorno del volumen)"')
    aplicar_estilo_tym(txt_intro, 11, True, True)
    doc_word.add_paragraph()

    # BLOQUE 1: Resumen General (15)
    h_res = doc_word.add_heading('🔍 Resumen General', level=2)
    aplicar_estilo_tym(h_res, 15, True)
    doc_word.add_paragraph()
    
    df_completos = df_datos[df_datos['CV'] != 'NC'].copy()
    txt_resumen = f"Total deportistas: {len(df_datos)}\nTriatletas completos: {len(df_completos)}\nHoras totales: {to_hhmmss_display(df_datos['T_Mins'].sum())}"
    p_resumen = doc_word.add_paragraph(txt_resumen)
    aplicar_estilo_tym(p_resumen, 11)
    
    # Gráfico Circular
    fig_pie, ax_pie = plt.subplots(figsize=(4,4))
    ax_pie.pie([df_datos['N_Mins'].sum(), df_datos['B_Mins'].sum(), df_datos['R_Mins'].sum()], 
               labels=['Natación', 'Ciclismo', 'Trote'], 
               autopct='%1.1f%%', 
               colors=['#1E90FF', '#32CD32', '#FF4500'])
    buf_img = io.BytesIO()
    plt.savefig(buf_img, format='png', bbox_inches='tight')
    plt.close(fig_pie)
    p_graf = doc_word.add_paragraph()
    p_graf.alignment = 1
    p_graf.add_run().add_picture(buf_img, width=Inches(3.5))

    # BLOQUE 2: Top 5 (Completos y Balanceados)
    lista_top5 = [
        ('🏅 TOP 5 TRIATLETAS COMPLETOS', df_completos.sort_values('T_Mins', ascending=False).head(5), 'Completos'),
        ('⚖️ TOP 5 TRIATLETAS MÁS BALANCEADOS', df_completos.sort_values('CV', ascending=True).head(5), 'CV')
    ]
    
    for titulo_t5, df_t5, cat_t5 in lista_top5:
        h_sec = doc_word.add_heading(titulo_t5, level=2)
        aplicar_estilo_tym(h_sec, 15, True)
        doc_word.add_paragraph()
        
        df_render_t5 = df_t5.copy()
        df_render_t5['#'] = range(1, len(df_render_t5)+1)
        cols_t5 = ['#', 'Deportista', 'Tiempo Total' if cat_t5=='Completos' else 'CV', 'Natación', 'Bicicleta', 'Trote']
        crear_tabla_profesional_tym(doc_word, df_render_t5, cols_t5)
        
        # Análisis Técnico
        p_an_t5 = doc_word.add_paragraph('Análisis del Desempeño:'); aplicar_estilo_tym(p_an_t5, 13, True)
        for _, fila_t5 in df_render_t5.iterrows():
            p_atleta = doc_word.add_paragraph(f"{fila_t5['#']}. {fila_t5['Deportista']}"); aplicar_estilo_tym(p_atleta, 11, True)
            doc_word.add_paragraph(generar_comentario(fila_t5, cat_t5, fila_t5['#']))

    # BLOQUE 3: Top 15 (General y por Disciplina)
    lista_top15 = [
        ('TIEMPO GENERAL', '🥇', 'T_Mins', 'Tiempo Total'),
        ('NATACIÓN', '🏊‍♂️', 'N_Mins', 'Natación'),
        ('CICLISMO', '🚴', 'B_Mins', 'Bicicleta'),
        ('TROTE', '🏃‍♂️', 'R_Mins', 'Trote')
    ]
    
    for tit_t15, icon_t15, col_m, col_t in lista_top15:
        doc_word.add_page_break()
        h_t15 = doc_word.add_heading(f'{icon_t15} TOP 15 {tit_t15}', level=1)
        aplicar_estilo_tym(h_t15, 15, True)
        doc_word.add_paragraph()
        
        df_render_t15 = df_datos[df_datos[col_m]>0].sort_values(col_m, ascending=False).head(15).copy()
        df_render_t15['#'] = range(1, len(df_render_t15)+1)
        
        columnas_t15 = ['#', 'Deportista', col_t, 'Tiempo Total', 'Natación', 'Bicicleta', 'Trote'] if tit_t15 == 'TIEMPO GENERAL' else ['#', 'Deportista', col_t, 'Tiempo Total']
        crear_tabla_profesional_tym(doc_word, df_render_t15, columnas_t15)
        
        # Análisis Podio T15
        p_an_t15 = doc_word.add_paragraph('Análisis del Podio:'); aplicar_estilo_tym(p_an_t15, 13, True)
        for _, fila_t15 in df_render_t15.head(3).iterrows():
            emoji_p = '🥇' if fila_t15['#']==1 else '🥈' if fila_t15['#']==2 else '🥉'
            p_top_atleta = doc_word.add_paragraph(f"{emoji_p} {fila_t15['Deportista']}"); aplicar_estilo_tym(p_top_atleta, 11, True)
            doc_word.add_paragraph(generar_comentario(fila_t15, col_t if col_t != 'Tiempo Total' else 'General', fila_t15['#']))

    # BLOQUE 4: PODIOS OCR FINALES
    doc_word.add_page_break()
    h_dist = doc_word.add_heading('📏 PODIO DISTANCIA TOTAL', level=1); aplicar_estilo_tym(h_dist, 15, True); doc_word.add_paragraph()
    for i, p_item in enumerate(podio_d): 
        doc_word.add_paragraph(f"{i+1}. {p_item['nombre']} ({p_item['valor']} km)")
    
    doc_word.add_paragraph()
    h_larga = doc_word.add_heading('⏱️ PODIO ACTIVIDAD MÁS LARGA', level=1); aplicar_estilo_tym(h_larga, 15, True); doc_word.add_paragraph()
    for i, p_item in enumerate(podio_l): 
        doc_word.add_paragraph(f"{i+1}. {p_item['nombre']} ({p_item['valor']})")
    
    byte_buf = io.BytesIO(); doc_word.save(byte_buf); byte_buf.seek(0); return byte_buf

# --- 7. INTERFAZ DE USUARIO ---
st.sidebar.header("📁 Gestión de Datos Históricos")
excel_maestro = st.sidebar.file_uploader("Cargar Excel Maestro (00 Estadísticas)", type=["xlsx"])

n_semana = st.text_input("Semana a procesar:", "08")
txt_strava = st.text_area("1. Datos de Tiempo Total (Strava):")
txt_ocr = st.text_area("2. Traducción OCR (Captura):")

if st.button("🚀 PROCESAR JORNADA Y ACTUALIZAR EXCEL"):
    if txt_strava.strip() and txt_ocr.strip() and excel_maestro:
        # Ejecutar análisis
        df_resultados = parse_raw_data(txt_strava)
        p_dist, p_larg = parse_ocr_data(txt_ocr)
        
        st.success(f"¡Semana {n_semana} procesada!")
        cw, ce = st.columns(2)
        
        # Word
        cw.download_button(
            label="📄 REPORTE WORD", 
            data=generar_reporte_word_completo(df_resultados, n_semana, p_dist, p_larg), 
            file_name=f"Reporte_TYM_Sem_{n_semana}.docx"
        )
        
        # Excel
        ce.download_button(
            label="📊 EXCEL ACTUALIZADO", 
            data=crear_excel_actualizado(excel_maestro, df_resultados, n_semana), 
            file_name=f"00_Estadisticas_Actualizado_Sem_{n_semana}.xlsx"
        )
    else:
        st.error("Error: Se requieren todos los archivos y datos para procesar.")
