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

# --- 2. UTILIDADES DE PROCESAMIENTO Y TIEMPO (BLINDADO) ---
def clean_string(text):
    """Normaliza nombres para eliminar tildes, mayúsculas y espacios extra."""
    if not text: return ""
    text = str(text).strip().upper()
    # Normalización para ignorar tildes y caracteres especiales de forma robusta
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
            # Maneja formatos HH:MM o HH:MM:SS
            return int(parts[0]) * 60 + int(parts[1])
        
        # Búsqueda por patrones de texto (h, min)
        h_match = re.search(r'(\d+)h', t_str)
        m_match = re.search(r'(\d+)min', t_str)
        h = int(h_match.group(1)) if h_match else 0
        m = int(m_match.group(1)) if m_match else 0
        return h * 60 + m
    except: 
        return 0

def to_hhmmss(mins):
    """Convierte minutos a formato estándar de texto HH:MM:00."""
    h = int(mins // 60)
    m = int(mins % 60)
    return f"{h:02d}:{m:02d}:00"

# --- 3. MOTOR DE COMENTARIOS TÉCNICOS (BLOQUEADO - NO SINTETIZAR) ---
def generar_comentario(row, cat, pos):
    """Genera análisis cualitativo para los podios del reporte Word."""
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
    
    tiempo = row.get(cat, "")
    if cat == 'Natación': 
        return f"Fuerza pura en el agua. {nombre} registra {tiempo}, liderando el podio con una técnica depurada. Sus hombros de acero dominan el volumen acuático."
    if cat == 'Bicicleta': 
        return f"Potencia pura sobre ruedas. {nombre} devoró la ruta con {tiempo}. Demuestra ser el motor del equipo en la carretera con promedios que intimidan."
    if cat == 'Trote': 
        return f"Resistencia inalcanzable. {nombre} domina el asfalto con {tiempo} y una fase de carrera soberbia."
    
    return "Desempeño técnico destacado durante la jornada."

# --- 4. PARSERS DE ENTRADA (BLINDADO) ---
def parse_raw_data(raw_text):
    """Parsea los datos de Tiempo Total pegados desde Strava."""
    data = []
    rank = 1
    raw_text = raw_text.replace('\xa0', ' ')
    lines = raw_text.strip().split('\n')
    
    for line in lines:
        if not line or 'Deportista' in line: 
            continue
        try:
            # Buscar patrones de tiempo
            times = re.findall(r'(\d+h\s*\d*min|\d+h|\d+min|--:--)', line)
            if not times: 
                continue
            
            # Limpiar nombre (antes del primer tiempo)
            name_part = line[:line.find(times[0])].strip()
            name = re.sub(r'^\d+\s*', '', name_part).strip()
            
            # Tiempos por disciplina
            t_m = to_mins(times[0])
            s_m = to_mins(times[1] if len(times) > 1 else 0)
            b_m = to_mins(times[2] if len(times) > 2 else 0)
            r_m = to_mins(times[3] if len(times) > 3 else 0)
            
            # Cálculo de Coeficiente de Variación (CV)
            vals = [s_m, b_m, r_m]
            cv = "NC" if 0 in vals else round(np.std(vals) / np.mean(vals), 4)
            
            # Número de actividades
            acts_match = re.search(r'\d+', line[line.find(times[0])+len(times[0]):])
            acts = int(acts_match.group()) if acts_match else 0
            
            data.append({
                '#': rank,
                'Deportista': name,
                'Tiempo Total': to_hhmmss(t_m),
                'Actividades': acts,
                'Natación': to_hhmmss(s_m),
                'Bicicleta': to_hhmmss(b_m),
                'Trote': to_hhmmss(r_m),
                'CV': cv,
                'T_Mins': t_m,
                'N_Mins': s_m,
                'B_Mins': b_m,
                'R_Mins': r_m
            })
            rank += 1
        except: 
            pass
    return pd.DataFrame(data)

def parse_ocr_data(ocr_text):
    """Parsea la tabla OCR filtrando mandatoriamente los encabezados."""
    dist, larg = [], []
    keywords_filtro = ["Nombre", "Distancia", "Actividad", "Tiempo", "Km"]
    lines = ocr_text.strip().split('\n')
    
    for line in lines:
        parts = line.split(';')
        if len(parts) >= 6:
            n_dist = parts[2].strip()
            v_dist = parts[3].strip()
            n_larg = parts[4].strip()
            v_larg = parts[5].strip()
            
            # Validar que no sea un encabezado
            es_encabezado_d = any(k in n_dist for k in keywords_filtro)
            es_encabezado_l = any(k in n_larg for k in keywords_filtro)
            
            if not es_encabezado_d and n_dist:
                dist.append({'nombre': n_dist, 'valor': v_dist})
            if not es_encabezado_l and n_larg:
                larg.append({'nombre': n_larg, 'valor': v_larg})
                
    return dist[:3], larg[:3]

# --- 5. ACTUALIZADOR DE EXCEL (ORDEN OPERATIVO Y TRASPLANTE DE VALORES) ---
def crear_excel_actualizado(archivo_maestro, df_semana, num_sem):
    """Actualiza el archivo Excel respetando el orden de hojas de trabajo."""
    xls = pd.ExcelFile(archivo_maestro)
    hojas_originales = xls.sheet_names
    nombre_nueva_hoja = f"Sem {num_sem.strip()}"
    
    # 🛡️ DEFINICIÓN DE ORDEN MANDATORIO:
    # 1. Hojas de Trabajo Técnicas (Número de Semana, Tiempo Total, Natación, Ciclismo, Trote, CV)
    hojas_trabajo = [h for h in hojas_originales if not h.startswith("Sem ")]
    # 2. Hojas Semanales Existentes (en orden descendente)
    hojas_historia = [h for h in hojas_originales if h.startswith("Sem ") and h != nombre_nueva_hoja]
    
    # El orden debe ser: Trabajo -> Nueva Semana -> Historia
    nuevo_orden_hojas = hojas_trabajo + [nombre_nueva_hoja] + hojas_historia

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for hoja in nuevo_orden_hojas:
            # ESCENARIO A: Crear la nueva hoja semanal detallada
            if hoja == nombre_nueva_hoja:
                df_export = df_semana[['#', 'Deportista', 'Tiempo Total', 'Actividades', 'Natación', 'Bicicleta', 'Trote', 'CV']].copy()
                df_export.rename(columns={'#': 'Clasificación'}, inplace=True)
                df_export.to_excel(writer, sheet_name=hoja, index=False)
            
            # ESCENARIO B: Actualizar columnas de hojas de trabajo (transcribiendo valores)
            elif hoja in ["Tiempo Total", "Natación", "Ciclismo", "Trote", "CV"]:
                df_original = pd.read_excel(xls, sheet_name=hoja)
                # Eliminación preventiva de columnas basura
                df_original = df_original.drop(columns=['Sem 51', 'Sem 52'], errors='ignore')
                
                # Identificar columna llave para el match
                col_key = next((c for c in df_original.columns if str(c).lower() in ['nombre', 'deportista']), df_original.columns[0])
                
                # Mapeo de datos para inyección
                mapeo_hojas = {
                    'Tiempo Total': 'Tiempo Total', 'Natación': 'Natación', 
                    'Ciclismo': 'Bicicleta', 'Trote': 'Trote', 'CV': 'CV'
                }
                col_datos_fuente = mapeo_hojas.get(hoja)
                
                if col_datos_fuente:
                    # Preparar llave de cruce limpia
                    df_upd_prep = df_semana[['Deportista', col_datos_fuente]].copy()
                    df_upd_prep['MatchKey'] = df_upd_prep['Deportista'].apply(clean_string)
                    df_original['MatchKey'] = df_original[col_key].astype(str).apply(clean_string)
                    
                    # Generar diccionario de actualización
                    dict_updates = df_upd_prep.set_index('MatchKey')[col_datos_fuente].to_dict()
                    
                    # Definir valor por defecto (00:00:00 o NC)
                    valor_defecto = 'NC' if hoja == 'CV' else '00:00:00'
                    
                    # Inyectar datos en la columna de la semana
                    df_original[nombre_nueva_hoja] = df_original['MatchKey'].map(dict_updates).fillna(valor_defecto)
                    df_original.drop(columns=['MatchKey'], inplace=True)
                
                df_original.to_excel(writer, sheet_name=hoja, index=False)
            
            # ESCENARIO C: Copiar el resto de hojas sin modificaciones (historial Sem 07...)
            else:
                df_hist = pd.read_excel(xls, sheet_name=hoja)
                df_hist.to_excel(writer, sheet_name=hoja, index=False)
                
    return output.getvalue()

# --- 6. GENERADOR DE REPORTE WORD (BLOQUEADO / NO MODIFICAR) ---
def aplicar_estilo_tym(parrafo, tamano, es_negrita=False, es_centrado=False):
    """Aplica el estilo oficial Calibri blindado."""
    run = parrafo.runs[0] if parrafo.runs else parrafo.add_run()
    run.font.name = 'Calibri'
    run.font.size = Pt(tamano)
    run.bold = es_negrita
    if es_centrado:
        parrafo.alignment = WD_ALIGN_PARAGRAPH.CENTER

def crear_tabla_blindada(documento, dataframe, columnas):
    """Crea tablas centradas con anchos de columna fijos (Protocolo TYM)."""
    tabla = documento.add_table(rows=1, cols=len(columnas))
    tabla.style = 'Light Grid Accent 1'
    tabla.alignment = 1 # Centrado horizontal
    tabla.autofit = False
    
    # Anchos milimétricos: Rank (0.4"), Nombres (2.8"), Datos (0.7")
    mapa_anchos = {
        '#': 0.4, 'Deportista': 2.8, 'Tiempo Total': 0.7, 
        'Natación': 0.7, 'Bicicleta': 0.7, 'Trote': 0.7, 'CV': 0.6
    }
    
    # Encabezados
    for i, col_name in enumerate(columnas):
        celda = tabla.rows[0].cells[i]
        celda.text = col_name
        celda.width = Inches(mapa_anchos.get(col_name, 0.7))
        aplicar_estilo_tym(celda.paragraphs[0], 9, True, True)
        
    # Datos
    for _, fila in dataframe.iterrows():
        celdas_fila = tabla.add_row().cells
        for i, col_name in enumerate(columnas):
            celdas_fila[i].text = str(fila[col_name])
            celdas_fila[i].width = Inches(mapa_anchos.get(col_name, 0.7))
            # Alineación: Nombres a la izquierda, el resto al centro
            aplicar_estilo_tym(celdas_fila[i].paragraphs[0], 9, False, col_name != 'Deportista')
    documento.add_paragraph() # Espaciador

def generar_reporte_word(df, sem, dist_p, larg_p):
    """Construye el reporte Word completo bajo el modelo funcional aceptado."""
    doc = Document()
    
    # Título Principal (Calibri 20)
    tit_principal = doc.add_heading(f'Reporte Semanal Club Tym Triatlón - Semana {sem}', 0)
    aplicar_estilo_tym(tit_principal, 20, True, True)
    doc.add_paragraph()
    
    intro = doc.add_paragraph('"(La semana de la simetría perfecta y el retorno del volumen)"')
    aplicar_estilo_tym(intro, 11, True, True)
    doc.add_paragraph()

    # SECCIÓN 1: Resumen General (Calibri 15)
    sec_resumen = doc.add_heading('🔍 Resumen General', level=2)
    aplicar_estilo_tym(sec_resumen, 15, True)
    doc.add_paragraph()
    
    df_completos = df[df['CV'] != 'NC'].copy()
    txt_res = f"Total deportistas registrados: {len(df)}\nTriatletas con reporte completo: {len(df_completos)}\nHoras totales de entrenamiento: {to_hhmmss(df['T_Mins'].sum())}"
    p_res = doc.add_paragraph(txt_res)
    aplicar_estilo_tym(p_res, 11)
    
    # Gráfico Circular Centrado
    fig, ax = plt.subplots(figsize=(4,4))
    ax.pie([df['N_Mins'].sum(), df['B_Mins'].sum(), df['R_Mins'].sum()], 
           labels=['Natación', 'Ciclismo', 'Trote'], 
           autopct='%1.1f%%', 
           colors=['#1E90FF', '#32CD32', '#FF4500'])
    img_stream = io.BytesIO()
    plt.savefig(img_stream, format='png', bbox_inches='tight')
    plt.close(fig)
    p_grafico = doc.add_paragraph()
    p_grafico.alignment = 1
    p_grafico.add_run().add_picture(img_stream, width=Inches(3.5))

    # SECCIÓN 2: TOP 5 (Completos y Balanceados)
    categorias_t5 = [
        ('🏅 TOP 5 TRIATLETAS COMPLETOS', df_completos.sort_values('T_Mins', ascending=False).head(5), 'Completos'),
        ('⚖️ TOP 5 TRIATLETAS MÁS BALANCEADOS', df_completos.sort_values('CV', ascending=True).head(5), 'CV')
    ]
    
    for titulo, df_p, cat_key in categorias_t5:
        h_t5 = doc.add_heading(titulo, level=2)
        aplicar_estilo_tym(h_t5, 15, True)
        doc.add_paragraph()
        
        df_render = df_p.copy()
        df_render['#'] = range(1, len(df_render)+1)
        columnas_t5 = ['#', 'Deportista', 'Tiempo Total' if cat_key=='Completos' else 'CV', 'Natación', 'Bicicleta', 'Trote']
        crear_tabla_blindada(doc, df_render, columnas_t5)
        
        # Análisis del Podio T5 (Calibri 13 y 11)
        h_analisis = doc.add_paragraph('Análisis del Desempeño:'); aplicar_estilo_tym(h_analisis, 13, True)
        for _, r in df_render.iterrows():
            p_nom_t5 = doc.add_paragraph(f"{r['#']}. {r['Deportista']}"); aplicar_estilo_tym(p_nom_t5, 11, True)
            doc.add_paragraph(generar_comentario(r, cat_key, r['#']))

    # SECCIÓN 3: TOP 15 (General y Disciplinas)
    disciplinas_t15 = [
        ('TIEMPO GENERAL', '🥇', 'T_Mins', 'Tiempo Total'),
        ('NATACIÓN', '🏊‍♂️', 'N_Mins', 'Natación'),
        ('CICLISMO', '🚴', 'B_Mins', 'Bicicleta'),
        ('TROTE', '🏃‍♂️', 'R_Mins', 'Trote')
    ]
    
    for nom_sec, icon, col_mins, col_txt in disciplinas_t15:
        doc.add_page_break()
        h_t15 = doc.add_heading(f'{icon} TOP 15 {nom_sec}', level=1)
        aplicar_estilo_tym(h_t15, 15, True)
        doc.add_paragraph()
        
        df_t15 = df[df[col_mins]>0].sort_values(col_mins, ascending=False).head(15).copy()
        df_t15['#'] = range(1, len(df_t15)+1)
        cols_final = ['#', 'Deportista', col_txt, 'Tiempo Total', 'Natación', 'Bicicleta', 'Trote'] if nom_sec == 'TIEMPO GENERAL' else ['#', 'Deportista', col_txt, 'Tiempo Total']
        crear_tabla_blindada(doc, df_t15, cols_final)
        
        # Análisis del Podio T15
        h_an_t15 = doc.add_paragraph('Análisis del Podio:'); aplicar_estilo_tym(h_an_t15, 13, True)
        for _, r in df_t15.head(3).iterrows():
            emoji = '🥇' if r['#']==1 else '🥈' if r['#']==2 else '🥉'
            p_p15 = doc.add_paragraph(f"{emoji} {r['Deportista']}"); aplicar_estilo_tym(p_p15, 11, True)
            doc.add_paragraph(generar_comentario(r, col_txt if col_txt != 'Tiempo Total' else 'General', r['#']))

    # SECCIÓN 4: PODIOS OCR FINAL
    doc.add_page_break()
    h_dist = doc.add_heading('📏 PODIO DISTANCIA TOTAL', level=1); aplicar_estilo_tym(h_dist, 15, True); doc.add_paragraph()
    for i, p in enumerate(dist_p): 
        doc.add_paragraph(f"{i+1}. {p['nombre']} ({p['valor']} km)")
    
    doc.add_paragraph()
    h_larg = doc.add_heading('⏱️ PODIO ACTIVIDAD MÁS LARGA', level=1); aplicar_estilo_tym(h_larg, 15, True); doc.add_paragraph()
    for i, p in enumerate(larg_p): 
        doc.add_paragraph(f"{i+1}. {p['nombre']} ({p['valor']})")
    
    buf = io.BytesIO(); doc.save(buf); buf.seek(0); return buf

# --- 7. INTERFAZ DE USUARIO ---
st.sidebar.header("📁 Gestión de Datos Históricos")
archivo_maestro = st.sidebar.file_uploader("Cargar Excel 00 Estadísticas TYM", type=["xlsx"])

sem_procesar = st.text_input("Número de Semana a procesar (Ej: 08):", "08")
data_strava = st.text_area("1. Pegar Datos de Tiempo Total (Strava):")
data_ocr = st.text_area("2. Pegar Traducción OCR (Distancia y Salida Larga):")

if st.button("🚀 PROCESAR JORNADA COMPLETA"):
    if data_strava.strip() and data_ocr.strip() and archivo_maestro:
        # Ejecutar Parsers
        df_semana = parse_raw_data(data_strava)
        podio_dist, podio_larg = parse_ocr_data(data_ocr)
        
        st.success(f"¡Semana {sem_procesar} procesada correctamente!")
        col_w, col_e = st.columns(2)
        
        # Botón para descargar Word
        col_w.download_button(
            label="📄 DESCARGAR REPORTE WORD", 
            data=generar_reporte_word(df_semana, sem_procesar, podio_dist, podio_larg), 
            file_name=f"Reporte_TYM_Sem_{sem_procesar}.docx"
        )
        
        # Botón para descargar Excel
        col_e.download_button(
            label="📊 DESCARGAR EXCEL ACTUALIZADO", 
            data=crear_excel_actualizado(archivo_maestro, df_semana, sem_procesar), 
            file_name=f"00_Estadisticas_Actualizado_Sem_{sem_procesar}.xlsx"
        )
    else:
        st.error("Error: Debes cargar el archivo maestro y completar ambos campos de datos.")
