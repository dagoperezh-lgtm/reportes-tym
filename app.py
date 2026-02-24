import streamlit as st
import pandas as pd
import numpy as np
import re
import matplotlib.pyplot as plt
import io
import os
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

# --- CONFIGURACIÓN DE PÁGINA (BLINDADO) ---
st.set_page_config(page_title="Inteligencia TYM", page_icon="🏆", layout="wide")

st.title("🏆 Plataforma de Inteligencia Deportiva - TYM")

HISTORICO_FILE = "historico_tym.csv"

# --- UTILIDADES DE TIEMPO (BLINDADO) ---
def to_mins(t_str):
    if pd.isna(t_str) or '--:--' in t_str or not t_str: return 0
    h, m = 0, 0
    hm = re.search(r'(\d+)h', str(t_str))
    if hm: h = int(hm.group(1))
    mm = re.search(r'(\d+)min', str(t_str))
    if mm: m = int(mm.group(1))
    return h * 60 + m

def to_hhmmss(mins):
    h = int(mins // 60)
    m = int(mins % 60)
    return f"{h:02d}:{m:02d}:00"

# --- MOTOR DE COMENTARIOS (BLINDADO) ---
def generar_comentario(row, df_contexto, categoria, pos):
    nombre = row['Deportista']
    if categoria in ['Completos', 'General']:
        if pos == 1: return f"Dominio absoluto de {nombre}. Se consolida en la cima del club con un volumen total envidiable. Su capacidad para sostener sesiones de alta intensidad demuestra una preparación de élite."
        elif pos == 2: return f"Una semana brillante para {nombre}. Se queda con la plata manteniendo una presión constante sobre el líder."
        else: return f"{nombre} cierra el podio con un rendimiento muy equilibrado. Demuestra que para estar en el grupo de avanzada no se puede regalar nada."
    if categoria == 'CV':
        if pos == 1: return f"¡El reloj suizo del club! {nombre} logra una simetría casi perfecta, demostrando una planificación milimétrica de sus cargas."
        return f"Excelente balance para {nombre}. Mantiene sus disciplinas en una armonía poco común."
    tiempo = row.get(categoria, "")
    if categoria == 'Natación': return f"Fuerza pura en el agua. {nombre} registra {tiempo}, liderando el podio con técnica depurada."
    if categoria == 'Bicicleta': return f"Potencia pura sobre ruedas. {nombre} devoró la ruta con {tiempo}."
    if categoria == 'Trote': return f"Resistencia inalcanzable. {nombre} domina el asfalto con {tiempo}."
    return "Desempeño destacado."

# --- PARSER DE DATOS (BLINDADO) ---
def parse_raw_data(raw_text):
    parsed_data = []
    rank_counter = 1 
    raw_text = raw_text.replace('\xa0', ' ')
    for line in raw_text.strip().split('\n'):
        line = line.strip()
        if not line or 'Deportista' in line: continue
        try:
            time_pattern = r'(\d+h\s*\d*min|\d+h|\d+min|--:--)'
            matches = list(re.finditer(time_pattern, line))
            if not matches: continue
            name_part = line[:matches[0].start()].strip()
            name = re.sub(r'^\d+\s*', '', name_part).strip()
            total_time = matches[0].group(1)
            rest = line[matches[0].end():].strip()
            rest = re.sub(r'\d+$', '', rest).strip() 
            rem_matches = list(re.finditer(time_pattern, rest))
            acts_str = rest[:rem_matches[0].start()].strip() if rem_matches else "0"
            acts = int(acts_str) if acts_str.isdigit() else 0
            times = [m.group(1) for m in rem_matches]
            s_str = times[0] if len(times) > 0 else "--:--"
            b_str = times[1] if len(times) > 1 else "--:--"
            r_str = times[2] if len(times) > 2 else "--:--"
            t_m, s_m, b_m, r_m = to_mins(total_time), to_mins(s_str), to_mins(b_str), to_mins(r_str)
            cv = "NC" if 0 in [s_m, b_m, r_m] else round(np.std([s_m, b_m, r_m]) / np.mean([s_m, b_m, r_m]), 4)
            parsed_data.append({'#': rank_counter, 'Deportista': name, 'Tiempo Total': to_hhmmss(t_m), 'Actividades': acts, 'Natación': to_hhmmss(s_m), 'Bicicleta': to_hhmmss(b_m), 'Trote': to_hhmmss(r_m), 'CV': cv, 'T_Mins': t_m, 'N_Mins': s_m, 'B_Mins': b_m, 'R_Mins': r_m})
            rank_counter += 1
        except: pass
    return pd.DataFrame(parsed_data)

def parse_ocr_data(ocr_text):
    distancia_podio, larga_podio = [], []
    for line in ocr_text.strip().split('\n'):
        parts = line.split(';')
        if len(parts) >= 6:
            distancia_podio.append({'nombre': parts[2].strip(), 'valor': parts[3].strip()})
            larga_podio.append({'nombre': parts[4].strip(), 'valor': parts[5].strip()})
    return distancia_podio[:3], larga_podio[:3]

# --- GESTIÓN DE EXCEL (REFORZADO CONTRA VALUEERROR) ---
def actualizar_excel_historico(df_actual, df_hist_previo, num_sem):
    col_nueva = f"Sem {num_sem}"
    df_nueva = df_actual[['Deportista', 'T_Mins']].rename(columns={'Deportista': 'Nombre', 'T_Mins': col_nueva})
    # Limpieza estricta de nombres para el cruce
    df_nueva['Nombre'] = df_nueva['Nombre'].astype(str).str.strip().str.upper()
    
    if df_hist_previo is not None:
        # Normalización de columnas y contenido
        df_hist_previo.columns = [str(c).strip().capitalize() for c in df_hist_previo.columns]
        if 'Nombre' not in df_hist_previo.columns:
            df_hist_previo.rename(columns={df_hist_previo.columns[0]: 'Nombre'}, inplace=True)
            
        # Forzamos tipo texto y limpieza en el historial previo
        df_hist_previo['Nombre'] = df_hist_previo['Nombre'].astype(str).str.strip().str.upper()
        
        if col_nueva in df_hist_previo.columns:
            df_hist_previo = df_hist_previo.drop(columns=[col_nueva])
            
        df_final = pd.merge(df_hist_previo, df_nueva, on='Nombre', how='outer').fillna(0)
    else:
        df_final = df_nueva
    
    cols_sem = [c for c in df_final.columns if "Sem" in c]
    df_final['Total Acumulado (mins)'] = df_final[cols_sem].sum(axis=1)
    df_final['Total Formateado'] = df_final['Total Acumulado (mins)'].apply(to_hhmmss)
    
    return df_final.sort_values('Total Acumulado (mins)', ascending=False)

# --- WORD GENERATOR (BLINDADO) ---
def aplicar_estilo(parrafo, size, negrita=False, center=False):
    run = parrafo.runs[0] if parrafo.runs else parrafo.add_run()
    run.font.name = 'Calibri'
    run.font.size = Pt(size)
    run.bold = negrita
    if center: parrafo.alignment = WD_ALIGN_PARAGRAPH.CENTER

def crear_tabla_centrada(doc, df, cols):
    table = doc.add_table(rows=1, cols=len(cols))
    table.style = 'Light Grid Accent 1'
    table.alignment = WD_ALIGN_PARAGRAPH.CENTER
    table.autofit = False
    anchos = {'#': 0.3, 'Deportista': 3.0, 'Tiempo Total': 0.7, 'Natación': 0.7, 'Bicicleta': 0.7, 'Trote': 0.7, 'CV': 0.5}
    for i, col_name in enumerate(cols):
        cell = table.rows[0].cells[i]
        cell.text = col_name
        cell.width = Inches(anchos.get(col_name, 0.7))
        aplicar_estilo(cell.paragraphs[0], 9, True, True)
    for _, row in df.iterrows():
        row_cells = table.add_row().cells
        for i, col_name in enumerate(cols):
            row_cells[i].text = str(row[col_name])
            row_cells[i].width = Inches(anchos.get(col_name, 0.7))
            aplicar_estilo(row_cells[i].paragraphs[0], 9, False, True if col_name != 'Deportista' else False)
    doc.add_paragraph()

def generar_word(df, sem, dist_p, larg_p):
    doc = Document()
    h0 = doc.add_heading(f'Reporte Semanal Club Tym Triatlón - Semana {sem}', 0)
    aplicar_estilo(h0, 20, True, True)
    doc.add_paragraph()
    p_intro = doc.add_paragraph('"(La semana de la simetría perfecta y el retorno del volumen)"')
    aplicar_estilo(p_intro, 11, True, True)
    doc.add_paragraph()
    h1 = doc.add_heading('🔍 Resumen General', level=2)
    aplicar_estilo(h1, 15, True)
    doc.add_paragraph()
    df_c = df[df['CV'] != 'NC'].copy()
    t_m = df['T_Mins'].sum()
    p_res = doc.add_paragraph(f'Total deportistas: {len(df)}\nTriatletas completos: {len(df_c)}\nHoras totales: {int(t_m//60)}h {int(t_m%60)}m')
    aplicar_estilo(p_res, 11)
    fig, ax = plt.subplots(figsize=(4,4))
    ax.pie([df['N_Mins'].sum(), df['B_Mins'].sum(), df['R_Mins'].sum()], labels=['Nat', 'Bici', 'Trote'], autopct='%1.1f%%', colors=['#1E90FF', '#32CD32', '#FF4500'])
    img_stream = io.BytesIO()
    plt.savefig(img_stream, format='png', bbox_inches='tight')
    plt.close(fig)
    p_img = doc.add_paragraph()
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img.add_run().add_picture(img_stream, width=Inches(3.5))
    h2 = doc.add_heading('🏅 TOP 5 TRIATLETAS COMPLETOS', level=2)
    aplicar_estilo(h2, 15, True)
    doc.add_paragraph()
    t5 = df_c.sort_values('T_Mins', ascending=False).head(5).copy()
    t5['#'] = range(1, 6)
    crear_tabla_centrada(doc, t5, ['#', 'Deportista', 'Tiempo Total', 'Natación', 'Bicicleta', 'Trote'])
    s1 = doc.add_paragraph('Análisis Ejecutivo')
    aplicar_estilo(s1, 13, True)
    for i, r in t5.iterrows():
        p = doc.add_paragraph(f"{r['#']}. {r['Deportista']}")
        aplicar_estilo(p, 11, True)
        doc.add_paragraph(generar_comentario(r, t5, 'Completos', r['#']))
    h3 = doc.add_heading('⚖️ TOP 5 TRIATLETAS MÁS BALANCEADOS', level=2)
    aplicar_estilo(h3, 15, True)
    doc.add_paragraph()
    df_c['CV_n'] = df_c['CV'].astype(float)
    b5 = df_c.sort_values('CV_n').head(5).copy()
    b5['#'] = range(1, 6)
    crear_tabla_centrada(doc, b5, ['#', 'Deportista', 'CV', 'Natación', 'Bicicleta', 'Trote'])
    s2 = doc.add_paragraph('Análisis de Simetría')
    aplicar_estilo(s2, 13, True)
    for i, r in b5.iterrows():
        p = doc.add_paragraph(f"{r['#']}. {r['Deportista']} (CV: {r['CV']})")
        aplicar_estilo(p, 11, True)
        doc.add_paragraph(generar_comentario(r, b5, 'CV', r['#']))
    doc.add_page_break()
    h4 = doc.add_heading('🥇 TOP 15 TIEMPO TOTAL GENERAL', level=1)
    aplicar_estilo(h4, 15, True); doc.add_paragraph()
    t15 = df.sort_values('T_Mins', ascending=False).head(15).copy()
    t15['#'] = range(1, 16)
    crear_tabla_centrada(doc, t15, ['#', 'Deportista', 'Tiempo Total', 'Natación', 'Bicicleta', 'Trote'])
    s3 = doc.add_paragraph('Análisis del Podio General')
    aplicar_estilo(s3, 13, True)
    for i, r in t15.head(3).iterrows():
        p = doc.add_paragraph(f"{r['#']}. {r['Deportista']}")
        aplicar_estilo(p, 11, True)
        doc.add_paragraph(generar_comentario(r, t15, 'General', r['#']))
    for disc, col, icono, m_col in [('NATACIÓN', 'Natación', '🏊‍♂️', 'N_Mins'), ('CICLISMO', 'Bicicleta', '🚴', 'B_Mins'), ('TROTE', 'Trote', '🏃‍♂️', 'R_Mins')]:
        doc.add_page_break()
        hd = doc.add_heading(f'{icono} TOP 15 {disc}', level=1)
        aplicar_estilo(hd, 15, True); doc.add_paragraph()
        d15 = df[df[m_col]>0].sort_values(m_col, ascending=False).head(15).copy()
        d15['#'] = range(1, len(d15)+1)
        crear_tabla_centrada(doc, d15, ['#', 'Deportista', col, 'Tiempo Total'])
        sd = doc.add_paragraph('Análisis del Podio')
        aplicar_estilo(sd, 13, True)
        for i, r in d15.head(3).iterrows():
            p = doc.add_paragraph(f"{r['#']}. {r['Deportista']} ({r[col]})")
            aplicar_estilo(p, 11, True)
            doc.add_paragraph(generar_comentario(r, d15, col, r['#']))
    doc.add_page_break()
    h5 = doc.add_heading('📏 PODIO DISTANCIA TOTAL', level=1); aplicar_estilo(h5, 15, True); doc.add_paragraph()
    for i, p in enumerate(dist_p): doc.add_paragraph(f"{i+1}. {p['nombre']} ({p['valor']} km)")
    h6 = doc.add_heading('⏱️ PODIO ACTIVIDAD MÁS LARGA', level=1); aplicar_estilo(h6, 15, True); doc.add_paragraph()
    for i, p in enumerate(larg_p): doc.add_paragraph(f"{i+1}. {p['nombre']} ({p['valor']})")
    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf

# --- UI ---
st.sidebar.header("📁 Gestión de Historial")
archivo_hist_cargado = st.sidebar.file_uploader("Subir Excel de Historial", type=["xlsx", "csv"])

sem_ui = st.text_input("Número de Semana:", "08")
raw_ui = st.text_area("1. Datos Tiempo Total:")
ocr_ui = st.text_area("2. Datos OCR Captura:")

if st.button("Generar Reporte Final"):
    if raw_ui.strip() and ocr_ui.strip():
        df_res = parse_raw_data(raw_ui)
        dist_p, larg_p = parse_ocr_data(ocr_ui)
        
        df_hist_previo = None
        if archivo_hist_cargado:
            df_hist_previo = pd.read_excel(archivo_hist_cargado) if archivo_hist_cargado.name.endswith('xlsx') else pd.read_csv(archivo_hist_cargado)
        
        df_hist_final = actualizar_excel_historico(df_res, df_hist_previo, sem_ui)
        
        st.success("¡Todo generado!")
        col1, col2 = st.columns(2)
        col1.download_button("📄 DESCARGAR REPORTE WORD", generar_word(df_res, sem_ui, dist_p, larg_p), f"Reporte_TYM_{sem_ui}.docx")
        
        excel_buf = io.BytesIO()
        df_hist_final.to_excel(excel_buf, index=False)
        col2.download_button("📊 DESCARGAR EXCEL HISTÓRICO", excel_buf.getvalue(), f"Historial_TYM_Sem_{sem_ui}.xlsx")
