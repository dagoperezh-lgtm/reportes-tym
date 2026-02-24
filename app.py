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

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Inteligencia TYM", page_icon="🏆", layout="wide")

st.title("🏆 Plataforma de Inteligencia Deportiva - TYM")
st.markdown("Generación automatizada de reportes con formato oficial y memoria histórica.")

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

# --- MOTOR DE COMENTARIOS INTELIGENTES (BLINDADO) ---
def generar_comentario(row, df_contexto, categoria, pos):
    nombre = row['Deportista']
    t_mins = row.get('T_Mins', 0)
    b_mins = row.get('B_Mins', 0)
    acts = row.get('Actividades', 0)
    
    pct_bici = (b_mins / t_mins * 100) if t_mins > 0 else 0
    
    if categoria == 'Completos':
        if pos == 1:
            return f"Dominio absoluto de {nombre}. Se consolida en la cima del club con un volumen total envidiable. Su capacidad para sostener {acts} sesiones demuestra una preparación de élite y una disciplina inquebrantable."
        elif pos == 2:
            return f"Una semana brillante para {nombre}. Se queda con la plata general manteniendo una presión constante sobre el líder. Su solidez en el ciclismo ({row['Bicicleta']}) fue el motor que lo mantuvo arriba."
        else:
            return f"{nombre} cierra el podio con un rendimiento muy equilibrado. Demuestra que para estar en el Top 5 no se puede regalar nada, sumando minutos de calidad en las tres disciplinas."

    if categoria == 'CV':
        cv_val = float(row['CV'])
        if pos == 1:
            return f"¡El reloj suizo del club! {nombre} logra una simetría de {cv_val}, una cifra casi perfecta. Requiere una planificación quirúrgica y un control total de las cargas."
        return f"Excelente balance para {nombre}. Mantiene sus disciplinas en una armonía poco común, progresando de forma integral sin sobrecargar una sola área."

    tiempo = row.get(categoria, "")
    if categoria == 'Natación':
        return f"Fuerza pura en el agua. {nombre} registra {tiempo}, liderando el podio con una técnica depurada. Sus hombros de acero dominan el volumen acuático con eficiencia."
    if categoria == 'Bicicleta':
        return f"Potencia pura sobre ruedas. {nombre} devoró la ruta con {tiempo}. Demuestra ser el gran motor del equipo en la carretera, manteniendo promedios que intimidan."
    if categoria == 'Trote':
        return f"Resistencia inalcanzable. {nombre} domina la carrera a pie con {tiempo}. Su fase de carrera es soberbia, demostrando una resiliencia cardiovascular de otro planeta."
    
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
            parsed_data.append({'Clasificación': rank_counter, 'Deportista': name, 'Tiempo Total': to_hhmmss(t_m), 'Actividades': acts, 'Natación': to_hhmmss(s_m), 'Bicicleta': to_hhmmss(b_m), 'Trote': to_hhmmss(r_m), 'CV': cv, 'T_Mins': t_m, 'N_Mins': s_m, 'B_Mins': b_m, 'R_Mins': r_m})
            rank_counter += 1
        except: pass
    return pd.DataFrame(parsed_data)

# --- NUEVA FUNCIONALIDAD: PARSER DE CAPTURA (OCR) ---
def parse_ocr_data(ocr_text):
    distancia_podio = []
    larga_podio = []
    lines = ocr_text.strip().split('\n')
    for line in lines:
        parts = line.split(';')
        if len(parts) >= 6:
            # Captura Distancia Total (Nombre en índice 2, Km en índice 3)
            distancia_podio.append({'nombre': parts[2].strip(), 'valor': parts[3].strip()})
            # Captura Salida más Larga (Nombre en índice 4, Tiempo en índice 5)
            larga_podio.append({'nombre': parts[4].strip(), 'valor': parts[5].strip()})
    return distancia_podio[:3], larga_podio[:3]

# --- WORD GENERATOR (ESTRUCTURA BLINDADA) ---
def crear_tabla(doc, df, cols):
    table = doc.add_table(rows=1, cols=len(cols))
    table.style = 'Light Grid Accent 1'
    for i, c in enumerate(cols): table.rows[0].cells[i].text = c
    for _, row in df.iterrows():
        row_cells = table.add_row().cells
        for i, c in enumerate(cols): row_cells[i].text = str(row[c])

def generar_word(df, sem, dist_p, larg_p):
    doc = Document()
    doc.add_heading(f'Reporte Semanal Club Tym Triatlón - Semana {sem}', 0).alignment = 1
    doc.add_paragraph('"(La semana de la simetría perfecta y el retorno del volumen)"').alignment = 1
    
    df_c = df[df['CV'] != 'NC'].copy()
    df_c['CV_n'] = df_c['CV'].astype(float)
    
    doc.add_heading('🔍 Resumen General', level=2)
    t_m = df['T_Mins'].sum()
    doc.add_paragraph(f'Total deportistas: {len(df)}\nTriatletas completos: {len(df_c)}\nHoras totales: {int(t_m//60)}h {int(t_m%60)}m')
    
    # Gráfico
    s_d = df['N_Mins'].sum() + df['B_Mins'].sum() + df['R_Mins'].sum()
    p_n, p_b, p_t = (df['N_Mins'].sum()/s_d*100), (df['B_Mins'].sum()/s_d*100), (df['R_Mins'].sum()/s_d*100)
    fig, ax = plt.subplots(figsize=(4,4))
    ax.pie([p_n, p_b, p_t], labels=['Nat', 'Bici', 'Trote'], autopct='%1.1f%%', colors=['#1E90FF', '#32CD32', '#FF4500'])
    img_s = io.BytesIO()
    plt.savefig(img_s, format='png')
    doc.add_paragraph().add_run().add_picture(img_s, width=Inches(3))
    
    # TOP 5 COMPLETOS
    doc.add_heading('🏅 TOP 5 TRIATLETAS COMPLETOS', level=2)
    t5 = df_c.sort_values('T_Mins', ascending=False).head(5)
    t5['#'] = range(1, 6)
    crear_tabla(doc, t5, ['#', 'Deportista', 'Tiempo Total', 'Actividades', 'Natación', 'Bicicleta', 'Trote'])
    for i, r in t5.iterrows():
        doc.add_paragraph(f"{r['#']}. {r['Deportista']}").bold = True
        doc.add_paragraph(f"{r['Tiempo Total']} | {r['Actividades']} act. | {r['Bicicleta']} bici").italic = True
        doc.add_paragraph(generar_comentario(r, t5, 'Completos', r['#']))

    # TOP 5 BALANCEADOS
    doc.add_heading('⚖️ TOP 5 TRIATLETAS MÁS BALANCEADOS', level=2)
    b5 = df_c.sort_values('CV_n').head(5)
    b5['#'] = range(1, 6)
    crear_tabla(doc, b5, ['#', 'Deportista', 'CV', 'Natación', 'Bicicleta', 'Trote'])
    for i, r in b5.iterrows():
        doc.add_paragraph(f"{r['#']}. {r['Deportista']} (CV: {r['CV']})").bold = True
        doc.add_paragraph(generar_comentario(r, b5, 'CV', r['#']))

    # TOP 15 GENERAL
    doc.add_heading('🥇 TOP 15 TIEMPO TOTAL GENERAL', level=2)
    t15 = df.sort_values('T_Mins', ascending=False).head(15)
    t15['#'] = range(1, 16)
    crear_tabla(doc, t15, ['#', 'Deportista', 'Tiempo Total', 'Natación', 'Bicicleta', 'Trote'])

    # OTRAS CATEGORÍAS (Llenado automático con OCR)
    doc.add_heading('🌟 OTRAS CATEGORÍAS DESTACADAS', level=2)
    doc.add_heading('🔄 MAYOR FRECUENCIA (ACTIVIDADES TOTALES)', level=3)
    for i, r in df.sort_values('Actividades', ascending=False).head(3).iterrows():
        doc.add_paragraph(f"{r['Deportista']} ({r['Actividades']} sesiones)")
        
    doc.add_heading('📏 PODIO DISTANCIA TOTAL', level=3)
    for i, p in enumerate(dist_p):
        doc.add_paragraph(f"{i+1}. {p['nombre']} ({p['valor']} km)")

    doc.add_heading('⏱️ PODIO ACTIVIDAD MÁS LARGA', level=3)
    for i, p in enumerate(larg_p):
        doc.add_paragraph(f"{i+1}. {p['nombre']} ({p['valor']})")

    # DISCIPLINAS (Una por página - BLINDADO)
    for disc, col, icono, m_col in [('NATACIÓN', 'Natación', '🏊‍♂️', 'N_Mins'), ('CICLISMO', 'Bicicleta', '🚴', 'B_Mins'), ('TROTE', 'Trote', '🏃‍♂️', 'R_Mins')]:
        doc.add_page_break()
        doc.add_heading(f'{icono} TOP 15 {disc}', level=1)
        d15 = df[df[m_col]>0].sort_values(m_col, ascending=False).head(15)
        d15['#'] = range(1, len(d15)+1)
        crear_tabla(doc, d15, ['#', 'Deportista', col, 'Tiempo Total'])
        doc.add_heading('Análisis del Podio:', level=3)
        for i, r in d15.head(3).iterrows():
            doc.add_paragraph(f"{'🥇' if r['#']==1 else '🥈' if r['#']==2 else '🥉'} {r['Deportista']} ({r[col]}): {generar_comentario(r, d15, col, r['#'])}")

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# --- UI ---
sem = st.text_input("Número de Semana:", "08")
raw = st.text_area("1. Pega los datos de Tiempo Total:")
ocr_input = st.text_area("2. Pega los datos traducidos de la Captura (OCR):")

if st.button("Generar Reporte Final"):
    if raw.strip() and ocr_input.strip():
        df_semana = parse_raw_data(raw)
        dist_p, larg_p = parse_ocr_data(ocr_input)
        
        st.download_button("📄 DESCARGAR REPORTE OFICIAL", generar_word(df_semana, sem, dist_p, larg_p), f"Reporte_TYM_Sem_{sem}.docx")
