import streamlit as st
import pandas as pd
import numpy as np
import re
import matplotlib.pyplot as plt
import io
import os
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Inteligencia TYM", page_icon="🏆", layout="wide")

st.title("🏆 Plataforma de Inteligencia Deportiva - TYM")
st.markdown("Generación automatizada de reportes con formato oficial Calibri.")

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
    
    if categoria in ['Completos', 'General']:
        if pos == 1:
            return f"Dominio absoluto de {nombre}. Se consolida en la cima del club con un volumen total envidiable. Su capacidad para sostener sesiones de alta intensidad demuestra una preparación de élite y una disciplina inquebrantable."
        elif pos == 2:
            return f"Una semana brillante para {nombre}. Se queda con la plata manteniendo una presión constante sobre el líder. Su solidez fue el motor que lo mantuvo en la parte más alta de la tabla."
        else:
            return f"{nombre} cierra el podio con un rendimiento muy equilibrado. Demuestra que para estar en el grupo de avanzada no se puede regalar nada, sumando minutos de calidad."

    if categoria == 'CV':
        cv_val = float(row['CV'])
        if pos == 1:
            return f"¡El reloj suizo del club! {nombre} logra una simetría de {cv_val}, una cifra casi perfecta. Requiere una planificación quirúrgica y un control total de las cargas."
        return f"Excelente balance para {nombre}. Mantiene sus disciplinas en una armonía poco común, progresando de forma integral."

    tiempo = row.get(categoria, "")
    if categoria == 'Natación': return f"Fuerza pura en el agua. {nombre} registra {tiempo}, liderando el podio con técnica depurada."
    if categoria == 'Bicicleta': return f"Potencia pura sobre ruedas. {nombre} devoró la ruta con {tiempo}, demostrando ser el motor del equipo."
    if categoria == 'Trote': return f"Resistencia inalcanzable. {nombre} domina el asfalto con {tiempo} y una fase de carrera soberbia."
    
    return "Desempeño destacado."

# --- PARSER DE DATOS ---
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

def parse_ocr_data(ocr_text):
    distancia_podio, larga_podio = [], []
    for line in ocr_text.strip().split('\n'):
        parts = line.split(';')
        if len(parts) >= 6:
            distancia_podio.append({'nombre': parts[2].strip(), 'valor': parts[3].strip()})
            larga_podio.append({'nombre': parts[4].strip(), 'valor': parts[5].strip()})
    return distancia_podio[:3], larga_podio[:3]

# --- WORD GENERATOR (ESTILO CALIBRI + ESPACIADO) ---
def aplicar_fuente(parrafo, size, negrita=False):
    run = parrafo.runs[0] if parrafo.runs else parrafo.add_run()
    run.font.name = 'Calibri'
    run.font.size = Pt(size)
    run.bold = negrita

def crear_tabla_profesional(doc, df, cols):
    table = doc.add_table(rows=1, cols=len(cols))
    table.style = 'Light Grid Accent 1'
    table.autofit = False
    
    # Configurar anchos (Damos más espacio a la columna Nombre)
    for i, col_name in enumerate(cols):
        cell = table.rows[0].cells[i]
        cell.text = col_name
        par = cell.paragraphs[0]
        aplicar_fuente(par, 11, True)
        if col_name == 'Deportista': cell.width = Inches(2.2)
        else: cell.width = Inches(0.8)

    for _, row in df.iterrows():
        row_cells = table.add_row().cells
        for i, col_name in enumerate(cols):
            row_cells[i].text = str(row[col_name])
            aplicar_fuente(row_cells[i].paragraphs[0], 11)
    
    doc.add_paragraph() # Renglón en blanco tras tabla

def generar_word(df, sem, dist_p, larg_p):
    doc = Document()
    # Estilo base
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(11)

    # --- PÁGINA 1 ---
    t = doc.add_heading(f'Reporte Semanal Club Tym Triatlón - Semana {sem}', 0)
    aplicar_fuente(t, 15, True)
    doc.add_paragraph() 
    
    p_intro = doc.add_paragraph('"(La semana de la simetría perfecta y el retorno del volumen)"')
    p_intro.alignment = WD_ALIGN_PARAGRAPH.CENTER
    aplicar_fuente(p_intro, 11, True)
    doc.add_paragraph()

    # Resumen
    h_res = doc.add_heading('🔍 Resumen General', level=2)
    aplicar_fuente(h_res, 15, True)
    doc.add_paragraph()
    
    t_m = df['T_Mins'].sum()
    df_c = df[df['CV'] != 'NC'].copy()
    res_text = f'Total deportistas: {len(df)}\nTriatletas completos: {len(df_c)}\nHoras totales: {int(t_m//60)}h {int(t_m%60)}m'
    p_res = doc.add_paragraph(res_text)
    aplicar_fuente(p_res, 11)

    # Gráfico
    s_d = df['N_Mins'].sum() + df['B_Mins'].sum() + df['R_Mins'].sum()
    fig, ax = plt.subplots(figsize=(4,4))
    ax.pie([df['N_Mins'].sum(), df['B_Mins'].sum(), df['R_Mins'].sum()], labels=['Nat', 'Bici', 'Trote'], autopct='%1.1f%%', colors=['#1E90FF', '#32CD32', '#FF4500'])
    img_s = io.BytesIO()
    plt.savefig(img_s, format='png')
    doc.add_paragraph().add_run().add_picture(img_s, width=Inches(3))
    plt.close(fig)

    # COMPLETOS
    h_comp = doc.add_heading('🏅 TOP 5 TRIATLETAS COMPLETOS', level=2)
    aplicar_fuente(h_comp, 15, True)
    doc.add_paragraph()
    t5 = df_c.sort_values('T_Mins', ascending=False).head(5)
    t5['#'] = range(1, 6)
    crear_tabla_profesional(doc, t5, ['#', 'Deportista', 'Tiempo Total', 'Natación', 'Bicicleta', 'Trote'])
    
    s_an = doc.add_paragraph('Análisis Ejecutivo')
    aplicar_fuente(s_an, 13, True)
    for i, r in t5.iterrows():
        p = doc.add_paragraph(f"{r['#']}. {r['Deportista']}")
        aplicar_fuente(p, 11, True)
        p_com = doc.add_paragraph(generar_comentario(r, t5, 'Completos', r['#']))
        aplicar_fuente(p_com, 11)

    # BALANCEADOS
    h_bal = doc.add_heading('⚖️ TOP 5 TRIATLETAS MÁS BALANCEADOS', level=2)
    aplicar_fuente(h_bal, 15, True)
    doc.add_paragraph()
    df_c['CV_n'] = df_c['CV'].astype(float)
    b5 = df_c.sort_values('CV_n').head(5)
    b5['#'] = range(1, 6)
    crear_tabla_profesional(doc, b5, ['#', 'Deportista', 'CV', 'Natación', 'Bicicleta', 'Trote'])
    
    s_anb = doc.add_paragraph('Análisis de Simetría')
    aplicar_fuente(s_anb, 13, True)
    for i, r in b5.iterrows():
        p = doc.add_paragraph(f"{r['#']}. {r['Deportista']} (CV: {r['CV']})")
        aplicar_fuente(p, 11, True)
        p_com = doc.add_paragraph(generar_comentario(r, b5, 'CV', r['#']))
        aplicar_fuente(p_com, 11)

    # --- PÁGINA 2: CLASIFICACIÓN GENERAL ---
    doc.add_page_break()
    h_gen = doc.add_heading('🥇 TOP 15 TIEMPO TOTAL GENERAL', level=1)
    aplicar_fuente(h_gen, 15, True)
    doc.add_paragraph()
    t15 = df.sort_values('T_Mins', ascending=False).head(15)
    t15['#'] = range(1, 16)
    crear_tabla_profesional(doc, t15, ['#', 'Deportista', 'Tiempo Total', 'Natación', 'Bicicleta', 'Trote'])
    
    s_ang = doc.add_paragraph('Análisis del Podio General')
    aplicar_fuente(s_ang, 13, True)
    for i, r in t15.head(3).iterrows():
        p = doc.add_paragraph(f"{'🥇' if r['#']==1 else '🥈' if r['#']==2 else '🥉'} {r['Deportista']}")
        aplicar_fuente(p, 11, True)
        p_com = doc.add_paragraph(generar_comentario(r, t15, 'General', r['#']))
        aplicar_fuente(p_com, 11)

    # --- DISCIPLINAS ---
    for disc, col, icono, m_col in [('NATACIÓN', 'Natación', '🏊‍♂️', 'N_Mins'), ('CICLISMO', 'Bicicleta', '🚴', 'B_Mins'), ('TROTE', 'Trote', '🏃‍♂️', 'R_Mins')]:
        doc.add_page_break()
        h_d = doc.add_heading(f'{icono} TOP 15 {disc}', level=1)
        aplicar_fuente(h_d, 15, True)
        doc.add_paragraph()
        d15 = df[df[m_col]>0].sort_values(m_col, ascending=False).head(15)
        d15['#'] = range(1, len(d15)+1)
        crear_tabla_profesional(doc, d15, ['#', 'Deportista', col, 'Tiempo Total'])
        
        s_and = doc.add_paragraph('Análisis del Podio')
        aplicar_fuente(s_and, 13, True)
        for i, r in d15.head(3).iterrows():
            p = doc.add_paragraph(f"{'🥇' if r['#']==1 else '🥈' if r['#']==2 else '🥉'} {r['Deportista']} ({r[col]})")
            aplicar_fuente(p, 11, True)
            p_com = doc.add_paragraph(generar_comentario(r, d15, col, r['#']))
            aplicar_fuente(p_com, 11)

    # --- FINAL ---
    doc.add_page_break()
    h_dis = doc.add_heading('📏 PODIO DISTANCIA TOTAL', level=1)
    aplicar_fuente(h_dis, 15, True)
    doc.add_paragraph()
    for i, p in enumerate(dist_p):
        par = doc.add_paragraph(f"{i+1}. {p['nombre']} ({p['valor']} km)")
        aplicar_fuente(par, 11)

    h_lar = doc.add_heading('⏱️ PODIO ACTIVIDAD MÁS LARGA', level=1)
    aplicar_fuente(h_lar, 15, True)
    doc.add_paragraph()
    for i, p in enumerate(larg_p):
        par = doc.add_paragraph(f"{i+1}. {p['nombre']} ({p['valor']})")
        aplicar_fuente(par, 11)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# --- UI ---
sem_i = st.text_input("Número de Semana:", "08")
raw_i = st.text_area("1. Datos de Tiempo Total:")
ocr_i = st.text_area("2. Datos de Captura (OCR):")

if st.button("Generar Reporte Final Calibri"):
    if raw_i.strip() and ocr_i.strip():
        df_s = parse_raw_data(raw_i)
        dist_p, larg_p = parse_ocr_data(ocr_i)
        st.download_button("📄 DESCARGAR REPORTE CALIBRI", generar_word(df_s, sem_i, dist_p, larg_p), f"Reporte_TYM_Sem_{sem_i}.docx")
