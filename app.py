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

st.title("🏆 Generador Oficial de Reportes - TYM")

HISTORICO_FILE = "historico_tym.csv"

# --- FUNCIONES DE TIEMPO ---
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

# --- MOTOR DE INTELIGENCIA DE COMENTARIOS ---
def obtener_comentario(row, df_contexto, categoria, pos):
    nombre = row['Deportista']
    tiempo = row.get(categoria, row.get('Tiempo Total'))
    t_mins = row.get('T_Mins', 0)
    b_mins = row.get('B_Mins', 0)
    acts = row.get('Actividades', 0)
    
    if categoria == 'Completos':
        if pos == 1:
            margen = t_mins - df_contexto.iloc[1]['T_Mins'] if len(df_contexto) > 1 else 0
            res = f"Dominio absoluto. {nombre} se consolida en la cima "
            res += f"con una ventaja de {int(margen//60)}h sobre su escolta. "
            res += "Su constancia en las tres disciplinas fue la llave del éxito."
            return res
        return f"Sólida posición de {nombre}. Mantiene un volumen de {tiempo} que lo sitúa en la élite del club gracias a su disciplina diaria."

    if categoria == 'CV':
        cv_val = float(row['CV'])
        if pos == 1: return f"¡Récord de simetría! {nombre} logra un balance casi perfecto ({cv_val}), demostrando una planificación milimétrica en sus cargas."
        return f"Equilibrio fantástico. {nombre} demuestra que la eficiencia es tan importante como el volumen, manteniendo sus disciplinas en perfecta armonía."

    if categoria == 'Natación':
        return f"Fuerza pura en el agua. {nombre} registra {tiempo}, liderando el podio con una técnica que marca la diferencia en el volumen acuático."
    if categoria == 'Bicicleta':
        return f"Potencia pura sobre ruedas. {nombre} devoró la ruta con {tiempo}, demostrando que es el gran motor del equipo en esta disciplina."
    if categoria == 'Trote':
        return f"Resistencia inalcanzable. {nombre} domina el asfalto con {tiempo}, cerrando la semana con una fase de carrera soberbia."
    
    return "Desempeño destacado y consistente."

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

def procesar_historico(df, sem):
    col = f"Sem {sem}"
    df_n = df[['Deportista', 'T_Mins']].rename(columns={'Deportista': 'Nombre', 'T_Mins': col})
    if os.path.exists(HISTORICO_FILE):
        df_h = pd.read_csv(HISTORICO_FILE)
        if col in df_h.columns: df_h = df_h.drop(columns=[col])
        df_h = pd.merge(df_h, df_n, on='Nombre', how='outer').fillna(0)
    else: df_h = df_n
    df_h['Tiempo Acumulado'] = df_h[[c for c in df_h.columns if 'Sem' in c]].sum(axis=1)
    df_h.to_csv(HISTORICO_FILE, index=False)

def crear_tabla(doc, df, cols):
    table = doc.add_table(rows=1, cols=len(cols))
    table.style = 'Light Grid Accent 1'
    for i, c in enumerate(cols): table.rows[0].cells[i].text = c
    for _, row in df.iterrows():
        row_cells = table.add_row().cells
        for i, c in enumerate(cols): row_cells[i].text = str(row[c])

def generar_word(df, sem):
    doc = Document()
    doc.add_heading(f'Reporte Semanal Club Tym Triatlón - Semana {sem}', 0).alignment = 1
    doc.add_paragraph('"(La semana de la simetría perfecta y el retorno del volumen)"').alignment = 1
    
    df_c = df[df['CV'] != 'NC'].copy()
    df_c['CV_n'] = df_c['CV'].astype(float)
    
    # Resumen e Imagen
    doc.add_heading('🔍 Resumen General', level=2)
    t_m = df['T_Mins'].sum()
    doc.add_paragraph(f'Total deportistas: {len(df)}\nTriatletas completos: {len(df_c)}\nHoras totales: {int(t_m//60)}h {int(t_m%60)}m')
    
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
        doc.add_paragraph(obtener_comentario(r, t5, 'Completos', r['#']))

    # TOP 5 BALANCEADOS
    doc.add_heading('⚖️ TOP 5 TRIATLETAS MÁS BALANCEADOS', level=2)
    b5 = df_c.sort_values('CV_n').head(5)
    b5['#'] = range(1, 6)
    crear_tabla(doc, b5, ['#', 'Deportista', 'CV', 'Natación', 'Bicicleta', 'Trote'])
    for i, r in b5.iterrows():
        doc.add_paragraph(f"{r['#']}. {r['Deportista']} (CV: {r['CV']})").bold = True
        doc.add_paragraph(obtener_comentario(r, b5, 'CV', r['#']))

    # TOP 15 GENERAL
    doc.add_heading('🥇 TOP 15 TIEMPO TOTAL GENERAL', level=2)
    t15 = df.sort_values('T_Mins', ascending=False).head(15)
    t15['#'] = range(1, 16)
    crear_tabla(doc, t15, ['#', 'Deportista', 'Tiempo Total', 'Natación', 'Bicicleta', 'Trote'])

    # DISCIPLINAS (Una por página)
    for disc, col, icono in [('NATACIÓN', 'Natación', '🏊‍♂️'), ('CICLISMO', 'Bicicleta', '🚴'), ('TROTE', 'Trote', '🏃‍♂️')]:
        doc.add_page_break()
        doc.add_heading(f'{icono} TOP 15 {disc}', level=1)
        m_col = 'N_Mins' if 'NAT' in disc else 'B_Mins' if 'CIC' in disc else 'R_Mins'
        d15 = df[df[m_col]>0].sort_values(m_col, ascending=False).head(15)
        d15['#'] = range(1, len(d15)+1)
        crear_tabla(doc, d15, ['#', 'Deportista', col, 'Tiempo Total'])
        doc.add_heading('Análisis del Podio:', level=3)
        for i, r in d15.head(3).iterrows():
            doc.add_paragraph(f"{'🥇' if r['#']==1 else '🥈' if r['#']==2 else '🥉'} {r['Deportista']} ({r[col]}): {obtener_comentario(r, d15, col, r['#'])}")

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# --- UI ---
sem = st.text_input("Semana:", "08")
raw = st.text_area("Datos:")
if st.button("Generar Reporte"):
    df = parse_raw_data(raw)
    if not df.empty:
        procesar_historico(df, sem)
        st.download_button("📄 DESCARGAR REPORTE OFICIAL", generar_word(df, sem), f"Reporte_TYM_Sem_{sem}.docx")
