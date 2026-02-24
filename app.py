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

# --- 1. CONFIGURACIÓN (BLINDADO) ---
st.set_page_config(page_title="Plataforma TYM", page_icon="🏆", layout="wide")
st.title("🏆 Gestión de Reportes y Estadísticas - Club TYM")

# --- 2. UTILIDADES DE TIEMPO ---
def to_mins(t_str):
    if pd.isna(t_str) or str(t_str).strip() in ['--:--', '0', '']: return 0
    try:
        if ':' in str(t_str):
            parts = str(t_str).split(':')
            return int(parts[0]) * 60 + int(parts[1])
        hm = re.search(r'(\d+)h', str(t_str))
        mm = re.search(r'(\d+)min', str(t_str))
        h = int(hm.group(1)) if hm else 0
        m = int(mm.group(1)) if mm else 0
        return h * 60 + m
    except: return 0

def to_hhmmss(mins):
    h, m = int(mins // 60), int(mins % 60)
    return f"{h:02d}:{m:02d}:00"

# --- 3. COMENTARIOS INTELIGENTES ---
def generar_comentario(row, df_contexto, categoria, pos):
    nombre = row['Deportista']
    if categoria in ['Completos', 'General']:
        if pos == 1: return f"Dominio absoluto de {nombre}. Se consolida en la cima con un volumen total envidiable, demostrando una preparación de élite."
        elif pos == 2: return f"Semana brillante para {nombre}. Presión constante sobre el líder y solidez en todas sus sesiones."
        else: return f"{nombre} cierra el podio con regularidad, sumando minutos de calidad en las tres disciplinas."
    if categoria == 'CV':
        cv = row.get('CV', 0)
        return f"¡Reloj suizo! {nombre} logra una simetría de {cv}, demostrando un control total de sus cargas de entrenamiento."
    return f"Rendimiento destacado de {nombre} en {categoria}, liderando el podio con técnica y consistencia."

# --- 4. PARSER DE DATOS ---
def parse_raw_data(raw_text):
    data = []
    rank = 1
    raw_text = raw_text.replace('\xa0', ' ')
    for line in raw_text.strip().split('\n'):
        if not line or 'Deportista' in line: continue
        try:
            times = re.findall(r'(\d+h\s*\d*min|\d+h|\d+min|--:--)', line)
            if not times: continue
            name = re.sub(r'^\d+\s*', '', line[:line.find(times[0])]).strip()
            t_m = to_mins(times[0])
            s_m = to_mins(times[1]) if len(times) > 1 else 0
            b_m = to_mins(times[2]) if len(times) > 2 else 0
            r_m = to_mins(times[3]) if len(times) > 3 else 0
            cv = "NC" if 0 in [s_m, b_m, r_m] else round(np.std([s_m, b_m, r_m])/np.mean([s_m, b_m, r_m]), 4)
            data.append({'#': rank, 'Deportista': name, 'Tiempo Total': to_hhmmss(t_m), 'Natación': to_hhmmss(s_m), 'Bicicleta': to_hhmmss(b_m), 'Trote': to_hhmmss(r_m), 'CV': cv, 'T_Mins': t_m, 'N_Mins': s_m, 'B_Mins': b_m, 'R_Mins': r_m})
            rank += 1
        except: pass
    return pd.DataFrame(data)

# --- 5. ACTUALIZADOR DE EXCEL (REPARADO Y BLINDADO) ---
def actualizar_excel_maestro(archivo_maestro, df_actual, num_sem):
    xls = pd.ExcelFile(archivo_maestro)
    output = io.BytesIO()
    col_nueva = f"Sem {num_sem.strip()}"
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for hoja in xls.sheet_names:
            df_h = pd.read_excel(xls, sheet_name=hoja)
            # Eliminar columnas basura si existen
            if 'Sem 51' in df_h.columns: df_h.drop(columns=['Sem 51'], inplace=True, errors='ignore')
            if 'Sem 52' in df_h.columns: df_h.drop(columns=['Sem 52'], inplace=True, errors='ignore')
            
            col_key = next((c for c in df_h.columns if str(c).lower() in ['nombre', 'deportista']), df_h.columns[0])
            mapeo = {'Tiempo Total': 'Tiempo Total', 'Natación': 'Natación', 'Ciclismo': 'Bicicleta', 'Trote': 'Trote', 'CV': 'CV'}
            col_dato = mapeo.get(hoja)
            
            if col_dato and col_nueva in df_h.columns:
                df_upd = df_actual[['Deportista', col_dato]].copy()
                df_upd['Key'] = df_upd['Deportista'].str.strip().str.upper()
                df_h['Key'] = df_h[col_key].astype(str).str.strip().str.upper()
                
                for idx, row in df_h.iterrows():
                    match = df_upd[df_upd['Key'] == row['Key']]
                    if not match.empty:
                        df_h.at[idx, col_nueva] = match.iloc[0][col_dato]
                df_h.drop(columns=['Key'], inplace=True)
            
            df_h.to_excel(writer, sheet_name=hoja, index=False)
    return output.getvalue()

# --- 6. GENERADOR DE WORD (CALIBRI 20/15/13/11) ---
def aplicar_estilo(p, size, bold=False, center=False):
    run = p.runs[0] if p.runs else p.add_run()
    run.font.name = 'Calibri'; run.font.size = Pt(size); run.bold = bold
    if center: p.alignment = 1

def crear_tabla(doc, df, cols):
    table = doc.add_table(rows=1, cols=len(cols))
    table.style = 'Light Grid Accent 1'; table.alignment = 1; table.autofit = False
    anchos = {'#': 0.3, 'Deportista': 2.8, 'Tiempo Total': 0.8, 'Natación': 0.8, 'Bicicleta': 0.8, 'Trote': 0.8, 'CV': 0.6}
    for i, c in enumerate(cols):
        cell = table.rows[0].cells[i]; cell.text = c; cell.width = Inches(anchos.get(c, 0.8))
        aplicar_estilo(cell.paragraphs[0], 9, True, True)
    for _, row in df.iterrows():
        row_cells = table.add_row().cells
        for i, c in enumerate(cols):
            row_cells[i].text = str(row[c]); aplicar_estilo(row_cells[i].paragraphs[0], 9, False, c != 'Deportista')
    doc.add_paragraph()

def generar_word(df, sem):
    doc = Document()
    # Título Principal
    t = doc.add_heading(f'Reporte Semanal Club Tym Triatlón - Semana {sem}', 0)
    aplicar_estilo(t, 20, True, True); doc.add_paragraph()
    
    # Resumen y Gráfico
    h1 = doc.add_heading('🔍 Resumen General', level=2); aplicar_estilo(h1, 15, True); doc.add_paragraph()
    df_c = df[df['CV'] != 'NC'].copy()
    p_res = doc.add_paragraph(f"Total deportistas: {len(df)}\nTriatletas completos: {len(df_c)}"); aplicar_estilo(p_res, 11)
    
    fig, ax = plt.subplots(figsize=(4,4))
    ax.pie([df['N_Mins'].sum(), df['B_Mins'].sum(), df['R_Mins'].sum()], labels=['Nat', 'Bici', 'Trote'], autopct='%1.1f%%', colors=['#1E90FF', '#32CD32', '#FF4500'])
    img = io.BytesIO(); plt.savefig(img, format='png', bbox_inches='tight'); plt.close(fig)
    p_img = doc.add_paragraph(); p_img.alignment = 1; p_img.add_run().add_picture(img, width=Inches(3.5))

    # Top 5 Completos y Balanceados
    for tit, d, cat in [('🏅 TOP 5 TRIATLETAS COMPLETOS', df_c.sort_values('T_Mins', ascending=False).head(5), 'Completos'), ('⚖️ TOP 5 TRIATLETAS MÁS BALANCEADOS', df_c.sort_values('CV').head(5), 'CV')]:
        h = doc.add_heading(tit, level=2); aplicar_estilo(h, 15, True); doc.add_paragraph()
        d['#'] = range(1, 6); crear_tabla(doc, d, ['#', 'Deportista', 'Tiempo Total' if cat=='Completos' else 'CV', 'Natación', 'Bicicleta', 'Trote'])
        doc.add_paragraph('Análisis:'); [doc.add_paragraph(generar_comentario(r, d, cat, r['#'])) for _, r in d.iterrows()]

    # Disciplinas Top 15
    for tit, icono, m_col, col_t in [('TIEMPO GENERAL', '🥇', 'T_Mins', 'Tiempo Total'), ('NATACIÓN', '🏊‍♂️', 'N_Mins', 'Natación'), ('CICLISMO', '🚴', 'B_Mins', 'Bicicleta'), ('TROTE', '🏃‍♂️', 'R_Mins', 'Trote')]:
        doc.add_page_break(); h = doc.add_heading(f'{icono} TOP 15 {tit}', level=1); aplicar_estilo(h, 15, True); doc.add_paragraph()
        d15 = df[df[m_col]>0].sort_values(m_col, ascending=False).head(15); d15['#'] = range(1, len(d15)+1)
        crear_tabla(doc, d15, ['#', 'Deportista', col_t, 'Tiempo Total'] if tit != 'TIEMPO GENERAL' else ['#', 'Deportista', 'Tiempo Total', 'Natación', 'Bicicleta', 'Trote'])

    buf = io.BytesIO(); doc.save(buf); buf.seek(0); return buf

# --- 7. INTERFAZ ---
st.sidebar.header("📁 Archivo Maestro")
archivo_maestro = st.sidebar.file_uploader("00 Estadísticas TYM (xlsx)", type=["xlsx"])
sem_num = st.text_input("Semana a procesar:", "08")
raw_data = st.text_area("Datos de Tiempo Total:")

if st.button("🚀 PROCESAR SEMANA"):
    if raw_data.strip() and archivo_maestro:
        df_semana = parse_raw_data(raw_data)
        st.success("¡Semana procesada!")
        col1, col2 = st.columns(2)
        col1.download_button("📄 DESCARGAR WORD", generar_word(df_semana, sem_num), f"Reporte_TYM_{sem_num}.docx")
        col2.download_button("📊 EXCEL ACTUALIZADO", actualizar_excel_maestro(archivo_maestro, df_semana, sem_num), "00_Estadisticas_TYM_Actualizado.xlsx")
