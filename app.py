import streamlit as st
import pandas as pd
import numpy as np
import re
import matplotlib.pyplot as plt
import io
import unicodedata
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

# --- 1. CONFIGURACIÓN (BLINDADO) ---
st.set_page_config(page_title="Plataforma TYM 2026", page_icon="🏆", layout="wide")
st.title("🏆 Gestión de Reportes y Estadísticas - Club TYM")

# --- 2. UTILIDADES DE TIEMPO Y LIMPIEZA ---
def clean_name(text):
    if not text: return ""
    text = str(text).strip().upper()
    return "".join(c for c in unicodedata.normalize('NFKD', text) if not unicodedata.combining(c))

def to_mins(t_str):
    if pd.isna(t_str) or str(t_str).strip() in ['--:--', '0', '', '00:00:00', 'NaN']: return 0
    try:
        t_str = str(t_str).strip()
        if ':' in t_str:
            parts = t_str.split(':')
            return int(parts[0]) * 60 + int(parts[1])
        hm = re.search(r'(\d+)h', t_str)
        mm = re.search(r'(\d+)min', t_str)
        return (int(hm.group(1)) if hm else 0) * 60 + (int(mm.group(1)) if mm else 0)
    except: return 0

def to_hhmmss(mins):
    return f"{int(mins // 60):02d}:{int(mins % 60):02d}:00"

# --- 3. MOTOR DE COMENTARIOS (FORMATO ORIGINAL) ---
def generar_comentario(row, cat, pos):
    nombre = row['Deportista']
    if cat in ['Completos', 'General']:
        if pos == 1: return f"Dominio absoluto de {nombre}. Se consolida en la cima con un volumen total envidiable, demostrando una preparación de élite."
        if pos == 2: return f"Semana brillante para {nombre}. Presión constante sobre el líder y solidez en todas sus sesiones."
        return f"{nombre} cierra el podio con regularidad, sumando minutos de calidad en las tres disciplinas."
    if cat == 'CV':
        return f"¡Reloj suizo! {nombre} logra una simetría de {row.get('CV', 0)}, demostrando un control total de sus cargas."
    return f"Rendimiento destacado de {nombre} en {cat}, liderando el podio con técnica y consistencia."

# --- 4. PARSER DE DATOS ---
def parse_raw_data(raw_text):
    data = []
    raw_text = raw_text.replace('\xa0', ' ')
    for line in raw_text.strip().split('\n'):
        if not line or 'Deportista' in line: continue
        try:
            times = re.findall(r'(\d+h\s*\d*min|\d+h|\d+min|--:--)', line)
            if not times: continue
            name = re.sub(r'^\d+\s*', '', line[:line.find(times[0])]).strip()
            t_m, s_m, b_m, r_m = to_mins(times[0]), to_mins(times[1] if len(times)>1 else 0), to_mins(times[2] if len(times)>2 else 0), to_mins(times[3] if len(times)>3 else 0)
            cv = "NC" if 0 in [s_m, b_m, r_m] else round(np.std([s_m, b_m, r_m])/np.mean([s_m, b_m, r_m]), 4)
            data.append({'Deportista': name, 'Tiempo Total': to_hhmmss(t_m), 'Natación': to_hhmmss(s_m), 'Bicicleta': to_hhmmss(b_m), 'Trote': to_hhmmss(r_m), 'CV': cv, 'T_Mins': t_m, 'N_Mins': s_m, 'B_Mins': b_m, 'R_Mins': r_m})
        except: pass
    return pd.DataFrame(data)

def parse_ocr_data(ocr_text):
    dist, larg = [], []
    for line in ocr_text.strip().split('\n'):
        p = line.split(';')
        if len(p) >= 6:
            dist.append({'nombre': p[2].strip(), 'valor': p[3].strip()})
            larg.append({'nombre': p[4].strip(), 'valor': p[5].strip()})
    return dist[:3], larg[:3]

# --- 5. ACTUALIZADOR DE EXCEL MAESTRO (RECONSTRUIDO) ---
def actualizar_excel_maestro(archivo_maestro, df_actual, num_sem):
    xls = pd.ExcelFile(archivo_maestro)
    output = io.BytesIO()
    col_nueva = f"Sem {num_sem.strip()}"
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for hoja in xls.sheet_names:
            df_h = pd.read_excel(xls, sheet_name=hoja)
            df_h = df_h.drop(columns=['Sem 51', 'Sem 52'], errors='ignore')
            
            # Identificar columna de nombres
            col_key = next((c for c in df_h.columns if str(c).lower() in ['nombre', 'deportista']), df_h.columns[0])
            mapeo = {'Tiempo Total': 'Tiempo Total', 'Natación': 'Natación', 'Ciclismo': 'Bicicleta', 'Trote': 'Trote', 'CV': 'CV'}
            col_dato = mapeo.get(hoja)
            
            if col_dato:
                # Asegurar que la columna nueva existe
                if col_nueva not in df_h.columns:
                    df_h[col_nueva] = '00:00:00' if hoja != 'CV' else 'NC'
                
                df_upd = df_actual[['Deportista', col_dato]].copy()
                df_upd['Match'] = df_upd['Deportista'].apply(clean_name)
                df_h['Match'] = df_h[col_key].apply(clean_name)
                
                updates = df_upd.set_index('Match')[col_dato].to_dict()
                # Solo actualiza la columna de la semana, manteniendo el resto del historial intacto
                df_h[col_nueva] = df_h['Match'].map(updates).fillna(df_h[col_nueva])
                df_h.drop(columns=['Match'], inplace=True)
            
            df_h.to_excel(writer, sheet_name=hoja, index=False)
    return output.getvalue()

# --- 6. GENERADOR DE WORD (FORMATO PROFESIONAL BLINDADO) ---
def aplicar_estilo(p, size, bold=False, center=False):
    run = p.runs[0] if p.runs else p.add_run()
    run.font.name = 'Calibri'; run.font.size = Pt(size); run.bold = bold
    if center: p.alignment = 1

def crear_tabla(doc, df, cols):
    table = doc.add_table(rows=1, cols=len(cols))
    table.style = 'Light Grid Accent 1'; table.alignment = 1; table.autofit = False
    anchos = {'#': 0.4, 'Deportista': 2.8, 'Tiempo Total': 0.7, 'Natación': 0.7, 'Bicicleta': 0.7, 'Trote': 0.7, 'CV': 0.6}
    for i, c in enumerate(cols):
        cell = table.rows[0].cells[i]; cell.text = c
        cell.width = Inches(anchos.get(c, 0.7))
        aplicar_estilo(cell.paragraphs[0], 9, True, True)
    for idx, row in df.iterrows():
        row_cells = table.add_row().cells
        for i, c in enumerate(cols):
            row_cells[i].text = str(row[c])
            row_cells[i].width = Inches(anchos.get(c, 0.7))
            aplicar_estilo(row_cells[i].paragraphs[0], 9, False, c != 'Deportista')
    doc.add_paragraph()

def generar_word(df, sem, dist_p, larg_p):
    doc = Document()
    # Título Principal (20)
    t = doc.add_heading(f'Reporte Semanal Club Tym Triatlón - Semana {sem}', 0)
    aplicar_estilo(t, 20, True, True); doc.add_paragraph()
    
    p_intro = doc.add_paragraph('"(La semana de la simetría perfecta y el retorno del volumen)"')
    aplicar_estilo(p_intro, 11, True, True); doc.add_paragraph()

    # Resumen (15)
    h1 = doc.add_heading('🔍 Resumen General', level=2); aplicar_estilo(h1, 15, True); doc.add_paragraph()
    df_c = df[df['CV'] != 'NC'].copy(); t_m = df['T_Mins'].sum()
    p_res = doc.add_paragraph(f"Total deportistas: {len(df)}\nTriatletas completos: {len(df_c)}\nHoras totales: {to_hhmmss(t_m)}"); aplicar_estilo(p_res, 11)
    
    fig, ax = plt.subplots(figsize=(4,4))
    ax.pie([df['N_Mins'].sum(), df['B_Mins'].sum(), df['R_Mins'].sum()], labels=['Nat', 'Bici', 'Trote'], autopct='%1.1f%%', colors=['#1E90FF', '#32CD32', '#FF4500'])
    img_s = io.BytesIO(); plt.savefig(img_s, format='png', bbox_inches='tight'); plt.close(fig)
    p_img = doc.add_paragraph(); p_img.alignment = 1; p_img.add_run().add_picture(img_s, width=Inches(3.5))

    # Podios Top 5
    for tit, d, cat in [('🏅 TOP 5 TRIATLETAS COMPLETOS', df_c.sort_values('T_Mins', ascending=False).head(5), 'Completos'), ('⚖️ TOP 5 TRIATLETAS MÁS BALANCEADOS', df_c.sort_values('CV', ascending=True).head(5), 'CV')]:
        h = doc.add_heading(tit, level=2); aplicar_estilo(h, 15, True); doc.add_paragraph()
        d = d.copy(); d['#'] = range(1, len(d)+1)
        crear_tabla(doc, d, ['#', 'Deportista', 'Tiempo Total' if cat=='Completos' else 'CV', 'Natación', 'Bicicleta', 'Trote'])
        h_an = doc.add_paragraph('Análisis:'); aplicar_estilo(h_an, 13, True)
        for _, r in d.iterrows(): doc.add_paragraph(f"{r['#']}. {generar_comentario(r, cat, r['#'])}")

    # Top 15 por Disciplina
    for tit, icono, m_col, col_t in [('TIEMPO GENERAL', '🥇', 'T_Mins', 'Tiempo Total'), ('NATACIÓN', '🏊‍♂️', 'N_Mins', 'Natación'), ('CICLISMO', '🚴', 'B_Mins', 'Bicicleta'), ('TROTE', '🏃‍♂️', 'R_Mins', 'Trote')]:
        doc.add_page_break(); h = doc.add_heading(f'{icono} TOP 15 {tit}', level=1); aplicar_estilo(h, 15, True); doc.add_paragraph()
        d15 = df[df[m_col]>0].sort_values(m_col, ascending=False).head(15).copy(); d15['#'] = range(1, len(d15)+1)
        crear_tabla(doc, d15, ['#', 'Deportista', col_t, 'Tiempo Total'] if tit != 'TIEMPO GENERAL' else ['#', 'Deportista', 'Tiempo Total', 'Natación', 'Bicicleta', 'Trote'])

    # Final OCR
    doc.add_page_break()
    h5 = doc.add_heading('📏 PODIO DISTANCIA TOTAL', level=1); aplicar_estilo(h5, 15, True); doc.add_paragraph()
    for i, p in enumerate(dist_p): doc.add_paragraph(f"{i+1}. {p['nombre']} ({p['valor']} km)")
    doc.add_paragraph()
    h6 = doc.add_heading('⏱️ PODIO ACTIVIDAD MÁS LARGA', level=1); aplicar_estilo(h6, 15, True); doc.add_paragraph()
    for i, p in enumerate(larg_p): doc.add_paragraph(f"{i+1}. {p['nombre']} ({p['valor']})")

    buf = io.BytesIO(); doc.save(buf); buf.seek(0); return buf

# --- 7. INTERFAZ ---
st.sidebar.header("📁 Gestión Histórica")
archivo_maestro = st.sidebar.file_uploader("Subir Archivo 00 Estadísticas (Excel)", type=["xlsx"])
sem_num = st.text_input("Semana a procesar (Ej: 08):", "08")
raw_data = st.text_area("1. Datos Tiempo Total:")
ocr_input = st.text_area("2. Datos OCR (Captura):")

if st.button("🚀 PROCESAR JORNADA"):
    if raw_data.strip() and ocr_input.strip() and archivo_maestro:
        df_sem = parse_raw_data(raw_data)
        dist_p, larg_p = parse_ocr_data(ocr_input)
        st.success("¡Datos procesados! Revisa los archivos abajo.")
        col1, col2 = st.columns(2)
        col1.download_button("📄 REPORTE WORD (Calibri)", generar_word(df_sem, sem_num, dist_p, larg_p), f"Reporte_TYM_{sem_num}.docx")
        col2.download_button("📊 EXCEL ACTUALIZADO", actualizar_excel_maestro(archivo_maestro, df_sem, sem_num), "00_Estadisticas_Actualizado.xlsx")
