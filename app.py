import streamlit as st
import pandas as pd
import numpy as np
import re
import plotly.express as px
import io
import os
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Reportes TYM", page_icon="🏆", layout="wide")

st.title("🏆 Plataforma de Reportes e Inteligencia - TYM")
st.markdown("Procesa la semana, actualiza el histórico y genera el reporte profesional.")

# --- ARCHIVO HISTÓRICO LOCAL ---
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

# --- GENERADOR DINÁMICO DE COMENTARIOS ---
def generar_comentario(posicion, nombre, tiempo_str, disciplina):
    horas = 0
    if "h" in tiempo_str or ":" in tiempo_str:
        if ":" in tiempo_str:
            horas = int(tiempo_str.split(":")[0])
        else:
            horas = int(re.search(r'(\d+)h', tiempo_str).group(1)) if re.search(r'(\d+)h', tiempo_str) else 0

    if posicion == 1:
        if horas >= 10: return f'Oro indiscutido. Una bestialidad de volumen con {tiempo_str}. {nombre} domina la disciplina aplastando los pedales/kilómetros y consolidando su posición como especialista absoluto de la semana.'
        else: return f'El primer lugar indiscutido de la semana. {nombre} marca el ritmo del equipo con {tiempo_str}, demostrando un nivel técnico y una constancia envidiable para llevarse el oro.'
    elif posicion == 2:
        return f'Plata muy sólida. {nombre} persigue la cima con {tiempo_str}, manteniendo una presión constante sobre el líder y asegurando puntos vitales en la clasificación general.'
    elif posicion == 3:
        return f'Bronce que vale oro. {nombre} cierra el podio de honor registrando {tiempo_str}, una demostración de resistencia pura que lo/la mete en la élite del club esta semana.'
    return f'Gran desempeño registrando {tiempo_str}.'

# --- LECTURA DE DATOS CRUDOS ---
def parse_raw_data(raw_text):
    parsed_data = []
    rank_counter = 1 
    raw_text = raw_text.replace('\xa0', ' ')
    
    for line in raw_text.strip().split('\n'):
        line = line.strip()
        if not line or 'Deportista' in line: continue
        
        try:
            match = re.search(r'(\d+h\s*\d*min|\d+min|--:--)', line)
            if not match: continue
            
            name_part = line[:match.start()].strip()
            name = re.sub(r'^\d+\s*', '', name_part) 
            
            rest = line[match.start():].strip()
            rest = re.sub(r'\d+$', '', rest).strip() 
            
            idx_min = rest.find('min')
            if idx_min != -1:
                total_time_str = rest[:idx_min+3]
                after_total = rest[idx_min+3:]
            else:
                total_time_str = rest
                after_total = ""
            
            acts, swim_time_str = 0, "--:--"
            m_dash = re.match(r'^(\d*)(--:--)(.*)', after_total)
            if m_dash:
                acts = int(m_dash.group(1)) if m_dash.group(1) else 0
                after_swim = m_dash.group(3)
            else:
                token_match = re.search(r'(h|min)', after_total)
                if token_match:
                    sep = token_match.start()
                    unit = token_match.group(1)
                    combined_nums = after_total[:sep]
                    if unit == 'h':
                        swim_val = combined_nums[-1]
                        try: acts = int(combined_nums[:-1])
                        except: acts = 0
                        rem = after_total[sep+1:]
                        min_match = re.match(r'^\s*(\d+)min', rem)
                        if min_match:
                            swim_time_str = f"{swim_val}h {min_match.group(1)}min"
                            after_swim = rem[min_match.end():]
                        else:
                            swim_time_str = f"{swim_val}h"
                            after_swim = rem
                    elif unit == 'min':
                        if len(combined_nums) > 2:
                            swim_val = combined_nums[-2:]
                            try: acts = int(combined_nums[:-2])
                            except: acts = 0
                        else:
                            swim_val, acts = combined_nums, 0 
                        swim_time_str = f"{swim_val}min"
                        after_swim = after_total[sep+3:]
                else:
                    after_swim = after_total
            
            times = re.findall(r'(\d+h\s*\d*min|\d+h|\d+min|--:--)', after_swim)
            bike_time_str = times[0] if len(times) > 0 else "--:--"
            run_time_str = times[1] if len(times) > 1 else "--:--"
            
            t_mins = to_mins(total_time_str)
            s_mins = to_mins(swim_time_str)
            b_mins = to_mins(bike_time_str)
            r_mins = to_mins(run_time_str)
            
            vals = [s_mins, b_mins, r_mins]
            cv = "NC" if 0 in vals else round(np.std(vals) / np.mean(vals), 4)

            parsed_data.append({
                'Clasificación': rank_counter,
                'Deportista': name,
                'Tiempo Total': to_hhmmss(t_mins),
                'Actividades': acts,
                'Natación': to_hhmmss(s_mins),
                'Bicicleta': to_hhmmss(b_mins),
                'Trote': to_hhmmss(r_mins),
                'CV': cv,
                'T_Mins': t_mins, 'N_Mins': s_mins, 'B_Mins': b_mins, 'R_Mins': r_mins
            })
            rank_counter += 1
        except Exception:
            pass
            
    return pd.DataFrame(parsed_data)

# --- ACTUALIZAR HISTÓRICO ---
def procesar_historico(df_semana, num_semana):
    col_semana = f"Sem {num_semana}"
    df_nueva = df_semana[['Deportista', 'T_Mins']].copy()
    df_nueva.rename(columns={'Deportista': 'Nombre', 'T_Mins': col_semana}, inplace=True)
    
    if os.path.exists(HISTORICO_FILE):
        df_hist = pd.read_csv(HISTORICO_FILE)
        # Unir datos
        df_hist = pd.merge(df_hist, df_nueva, on='Nombre', how='outer')
        df_hist[col_semana] = df_hist[col_semana].fillna(0)
    else:
        df_hist = df_nueva.copy()
        df_hist['Tiempo Acumulado (Mins)'] = 0

    # Recalcular acumulado total en minutos
    cols_semanas = [c for c in df_hist.columns if c.startswith('Sem ')]
    df_hist[cols_semanas] = df_hist[cols_semanas].fillna(0)
    df_hist['Tiempo Acumulado (Mins)'] = df_hist[cols_semanas].sum(axis=1)
    
    # Formatear para exportación
    df_hist['Tiempo Total Formateado'] = df_hist['Tiempo Acumulado (Mins)'].apply(to_hhmmss)
    df_hist = df_hist.sort_values('Tiempo Acumulado (Mins)', ascending=False).reset_index(drop=True)
    
    df_hist.to_csv(HISTORICO_FILE, index=False)
    return df_hist

# --- GENERADOR DE WORD PROFESIONAL ---
def crear_tabla_word(doc, df_datos, columnas, anchos=None):
    table = doc.add_table(rows=1, cols=len(columnas))
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    for i, col_name in enumerate(columnas):
        hdr_cells[i].text = col_name
        hdr_cells[i].paragraphs[0].runs[0].bold = True
        
    for index, row in df_datos.iterrows():
        row_cells = table.add_row().cells
        for i, col_name in enumerate(columnas):
            row_cells[i].text = str(row[col_name])
    return table

def generar_word(df, semana_num):
    doc = Document()
    
    # --- PÁGINA 1: RESUMEN GENERAL ---
    titulo = doc.add_heading(f'🏆 REPORTE SEMANAL CLUB TYM TRIATLÓN - SEMANA {semana_num}', 0)
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph('"(La semana de la disciplina y el volumen sostenido)"').alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()
    
    total_dep = len(df)
    df_comp = df[df['CV'] != 'NC'].copy()
    df_comp['CV_num'] = df_comp['CV'].astype(float)
    total_comp = len(df_comp)
    total_act = df['Actividades'].sum()
    
    doc.add_heading('🔍 Resumen General', level=2)
    doc.add_paragraph(f'Total deportistas activos: {total_dep}')
    doc.add_paragraph(f'Triatletas completos (CV válido en las 3 disciplinas): {total_comp}')
    doc.add_paragraph(f'Actividades registradas: {total_act} sesiones de entrenamiento')
    
    doc.add_heading('🏅 TOP 15 TRIATLETAS COMPLETOS', level=2)
    top15 = df_comp.sort_values('T_Mins', ascending=False).head(15)
    top15['Pos'] = range(1, len(top15) + 1)
    crear_tabla_word(doc, top15, ['Pos', 'Deportista', 'Tiempo Total', 'Natación', 'Bicicleta', 'Trote'])
    
    doc.add_heading('⚖️ TOP 15 MÁS BALANCEADOS (CV)', level=2)
    bal15 = df_comp.sort_values('CV_num', ascending=True).head(15)
    bal15['Pos'] = range(1, len(bal15) + 1)
    crear_tabla_word(doc, bal15, ['Pos', 'Deportista', 'CV', 'Natación', 'Bicicleta', 'Trote'])
    
    # --- PÁGINA 2: NATACIÓN ---
    doc.add_page_break()
    doc.add_heading('🏊‍♂️ ANÁLISIS POR DISCIPLINA: NATACIÓN', level=1)
    doc.add_paragraph('El podio de los atletas con mayor volumen en el agua durante esta semana.')
    
    top_nat = df.sort_values('N_Mins', ascending=False).head(10)
    top_nat['Pos'] = range(1, len(top_nat) + 1)
    crear_tabla_word(doc, top_nat, ['Pos', 'Deportista', 'Natación', 'Tiempo Total'])
    
    doc.add_heading('Análisis del Podio:', level=3)
    for i, row in top_nat.head(3).iterrows():
        comentario = generar_comentario(row['Pos'], row['Deportista'], row['Natación'], 'Natación')
        p = doc.add_paragraph()
        p.add_run(f"🥇 {row['Deportista']} ({row['Natación']}): ").bold = True if row['Pos']==1 else False
        p.add_run(f"🥈 {row['Deportista']} ({row['Natación']}): ").bold = True if row['Pos']==2 else False
        p.add_run(f"🥉 {row['Deportista']} ({row['Natación']}): ").bold = True if row['Pos']==3 else False
        p.add_run(comentario)

    # --- PÁGINA 3: CICLISMO ---
    doc.add_page_break()
    doc.add_heading('🚴‍♂️ ANÁLISIS POR DISCIPLINA: CICLISMO', level=1)
    doc.add_paragraph('El podio de los ruteros y rodadores que sumaron más horas de pedaleo.')
    
    top_bic = df.sort_values('B_Mins', ascending=False).head(10)
    top_bic['Pos'] = range(1, len(top_bic) + 1)
    crear_tabla_word(doc, top_bic, ['Pos', 'Deportista', 'Bicicleta', 'Tiempo Total'])
    
    doc.add_heading('Análisis del Podio:', level=3)
    for i, row in top_bic.head(3).iterrows():
        comentario = generar_comentario(row['Pos'], row['Deportista'], row['Bicicleta'], 'Bicicleta')
        p = doc.add_paragraph()
        if row['Pos']==1: p.add_run(f"🥇 {row['Deportista']} ({row['Bicicleta']}): ").bold = True
        elif row['Pos']==2: p.add_run(f"🥈 {row['Deportista']} ({row['Bicicleta']}): ").bold = True
        elif row['Pos']==3: p.add_run(f"🥉 {row['Deportista']} ({row['Bicicleta']}): ").bold = True
        p.add_run(comentario)

    # --- PÁGINA 4: TROTE ---
    doc.add_page_break()
    doc.add_heading('🏃‍♂️ ANÁLISIS POR DISCIPLINA: TROTE', level=1)
    doc.add_paragraph('El ranking de los corredores con mayor resistencia en el asfalto.')
    
    top_tro = df.sort_values('R_Mins', ascending=False).head(10)
    top_tro['Pos'] = range(1, len(top_tro) + 1)
    crear_tabla_word(doc, top_tro, ['Pos', 'Deportista', 'Trote', 'Tiempo Total'])
    
    doc.add_heading('Análisis del Podio:', level=3)
    for i, row in top_tro.head(3).iterrows():
        comentario = generar_comentario(row['Pos'], row['Deportista'], row['Trote'], 'Trote')
        p = doc.add_paragraph()
        if row['Pos']==1: p.add_run(f"🥇 {row['Deportista']} ({row['Trote']}): ").bold = True
        elif row['Pos']==2: p.add_run(f"🥈 {row['Deportista']} ({row['Trote']}): ").bold = True
        elif row['Pos']==3: p.add_run(f"🥉 {row['Deportista']} ({row['Trote']}): ").bold = True
        p.add_run(comentario)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- INTERFAZ WEB STREAMLIT ---
with st.sidebar:
    st.header("⚙️ Memoria Histórica")
    st.write("La app guarda el histórico automáticamente.")
    if os.path.exists(HISTORICO_FILE):
        df_historico = pd.read_csv(HISTORICO_FILE)
        st.success(f"Histórico activo: {len(df_historico)} atletas registrados.")
        csv_hist = df_historico.to_csv(index=False).encode('utf-8')
        st.download_button("📥 Descargar Base Histórica", data=csv_hist, file_name="Historial_Completo_TYM.csv", mime="text/csv")
    else:
        st.warning("No hay histórico guardado aún. Se creará al procesar la primera semana.")

semana_input = st.text_input("Ingresa el número de semana (Ej: 08):", "08")
datos_crudos = st.text_area("Pega los datos de la semana:", height=200)

if st.button("Procesar Semana y Generar Reporte"):
    if datos_crudos.strip():
        df = parse_raw_data(datos_crudos)
        
        if not df.empty:
            st.success("¡Datos leídos! Actualizando memoria y generando documentos...")
            
            # Actualizar y guardar histórico
            df_hist_actualizado = procesar_historico(df, semana_input)
            
            # Mostrar Resumen en Pantalla
            st.subheader("📋 Vista Previa del Top 5 General")
            st.dataframe(df[['Clasificación', 'Deportista', 'Tiempo Total', 'CV']].head(5))
            
            # Botones
            col1, col2 = st.columns(2)
            word_file = generar_word(df, semana_input)
            col1.download_button("📄 DESCARGAR REPORTE WORD", data=word_file, file_name=f"Reporte_Semana_{semana_input}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            
            csv_semana = df.to_csv(index=False).encode('utf-8')
            col2.download_button("📊 Descargar Datos de la Semana (CSV)", data=csv_semana, file_name=f"Datos_Semana_{semana_input}.csv", mime="text/csv")
            
        else:
            st.error("No se detectaron datos válidos.")
    else:
        st.warning("Por favor, pega los datos primero.")
