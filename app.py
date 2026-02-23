import streamlit as st
import pandas as pd
import numpy as np
import re
import plotly.express as px
import io
from docx import Document

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Generador de Reportes - TYM", page_icon="🏆", layout="wide")

st.title("🏆 Generador de Reportes Semanales - TYM Triatlón")
st.markdown("Pega los datos de la semana para generar tablas, gráficos y el reporte en Word.")

# --- FUNCIONES DE LIMPIEZA Y CÁLCULO ---
def to_mins(t_str):
    if pd.isna(t_str) or '--:--' in t_str or not t_str: return 0
    h, m = 0, 0
    hm = re.search(r'(\d+)h', t_str)
    if hm: h = int(hm.group(1))
    mm = re.search(r'(\d+)min', t_str)
    if mm: m = int(mm.group(1))
    return h * 60 + m

def to_hhmmss(mins):
    h = int(mins // 60)
    m = int(mins % 60)
    return f"{h:02d}:{m:02d}:00"

def parse_raw_data(raw_text):
    parsed_data = []
    for line in raw_text.strip().split('\n'):
        if not line.strip(): continue
        try:
            match_start = re.search(r'^(\d+)\s+(.*?)(\d+h\s*\d*min|\d+min)', line)
            if not match_start: continue
                
            rank = match_start.group(1)
            name = match_start.group(2).strip()
            rest = line[match_start.end(2):].strip()
            
            idx_min = rest.find('min')
            total_time_str = rest[:idx_min+3]
            after_total = rest[idx_min+3:]
            
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
                'Clasificación': rank,
                'Deportista': name,
                'Tiempo Total': to_hhmmss(t_mins),
                'Actividades': acts,
                'Natación': to_hhmmss(s_mins),
                'Bicicleta': to_hhmmss(b_mins),
                'Trote': to_hhmmss(r_mins),
                'CV': cv,
                'T_Mins': t_mins, 'N_Mins': s_mins, 'B_Mins': b_mins, 'R_Mins': r_mins
            })
        except Exception:
            pass
    return pd.DataFrame(parsed_data)

# --- GENERADOR DE WORD ---
def generar_word(df, semana_num):
    doc = Document()
    doc.add_heading(f'🏆 REPORTE SEMANAL CLUB TYM TRIATLÓN - SEMANA {semana_num}', 0)
    
    total_dep = len(df)
    df_comp = df[df['CV'] != 'NC'].copy()
    total_comp = len(df_comp)
    
    doc.add_heading('🔍 Resumen General', level=2)
    doc.add_paragraph(f'Total deportistas: {total_dep}')
    doc.add_paragraph(f'Triatletas completos (CV válido): {total_comp}')
    
    doc.add_heading('🏅 TOP 5 TRIATLETAS COMPLETOS', level=2)
    if not df_comp.empty:
        df_comp['CV_num'] = df_comp['CV'].astype(float)
        top5 = df_comp.sort_values('T_Mins', ascending=False).head(5)
        for i, row in top5.iterrows():
            doc.add_paragraph(f"{row['Deportista']} - {row['Tiempo Total']} ({row['Actividades']} act.)")
            
    doc.add_heading('⚖️ TRIATLETAS MÁS BALANCEADOS', level=2)
    if not df_comp.empty:
        bal = df_comp.sort_values('CV_num', ascending=True).head(3)
        for i, row in bal.iterrows():
            doc.add_paragraph(f"{row['Deportista']} - CV: {row['CV']}")
            
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- INTERFAZ WEB ---
semana_input = st.text_input("Ingresa el número de semana (Ej: 08):", "08")
datos_crudos = st.text_area("Pega aquí los datos de la semana (texto tal cual sale de tu plataforma):", height=200)

if st.button("Procesar Datos"):
    if datos_crudos.strip():
        df = parse_raw_data(datos_crudos)
        
        if not df.empty:
            st.success("¡Datos procesados correctamente!")
            
            # Gráfico de Distribución
            st.subheader("📊 Distribución de Disciplinas")
            total_n = df['N_Mins'].sum()
            total_b = df['B_Mins'].sum()
            total_r = df['R_Mins'].sum()
            
            fig = px.pie(
                values=[total_n, total_b, total_r], 
                names=['Natación', 'Ciclismo', 'Trote'],
                color_discrete_sequence=['#1E90FF', '#32CD32', '#FF4500'],
                hole=0.4
            )
            st.plotly_chart(fig, use_container_width=True)
            
            # Tabla Principal
            st.subheader("📋 Tabla General Estandarizada")
            df_mostrar = df[['Clasificación', 'Deportista', 'Tiempo Total', 'Actividades', 'Natación', 'Bicicleta', 'Trote', 'CV']]
            st.dataframe(df_mostrar)
            
            # Botones de Descarga
            col1, col2 = st.columns(2)
            
            # Descargar CSV
            csv = df_mostrar.to_csv(index=False).encode('utf-8')
            col1.download_button("⬇️ Descargar Tabla (Excel/CSV)", data=csv, file_name=f"Semana_{semana_input}_TYM.csv", mime="text/csv")
            
            # Descargar Word
            word_file = generar_word(df, semana_input)
            col2.download_button("⬇️ Descargar Reporte Base (Word)", data=word_file, file_name=f"Reporte_Semana_{semana_input}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            
        else:
            st.error("No se pudieron leer los datos. Asegúrate de pegarlos con el formato habitual.")
    else:
        st.warning("Por favor, pega los datos primero.")
