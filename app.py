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
    rank_counter = 1 
    
    # Limpiamos espacios en blanco invisibles que vienen de la web
    raw_text = raw_text.replace('\xa0', ' ')
    
    for line in raw_text.strip().split('\n'):
        line = line.strip()
        # Ignoramos líneas vacías o la cabecera
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

# --- GENERADOR DE WORD ---
def generar_word(df, semana_num):
    doc = Document()
    
    # --- TÍTULO PRINCIPAL ---
    doc.add_heading(f'🏆 REPORTE SEMANAL CLUB TYM TRIATLÓN - SEMANA {semana_num}', 0)
    doc.add_paragraph('"(Frase destacada de la semana)"')
    doc.add_paragraph('Breve introducción o resumen general de la semana...')
    
    # --- CÁLCULOS GENERALES ---
    total_dep = len(df)
    df_comp = df[df['CV'] != 'NC'].copy()
    df_comp['CV_num'] = df_comp['CV'].astype(float)
    total_comp = len(df_comp)
    total_act = df['Actividades'].sum()
    
    total_mins = df['T_Mins'].sum()
    horas_totales = int(total_mins // 60)
    mins_totales = int(total_mins % 60)
    
    t_nat = df['N_Mins'].sum()
    t_bic = df['B_Mins'].sum()
    t_tro = df['R_Mins'].sum()
    suma_disc = t_nat + t_bic + t_tro
    
    pct_nat = (t_nat / suma_disc * 100) if suma_disc > 0 else 0
    pct_bic = (t_bic / suma_disc * 100) if suma_disc > 0 else 0
    pct_tro = (t_tro / suma_disc * 100) if suma_disc > 0 else 0
    
    # --- RESUMEN GENERAL ---
    doc.add_heading('🔍 Resumen General', level=2)
    doc.add_paragraph(f'Total deportistas: {total_dep}')
    doc.add_paragraph(f'Triatletas completos (CV válido): {total_comp}')
    doc.add_paragraph(f'Horas totales acumuladas: {horas_totales} horas y {mins_totales} minutos')
    doc.add_paragraph(f'Actividades registradas: {total_act} actividades')
    doc.add_paragraph('Distribución aproximada:')
    doc.add_paragraph(f'🏊‍♂️ Natación: {pct_nat:.1f}%', style='List Bullet')
    doc.add_paragraph(f'🚴‍♂️ Ciclismo: {pct_bic:.1f}%', style='List Bullet')
    doc.add_paragraph(f'🏃‍♂️ Trote: {pct_tro:.1f}%', style='List Bullet')

    # --- TOP 15 COMPLETOS ---
    doc.add_heading('🏅 TOP 15 TRIATLETAS COMPLETOS', level=2)
    doc.add_paragraph('(Clasificación por tiempo total acumulado, filtrando estrictamente CV válido)')
    
    if not df_comp.empty:
        top15 = df_comp.sort_values('T_Mins', ascending=False).head(15)
        
        table = doc.add_table(rows=1, cols=7)
        table.style = 'Table Grid'
        hdr_cells = table.rows[0].cells
        encabezados = ['#', 'Deportista', 'Tiempo Total', 'Act.', 'Natación', 'Bicicleta', 'Carrera']
        for i, header in enumerate(encabezados):
            hdr_cells[i].text = header
            
        for i, row in top15.reset_index().iterrows():
            row_cells = table.add_row().cells
            row_cells[0].text = str(i + 1)
            row_cells[1].text = str(row['Deportista'])
            row_cells[2].text = str(row['Tiempo Total'])
            row_cells[3].text = str(row['Actividades'])
            row_cells[4].text = str(row['Natación'])
            row_cells[5].text = str(row['Bicicleta'])
            row_cells[6].text = str(row['Trote'])
            
    doc.add_heading('Análisis Ejecutivo (El Podio)', level=3)
    doc.add_paragraph('🥇 1. [Nombre] – "[Apodo o frase corta]"')
    doc.add_paragraph('[Comentario entretenido del primer lugar...]')
    doc.add_paragraph('🥈 2. [Nombre] – "[Apodo o frase corta]"')
    doc.add_paragraph('[Comentario entretenido del segundo lugar...]')
    doc.add_paragraph('🥉 3. [Nombre] – "[Apodo o frase corta]"')
    doc.add_paragraph('[Comentario entretenido del tercer lugar...]')

    # --- ATLETAS BALANCEADOS ---
    doc.add_heading('⚖️ TOP 15 TRIATLETAS MÁS BALANCEADOS', level=2)
    doc.add_paragraph('(Menor Coeficiente de Variación - La proporción perfecta entre disciplinas)')
    
    if not df_comp.empty:
        bal = df_comp.sort_values('CV_num', ascending=True).head(15)
        table2 = doc.add_table(rows=1, cols=6)
        table2.style = 'Table Grid'
        hdr_cells2 = table2.rows[0].cells
        encabezados2 = ['#', 'Deportista', 'CV', 'Natación', 'Bicicleta', 'Carrera']
        for i, header in enumerate(encabezados2):
            hdr_cells2[i].text = header
            
        for i, row in bal.reset_index().iterrows():
            row_cells2 = table2.add_row().cells
            row_cells2[0].text = str(i + 1)
            row_cells2[1].text = str(row['Deportista'])
            row_cells2[2].text = str(row['CV'])
            row_cells2[3].text = str(row['Natación'])
            row_cells2[4].text = str(row['Bicicleta'])
            row_cells2[5].text = str(row['Trote'])

    doc.add_heading('Análisis breve del Podio de Simetría:', level=3)
    doc.add_paragraph('🥇 1. [Nombre] ([CV]) – [Comentario entretenido sobre su balance]')
    doc.add_paragraph('🥈 2. [Nombre] ([CV]) – [Comentario entretenido sobre su balance]')
    doc.add_paragraph('🥉 3. [Nombre] ([CV]) – [Comentario entretenido sobre su balance]')

    # --- PODIOS POR DISCIPLINA ---
    def agregar_podio(titulo, icono, col_orden, formato_str):
        doc.add_heading(f'{icono} {titulo}', level=2)
        top3 = df.sort_values(col_orden, ascending=False).head(3)
        for i, row in top3.reset_index().iterrows():
            tiempo = row[formato_str]
            doc.add_paragraph(f"{i+1}. {row['Deportista']} ({tiempo}) – [Comentario entretenido]")

    doc.add_heading('🥇 PODIO TIEMPO TOTAL GENERAL (Incluyendo especialistas)', level=2)
    top_general = df.sort_values('T_Mins', ascending=False).head(3)
    for i, row in top_general.reset_index().iterrows():
        doc.add_paragraph(f"{i+1}. {row['Deportista']} ({row['Tiempo Total']}) – [Comentario entretenido]")

    agregar_podio('PODIO NATACIÓN', '🏊‍♂️', 'N_Mins', 'Natación')
    agregar_podio('PODIO CICLISMO', '🚴', 'B_Mins', 'Bicicleta')
    agregar_podio('PODIO TROTE', '🏃‍♂️', 'R_Mins', 'Trote')

    # --- OTRAS CATEGORÍAS ---
    doc.add_heading('🔄 MAYOR FRECUENCIA (ACTIVIDADES TOTALES)', level=2)
    top_act = df.sort_values('Actividades', ascending=False).head(3)
    for i, row in top_act.reset_index().iterrows():
        doc.add_paragraph(f"{i+1}. {row['Deportista']} ({row['Actividades']} actividades) – [Comentario entretenido]")

    doc.add_heading('📊 RESUMEN DE CAMBIOS Y CONCLUSIONES', level=2)
    doc.add_paragraph('[Escribir comparativa con la semana anterior...]')

    doc.add_heading('💡 INSIGHTS ESTRATÉGICOS', level=2)
    doc.add_paragraph('[Escribir reflexiones finales...]')
    
    # Guardar en buffer
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
            
            # Tabla Principal (Mostramos el Top 15 Completo en la app también)
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
