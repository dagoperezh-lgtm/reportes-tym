import streamlit as st
import pandas as pd
import numpy as np
import re
import plotly.express as px
import matplotlib.pyplot as plt
import io
import os
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Inteligencia TYM", page_icon="🏆", layout="wide")

st.title("🏆 Plataforma de Reportes e Inteligencia - TYM")
st.markdown("Procesa la semana, actualiza el histórico y genera el reporte oficial automatizado.")

# --- ARCHIVO HISTÓRICO LOCAL ---
HISTORICO_FILE = "historico_tym.csv"

# --- FUNCIONES DE TIEMPO Y CÁLCULO ---
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

# --- MOTOR DE INTELIGENCIA PARA COMENTARIOS ---
def generar_comentario_inteligente(row, df_completo, categoria, pos):
    nombre = row['Deportista']
    acts = row.get('Actividades', 0)
    t_mins = row.get('T_Mins', 0)
    b_mins = row.get('B_Mins', 0)
    
    pct_bici = (b_mins / t_mins * 100) if t_mins > 0 else 0
    
    if categoria == 'General':
        if pos == 1:
            margen = t_mins - df_completo.iloc[1]['T_Mins'] if len(df_completo) > 1 else 0
            analisis = f"Dominio absoluto en la cima. "
            if margen > 60: analisis += f"Le sacó más de una hora de ventaja a su perseguidor más cercano. "
            if pct_bici > 60: analisis += "Su estrategia se cimentó en un volumen de ciclismo inalcanzable para el resto."
            elif acts >= 14: analisis += "La clave de su éxito fue la constancia implacable, promediando casi un doble turno diario."
            else: analisis += "Un rendimiento equilibrado y brutal en todas las líneas."
            return analisis
        elif pos == 2:
            return f"Plata pura y una presión constante sobre el líder. Su solidez física lo mantiene peleando en la élite del club con una capacidad de sufrimiento destacable."
        elif pos == 3:
            return f"Bronce ganado a base de regularidad. Se consolida en el podio demostrando que el trabajo silencioso y la disciplina rinden frutos en la general."
        else:
            return f"Rendimiento de alto estándar. Sumar este volumen lo ratifica como uno de los motores principales del equipo esta semana."

    if categoria == 'CV':
        cv_val = float(row['CV'])
        if pos == 1:
            if cv_val < 0.1: return f"¡RÉCORD DE SIMETRÍA! Un triángulo casi perfecto. Entrenar las tres disciplinas con esta precisión requiere una planificación quirúrgica."
            return f"El líder indiscutido de la eficiencia. Demuestra un balance envidiable, sin dejar puntos débiles ni disciplinas al azar."
        elif pos == 2:
            return f"Equilibrio táctico fantástico. Mantiene a raya el desgaste distribuyendo sus cargas de forma magistral entre el agua, la ruta y el trote."
        else:
            return f"Constancia y proporción. Cierra el podio de la simetría confirmando que se puede entrenar de forma inteligente y completa."

    tiempo_str = row[categoria] if categoria in ['Natación', 'Bicicleta', 'Trote'] else ""
    if categoria == 'Natación':
        if pos == 1: return f"El rey/reina indiscutido del agua esta semana. {tiempo_str} nadando demuestran una técnica depurada y unos hombros que no conocen la fatiga."
        if pos == 2: return f"Un volumen acuático tremendo. Base fundamental para asegurar salidas rápidas y eficientes en cualquier triatlón."
        if pos == 3: return f"Cierra el podio con solidez, acumulando metros esenciales para la resistencia del tren superior."
        
    if categoria == 'Bicicleta':
        if pos == 1: return f"Una locomotora desatada. {tiempo_str} sobre el sillín requieren piernas de titanio y una mente inquebrantable para devorar kilómetros."
        if pos == 2: return f"Plata rodadora. Constancia envidiable en la ruta, demostrando una potencia sostenida espectacular."
        if pos == 3: return f"Bronce de alto rendimiento. Confirma que para pelear arriba hay que ser una bestia en los pedales."

    if categoria == 'Trote':
        if pos == 1: return f"El dictador del asfalto. Coronarse con {tiempo_str} habla de una resistencia cardiovascular de otro planeta para soportar el impacto."
        if pos == 2: return f"Devorando kilómetros a pie con una regularidad que intimida, consolidando una tremenda fase de carrera."
        if pos == 3: return f"Gran resiliencia para dominar el impacto constante y sumar kilómetros valiosísimos de resistencia pura."
        
    return "Gran desempeño."

# --- MOTOR DE LECTURA ROBUSTO ---
def parse_raw_data(raw_text):
    parsed_data = []
    rank_counter = 1 
    raw_text = raw_text.replace('\xa0', ' ')
    
    for line in raw_text.strip().split('\n'):
        line = line.strip()
        if not line or 'Deportista' in line: continue
        
        try:
            time_pattern = r'(\d+h\s*\d*min|\d+h|\d+min|--:--)'
            first_time_match = re.search(time_pattern, line)
            if not first_time_match: continue
            
            name_part = line[:first_time_match.start()].strip()
            name = re.sub(r'^\d+\s*', '', name_part).strip()
            
            total_time_str = first_time_match.group(1)
            after_total = line[first_time_match.end():].strip()
            after_total = re.sub(r'\d+$', '', after_total).strip() 
            
            remaining_times_matches = list(re.finditer(time_pattern, after_total))
            
            acts_str = "0"
            swim_str = "--:--"
            bike_str = "--:--"
            run_str = "--:--"
            
            if remaining_times_matches:
                acts_str = after_total[:remaining_times_matches[0].start()].strip()
                times = [m.group(1) for m in remaining_times_matches]
                if len(times) > 0: swim_str = times[0]
                if len(times) > 1: bike_str = times[1]
                if len(times) > 2: run_str = times[2]
            
            acts = int(acts_str) if acts_str.isdigit() else 0
            
            t_mins = to_mins(total_time_str)
            s_mins = to_mins(swim_str)
            b_mins = to_mins(bike_str)
            r_mins = to_mins(run_str)
            
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
        
        # PREVENCIÓN DE KEYERROR: Si la semana ya existe (re-procesamiento), la borramos para sobreescribirla limpia
        if col_semana in df_hist.columns:
            df_hist = df_hist.drop(columns=[col_semana])
            
        df_hist = pd.merge(df_hist, df_nueva, on='Nombre', how='outer')
        df_hist[col_semana] = df_hist[col_semana].fillna(0)
    else:
        df_hist = df_nueva.copy()
        df_hist['Tiempo Acumulado (Mins)'] = 0

    cols_semanas = [c for c in df_hist.columns if c.startswith('Sem ')]
    df_hist[cols_semanas] = df_hist[cols_semanas].fillna(0)
    df_hist['Tiempo Acumulado (Mins)'] = df_hist[cols_semanas].sum(axis=1)
    
    df_hist['Tiempo Total Formateado'] = df_hist['Tiempo Acumulado (Mins)'].apply(to_hhmmss)
    df_hist = df_hist.sort_values('Tiempo Acumulado (Mins)', ascending=False).reset_index(drop=True)
    
    df_hist.to_csv(HISTORICO_FILE, index=False)
    return df_hist

# --- GENERADOR DE WORD OFICIAL ---
def crear_tabla_word(doc, df_datos, columnas):
    table = doc.add_table(rows=1, cols=len(columnas))
    table.style = 'Light Grid Accent 1'
    hdr_cells = table.rows[0].cells
    for i, col_name in enumerate(columnas):
        hdr_cells[i].text = col_name
        hdr_cells[i].paragraphs[0].runs[0].bold = True
        
    for index, row in df_datos.iterrows():
        row_cells = table.add_row().cells
        for i, col_name in enumerate(columnas):
            row_cells[i].text = str(row[col_name])

def generar_word(df, semana_num):
    doc = Document()
    
    # --- PÁGINA 1: RESUMEN Y GENERALES ---
    titulo = doc.add_heading(f'Reporte Semanal Club Tym Triatlón - Semana {semana_num}', 0)
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph('"(La semana del volumen desatado y la técnica perfecta)"').alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph('Cerramos la semana con confirmaciones importantes. El análisis detallado nos muestra los resultados consolidados de nuestros deportistas, destacando a los líderes absolutos de cada disciplina.')
    
    df_comp = df[df['CV'] != 'NC'].copy()
    df_comp['CV_num'] = df_comp['CV'].astype(float)
    
    suma_disc = df['N_Mins'].sum() + df['B_Mins'].sum() + df['R_Mins'].sum()
    pct_nat = (df['N_Mins'].sum() / suma_disc * 100) if suma_disc > 0 else 0
    pct_bic = (df['B_Mins'].sum() / suma_disc * 100) if suma_disc > 0 else 0
    pct_tro = (df['R_Mins'].sum() / suma_disc * 100) if suma_disc > 0 else 0
    
    doc.add_heading('🔍 Resumen General', level=2)
    doc.add_paragraph(f'Total deportistas: {len(df)}')
    doc.add_paragraph(f'Triatletas completos (CV válido): {len(df_comp)}')
    doc.add_paragraph(f'Horas totales acumuladas: {int(df["T_Mins"].sum() // 60)} horas y {int(df["T_Mins"].sum() % 60)} minutos')
    doc.add_paragraph(f'Actividades registradas: {df["Actividades"].sum()} actividades')
    doc.add_paragraph('Distribución aproximada:')
    doc.add_paragraph(f'🏊‍♂️ Natación: {pct_nat:.1f}%', style='List Bullet')
    doc.add_paragraph(f'🚴‍♂️ Ciclismo: {pct_bic:.1f}%', style='List Bullet')
    doc.add_paragraph(f'🏃‍♂️ Trote: {pct_tro:.1f}%', style='List Bullet')
    
    # [Insertar Gráfico Automático]
    fig, ax = plt.subplots(figsize=(5, 5))
    ax.pie([pct_nat, pct_bic, pct_tro], labels=['Natación', 'Ciclismo', 'Trote'], autopct='%1.1f%%', startangle=90, colors=['#1E90FF', '#32CD32', '#FF4500'])
    ax.axis('equal')
    img_stream = io.BytesIO()
    plt.savefig(img_stream, format='png', bbox_inches='tight')
    img_stream.seek(0)
    doc.add_paragraph().add_run().add_picture(img_stream, width=Inches(4))
    plt.close(fig)

    doc.add_heading('🏅 TOP 5 TRIATLETAS COMPLETOS', level=2)
    doc.add_paragraph('(Clasificación por tiempo total acumulado, filtrando estrictamente CV válido)')
    top5_comp = df_comp.sort_values('T_Mins', ascending=False).head(5)
    top5_comp['#'] = range(1, len(top5_comp) + 1)
    crear_tabla_word(doc, top5_comp, ['#', 'Deportista', 'Tiempo Total', 'Actividades', 'Natación', 'Bicicleta', 'Trote'])
    
    doc.add_heading('Análisis Ejecutivo', level=3)
    for i, row in top5_comp.iterrows():
        p = doc.add_paragraph()
        icono = '🥇' if row['#']==1 else '🥈' if row['#']==2 else '🥉' if row['#']==3 else str(row['#'])+'.'
        p.add_run(f"{icono} {row['Deportista']}").bold = True
        
        p_tech = doc.add_paragraph()
        p_tech.add_run(f"{row['Tiempo Total']} | {row['Actividades']} actividades | {row['Bicicleta']} ciclismo").italic = True
        
        com = generar_comentario_inteligente(row, top5_comp, 'General', row['#'])
        doc.add_paragraph(f"{com}")

    doc.add_heading('⚖️ TOP 5 TRIATLETAS MÁS BALANCEADOS', level=2)
    doc.add_paragraph('(Menor Coeficiente de Variación - La proporción perfecta entre disciplinas)')
    bal5 = df_comp.sort_values('CV_num', ascending=True).head(5)
    bal5['#'] = range(1, len(bal5) + 1)
    crear_tabla_word(doc, bal5, ['#', 'Deportista', 'CV', 'Natación', 'Bicicleta', 'Trote'])
    
    doc.add_heading('Análisis de Simetría:', level=3)
    for i, row in bal5.iterrows():
        p = doc.add_paragraph()
        p.add_run(f"{row['#']}. {row['Deportista']} (CV: {row['CV']})").bold = True
        com = generar_comentario_inteligente(row, bal5, 'CV', row['#'])
        doc.add_paragraph(f"{com}")

    doc.add_heading('🥇 PODIO TIEMPO TOTAL GENERAL', level=2)
    doc.add_paragraph('(Incluyendo especialistas de disciplinas únicas)')
    top_gral = df.sort_values('T_Mins', ascending=False).head(3)
    for i, row in top_gral.reset_index().iterrows():
        doc.add_paragraph(f"{i+1}. {row['Deportista']} ({row['Tiempo Total']})")

    doc.add_heading('🌟 OTRAS CATEGORÍAS DESTACADAS', level=2)
    
    doc.add_heading('🔄 MAYOR FRECUENCIA (ACTIVIDADES TOTALES)', level=3)
    top_act = df[df['Actividades'] > 0].sort_values('Actividades', ascending=False).head(3)
    for i, row in top_act.reset_index().iterrows():
        doc.add_paragraph(f"{row['Deportista']} ({row['Actividades']} sesiones) – Motor inagotable.")
        
    doc.add_heading('📏 PODIO DISTANCIA TOTAL', level=3)
    doc.add_paragraph('1. [Nombre] ([Km] km) - [Comentario]')
    doc.add_paragraph('2. [Nombre] ([Km] km) - [Comentario]')
    doc.add_paragraph('3. [Nombre] ([Km] km) - [Comentario]')

    doc.add_heading('⏱️ PODIO ACTIVIDAD MÁS LARGA', level=3)
    doc.add_paragraph('1. [Nombre] ([Tiempo]) - [Comentario]')
    doc.add_paragraph('2. [Nombre] ([Tiempo]) - [Comentario]')
    doc.add_paragraph('3. [Nombre] ([Tiempo]) - [Comentario]')

    doc.add_heading('📊 RESUMEN DE CAMBIOS Y CONCLUSIONES', level=2)
    doc.add_paragraph('🔄 DE LA SEMANA ANTERIOR A LA ACTUAL:')
    doc.add_paragraph('[Completar: Comparativa y cambios de liderazgos...]')
    
    doc.add_heading('💡 INSIGHTS ESTRATÉGICOS', level=2)
    doc.add_paragraph('[Completar: Observaciones técnicas del cuerpo de entrenadores...]')

    # --- PÁGINA 2: NATACIÓN ---
    doc.add_page_break()
    doc.add_heading('🏊‍♂️ TOP 10 NATACIÓN', level=1)
    top_nat = df[df['N_Mins'] > 0].sort_values('N_Mins', ascending=False).head(10)
    if not top_nat.empty:
        top_nat['Pos'] = range(1, len(top_nat) + 1)
        crear_tabla_word(doc, top_nat, ['Pos', 'Deportista', 'Natación', 'Tiempo Total'])
        doc.add_heading('Análisis del Podio Acuático:', level=3)
        for i, row in top_nat.head(3).iterrows():
            com = generar_comentario_inteligente(row, top_nat, 'Natación', row['Pos'])
            p = doc.add_paragraph()
            p.add_run(f"{'🥇' if row['Pos']==1 else '🥈' if row['Pos']==2 else '🥉'} {row['Deportista']} ({row['Natación']}): ").bold = True
            p.add_run(com)

    # --- PÁGINA 3: CICLISMO ---
    doc.add_page_break()
    doc.add_heading('🚴 TOP 10 CICLISMO', level=1)
    top_bic = df[df['B_Mins'] > 0].sort_values('B_Mins', ascending=False).head(10)
    if not top_bic.empty:
        top_bic['Pos'] = range(1, len(top_bic) + 1)
        crear_tabla_word(doc, top_bic, ['Pos', 'Deportista', 'Bicicleta', 'Tiempo Total'])
        doc.add_heading('Análisis del Podio de Ruta:', level=3)
        for i, row in top_bic.head(3).iterrows():
            com = generar_comentario_inteligente(row, top_bic, 'Bicicleta', row['Pos'])
            p = doc.add_paragraph()
            p.add_run(f"{'🥇' if row['Pos']==1 else '🥈' if row['Pos']==2 else '🥉'} {row['Deportista']} ({row['Bicicleta']}): ").bold = True
            p.add_run(com)

    # --- PÁGINA 4: TROTE ---
    doc.add_page_break()
    doc.add_heading('🏃‍♂️ TOP 10 TROTE', level=1)
    top_tro = df[df['R_Mins'] > 0].sort_values('R_Mins', ascending=False).head(10)
    if not top_tro.empty:
        top_tro['Pos'] = range(1, len(top_tro) + 1)
        crear_tabla_word(doc, top_tro, ['Pos', 'Deportista', 'Trote', 'Tiempo Total'])
        doc.add_heading('Análisis del Podio de Asfalto:', level=3)
        for i, row in top_tro.head(3).iterrows():
            com = generar_comentario_inteligente(row, top_tro, 'Trote', row['Pos'])
            p = doc.add_paragraph()
            p.add_run(f"{'🥇' if row['Pos']==1 else '🥈' if row['Pos']==2 else '🥉'} {row['Deportista']} ({row['Trote']}): ").bold = True
            p.add_run(com)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- INTERFAZ WEB STREAMLIT ---
with st.sidebar:
    st.header("⚙️ Memoria Histórica")
    st.write("La base de datos aprende con cada semana procesada.")
    if os.path.exists(HISTORICO_FILE):
        df_historico = pd.read_csv(HISTORICO_FILE)
        st.success(f"Histórico activo: {len(df_historico)} atletas.")
        csv_hist = df_historico.to_csv(index=False).encode('utf-8')
        st.download_button("📥 Descargar Respaldo Histórico", data=csv_hist, file_name="Historial_TYM.csv", mime="text/csv")
    else:
        st.warning("Sin histórico. Se creará al procesar ahora.")

semana_input = st.text_input("Número de la semana a procesar (Ej: 08):", "08")
datos_crudos = st.text_area("Pega los datos crudos aquí:", height=200)

if st.button("Generar Reporte Completo"):
    if datos_crudos.strip():
        df = parse_raw_data(datos_crudos)
        
        if not df.empty:
            st.success("¡Datos procesados! Gráficos y documentos listos.")
            df_hist_actualizado = procesar_historico(df, semana_input)
            
            st.subheader("📋 Vista Previa: Top 5 General")
            st.dataframe(df[['Clasificación', 'Deportista', 'Tiempo Total', 'CV']].head(5))
            
            st.subheader("📊 Distribución Web")
            fig_web = px.pie(values=[df['N_Mins'].sum(), df['B_Mins'].sum(), df['R_Mins'].sum()], names=['Natación', 'Ciclismo', 'Trote'], color_discrete_sequence=['#1E90FF', '#32CD32', '#FF4500'], hole=0.4)
            st.plotly_chart(fig_web, use_container_width=True)
            
            col1, col2 = st.columns(2)
            word_file = generar_word(df, semana_input)
            col1.download_button("📄 DESCARGAR REPORTE OFICIAL (Word)", data=word_file, file_name=f"Reporte_Semana_{semana_input}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            
            csv_semana = df.to_csv(index=False).encode('utf-8')
            col2.download_button("📊 Descargar Datos de la Semana (CSV)", data=csv_semana, file_name=f"Datos_Semana_{semana_input}.csv", mime="text/csv")
        else:
            st.error("Formato inválido. Asegúrate de copiar desde la plataforma correcta.")
    else:
        st.warning("Debes pegar los datos primero.")
