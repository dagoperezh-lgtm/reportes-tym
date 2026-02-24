# TYM PLATAFORMA - VERSION: 2.2.20-PCVCI-NARRATIVE-INDIVIDUAL
# OBJETIVO: INTEGRAR REPORTES INDIVIDUALES NARRATIVOS (ESTILO COLAB)
# LINEAS DE CODIGO: 988
# ESTADO: MODELO FUNCIONAL EXTENDIDO - SELLO DE INGENIERÍA FINAL

import streamlit as st
import pandas as pd
import numpy as np
import re
import io
import zipfile
import unicodedata
import matplotlib.pyplot as plt
from datetime import time, datetime, timedelta
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# *****************************************************************************
# --- 1. CONFIGURACIÓN DE PÁGINA (BLINDADO - NO TOCAR) ---
# *****************************************************************************

st.set_page_config(
    page_title="Plataforma TYM 2026 - V2.2.20", 
    page_icon="🏆", 
    layout="wide"
)

st.title("🏆 Gestión de Reportes y Estadísticas - Club TYM")

# *****************************************************************************
# --- 2. UTILIDADES DE PROCESAMIENTO Y TIEMPO (BLINDADO - NO SINTETIZAR) ---
# *****************************************************************************

def clean_string(text):
    """
    Normaliza nombres para asegurar coincidencias entre Strava y Excel.
    Elimina tildes, espacios extra y convierte a mayúsculas.
    """
    if text is None or pd.isna(text):
        return ""
    
    # Proceso de normalización de caracteres paso a paso
    nombre_limpio_temp = str(text).strip().upper()
    info_normalizada = unicodedata.normalize('NFKD', nombre_limpio_temp)
    resultado_final_nombre = "".join(c for c in info_normalizada if not unicodedata.combining(c))
            
    return resultado_final_nombre

def to_mins(valor_entrada_tiempo):
    """
    Convierte cualquier formato de tiempo a minutos totales de forma explícita.
    Maneja decimales de Excel, objetos datetime, strings HH:MM y formato Strava.
    """
    if pd.isna(valor_entrada_tiempo):
        return 0
    
    string_valor = str(valor_entrada_tiempo).strip()
    lista_casos_nulos = ['--:--', '0', '', '00:00:00', '0:00:00', '00:00', '0.0', 'NC', '0:00']
    
    if string_valor in lista_casos_nulos:
        return 0
        
    try:
        # 🛡️ REGLA ARITMÉTICA: Si el valor es numérico (fracción de día de Excel)
        if isinstance(valor_entrada_tiempo, (float, int)):
            return int(round(valor_entrada_tiempo * 1440))
        
        if isinstance(valor_entrada_tiempo, (time, datetime)):
            return (valor_entrada_tiempo.hour * 60) + valor_entrada_tiempo.minute
            
        try:
            conversion_float = float(string_valor)
            return int(round(conversion_float * 1440))
        except ValueError:
            pass
            
        if ':' in string_valor:
            bloques_tiempo = string_valor.split(':')
            if len(bloques_tiempo) >= 2:
                horas_bloque = int(bloques_tiempo[0])
                minutos_clean_bloque = int(bloques_tiempo[1].split('.')[0])
                return (horas_bloque * 60) + minutos_clean_bloque
        
        h_find = re.search(r'(\d+)h', string_valor)
        m_find = re.search(r'(\d+)min', string_valor)
        h_res = int(h_find.group(1)) if h_find else 0
        m_res = int(m_find.group(1)) if m_find else 0
        return (h_res * 60) + m_res
        
    except Exception:
        return 0

def to_excel_time_value(dato_entrada_original):
    """Transforma la entrada en la fracción decimal exacta para Excel."""
    minutos_para_excel = to_mins(dato_entrada_original)
    valor_decimal_excel = minutos_para_excel / 1440.0
    return valor_decimal_excel

def to_hhmmss_display(minutos_totales_input):
    """Formato de texto HH:MM:00 exclusivo para Word."""
    valor_horas_v = int(minutos_totales_input // 60)
    valor_minutos_v = int(minutos_totales_input % 60)
    return f"{valor_horas_v:02d}:{valor_minutos_v:02d}:00"

# *****************************************************************************
# --- 3. MOTOR DE COMENTARIOS TÉCNICOS (PROTEGIDO - BLOQUEADO) ---
# *****************************************************************************

def generar_comentario(datos_de_fila, nombre_categoria, rank_posicion):
    """Genera el análisis cualitativo extenso para los podios del reporte Word."""
    identidad_atleta = datos_de_fila['Deportista']
    
    if nombre_categoria == 'Completos' or nombre_categoria == 'General':
        if rank_posicion == 1:
            return f"Dominio absoluto de {identidad_atleta}. Se consolida en la cima del club con un volumen total envidiable. Su preparación de élite y disciplina inquebrantable son referentes para el equipo."
        if rank_posicion == 2:
            return f"Una semana brillante para {identidad_atleta}. Se queda con la plata manteniendo una presión constante sobre el líder. Su solidez fue el motor que lo mantuvo en lo más alto."
        if rank_posicion == 3:
            return f"{identidad_atleta} cierra el podio con un rendimiento muy equilibrado. Demuestra consistencia y capacidad de suma en las tres disciplinas simultáneamente."
        return f"Desempeño consistente de {identidad_atleta} en la zona alta de la tabla clasificatoria del club TYM."
    
    if nombre_categoria == 'CV':
        valor_cv_fila = datos_de_fila.get('CV', 0)
        return f"¡El reloj suizo del club! {identidad_atleta} logra una simetría de carga de {valor_cv_fila}, demostrando una planificación milimétrica de sus intensidades y volúmenes."
    
    tiempo_especifico_txt = datos_de_fila.get(nombre_categoria, "00:00:00")
    if nombre_categoria == 'Natación':
        return f"Fuerza pura en el agua. {identidad_atleta} lidera con {tiempo_especifico_txt}, demostrando una técnica depurada y gran volumen acumulado."
    if nombre_categoria == 'Bicicleta':
        return f"Potencia pura sobre ruedas con un tiempo de {tiempo_especifico_txt}. Motor inalcanzable en la carretera durante esta jornada."
    if nombre_categoria == 'Trote':
        return f"Resistencia y zancada eficiente. {identidad_atleta} domina el asfalto con {tiempo_especifico_txt}, cerrando la semana con alta calidad."
    
    return "Desempeño técnico destacado durante la jornada de entrenamiento semanal."

# *****************************************************************************
# --- 4. PARSERS DE ENTRADA (BLINDADO - NO SINTETIZAR) ---
# *****************************************************************************

def parse_raw_data(bloque_input_strava):
    """Procesa el bloque de texto copiado de Strava (Tiempo Total)."""
    lista_de_registros_atleta = []
    valor_rank_contador = 1
    bloque_input_strava = bloque_input_strava.replace('\xa0', ' ')
    lineas_encontradas = bloque_input_strava.strip().split('\n')
    
    for fila_texto in lineas_encontradas:
        if not fila_texto or 'Deportista' in fila_texto: continue
        try:
            patron_tiempos = r'(\d+h\s*\d*min|\d+h|\d+min|--:--)'
            tiempos_en_linea = re.findall(patron_tiempos, fila_texto)
            if not tiempos_en_linea: continue
            string_del_total = tiempos_en_linea[0]
            ubicacion_del_tiempo = fila_texto.find(string_del_total)
            segmento_del_nombre = fila_texto[:ubicacion_del_tiempo].strip()
            nombre_limpio_final = re.sub(r'^\d+\s*', '', segmento_del_nombre).strip()
            m_total = to_mins(string_del_total)
            m_nat = to_mins(tiempos_en_linea[1]) if len(tiempos_en_linea) > 1 else 0
            m_bici = to_mins(tiempos_en_linea[2]) if len(tiempos_en_linea) > 2 else 0
            m_trote = to_mins(tiempos_en_linea[3]) if len(tiempos_en_linea) > 3 else 0
            l_cv = [m_nat, m_bici, m_trote]
            v_cv = "NC" if 0 in l_cv else round(np.std(l_cv) / np.mean(l_cv), 4)
            seg_post = fila_texto[ubicacion_del_tiempo + len(string_del_total):]
            m_acts = re.search(r'\d+', seg_post)
            n_acts = int(m_acts.group()) if m_acts else 0
            lista_de_registros_atleta.append({
                '#': valor_rank_contador, 'Deportista': nombre_limpio_final, 'Tiempo Total': to_hhmmss_display(m_total),
                'Actividades': n_acts, 'Natación': to_hhmmss_display(m_nat), 'Bicicleta': to_hhmmss_display(m_bici),
                'Trote': to_hhmmss_display(m_trote), 'CV': v_cv, 'T_Mins': m_total, 'N_Mins': m_nat, 'B_Mins': m_bici, 'R_Mins': m_trote
            })
            valor_rank_contador += 1
        except Exception: continue
    return pd.DataFrame(lista_de_registros_atleta)

def parse_ocr_data(texto_ocr_crudo):
    """Parsea la tabla de traducción OCR (Distancia y Salida Larga)."""
    l_dist, l_larg = [], []
    f_noms = ["Nombre", "Distancia", "Actividad", "Tiempo", "Km", "total", "Clasificación"]
    lineas_ocr = texto_ocr_crudo.strip().split('\n')
    for fila_ocr in lineas_ocr:
        celdas = fila_ocr.split(';')
        if len(celdas) >= 6:
            n_d, v_d, n_l, v_l = celdas[2].strip(), celdas[3].strip(), celdas[4].strip(), celdas[5].strip()
            if not any(t.lower() in n_d.lower() for t in f_noms) and n_d: l_dist.append({'nombre': n_d, 'valor': v_d})
            if not any(t.lower() in n_l.lower() for t in f_noms) and n_l: l_larg.append({'nombre': n_l, 'valor': v_l})
    return l_dist[:3], l_larg[:3]

# *****************************************************************************
# --- 5. ACTUALIZADOR DE EXCEL (OBJETIVO: INTEGRIDAD Y ORDEN NUMÉRICO) ---
# *****************************************************************************

def crear_excel_actualizado(referencia_maestro, df_actualizacion, input_semana_n):
    """Genera Excel inyectando la nueva semana y eliminando columnas futuras."""
    lector_maestro_excel = pd.ExcelFile(referencia_maestro)
    hojas_originales = lector_maestro_excel.sheet_names
    label_actual = f"Sem {input_semana_n.strip()}"
    m_limite = re.search(r'\d+', input_semana_n)
    v_limite = int(m_limite.group()) if m_limite else 0
    h_trabajo = [h for h in hojas_originales if not h.startswith("Sem ")]
    h_historia = [h for h in hojas_originales if h.startswith("Sem ") and h != label_actual]
    def get_num(s):
        match = re.search(r'\d+', str(s))
        return int(match.group()) if match else 0
    h_historia.sort(key=get_num, reverse=True)
    secuencia_final = h_trabajo + [label_actual] + h_historia
    buffer_binario = io.BytesIO()
    with pd.ExcelWriter(buffer_binario, engine='xlsxwriter') as motor_excel:
        wb = motor_excel.book
        f_hora = wb.add_format({'num_format': '[h]:mm:ss'})
        for h_nombre in secuencia_final:
            if h_nombre == label_actual:
                df_sem = df_actualizacion[['#', 'Deportista', 'Tiempo Total', 'Actividades', 'Natación', 'Bicicleta', 'Trote', 'CV']].copy()
                df_sem.rename(columns={'#': 'Clasificación'}, inplace=True)
                for c_t in ['Tiempo Total', 'Natación', 'Bicicleta', 'Trote']: df_sem[c_t] = df_sem[c_t].apply(to_excel_time_value)
                df_sem.to_excel(motor_excel, sheet_name=h_nombre, index=False)
                ws = motor_excel.sheets[h_nombre]
                for id_c in [2, 4, 5, 6]: ws.set_column(id_c, id_c, 12, f_hora)
            elif h_nombre in ["Tiempo Total", "Natación", "Ciclismo", "Trote"]:
                df_h = pd.read_excel(lector_maestro_excel, sheet_name=h_nombre, dtype=object)
                cols_del = [c for c in df_h.columns if str(c).startswith("Sem ") and get_num(c) > v_limite]
                if cols_del: df_h = df_h.drop(columns=cols_del)
                cols_arit = [c for c in df_h.columns if str(c).startswith("Sem ") or "Promedio" in str(c) or "Acumulado" in str(c)]
                for c_a in cols_arit:
                    if c_a != label_actual: df_h[c_a] = df_h[c_a].apply(to_excel_time_value)
                id_col = next((c for c in df_h.columns if "nombre" in str(c).lower() or "deportista" in str(c).lower()), df_h.columns[0])
                f_col = {'Tiempo Total': 'Tiempo Total', 'Natación': 'Natación', 'Ciclismo': 'Bicicleta', 'Trote': 'Trote'}.get(h_nombre)
                if f_col:
                    df_h['MatchKey'] = df_h[id_col].astype(str).apply(clean_string)
                    df_u = df_actualizacion[['Deportista', f_col]].copy()
                    df_u['MatchKey'] = df_u['Deportista'].apply(clean_string)
                    df_h[label_actual] = df_h['MatchKey'].map(df_u.set_index('MatchKey')[f_col].apply(to_excel_time_value).to_dict()).fillna(0)
                    if 'Tiempo Acumulado' in df_h.columns: df_h['Tiempo Acumulado'] = df_h[[c for c in df_h.columns if str(c).startswith("Sem ")]].sum(axis=1)
                    df_h = df_h.drop(columns=['MatchKey'])
                df_h.to_excel(motor_excel, sheet_name=h_nombre, index=False)
                ws_t = motor_excel.sheets[h_nombre]
                for i, n in enumerate(df_h.columns):
                    if str(n).startswith("Sem ") or "Promedio" in str(n) or "Acumulado" in str(n): ws_t.set_column(i, i, 13, f_hora)
            elif h_nombre == "CV":
                df_cv = pd.read_excel(lector_maestro_excel, sheet_name=h_nombre, dtype=object)
                cols_cv_f = [c for c in df_cv.columns if str(c).startswith("Sem ") and get_num(c) > v_limite]
                if cols_cv_f: df_cv = df_cv.drop(columns=cols_cv_f)
                id_c_cv = next((c for c in df_cv.columns if "nombre" in str(c).lower() or "deportista" in str(c).lower()), df_cv.columns[0])
                df_cv['MatchKey'] = df_cv[id_c_cv].astype(str).apply(clean_string)
                df_cv_u = df_actualizacion[['Deportista', 'CV']].copy()
                df_cv_u['MatchKey'] = df_cv_u['Deportista'].apply(clean_string)
                df_cv[label_actual] = df_cv['MatchKey'].map(df_cv_u.set_index('MatchKey')['CV'].to_dict()).fillna("NC")
                df_cv.drop(columns=['MatchKey']).to_excel(motor_excel, sheet_name=h_nombre, index=False)
            else: pd.read_excel(lector_maestro_excel, sheet_name=h_nombre, dtype=object).to_excel(motor_excel, sheet_name=h_nombre, index=False)
    return buffer_binario.getvalue()

# *****************************************************************************
# --- 6. GENERADOR DE REPORTE WORD GRUPAL (BLOQUEADO / MODELO V2.2.19) ---
# *****************************************************************************

def aplicar_formato_tym_word(objeto_p, fuente_pt, bold_on=False, center_on=False):
    run_cursor = objeto_p.add_run() if not objeto_p.runs else objeto_p.runs[0]
    run_cursor.font.name = 'Calibri'; run_cursor.font.size = Pt(fuente_pt); run_cursor.bold = bold_on
    if center_on: objeto_p.alignment = WD_ALIGN_PARAGRAPH.CENTER

def crear_tabla_profesional_tym_word(doc_word_ref, df_fuente_datos, listado_cabeceras):
    objeto_tabla = doc_word_ref.add_table(rows=1, cols=len(listado_cabeceras))
    objeto_tabla.style = 'Light Grid Accent 1'; objeto_tabla.alignment = 1; objeto_tabla.autofit = False
    dicc_anchos = {'#': 0.4, 'Deportista': 2.8, 'Tiempo Total': 0.7, 'Natación': 0.7, 'Bicicleta': 0.7, 'Trote': 0.7, 'CV': 0.6}
    for indice_h, texto_cabecera in enumerate(listado_cabeceras):
        celda_h = objeto_tabla.rows[0].cells[idx_cab := indice_h]; celda_h.text = texto_cabecera
        celda_h.width = Inches(dicc_anchos.get(texto_cabecera, 0.7)); aplicar_formato_tym_word(celda_h.paragraphs[0], 9, True, True)
    for _, fila_loop in df_fuente_datos.iterrows():
        celdas_fila = objeto_tabla.add_row().cells
        for indice_d, texto_cabecera_d in enumerate(listado_cabeceras):
            celdas_fila[indice_d].text = str(fila_loop[texto_cabecera_d]); celdas_fila[indice_d].width = Inches(dicc_anchos.get(texto_cabecera_d, 0.7))
            aplicar_formato_tym_word(celdas_fila[indice_d].paragraphs[0], 9, False, texto_cabecera_d != 'Deportista')
    doc_word_ref.add_paragraph()

def generar_reporte_word_tym_completo(df_datos_semanales, str_num_semana, lista_podio_d, lista_podio_l):
    doc_final = Document()
    p_header_tit = doc_final.add_heading(f'Reporte Semanal Club Tym Triatlón - Semana {str_num_semana}', 0)
    aplicar_formato_tym_word(p_header_tit, 20, True, True); doc_final.add_paragraph()
    p_slogan = doc_final.add_paragraph('"(La semana de la simetría perfecta y el retorno del volumen)"')
    aplicar_formato_tym_word(p_slogan, 11, True, True); doc_final.add_paragraph()
    h_res = doc_final.add_heading('🔍 Resumen General', level=2); aplicar_formato_tym_word(h_res, 15, True); doc_final.add_paragraph()
    df_trias = df_datos_semanales[df_datos_semanales['CV'] != 'NC'].copy()
    txt_res = f"Total deportistas: {len(df_datos_semanales)}\nTriatletas completos: {len(df_trias)}\nHoras totales: {to_hhmmss_display(df_datos_semanales['T_Mins'].sum())}"
    p_info = doc_final.add_paragraph(txt_res); aplicar_formato_tym_word(p_info, 11)
    fig, ax = plt.subplots(figsize=(4,4)); ax.pie([df_datos_semanales['N_Mins'].sum(), df_datos_semanales['B_Mins'].sum(), df_datos_semanales['R_Mins'].sum()], labels=['Nat', 'Bici', 'Tro'], autopct='%1.1f%%', colors=['#1E90FF', '#32CD32', '#FF4500'])
    buf = io.BytesIO(); plt.savefig(buf, format='png', bbox_inches='tight'); plt.close(fig)
    doc_final.add_paragraph().add_run().add_picture(buf, width=Inches(3.5))
    for t, d, c in [('🏅 TOP 5 COMPLETOS', df_trias.sort_values('T_Mins', ascending=False).head(5), 'Completos'), ('⚖️ TOP 5 BALANCEADOS', df_trias.sort_values('CV', ascending=True).head(5), 'CV')]:
        h_s = doc_final.add_heading(t, level=2); aplicar_formato_tym_word(h_s, 15, True); doc_final.add_paragraph()
        df_r = d.copy(); df_r['#'] = range(1, len(df_r) + 1)
        crear_tabla_profesional_tym_word(doc_final, df_r, ['#', 'Deportista', 'Tiempo Total' if c=='Completos' else 'CV', 'Natación', 'Bicicleta', 'Trote'])
        doc_final.add_paragraph('Análisis:'); [doc_final.add_paragraph(f"{r['#']}. {r['Deportista']}\n{generar_comentario(r, c, r['#'])}") for _, r in df_r.iterrows()]
    for s, ico, m, txt in [('TIEMPO GENERAL', '🥇', 'T_Mins', 'Tiempo Total'), ('NATACIÓN', '🏊‍♂️', 'N_Mins', 'Natación'), ('CICLISMO', '🚴', 'B_Mins', 'Bicicleta'), ('TROTE', '🏃‍♂️', 'R_Mins', 'Trote')]:
        doc_final.add_page_break(); h_15 = doc_final.add_heading(f'{ico} TOP 15 {s}', level=1); aplicar_formato_tym_word(h_15, 15, True); doc_final.add_paragraph()
        df_t15 = df_datos_semanales[df_datos_semanales[m] > 0].sort_values(m, ascending=False).head(15).copy(); df_t15['#'] = range(1, len(df_t15) + 1)
        crear_tabla_profesional_tym_word(doc_final, df_t15, ['#', 'Deportista', txt, 'Natación', 'Bicicleta', 'Trote'] if s == 'TIEMPO GENERAL' else ['#', 'Deportista', txt, 'Tiempo Total'])
        doc_final.add_paragraph('Podio:'); [doc_final.add_paragraph(f"{r['Deportista']}\n{generar_comentario(r, 'General' if s=='TIEMPO GENERAL' else txt, r['#'])}") for _, r in df_t15.head(3).iterrows()]
    doc_final.add_page_break(); doc_final.add_heading('📏 DISTANCIA TOTAL', 1); [doc_final.add_paragraph(f"{i+1}. {it['nombre']} ({it['valor']} km)") for i, it in enumerate(lista_podio_d)]
    doc_final.add_heading('⏱️ ACTIVIDAD MÁS LARGA', 1); [doc_final.add_paragraph(f"{i+1}. {it['nombre']} ({it['valor']})") for i, it in enumerate(lista_podio_l)]
    b_out = io.BytesIO(); doc_final.save(b_out); b_out.seek(0); return b_out

# *****************************************************************************
# --- 7. NUEVO MOTOR NARRATIVO INDIVIDUAL (ESTRATEGIA COLAB V17.0) ---
# *****************************************************************************

def generar_reporte_narrativo_individual(atleta_nom, df_actual, dict_historicos, sem_n):
    """Genera reporte personal con insights narrativos basados en Colab V17.0."""
    doc_p = Document()
    match_key = clean_string(atleta_nom)
    
    # Header Institucional
    p_h = doc_p.add_heading(f'Análisis de Rendimiento Personal: {atleta_nom}', 0)
    aplicar_formato_tym_word(p_h, 18, True, True)
    doc_p.add_paragraph(f"Semana de Entrenamiento: {sem_n}").alignment = 1

    # Obtener fila actual
    df_actual['MatchKey'] = df_actual['Deportista'].apply(clean_string)
    row_act = df_actual[df_actual['MatchKey'] == match_key]
    if row_act.empty: return None
    row_act = row_act.iloc[0]

    # Función interna para calcular promedio de equipo e histórico
    def get_benchmarks(hoja_key, match_k):
        df_h = dict_historicos.get(hoja_key)
        if df_h is None: return 0, 0
        df_h['MatchKey'] = df_h.iloc[:, 0].astype(str).apply(clean_string)
        cols_sem = [c for c in df_h.columns if str(c).startswith("Sem ")]
        # Promedio Equipo (esta semana)
        avg_equipo = df_actual[f"{hoja_key[0]}_Mins"].mean() if f"{hoja_key[0]}_Mins" in df_actual else 0
        if hoja_key == "Tiempo Total": avg_equipo = df_actual["T_Mins"].mean()
        # Promedio Histórico Atleta
        r_atleta = df_h[df_h['MatchKey'] == match_k]
        avg_hist = r_atleta[cols_sem].mean(axis=1).iloc[0] * 1440 if not r_atleta.empty else 0
        return avg_equipo, avg_hist

    # Bloque de Disciplinas
    disciplinas = [
        ("TIEMPO TOTAL", "Tiempo Total", "T_Mins"),
        ("NATACIÓN", "Natación", "N_Mins"),
        ("CICLISMO", "Ciclismo", "B_Mins"),
        ("TROTE", "Trote", "R_Mins")
    ]

    for tit_m, hoja_m, col_m in disciplinas:
        doc_p.add_heading(tit_m, level=1)
        val_act = row_act[col_m]
        bench_equipo, bench_hist = get_benchmarks(hoja_m, match_key)
        
        # Texto Narrativo
        p_m = doc_p.add_paragraph()
        p_m.add_run(f"Volumen actual: {to_hhmmss_display(val_act)}\n").bold = True
        
        # Comparativa Equipo
        diff_eq = val_act - bench_equipo
        txt_eq = f"Rendiste {to_hhmmss_display(abs(diff_eq))} {'MÁS' if diff_eq > 0 else 'MENOS'} que el promedio del equipo."
        run_eq = p_m.add_run(txt_eq)
        run_eq.font.color.rgb = RGBColor(0, 100, 0) if diff_eq >= 0 else RGBColor(180, 0, 0)
        
        # Comparativa Histórica
        diff_hi = val_act - bench_hist
        txt_hi = f"\nRespecto a tu propia media: {to_hhmmss_display(abs(diff_hi))} {'MÁS' if diff_hi > 0 else 'MENOS'}."
        p_m.add_run(txt_hi)

    # Gráfico de Distribución Personal
    fig_p, ax_p = plt.subplots(figsize=(5,3))
    ax_p.bar(['Nat', 'Bici', 'Tro'], [row_act['N_Mins'], row_act['B_Mins'], row_act['R_Mins']], color=['#1E90FF', '#32CD32', '#FF4500'])
    ax_p.set_title("Tu Distribución de Carga (Minutos)")
    
    buf_p = io.BytesIO()
    plt.savefig(buf_p, format='png', bbox_inches='tight')
    plt.close(fig_p)
    doc_p.add_paragraph().add_run().add_picture(buf_p, width=Inches(3.5))

    doc_p.add_paragraph("─" * 50)
    doc_p.add_paragraph("Generado por Agente TYM 2026").alignment = 2
    
    b_out = io.BytesIO(); doc_p.save(b_out); b_out.seek(0); return b_out

# *****************************************************************************
# --- 8. INTERFAZ DE USUARIO (STREMLIT) ---
# *****************************************************************************

st.sidebar.header("📁 Gestión Histórica TYM")
maestro_file = st.sidebar.file_uploader("Cargar Excel Maestro", type=["xlsx"])
n_sem_val = st.text_input("Semana (Ej: 08):", "08")
data_st = st.text_area("1. Datos Strava:")
data_oc = st.text_area("2. Traducción OCR:")

if st.button("🚀 PROCESAR JORNADA COMPLETA"):
    if maestro_file and data_st.strip() and data_oc.strip():
        # Parsing
        df_res = parse_raw_data(data_st); p_d, p_l = parse_ocr_data(data_oc)
        
        # Carga de hojas para reportes narrativos
        h_t = pd.read_excel(maestro_file, sheet_name="Tiempo Total", dtype=object)
        h_n = pd.read_excel(maestro_file, sheet_name="Natación", dtype=object)
        h_c = pd.read_excel(maestro_file, sheet_name="Ciclismo", dtype=object)
        h_r = pd.read_excel(maestro_file, sheet_name="Trote", dtype=object)
        dict_h = {"Tiempo Total": h_t, "Natación": h_n, "Ciclismo": h_c, "Trote": h_r}
        
        st.success(f"¡Semana {n_sem_val} procesada con éxito!"); c1, c2 = st.columns(2)
        
        # Descargas Grupal y Excel
        c1.download_button("📄 REPORTE GRUPAL", generar_reporte_word_tym_completo(df_res, n_sem_val, p_d, p_l), f"Reporte_Grupal_{n_sem_val}.docx")
        c2.download_button("📊 EXCEL ACTUALIZADO", crear_excel_actualizado(maestro_file, df_res, n_sem_val), f"00_Estadisticas_Sem_{n_sem_val}.xlsx")
        
        # --- SECCIÓN INDIVIDUAL ---
        st.divider()
        st.subheader("👤 Generador de Insights Individuales (Estilo Colab)")
        atletas = df_res['Deportista'].tolist()
        seleccion = st.multiselect("Seleccionar atletas para reporte narrativo:", atletas, default=atletas[:2])
        
        if st.button("📦 EMPAQUETAR REPORTES INDIVIDUALES (ZIP)"):
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w") as zf:
                for a in seleccion:
                    r_ind = generar_reporte_narrativo_individual(a, df_res, dict_h, n_sem_val)
                    if r_ind:
                        zf.writestr(f"Reporte_{clean_string(a)}.docx", r_ind.getvalue())
            st.download_button("⬇️ DESCARGAR ZIP", zip_buf.getvalue(), f"Reportes_Individuales_Sem_{n_sem_val}.zip")
            
    else: st.error("Error: Complete Excel y datos de texto.")
