import streamlit as st
import pandas as pd
import numpy as np
import re
import io
import unicodedata
from datetime import time

# --- 1. CONFIGURACIÓN (BLINDADO) ---
st.set_page_config(page_title="Plataforma TYM 2026", page_icon="🏆", layout="wide")
st.title("🏆 Gestión de Estadísticas - Club TYM")

# --- 2. UTILIDADES DE PROCESAMIENTO (BLINDADO) ---
def clean_string(text):
    if not text: return ""
    text = str(text).strip().upper()
    return "".join(c for c in unicodedata.normalize('NFKD', text) if not unicodedata.combining(c))

def to_mins(t_str):
    if pd.isna(t_str) or str(t_str).strip() in ['--:--', '0', '', '00:00:00']: return 0
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

# --- 3. PARSER DE DATOS PARA LA HOJA SEMANAL ---
def parse_raw_data_for_excel(raw_text):
    data = []
    rank = 1
    raw_text = raw_text.replace('\xa0', ' ')
    for line in raw_text.strip().split('\n'):
        if not line or 'Deportista' in line: continue
        try:
            times = re.findall(r'(\d+h\s*\d*min|\d+h|\d+min|--:--)', line)
            if not times: continue
            name = re.sub(r'^\d+\s*', '', line[:line.find(times[0])]).strip()
            t_m, s_m, b_m, r_m = to_mins(times[0]), to_mins(times[1] if len(times)>1 else 0), to_mins(times[2] if len(times)>2 else 0), to_mins(times[3] if len(times)>3 else 0)
            cv = "NC" if 0 in [s_m, b_m, r_m] else round(np.std([s_m, b_m, r_m])/np.mean([s_m, b_m, r_m]), 4)
            
            # Captura de actividades (número después del tiempo total)
            acts_match = re.search(r'\d+', line[line.find(times[0])+len(times[0]):])
            acts = int(acts_match.group()) if acts_match else 0
            
            data.append({
                'Clasificación': rank,
                'Deportista': name,
                'Tiempo Total': to_hhmmss(t_m),
                'Actividades': acts,
                'Natación': to_hhmmss(s_m),
                'Bicicleta': to_hhmmss(b_m),
                'Trote': to_hhmmss(r_m),
                'CV': cv
            })
            rank += 1
        except: pass
    return pd.DataFrame(data)

# --- 4. FUNCIÓN EXCLUSIVA: CREAR HOJA SEM 08 ---
def crear_hoja_semanal(archivo_maestro, df_semana, num_sem):
    xls = pd.ExcelFile(archivo_maestro)
    hojas_existentes = xls.sheet_names
    nombre_nueva_hoja = f"Sem {num_sem.strip()}"
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Copiamos todas las hojas existentes sin cambios
        for hoja in hojas_existentes:
            # Si por error ya existe la Sem 08, la saltamos para sobrescribirla con la nueva
            if hoja == nombre_nueva_hoja: continue
            df_temp = pd.read_excel(xls, sheet_name=hoja)
            df_temp.to_excel(writer, sheet_name=hoja, index=False)
        
        # AÑADIMOS EXCLUSIVAMENTE LA NUEVA HOJA SEM 08
        df_semana.to_excel(writer, sheet_name=nombre_nueva_hoja, index=False)
        
    return output.getvalue()

# --- 5. INTERFAZ ENFOCADA ---
st.sidebar.header("📁 Carga de Archivo Histórico")
archivo_maestro = st.sidebar.file_uploader("Subir 00 Estadísticas TYM (xlsx)", type=["xlsx"])

sem_num = st.text_input("Número de semana a crear (ej: 08):", "08")
raw_data = st.text_area("Pegar datos de la semana para el Excel:")

if st.button("CREAR HOJA SEMANAL EN EXCEL"):
    if raw_data.strip() and archivo_maestro:
        df_sem = parse_raw_data_for_excel(raw_data)
        excel_actualizado = crear_hoja_semanal(archivo_maestro, df_sem, sem_num)
        
        st.success(f"Hoja '{sem_num}' añadida correctamente al archivo.")
        st.download_button(
            label="📊 DESCARGAR EXCEL CON HOJA NUEVA",
            data=excel_actualizado,
            file_name=f"00_Estadisticas_TYM_con_Sem_{sem_num}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("Por favor, sube el archivo Excel y pega los datos de la semana.")
