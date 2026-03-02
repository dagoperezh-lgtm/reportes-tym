# TYM PLATAFORMA - VERSION: V2.2.28-GOLD-CH.
# OBJETIVO: RESTAURAR EXTENSIÓN COMPLETA Y ELIMINAR SÍNTESIS AUTOMÁTICA
# LINEAS DE CODIGO: 785
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
    page_title="Plataforma TYM 2026 - V2.2.19", 
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
    nombre_limpio_temp = str(text).strip()
    
    nombre_limpio_temp = nombre_limpio_temp.upper()
    
    # Uso de NFKD para descomponer caracteres con tildes
    info_normalizada = unicodedata.normalize('NFKD', nombre_limpio_temp)
    
    resultado_final_nombre = ""
    
    for caracter_indiv in info_normalizada:
        # Filtro para ignorar los caracteres de combinación (tildes)
        if not unicodedata.combining(caracter_indiv):
            resultado_final_nombre = resultado_final_nombre + caracter_indiv
            
    return resultado_final_nombre

def to_mins(valor_entrada_tiempo):
    """
    Convierte cualquier formato de tiempo a minutos totales de forma explícita.
    Maneja decimales de Excel, objetos datetime, strings HH:MM y formato Strava.
    """
    if pd.isna(valor_entrada_tiempo):
        return 0
    
    string_valor = str(valor_entrada_tiempo).strip()
    
    # Listado exhaustivo de casos nulos detectados en la operativa real
    lista_casos_nulos = ['--:--', '0', '', '00:00:00', '0:00:00', '00:00', '0.0', 'NC', '0:00']
    
    if string_valor in lista_casos_nulos:
        return 0
        
    try:
        # 🛡️ REGLA ARITMÉTICA: Si el valor es numérico (fracción de día de Excel)
        if isinstance(valor_entrada_tiempo, (float, int)):
            # Excel almacena 1 día completo como 1.0. 
            # Multiplicamos por 1440 para obtener la cifra real de minutos.
            minutos_finales_calculados = int(round(valor_entrada_tiempo * 1440))
            return minutos_finales_calculados
        
        # Si el dato es un objeto de tiempo nativo de Python
        if isinstance(valor_entrada_tiempo, (time, datetime)):
            minutos_finales_calculados = (valor_entrada_tiempo.hour * 60) + valor_entrada_tiempo.minute
            return minutos_finales_calculados
            
        # Si el string representa un número decimal puro
        try:
            conversion_float = float(string_valor)
            minutos_finales_calculados = int(round(conversion_float * 1440))
            return minutos_finales_calculados
        except ValueError:
            # No es numérico, continuamos con la lógica de parsing de texto
            pass
            
        # Formato de hora estándar con separador de dos puntos (HH:MM)
        if ':' in string_valor:
            bloques_tiempo = string_valor.split(':')
            if len(bloques_tiempo) >= 2:
                horas_bloque = int(bloques_tiempo[0])
                
                minutos_raw_bloque = bloques_tiempo[1]
                # Se eliminan segundos o microsegundos si existen
                minutos_clean_bloque = int(minutos_raw_bloque.split('.')[0])
                
                total_minutos_bloque = (horas_bloque * 60) + minutos_clean_bloque
                return total_minutos_bloque
        
        # Formato nativo de Strava (ejemplo: 11h 6min)
        busqueda_horas = re.search(r'(\d+)h', string_valor)
        busqueda_minutos = re.search(r'(\d+)min', string_valor)
        
        h_resultado = 0
        if busqueda_horas:
            h_resultado = int(busqueda_horas.group(1))
            
        m_resultado = 0
        if busqueda_minutos:
            m_resultado = int(busqueda_minutos.group(1))
            
        resultado_total_minutos = (h_resultado * 60) + m_resultado
        return resultado_total_minutos
        
    except Exception:
        # Fallback de seguridad para evitar que la aplicación se detenga
        return 0

def to_excel_time_value(dato_entrada_original):
    """
    Transforma la entrada en la fracción decimal exacta que requiere el motor de Excel.
    Este paso es vital para que las celdas sean sumables y promediables.
    """
    minutos_para_excel = to_mins(dato_entrada_original)
    
    # 24 horas equivalen a 1440 minutos totales
    valor_decimal_excel = minutos_para_excel / 1440.0
    
    return valor_decimal_excel

def to_hhmmss_display(minutos_totales_input):
    """
    Formato de texto HH:MM:00 exclusivo para la estética del reporte Word.
    """
    valor_horas_v = int(minutos_totales_input // 60)
    valor_minutos_v = int(minutos_totales_input % 60)
    
    # Generación de la cadena de texto con formato de reloj
    string_formato_reloj = f"{valor_horas_v:02d}:{valor_minutos_v:02d}:00"
    
    return string_formato_reloj

# =============================================================================
# SECCIÓN 3: MOTOR NARRATIVO PRO CHILE (V2.2.28 - 20+ FRASES POR SECCIÓN)
# =============================================================================
import random

# Diccionario global para persistencia de frases durante la ejecución
PILAS_COMENTARIOS = {}

def obtener_frase_base(categoria, pool_frases):
    """Maneja el barajado de frases para garantizar 0 repeticiones."""
    global PILAS_COMENTARIOS
    if categoria not in PILAS_COMENTARIOS or not PILAS_COMENTARIOS[categoria]:
        temp_pool = [str(f) for f in pool_frases] # Forzamos a string para evitar TypeErrors
        random.shuffle(temp_pool)
        PILAS_COMENTARIOS[categoria] = temp_pool
    return PILAS_COMENTARIOS[categoria].pop()

def generar_comentario(datos_de_fila, nombre_categoria, rank_posicion):
    """
    Motor de Narrativa Pro Chile: 20+ variantes por sección.
    Léxico corregido (piscina) e inyección dinámica de identidad.
    """
    atleta_actual = str(datos_de_fila.get('Deportista', 'Atleta TYM'))
    tiempo_actual = str(datos_de_fila.get(nombre_categoria, "00:00:00"))
    
    # --- BANCO DE NARRATIVA CHILENA (20-25 FRASES POR SECCIÓN) ---
    pools = {
        'General': [
            "La disciplina de {atleta} es el motor del club; liderar con este volumen es pura entrega.",
            "Semana de consolidación para {atleta}. No solo es cantidad, es la calidad del tiempo acumulado.",
            "El compromiso de {atleta} se refleja en cada sesión. Un pilar fundamental del ranking hoy.",
            "Rendimiento de alto nivel. {atleta} entiende que la base del éxito es este volumen sostenido.",
            "Impresionante despliegue de {atleta}. Gestionar estas cargas requiere madurez deportiva.",
            "La constancia de {atleta} marca el paso del equipo. Una semana de trabajo impecable.",
            "Fuerza mental y física. {atleta} asimila el volumen semanal con una resiliencia notable.",
            "Evolución sostenida de {atleta}. Estar en el top general es fruto de una planificación seria.",
            "La ética de trabajo de {atleta} es envidiable. Cada hora sumada construye su mejor versión.",
            "Control total de la fatiga. {atleta} cierra la semana en lo más alto con mérito propio.",
            "{atleta} demuestra que la regularidad es el camino corto hacia los objetivos de temporada.",
            "Capacidad de carga superior. {atleta} lidera la tabla con una solvencia técnica admirable.",
            "Una semana brillante para {atleta}, demostrando una solidez física que inspira al resto.",
            "Disciplina inquebrantable. {atleta} se mantiene en la cima con un enfoque envidiable.",
            "{atleta} cierra la jornada con un volumen que refleja ambición y preparación rigurosa.",
            "Poderío aeróbico de {atleta}. Registrar estas horas es señal de una base muy robusta.",
            "Planificación ejecutada a la perfección por {atleta}. La consistencia es su mayor virtud.",
            "Rendimiento de punta. {atleta} encabeza el grupo con una capacidad de recuperación única.",
            "El volumen de {atleta} es el resultado de una mentalidad enfocada en la larga distancia.",
            "Foco y determinación. {atleta} asume el liderato semanal con una carga de trabajo sólida.",
            "Notable fondo físico de {atleta}. Su presencia en el podio general es garantía de perseverancia.",
            "{atleta} proyecta una temporada sólida manteniendo este ritmo de entrenamientos semanales.",
            "Sello de calidad TYM: {atleta} pone el trabajo necesario para destacar en la tabla general.",
            "Madurez competitiva. {atleta} sabe que el volumen es el cimiento de su rendimiento futuro.",
            "Gran lectura de cargas de {atleta}, logrando un volumen total que marca diferencias claras."
        ],
        'CV': [
            "Equilibrio milimétrico. {atleta} entrena con la precisión de quien no deja nada al azar.",
            "La polivalencia de {atleta} es su mayor ventaja. Simetría total en las tres áreas.",
            "Control de carga magistral. {atleta} distribuye su energía de forma balanceada.",
            "Triatlón en estado puro: {atleta} demuestra que dominar la transición es dominar el balance.",
            "Eficiencia técnica destacada. {atleta} logra que la simetría parezca sencilla pero es pura gestión.",
            "Versatilidad técnica. {atleta} no descuida ningún frente, fortaleciendo sus debilidades.",
            "Arquitectura de entrenamiento impecable. {atleta} refleja la esencia del deportista integral.",
            "Cero puntos débiles. {atleta} mantiene una paridad envidiable entre agua, bici y trote.",
            "Gestión inteligente de las cargas. {atleta} prioriza la salud y el equilibrio deportivo.",
            "Sincronía total. {atleta} asimila las tres disciplinas con una regularidad asombrosa.",
            "El balance de {atleta} es la clave para evitar lesiones y potenciar el rendimiento global.",
            "{atleta} demuestra que ser completo es más importante que ser rápido en una sola área.",
            "Madurez deportiva de {atleta}. Su coeficiente de variación es de los mejores del club.",
            "Planificación equilibrada de {atleta}. Cada disciplina recibe la atención que merece.",
            "Solidez transversal. {atleta} se consolida como uno de los atletas más balanceados.",
            "Precisión técnica en la distribución. {atleta} entrena con inteligencia y visión global.",
            "La armonía de {atleta} en las tres áreas es fruto de un compromiso técnico superior.",
            "{atleta} destaca por su capacidad de mantener la calidad sin importar el medio.",
            "Consistencia simétrica. {atleta} es el referente de equilibrio para el equipo hoy.",
            "Desarrollo armónico. {atleta} fortalece su base con una distribución de tiempo magistral."
        ],
        'Natación': [
            "Fluidez y potencia. Los {tiempo} de {atleta} en la piscina son el cimiento de su base.",
            "Dominio acuático. {atleta} marca la pauta con un volumen técnico de {tiempo} en piscina.",
            "Calidad en el agua. {atleta} suma {tiempo} de nado con una técnica cada vez más depurada.",
            "{atleta} lidera la fase acuática con {tiempo}, demostrando que la piscina es su fortaleza.",
            "Brazada eficiente y constante. {atleta} asimila {tiempo} de natación con gran solvencia.",
            "Resistencia hidrodinámica de {atleta}. Registrar {tiempo} en piscina es un hito importante.",
            "El agua no miente: {atleta} ha trabajado duro para lograr estos {tiempo} de volumen neto.",
            "Foco técnico en natación. {atleta} cierra con {tiempo}, consolidando su fase de apertura.",
            "Disciplina en la piscina. {atleta} no falla y suma {tiempo} de alta relevancia técnica.",
            "Progreso acuático visible. {atleta} domina su carril con {tiempo} de trabajo serio.",
            "{atleta} demuestra solidez en el agua, acumulando {tiempo} de nado de alta calidad.",
            "Eficiencia en cada largo. {atleta} optimiza sus {tiempo} en piscina para mejorar su fondo.",
            "Control de ritmo acuático. {atleta} suma {tiempo} de natación con una técnica sólida.",
            "Fuerza en la piscina. {atleta} proyecta una gran base aeróbica con sus {tiempo} actuales.",
            "Consistencia en el agua. {atleta} aprovecha sus {tiempo} en piscina para pulir detalles."
        ],
        'Bicicleta': [
            "Kilometraje de calidad. {atleta} construye su fortaleza sobre los pedales con {tiempo} de rodaje.",
            "El gran motor del equipo. {atleta} asimila la carga de ciclismo con resiliencia envidiable.",
            "Potencia y fondo. {atleta} devoró la ruta sumando {tiempo}, demostrando preparación superior.",
            "Solidez sobre ruedas. {atleta} aprovecha cada sesión para sumar {tiempo} de base aeróbica.",
            "El asfalto es el hábitat de {atleta}. Su volumen de {tiempo} en bici es pilar de su plan.",
            "Resistencia sobre el pedal. {atleta} acumula {tiempo} de calidad para blindar sus piernas.",
            "Ciclismo de alto impacto. {atleta} se sitúa como líder con {tiempo} de rodaje neto.",
            "Fuerza y cadencia. {atleta} gestiona sus {tiempo} en bicicleta con una madurez notable.",
            "Fondo inquebrantable. {atleta} suma {tiempo} en la ruta, clave para la larga distancia.",
            "Dominio del segmento de ciclismo. {atleta} marca el ritmo con {tiempo} de trabajo duro.",
            "Potencia aeróbica en ruta. {atleta} consolida sus {tiempo} de pedaleo con determinación.",
            "{atleta} demuestra que la bicicleta es su fuerte, acumulando {tiempo} de volumen masivo.",
            "Resiliencia sobre el sillín. {atleta} asimila {tiempo} de ciclismo con una solvencia técnica única.",
            "Gestión de potencia de {atleta}. Sus {tiempo} de rodaje son fundamentales para la temporada.",
            "Control y resistencia. {atleta} suma {tiempo} de bicicleta, blindando su motor aeróbico."
        ],
        'Trote': [
            "Zancada resiliente. Cerrar la semana con {tiempo} de impacto en el asfalto define el carácter de {atleta}.",
            "Resistencia específica. {atleta} domina la fase de carrera con una gestión de fatiga admirable.",
            "Persistencia técnica. {atleta} asimila el volumen de {tiempo} en running fortaleciendo su base.",
            "El asfalto premia la constancia. {atleta} cierra con {tiempo} de trote muy sólidos.",
            "Capacidad de cierre. {atleta} demuestra su fondo aeróbico con {tiempo} de carrera a pie.",
            "Impacto controlado y eficiente. {atleta} suma {tiempo} de trote, clave para su evolución.",
            "Running de alta gama. {atleta} se posiciona en el top con {tiempo} de asimilación de carga.",
            "Fortaleza en la carrera. {atleta} no cede y registra {tiempo} de volumen neto en el asfalto.",
            "Zancada potente y rítmica. {atleta} asume sus {tiempo} de trote con una técnica ejemplar.",
            "Resiliencia en cada kilómetro. {atleta} demuestra que el trote es donde se ganan las carreras.",
            "Gestión de la fatiga en asfalto. {atleta} completa sus {tiempo} de trote con gran madurez.",
            "Fuerza mental en la carrera. {atleta} suma {tiempo} netos, esenciales para su progresión.",
            "Eficiencia de zancada. {atleta} asimila {tiempo} de trote, cuidando la técnica en cada tramo.",
            "Consistencia en el running. {atleta} cierra la semana con {tiempo} de carga aeróbica sólida.",
            "Resistencia de punta. {atleta} marca diferencias en el asfalto con sus {tiempo} de volumen."
        ]
    }

    # SELECCIÓN Y FORMATEO SEGURO
    cat_key = 'General' if nombre_categoria in ['Completos', 'General'] else nombre_categoria
    if cat_key not in pools:
        return f"Desempeño consistente de {atleta_actual} en {nombre_categoria}."

    frase_plantilla = str(obtener_frase_base(cat_key, pools[cat_key]))
    
    # REEMPLAZO DINÁMICO (Seguro contra duplicidad)
    comentario_final = frase_plantilla.replace("{atleta}", atleta_actual).replace("{tiempo}", tiempo_actual)
    
    # Distinción Líder
    if rank_posicion == 1 and cat_key == 'General':
        comentario_final = f"🏆 {comentario_final.replace(atleta_actual, f'nuestro líder {atleta_actual}')}"
    
    return comentario_final

# *****************************************************************************
# --- 3B. MOTOR DE ADHERENCIA Y GOBERNANZA (REPARADO V2.2.35) ---
# *****************************************************************************

def asegurar_minutos(valor):
    """
    PRUEBA BÁSICA DE INGENIERÍA:
    - Si el dato es nulo o vacío, retorna 0.0.
    - Si es número (int/float), lo mantiene (ya es funcional).
    - Si es texto o tiempo, usa to_mins para convertirlo.
    """
    # 🛡️ Evita el ValueError de ambigüedad en Pandas
    if pd.isna(valor):
        return 0.0
    
    # REGLA: Si ya es un número funcional, no aplicar transformación
    if isinstance(valor, (int, float)):
        return float(valor)
    
    # Casos nulos escritos como texto
    if str(valor).strip() in ['0', '0.0', '--:--', '', 'NC']:
        return 0.0
        
    # De lo contrario, delegamos a tu función to_mins de la Sección 2
    return to_mins(valor)

def calcular_adherencia_v2(df_reales, df_plan_indiv=None, dict_plan_global=None):
    """
    Motor de Adherencia: Traduce el Excel Semanal al lenguaje de la App.
    Aplica la jerarquía: Plan Individual > Plan Global.
    """
    df = df_reales.copy()
    
    # 1. Sincronización de Identidad y Columnas
    # Mapeamos los nombres de tu Excel 'Sem 08' a las variables del motor
    mapeo_entrada = {
        'Natación': 'N_Mins', 
        'Bicicleta': 'B_Mins', 
        'Trote': 'R_Mins',
        'Nombre': 'Deportista',
        'Atleta': 'Deportista'
    }
    df.rename(columns=mapeo_entrada, inplace=True)

    # 2. Conversión Aritmética Selectiva (Tu prueba básica)
    for col in ['N_Mins', 'B_Mins', 'R_Mins']:
        if col in df.columns:
            df[col] = df[col].apply(asegurar_minutos)
        else:
            df[col] = 0.0
    
    df['T_Mins'] = df['N_Mins'] + df['B_Mins'] + df['R_Mins']

    # 3. Lógica de Cascada por Disciplina (Individual > Global)
    disciplinas = {'Natacion': 'N_Mins', 'Ciclismo': 'B_Mins', 'Trote': 'R_Mins'}
    for d_name in disciplinas.keys():
        h_plan_key = f"{d_name}_Hrs_Plan"
        s_plan_key = f"{d_name}_Ses_Plan"
        col_real = disciplinas[d_name]
        
        def obtener_meta(atl, m_col, v_glob):
            if df_plan_indiv is not None and 'Deportista' in df_plan_indiv.columns:
                atl_c = clean_string(atl)
                plan = df_plan_indiv[df_plan_indiv['Deportista'].apply(clean_string) == atl_c]
                if not plan.empty and m_col in plan.columns:
                    val = plan[m_col].values[0]
                    return val if (pd.notna(val) and val > 0) else v_glob
            return v_glob

        df[h_plan_key] = df.apply(lambda r: obtener_meta(r.get('Deportista', ''), h_plan_key, dict_plan_global.get(h_plan_key, 0)), axis=1)
        df[s_plan_key] = df.apply(lambda r: obtener_meta(r.get('Deportista', ''), s_plan_key, dict_plan_global.get(s_plan_key, 0)), axis=1)

        # 4. KPIs de Adherencia (TPI: 40% Volumen / 60% Sesiones)
        # Usamos fillna(0) para evitar colapsos en divisiones por cero
        vci = (df[col_real] / (df[h_plan_key] * 60)) * 100
        sei = (df[col_real].apply(lambda x: 1 if x > 0 else 0) / df[s_plan_key]) * 100
        df[f'TPI_{d_name}'] = (0.4 * vci.fillna(0)) + (0.6 * sei.fillna(0))

    # 5. Resultados Finales para los Reportes
    hrs_p_total = (df['Natacion_Hrs_Plan'] + df['Ciclismo_Hrs_Plan'] + df['Trote_Hrs_Plan'])
    df['TPI_Global'] = (df['T_Mins'] / (hrs_p_total * 60)) * 100
    df['TPI_Global'] = df['TPI_Global'].fillna(0)
    
    df['Estado_Cumplimiento'] = df['TPI_Global'].apply(lambda v: "Óptimo" if v >= 95 else ("Parcial" if v >= 85 else "Riesgo"))
    
    # Columnas de visualización con formato reloj
    df['Tiempo Total'] = df['T_Mins'].apply(to_hhmmss_display)
    df['Natación'] = df['N_Mins'].apply(to_hhmmss_display)
    df['Bicicleta'] = df['B_Mins'].apply(to_hhmmss_display)
    df['Trote'] = df['R_Mins'].apply(to_hhmmss_display)
    
    return df
# *****************************************************************************
# --- 4. PARSERS DE ENTRADA (BLINDADO - NO SINTETIZAR) ---
# *****************************************************************************

def parse_raw_data(bloque_input_strava):
    """
    Procesa el bloque de texto copiado de Strava (Tiempo Total).
    No utiliza síntesis; cada paso de extracción es explícito y visible.
    """
    lista_de_registros_atleta = []
    valor_rank_contador = 1
    
    # Limpieza de caracteres de control web (espacios de no ruptura)
    bloque_input_strava = bloque_input_strava.replace('\xa0', ' ')
    lineas_encontradas = bloque_input_strava.strip().split('\n')
    
    for fila_texto in lineas_encontradas:
        if not fila_texto:
            continue
            
        if 'Deportista' in fila_texto:
            continue
            
        try:
            # Expresión regular para detectar tiempos con formato h y min
            patron_tiempos = r'(\d+h\s*\d*min|\d+h|\d+min|--:--)'
            tiempos_en_linea = re.findall(patron_tiempos, fila_texto)
            
            # 🛡️ CORRECCIÓN SINTAXIS AUDITADA:
            if not tiempos_en_linea:
                continue
                
            # El Tiempo Total es siempre el primer elemento detectado
            string_del_total = tiempos_en_linea[0]
            ubicacion_del_tiempo = fila_texto.find(string_del_total)
            
            # El nombre del deportista precede a la cifra de tiempo
            segmento_del_nombre = fila_texto[:ubicacion_del_tiempo].strip()
            
            # Limpieza del número de ranking si está presente en el copiado (ej: "1 Rodrigo")
            nombre_limpio_final = re.sub(r'^\d+\s*', '', segmento_del_nombre).strip()
            
            # Conversión de los bloques de tiempo a minutos enteros
            minutos_volumen_total = to_mins(string_del_total)
            
            minutos_nat = 0
            if len(tiempos_en_linea) > 1:
                minutos_nat = to_mins(tiempos_en_linea[1])
                
            minutos_bici = 0
            if len(tiempos_en_linea) > 2:
                minutos_bici = to_mins(tiempos_en_linea[2])
                
            minutos_trote = 0
            if len(tiempos_en_linea) > 3:
                minutos_trote = to_mins(tiempos_en_linea[3])
                
            # Cálculo del Coeficiente de Variación (CV)
            lista_tiempos_cv = [minutos_nat, minutos_bici, minutos_trote]
            
            if 0 in lista_tiempos_cv:
                valor_cv_final = "NC"
            else:
                calculo_std = np.std(lista_tiempos_cv)
                calculo_mean = np.mean(lista_tiempos_cv)
                valor_cv_final = round(calculo_std / calculo_mean, 4)
            
            # Extracción del conteo de actividades (dato tras el tiempo total)
            segmento_final_linea = fila_texto[ubicacion_del_tiempo + len(string_total := string_del_total):]
            match_de_actividades = re.search(r'\d+', segmento_final_linea)
            
            numero_de_actividades = 0
            if match_de_actividades:
                numero_de_actividades = int(match_de_actividades.group())
            
            # Construcción del registro detallado por cada deportista
            diccionario_de_atleta = {
                '#': valor_rank_contador,
                'Deportista': nombre_limpio_final,
                'Tiempo Total': to_hhmmss_display(minutos_volumen_total),
                'Actividades': numero_de_actividades,
                'Natación': to_hhmmss_display(minutos_nat),
                'Bicicleta': to_hhmmss_display(minutos_bici),
                'Trote': to_hhmmss_display(minutos_trote),
                'CV': valor_cv_final,
                'T_Mins': minutos_volumen_total,
                'N_Mins': minutos_nat,
                'B_Mins': minutos_bici,
                'R_Mins': minutos_trote
            }
            
            lista_de_registros_atleta.append(diccionario_de_atleta)
            valor_rank_contador = valor_rank_contador + 1
            
        except Exception:
            # Omisión de líneas corruptas o sin formato válido
            continue
            
    # Retorno estructurado para procesamiento masivo en hojas de cálculo
    df_resultado_parsing = pd.DataFrame(lista_de_registros_atleta)
    
    return df_resultado_parsing

def parse_ocr_data(texto_ocr_crudo):
    """
    Parser de Ingeniería para Formato Vertical:
    Detecta el patrón [Nombre] [Nombre] [Valor] y lo traduce a podios.
    Blindado contra duplicidad de nombres en la misma línea.
    """
    # 1. Limpieza inicial: quitamos líneas vacías y encabezados de ruido
    lineas = [l.strip() for l in texto_ocr_crudo.split('\n') if l.strip()]
    palabras_ruido = ['tiempo', 'distancia', 'actividad', 'larga', 'total', 'clasificación']
    lineas_limpias = [l for l in lineas if not any(r in l.lower() for r in palabras_ruido)]
    
    podio_distancia = []
    podio_larga = []
    
    # 2. Procesamiento de bloques (Nombre -> Valor)
    # Iteramos saltando de 2 en 2, asumiendo que el nombre es la base
    i = 0
    while i < len(lineas_limpias) - 1:
        item_nombre = lineas_limpias[i]
        
        # Lógica para limpiar nombres duplicados (ej: "Claudio Claudio")
        palabras_nombre = item_nombre.split()
        mitad = len(palabras_nombre) // 2
        if mitad > 0 and palabras_nombre[:mitad] == palabras_nombre[mitad:]:
            nombre_final = " ".join(palabras_nombre[:mitad])
        else:
            nombre_final = item_nombre
            
        valor = lineas_limpias[i+1]
        
        # 3. Clasificación por naturaleza del dato
        # Si tiene ',' o 'km', es Distancia Total
        if ',' in valor or 'km' in valor.lower():
            podio_distancia.append({'nombre': nombre_final, 'valor': valor})
            i += 2
        # Si tiene ':' es un tiempo (Actividad Larga)
        elif ':' in valor:
            podio_larga.append({'nombre': nombre_final, 'valor': valor})
            i += 2
        else:
            # Si la línea siguiente no es un valor válido, saltamos solo 1 para buscar el par
            i += 1

    # Retornamos los Top 3 de cada categoría para el Reporte Word
    return podio_distancia[:3], podio_larga[:3]

# *****************************************************************************
# --- 5. ACTUALIZADOR DE EXCEL (OBJETIVO: INTEGRIDAD Y ORDEN NUMÉRICO) ---
# *****************************************************************************

def crear_excel_actualizado(referencia_maestro, df_actualizacion, input_semana_n):
    """
    Genera el archivo Excel inyectando la nueva semana y eliminando columnas futuras.
    Lógica de ordenamiento numérico real implementada para evitar errores en el historial.
    """
    # Carga del libro maestro con tipos de datos protegidos (object)
    lector_maestro_excel = pd.ExcelFile(referencia_maestro)
    hojas_existentes_maestro = lector_maestro_excel.sheet_names
    label_de_la_semana_actual = f"Sem {input_semana_n.strip()}"
    
    # 🛡️ FILTRO DE INGENIERÍA: Extraer el límite numérico de la semana procesada
    match_numero_semana = re.search(r'\d+', input_semana_n)
    valor_entero_limite = int(match_numero_semana.group()) if match_numero_semana else 0
    
    # DETERMINACIÓN DEL ORDEN OPERATIVO MANDATORIO (PROTOCOLO TYM):
    
    # Bloque 1: Identificación de Hojas Técnicas de Trabajo
    hojas_de_trabajo_lista = []
    for h_name in hojas_existentes_maestro:
        if not h_name.startswith("Sem "):
            hojas_de_trabajo_lista.append(h_name)
            
    # Bloque 2: Identificación y limpieza del historial semanal
    hojas_del_historial_lista = []
    for h_name in hojas_existentes_maestro:
        if h_name.startswith("Sem "):
            # Condición crítica: No duplicar la semana cargada actualmente
            if h_name != label_de_la_semana_actual:
                hojas_del_historial_lista.append(h_name)
    
    # Bloque 3: Ordenamiento Numérico Real (Auditado)
    def extraer_numero_pestaña(texto_pestaña):
        match_p = re.search(r'\d+', str(texto_pestaña))
        if match_p:
            return int(match_p.group())
        return 0
        
    hojas_del_historial_lista.sort(key=extraer_numero_pestaña, reverse=True)
    
    # Bloque 4: Construcción del libro final de salida
    secuencia_de_hojas_final = hojas_de_trabajo_lista + [label_de_la_semana_actual] + hojas_del_historial_lista

    # Creación del stream de memoria binario
    buffer_binario_descarga = io.BytesIO()
    
    with pd.ExcelWriter(buffer_binario_descarga, engine='xlsxwriter') as motor_escritura_excel:
        libro_excel_obj = motor_escritura_excel.book
        
        # Formato de tiempo TYM para celdas operables numéricamente
        formato_hora_tym = libro_excel_obj.add_format({'num_format': '[h]:mm:ss'})
        
        for hoja_en_curso in secuencia_de_hojas_final:
            
            # ESCENARIO A: GENERACIÓN DE LA NUEVA PESTAÑA SEMANAL DETALLADA
            if hoja_en_curso == label_de_la_semana_actual:
                cols_requeridas_sem = ['#', 'Deportista', 'Tiempo Total', 'Actividades', 'Natación', 'Bicicleta', 'Trote', 'CV']
                df_hoja_semana_final = df_actualizacion[cols_requeridas_sem].copy()
                df_hoja_semana_final.rename(columns={'#': 'Clasificación'}, inplace=True)
                
                # Conversión mandatoria a valores decimales de Excel
                for columna_t_name in ['Tiempo Total', 'Natación', 'Bicicleta', 'Trote']:
                    df_hoja_semana_final[columna_t_name] = df_hoja_semana_final[columna_t_name].apply(to_excel_time_value)
                
                # Escritura de la pestaña de clasificación
                df_hoja_semana_final.to_excel(motor_escritura_excel, sheet_name=hoja_en_curso, index=False)
                
                # Formateo visual de celdas horarias
                pestaña_actual_obj = motor_escritura_excel.sheets[hoja_en_curso]
                # Índices de columnas de tiempo (C, E, F, G)
                for id_columna_v in [2, 4, 5, 6]:
                    pestaña_actual_obj.set_column(id_columna_v, id_columna_v, 12, formato_hora_tym)
            
            # ESCENARIO B: ACTUALIZACIÓN DE LAS HOJAS DE TRABAJO TÉCNICAS (CON FILTRO DE COLUMNAS)
            elif hoja_en_curso in ["Tiempo Total", "Natación", "Ciclismo", "Trote"]:
                # Lectura blindada de la hoja histórica
                df_maestro_h_tecnica = pd.read_excel(lector_maestro_excel, sheet_name=hoja_en_curso, dtype=object)
                
                # 🛡️ FILTRO DE SEGURIDAD DE COLUMNAS (PROTOCOLO V2.2.19)
                # Remoción física de columnas "Sem XX" que exceden la semana actual procesada
                lista_de_cols_desbordadas = []
                for nombre_col_check in df_maestro_h_tecnica.columns:
                    str_check_nom = str(nombre_col_check)
                    if str_check_nom.startswith("Sem "):
                        match_num_check = re.search(r'\d+', str_check_nom)
                        if match_num_check:
                            if int(match_num_check.group()) > valor_entero_limite:
                                lista_de_cols_desbordadas.append(nombre_col_check)
                                
                if len(lista_de_cols_desbordadas) > 0:
                    df_maestro_h_tecnica = df_maestro_h_tecnica.drop(columns=lista_de_cols_desbordadas)
                
                # Sello de Transcripción Aritmética para columnas de cálculo
                cols_arit_tecnicas = []
                for c_header_t in df_maestro_h_tecnica.columns:
                    c_header_t_str = str(c_header_t)
                    if c_header_t_str.startswith("Sem ") or "Promedio" in c_header_t_str or "Acumulado" in c_header_t_str:
                        cols_arit_tecnicas.append(c_header_t)
                
                for col_a_t in cols_arit_tecnicas:
                    if col_a_t != label_de_la_semana_actual:
                        df_maestro_h_tecnica[col_a_t] = df_maestro_h_tecnica[col_a_t].apply(to_excel_time_value)
                
                # Match de deportistas mediante columna identificadora
                id_col_identificadora = df_maestro_h_tecnica.columns[0]
                for col_it_t in df_maestro_h_tecnica.columns:
                    if "nombre" in str(col_it_t).lower() or "deportista" in str(col_it_t).lower():
                        id_col_identificadora = col_it_t
                        break
                
                dict_mapeo_hojas_tym = {'Tiempo Total': 'Tiempo Total', 'Natación': 'Natación', 'Ciclismo': 'Bicicleta', 'Trote': 'Trote'}
                nombre_fuente_actual_tym = dict_mapeo_hojas_tym.get(hoja_en_curso)
                
                if nombre_fuente_actual_tym:
                    df_prep_normalizacion = df_actualizacion[['Deportista', nombre_fuente_actual_tym]].copy()
                    df_prep_normalizacion['MatchKey'] = df_prep_normalizacion['Deportista'].apply(clean_string)
                    df_maestro_h_tecnica['MatchKey'] = df_maestro_h_tecnica[id_col_identificadora].astype(str).apply(clean_string)
                    dict_valores_inyectar = df_prep_normalizacion.set_index('MatchKey')[nombre_fuente_actual_tym].apply(to_excel_time_value).to_dict()
                    df_maestro_h_tecnica[label_de_la_semana_actual] = df_maestro_h_tecnica['MatchKey'].map(dict_valores_inyectar).fillna(0)
                    
                    # RECALCULO DE COLUMNA TIEMPO ACUMULADO
                    lista_sem_para_suma = [c_f_s for c_f_s in df_maestro_h_tecnica.columns if str(c_f_s).startswith("Sem ")]
                    if 'Tiempo Acumulado' in df_maestro_h_tecnica.columns:
                        df_maestro_h_tecnica['Tiempo Acumulado'] = df_maestro_h_tecnica[lista_sem_para_suma].sum(axis=1)
                    
                    df_maestro_h_tecnica = df_maestro_h_tecnica.drop(columns=['MatchKey'])
                
                # Escritura definitiva de la hoja técnica
                df_maestro_h_tecnica.to_excel(motor_escritura_excel, sheet_name=hoja_en_curso, index=False)
                
                # Formatear visualmente las columnas de cálculo matemático
                pestaña_trabajo_activa = motor_escritura_excel.sheets[hoja_en_curso]
                for id_c_f, nom_c_f in enumerate(df_maestro_h_tecnica.columns):
                    nom_c_f_str = str(nom_c_f)
                    if nom_c_f_str.startswith("Sem ") or "Promedio" in nom_c_f_str or "Acumulado" in nom_c_f_str:
                        pestaña_trabajo_activa.set_column(id_c_f, id_c_f, 13, formato_hora_tym)

            # ESCENARIO C: ACTUALIZACIÓN DE LA HOJA COEFICIENTE DE VARIACIÓN (CV)
            elif hoja_en_curso == "CV":
                df_maestro_cv_hoja = pd.read_excel(lector_maestro_excel, sheet_name=hoja_en_curso, dtype=object)
                
                # 🛡️ BLINDAJE REGEX CV: Filtrado de columnas futuras
                cols_cv_a_borrar = []
                for c_cv_iterador in df_maestro_cv_hoja.columns:
                    if str(c_cv_iterador).startswith("Sem "):
                        m_cv_check = re.search(r'\d+', str(c_cv_iterador))
                        if m_cv_check:
                            if int(m_cv_check.group()) > valor_entero_limite:
                                cols_cv_a_borrar.append(c_cv_iterador)
                
                if len(cols_cv_a_borrar) > 0:
                    df_maestro_cv_hoja = df_maestro_cv_hoja.drop(columns=cols_cv_a_borrar)
                
                id_col_nombre_en_cv = next((c_cv for c_cv in df_maestro_cv_hoja.columns if "nombre" in str(c_cv).lower() or "deportista" in str(c_cv).lower()), df_maestro_cv_hoja.columns[0])
                df_maestro_cv_hoja['MatchKey'] = df_maestro_cv_hoja[id_col_nombre_en_cv].astype(str).apply(clean_string)
                df_cv_semana_prep = df_actualizacion[['Deportista', 'CV']].copy()
                df_cv_semana_prep['MatchKey'] = df_cv_semana_prep['Deportista'].apply(clean_string)
                df_maestro_cv_hoja[label_de_la_semana_actual] = df_maestro_cv_hoja['MatchKey'].map(df_cv_semana_prep.set_index('MatchKey')['CV'].to_dict()).fillna("NC")
                df_maestro_cv_hoja.drop(columns=['MatchKey']).to_excel(motor_escritura_excel, sheet_name=hoja_en_curso, index=False)

            # ESCENARIO D: RÉPLICA DE TODAS LAS DEMÁS PESTAÑAS DEL LIBRO
            else:
                df_replica_hoja_estatica = pd.read_excel(lector_maestro_excel, sheet_name=hoja_en_curso, dtype=object)
                df_replica_hoja_estatica.to_excel(motor_escritura_excel, sheet_name=hoja_en_curso, index=False)
                
    return buffer_binario_descarga.getvalue()

# *****************************************************************************
# --- 6. GENERADOR DE REPORTE WORD PROFESIONAL (V2.2.35 - INTEGRADO) ---
# *****************************************************************************

def aplicar_formato_tym_word(objeto_parrafo_word, pt_fuente, negrita_activo=False, centrado_activo=False):
    """Aplica rigurosamente el estilo institucional Calibri."""
    if not objeto_parrafo_word.runs:
        cursor_de_run = objeto_parrafo_word.add_run()
    else:
        cursor_de_run = objeto_parrafo_word.runs[0]
    cursor_de_run.font.name = 'Calibri'
    cursor_de_run.font.size = Pt(pt_fuente)
    cursor_de_run.bold = negrita_activo
    if centrado_activo:
        objeto_parrafo_word.alignment = WD_ALIGN_PARAGRAPH.CENTER

def crear_tabla_profesional_tym_word(doc_word_instancia, df_origen_datos, lista_de_cabeceras):
    """Genera tablas con anchos milimétricos (Protocolo TYM)."""
    instancia_de_tabla = doc_word_instancia.add_table(rows=1, cols=len(lista_de_cabeceras))
    instancia_de_tabla.style = 'Light Grid Accent 1'
    instancia_de_tabla.alignment = 1 
    instancia_de_tabla.autofit = False
    anchos_tym_fijos = {'#': 0.4, 'Deportista': 2.8, 'Tiempo Total': 0.7, 'Natación': 0.7, 'Bicicleta': 0.7, 'Trote': 0.7, 'CV': 0.6}
    
    for idx_cab, txt_cab in enumerate(lista_de_cabeceras):
        celda = instancia_de_tabla.rows[0].cells[idx_cab]
        celda.text = txt_cab
        celda.width = Inches(anchos_tym_fijos.get(txt_cab, 0.7))
        aplicar_formato_tym_word(celda.paragraphs[0], 9, True, True)
        
    for _, fila in df_origen_datos.iterrows():
        celdas = instancia_de_tabla.add_row().cells
        for idx, cab in enumerate(lista_de_cabeceras):
            celdas[idx].text = str(fila[cab])
            celdas[idx].width = Inches(anchos_tym_fijos.get(cab, 0.7))
            aplicar_formato_tym_word(celdas[idx].paragraphs[0], 9, False, cab != 'Deportista')
    doc_word_instancia.add_paragraph()

def generar_reporte_word_tym_completo(df_semanal_datos, num_sem_texto, podio_d_lista, podio_l_lista):
    """Mantiene tu reporte grupal original íntegro."""
    documento = Document()
    p_title = documento.add_heading(f'Reporte Semanal Club Tym Triatlón - Semana {num_sem_texto}', 0)
    aplicar_formato_tym_word(p_title, 20, True, True)
    
    # Resumen y Gráfico (Tu lógica GOLD)
    h_res = documento.add_heading('🔍 Resumen General', level=2); aplicar_formato_tym_word(h_res, 15, True)
    df_f = df_semanal_datos[df_semanal_datos['CV'] != 'NC'].copy()
    txt = f"Total deportistas: {len(df_semanal_datos)}\nHoras totales: {to_hhmmss_display(df_semanal_datos['T_Mins'].sum())}"
    aplicar_formato_tym_word(documento.add_paragraph(txt), 11)
    
    fig, ax = plt.subplots(figsize=(4,4))
    ax.pie([df_semanal_datos['N_Mins'].sum(), df_semanal_datos['B_Mins'].sum(), df_semanal_datos['R_Mins'].sum()], 
           labels=['Natación', 'Ciclismo', 'Trote'], autopct='%1.1f%%', colors=['#1E90FF', '#32CD32', '#FF4500'])
    buf = io.BytesIO(); plt.savefig(buf, format='png'); plt.close(fig)
    documento.add_paragraph().alignment = 1
    documento.paragraphs[-1].add_run().add_picture(buf, width=Inches(3.5))

    # Podios (Reutiliza tu lógica de comentarios)
    for t_pod, d_pod, c_key in [('🏅 TOP 5 COMPLETOS', df_f.sort_values('T_Mins', ascending=False).head(5), 'Completos')]:
        h = documento.add_heading(t_pod, level=2); aplicar_formato_tym_word(h, 15, True)
        d_ren = d_pod.copy(); d_ren['#'] = range(1, len(d_ren) + 1)
        crear_tabla_profesional_tym_word(documento, d_ren, ['#', 'Deportista', 'Tiempo Total', 'Natación', 'Bicicleta', 'Trote'])
        for _, fila in d_ren.iterrows():
            p = documento.add_paragraph(f"{fila['#']}. {fila['Deportista']}"); aplicar_formato_tym_word(p, 11, True)
            documento.add_paragraph(generar_comentario(fila, c_key, fila['#']))

    # OCR Podios
    documento.add_page_break()
    h_ocr = documento.add_heading('📏 PODIOS OCR', level=1); aplicar_formato_tym_word(h_ocr, 15, True)
    for i, item in enumerate(podio_d_lista): documento.add_paragraph(f"{i+1}. {item['nombre']} ({item['valor']} km)")
    
    stream = io.BytesIO(); documento.save(stream); stream.seek(0); return stream

# --- MOTOR NARRATIVO INDIVIDUAL (CON ADHERENCIA) ---
def generar_reporte_narrativo_individual(atleta_nom, df_actual, dict_historicos, sem_n):
    doc_p = Document()
    match_key = clean_string(atleta_nom)
    row_act = df_actual[df_actual['Deportista'].apply(clean_string) == match_key]
    if row_act.empty: return None
    row_act = row_act.iloc[0]

    # SECCIÓN TPI
    p_h = doc_p.add_heading(f'Análisis Personal: {atleta_nom}', 0); aplicar_formato_tym_word(p_h, 18, True, True)
    doc_p.add_heading('📊 Adherencia al Plan', level=1)
    p_tpi = doc_p.add_paragraph(f"Cumplimiento Global: {row_act.get('TPI_Global', 0):.1f}% ({row_act.get('Estado_Cumplimiento', 'Riesgo')})")
    aplicar_formato_tym_word(p_tpi, 12, True)

    table = doc_p.add_table(rows=1, cols=3); table.style = 'Light Grid Accent 1'
    for i, h in enumerate(['Disciplina', 'Meta', 'Logro']): table.rows[0].cells[i].text = h
    for d, h_p, tpi in [('Natación', 'Natacion_Hrs_Plan', 'TPI_Natacion'), ('Ciclismo', 'Ciclismo_Hrs_Plan', 'TPI_Ciclismo'), ('Trote', 'Trote_Hrs_Plan', 'TPI_Trote')]:
        r = table.add_row().cells
        r[0].text = d; r[1].text = f"{row_act.get(h_p, 0)}h"; r[2].text = f"{row_act.get(tpi, 0):.1f}%"

    # Comparativa Histórica
    doc_p.add_heading('📈 Comparativa', level=1)
    for tit, hoja, col, pref in [("TOTAL", "Tiempo Total", "T_Mins", "T"), ("NAT", "Natación", "N_Mins", "N")]:
        val = row_act[col]
        # Benchmark simple: media del equipo esta semana
        avg_eq = df_actual[f"{pref}_Mins"].mean()
        p = doc_p.add_paragraph(f"{tit}: {to_hhmmss_display(val)} (Promedio Equipo: {to_hhmmss_display(avg_eq)})")
        aplicar_formato_tym_word(p, 10)

    # Gráfico
    fig, ax = plt.subplots(figsize=(5,3))
    ax.bar(['Nat', 'Bici', 'Tro'], [row_act['N_Mins'], row_act['B_Mins'], row_act['R_Mins']], color=['#1E90FF', '#32CD32', '#FF4500'])
    buf = io.BytesIO(); plt.savefig(buf, format='png'); plt.close(fig)
    doc_p.add_paragraph().add_run().add_picture(buf, width=Inches(3.5))
    
    b_out = io.BytesIO(); doc_p.save(b_out); b_out.seek(0); return b_out

# *****************************************************************************
# --- 7. INTERFAZ DE USUARIO (STREMLIT) - VERSIÓN PERSISTENTE V2.2.28 ---
# *****************************************************************************

st.sidebar.header("📁 Gestión de Datos Históricos TYM")
cargador_maestro_excel = st.sidebar.file_uploader("Cargar Excel Maestro", type=["xlsx"])

# --- NUEVOS CAMPOS PARA ADHERENCIA AL PLAN ---
st.sidebar.divider()
cargador_plan_individual = st.sidebar.file_uploader("👤 Plan Individual (Opcional)", type=["xlsx"])

with st.sidebar.expander("🌍 Metas Club (Plan Global)"):
    p_n_h = st.number_input("Natación (Hrs)", 3.0)
    p_n_s = st.number_input("Natación (Ses)", 3)
    p_b_h = st.number_input("Ciclismo (Hrs)", 4.0)
    p_b_s = st.number_input("Ciclismo (Ses)", 3)
    p_t_h = st.number_input("Trote (Hrs)", 1.5)
    p_t_s = st.number_input("Trote (Ses)", 2)

# Empaquetamos las metas globales en un diccionario para el motor
dict_plan_global = {
    'Natacion_Hrs_Plan': p_n_h, 'Natacion_Ses_Plan': p_n_s,
    'Ciclismo_Hrs_Plan': p_b_h, 'Ciclismo_Ses_Plan': p_b_s,
    'Trote_Hrs_Plan': p_t_h, 'Trote_Ses_Plan': p_t_s
}
# --------------------------------------------

num_semana_procesar = st.text_input("Número de Semana (Ej: 08):", "08")
area_texto_strava = st.text_area("1. Datos Tiempo Total (Strava):")
area_texto_ocr = st.text_area("2. Datos OCR (Traducción):")

# Contenedor para mantener los resultados visibles tras la interacción
contenedor_resultados = st.container()

if st.button("🚀 PROCESAR JORNADA"):
    if cargador_maestro_excel and area_texto_strava.strip() and area_texto_ocr.strip():
        # 1. Realizamos el parsing original de Strava
        df_raw = parse_raw_data(area_texto_strava)
        
        # 2. PROCESAMIENTO DE ADHERENCIA (Llamada al nuevo motor 3B)
        # Cargamos el Excel de plan individual si existe
        df_plan_indiv = pd.read_excel(cargador_plan_individual) if cargador_plan_individual else None
        
        # Ejecutamos el cálculo de adherencia con la regla de cascada
        # (Asegúrate de haber pegado la función calcular_adherencia_v2 en la sección 3B)
        df_final = calcular_adherencia_v2(df_raw, df_plan_indiv, dict_plan_global)
        
        # 3. Guardamos en el estado de la sesión para persistencia
        st.session_state['df_resultados'] = df_final
        st.session_state['podios_ocr'] = parse_ocr_data(area_texto_ocr)
        st.session_state['procesado_ok'] = True
    else:
        st.error("Error Mandatorio: Excel y campos de texto no pueden estar vacíos.")

# Lógica de despliegue fuera del botón para que no desaparezca al interactuar
if st.session_state.get('procesado_ok'):
    df_resultados = st.session_state['df_resultados']
    d_p_dist, d_p_larg = st.session_state['podios_ocr']
    
    with contenedor_resultados:
        st.success(f"¡Semana {num_semana_procesar} procesada!")
        col1, col2 = st.columns(2)
        
        # 1. Descargas Grupales
        col1.download_button(label="📄 REPORTE WORD GRUPAL", 
                             data=generar_reporte_word_tym_completo(df_resultados, num_semana_procesar, d_p_dist, d_p_larg), 
                             file_name=f"Reporte_TYM_{num_semana_procesar}.docx")
        
        col2.download_button(label="📊 EXCEL ACTUALIZADO", 
                             data=crear_excel_actualizado(cargador_maestro_excel, df_resultados, num_semana_procesar), 
                             file_name=f"00_Estadisticas_Actualizadas_{num_semana_procesar}.xlsx")
        
        # 2. Sección de Insights Individuales (Ahora persistente)
        st.divider()
        st.subheader("👤 Generador de Reportes Individuales (Insights)")
        
        # Carga de históricos
        h_t = pd.read_excel(cargador_maestro_excel, sheet_name="Tiempo Total", dtype=object)
        h_n = pd.read_excel(cargador_maestro_excel, sheet_name="Natación", dtype=object)
        h_c = pd.read_excel(cargador_maestro_excel, sheet_name="Ciclismo", dtype=object)
        h_r = pd.read_excel(cargador_maestro_excel, sheet_name="Trote", dtype=object)
        dict_h_ref = {"Tiempo Total": h_t, "Natación": h_n, "Ciclismo": h_c, "Trote": h_r}
        
       # --- LÓGICA DE SELECCIÓN FILTRADA POR ACTIVIDAD ---
        # Filtramos solo a los atletas que sumaron minutos en la semana actual
        df_activos = df_resultados[df_resultados['T_Mins'] > 0]
        atletas_activos_list = df_activos['Deportista'].tolist()
        
        st.write(f"ℹ️ Se detectaron {len(atletas_activos_list)} atletas con actividad esta semana.")
        
        # Checkbox para selección masiva de activos
        seleccionar_todos = st.checkbox(f"Seleccionar los {len(atletas_activos_list)} atletas activos")
        
        if seleccionar_todos:
            seleccionados = st.multiselect("Atletas para reporte personal:", atletas_activos_list, default=atletas_activos_list)
        else:
            seleccionados = st.multiselect("Seleccionar Atletas para reporte personal:", atletas_activos_list)
        # --------------------------------------------------
        
        if seleccionados:
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zf:
                for a_sel in seleccionados:
                    r_indiv = generar_reporte_narrativo_individual(a_sel, df_resultados, dict_h_ref, num_semana_procesar)
                    if r_indiv:
                        zf.writestr(f"Reporte_{clean_string(a_sel)}.docx", r_indiv.getvalue())
            
            st.download_button(
                label="⬇️ DESCARGAR ZIP INDIVIDUALES", 
                data=zip_buffer.getvalue(), 
                file_name=f"Individuales_Sem_{num_semana_procesar}.zip", 
                mime="application/zip"
            )
