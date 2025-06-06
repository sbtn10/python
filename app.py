import pandas as pd
import re
from datetime import datetime, timedelta
import unicodedata
from flask import Flask, request, jsonify
import os

app = Flask(__name__)

# --- Funciones Auxiliares ---

def convertir_fraccion_a_tiempo(fraction_of_day):
    if pd.isna(fraction_of_day):
        return "N/A"
    total_seconds = fraction_of_day * 24 * 3600
    td = timedelta(seconds=total_seconds)
    hours, remainder = divmod(td.seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    hours += td.days * 24
    return f"{hours:02d}:{minutes:02d}:{seconds:02d}"

def extraer_fecha(pregunta):
    fechas = re.findall(r"\d{4}-\d{2}-\d{2}", pregunta)
    return [datetime.strptime(f, "%Y-%m-%d").date() for f in fechas] if fechas else []

def extraer_hora(pregunta):
    horas = re.findall(r"\b([0-2]?[0-9]:[0-5][0-9])\b", pregunta)
    return [datetime.strptime(h, "%H:%M").time() for h in horas] if horas else []

def extraer_campana(pregunta):
    pregunta_lower = pregunta.lower()
    campanas_conocidas = ["preventiva", "cobranza", "cobranzas", "wow preventiva", "wow cobranza"]
    for campana_nombre in campanas_conocidas:
        if campana_nombre in pregunta_lower:
            if "preventiva" in campana_nombre:
                return "WOW PREVENTIVA"
            elif "cobranza" in campana_nombre:
                return "WOW COBRANZAS"
    return None

def quitar_tildes(s):
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')

# --- Carga de Datos ---
try:
    df = pd.read_excel("ejemplo.xlsx")
except FileNotFoundError:
    print("Error: archivo no encontrado.")
    exit()
except Exception as e:
    print(f"Error al cargar Excel: {e}")
    exit()

df["NOMBRES_BD"] = df["NOMBRES_BD"].str.upper().apply(quitar_tildes)
df["Campaña"] = df["Campaña"].str.upper().apply(quitar_tildes)
df["FECHA"] = pd.to_datetime(df["FECHA"]).dt.date

if not pd.api.types.is_object_dtype(df["HORA"]):
    df["HORA"] = pd.to_datetime(df["HORA"], format="%H:%M:%S", errors='coerce').dt.time
    if df["HORA"].isnull().any():
        df["HORA"] = pd.to_datetime(df["HORA"], format="%H", errors='coerce').dt.time

TIEMPO_COLUMNAS = ["T_LOGIN", "T_PAUSA", "T_ESPERA", "T_DISPONIBLE", "T_HABLADO", "T_ACW"]

PALABRAS_CLAVE = {
    "llamadas": "Q_LLA", "realizo": "Q_LLA", "cuantas llamadas": "Q_LLA",
    "hablado": "T_HABLADO", "tiempo hablado": "T_HABLADO", "hablo": "T_HABLADO",
    "pausa": "T_PAUSA", "espera": "T_ESPERA", "disponible": "T_DISPONIBLE",
    "acw": "T_ACW", "post llamada": "T_ACW", "dias trabajados": "Q_DIAS_TRABAJADOS",
    "login": "T_LOGIN", "tiempo login": "T_LOGIN", "logueado": "T_LOGIN",
    "monto cumplido": "Q_PDP_CUMPLIDA_MONTO", "promesas cumplidas": "Q_PDP_CUMPLIDA",
    "pdp": "Q_PDP", "monto pdp": "Q_PDP_MONTO", "promesas": "Q_PDP",
    "cet": "Q_CET", "contactos efectivos telefono": "Q_CET",
    "ctr": "Q_CTR", "contactos efectivos referido": "Q_CTR"
}

# --- Filtro ---
def filtrar_datos(df_base, fechas, horas, campana):
    df_filtrado = df_base.copy()
    if fechas:
        if len(fechas) == 1:
            df_filtrado = df_filtrado[df_filtrado["FECHA"] == fechas[0]]
        elif len(fechas) >= 2:
            fechas_ordenadas = sorted(fechas)
            df_filtrado = df_filtrado[(df_filtrado["FECHA"] >= fechas_ordenadas[0]) & (df_filtrado["FECHA"] <= fechas_ordenadas[1])]
    if horas:
        df_filtrado = df_filtrado[df_filtrado['HORA'].notna()]
        if len(horas) == 1:
            df_filtrado = df_filtrado[df_filtrado["HORA"] == horas[0]]
        elif len(horas) >= 2:
            horas_ordenadas = sorted(horas)
            df_filtrado = df_filtrado[(df_filtrado["HORA"] >= horas_ordenadas[0]) & (df_filtrado["HORA"] <= horas_ordenadas[1])]
    if campana:
        df_filtrado = df_filtrado[df_filtrado["Campaña"] == campana]
    return df_filtrado

# --- Respuesta ---
def responder(pregunta):
    pregunta_sin_tilde = quitar_tildes(pregunta.lower())
    fechas = extraer_fecha(pregunta)
    horas = extraer_hora(pregunta)
    campana = extraer_campana(pregunta)

    nombre_agente_encontrado = None
    for nombre_df in df["NOMBRES_BD"].unique():
        if re.search(r'\b' + re.escape(quitar_tildes(nombre_df.lower())) + r'\b', pregunta_sin_tilde):
            nombre_agente_encontrado = nombre_df
            break

    df_filtrado_global = filtrar_datos(df, fechas, horas, campana)

    if df_filtrado_global.empty:
        contexto_vacio = []
        if campana: contexto_vacio.append(f"la campaña '{campana.title()}'")
        if fechas:
            if len(fechas) == 1: contexto_vacio.append(f"la fecha {fechas[0]}")
            else: contexto_vacio.append(f"el rango de fechas entre {fechas[0]} y {fechas[1]}")
        if horas:
            if len(horas) == 1: contexto_vacio.append(f"la hora {horas[0].strftime('%H:%M')}")
            else: contexto_vacio.append(f"el rango de horas entre {horas[0].strftime('%H:%M')} y {horas[1].strftime('%H:%M')}")
        return f"No se encontraron datos para {' y '.join(contexto_vacio)}." if contexto_vacio else "No se encontraron datos."

    # --- Porcentaje de cierre ---
    if "porcentaje de cierre" in pregunta_sin_tilde or "tasa de cierre" in pregunta_sin_tilde:
        df_para_calculo = df_filtrado_global.copy()
        if nombre_agente_encontrado:
            df_para_calculo = df_para_calculo[df_para_calculo["NOMBRES_BD"] == nombre_agente_encontrado]
            if df_para_calculo.empty:
                return f"No se encontraron datos para {nombre_agente_encontrado.title()}."

        total_q_pdp = df_para_calculo["Q_PDP"].sum()
        total_q_cet = df_para_calculo["Q_CET"].sum()
        total_q_ctr = df_para_calculo["Q_CTR"].sum()
        denominador = total_q_cet + total_q_ctr

        if denominador == 0:
            return "No se puede calcular el porcentaje de cierre porque no hay contactos efectivos (CET + CTR = 0)."

        porcentaje = (total_q_pdp / denominador) * 100
        contexto_str = ""
        if nombre_agente_encontrado: contexto_str += f"{nombre_agente_encontrado.title()} - "
        if campana: contexto_str += f"{campana.title()} - "
        if fechas: contexto_str += f"Fechas: {fechas[0]}{' a ' + str(fechas[1]) if len(fechas) > 1 else ''}"

        return f"{contexto_str.strip()} → Porcentaje de cierre: {porcentaje:.2f}%"

    # --- Métricas ---
    for palabra, columna in PALABRAS_CLAVE.items():
        if palabra in pregunta_sin_tilde:
            df_filtrado = df_filtrado_global.copy()
            if nombre_agente_encontrado:
                df_filtrado = df_filtrado[df_filtrado["NOMBRES_BD"] == nombre_agente_encontrado]

            if df_filtrado.empty:
                return f"No se encontraron registros para {nombre_agente_encontrado.title()} en ese rango."

            total = df_filtrado[columna].sum()
            if columna in TIEMPO_COLUMNAS:
                total = convertir_fraccion_a_tiempo(total)
            else:
                total = int(total)

            return f"{columna} de {nombre_agente_encontrado.title() if nombre_agente_encontrado else 'todos'}: {total}"

    return "No entendí la métrica solicitada. Intenta con otra pregunta."

# --- Ruta principal ---
@app.route('/consultar', methods=['POST'])
def consultar():
    data = request.get_json()
    pregunta = data.get("pregunta", "")
    respuesta = responder(pregunta)
    return jsonify({"respuesta": respuesta})

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=True)
