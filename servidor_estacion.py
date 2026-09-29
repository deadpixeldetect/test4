# -*- coding: utf-8 -*-
"""
Estación Anemométrica Loyola - Servidor Local Todo-en-Uno
==========================================================
- DASHBOARD (KPIs, brújula y gráficas): se actualiza CADA SEGUNDO consultando el
  datalogger en vivo. Solo MUESTRA: no escribe nada en disco. Los puntos de las gráficas
  viven únicamente en memoria (buffer circular) y se pierden al cerrar el servidor.
- TABLA / EXCEL / PDF: histórico minuto a minuto guardado en SQLite. Al arrancar se
  descargan ~3 años (backfill XML) y luego, cada minuto, solo lo nuevo. Una fila por
  minuto, sin repetidos.
- IP DEL DATALOGGER: configurable desde el propio dashboard (sección "Configuración"),
  sin recompilar. Se guarda en config.json, junto al programa.

Requisitos:  pip install flask requests openpyxl reportlab
Ejecutar:    python servidor_estacion.py
Luego abrir: http://localhost:5000

Para compilar un .exe de Windows, ver build_windows.bat / instrucciones adjuntas.
"""

import io
import os
import sys
import json
import math
import time
import sqlite3
import threading
import logging
from collections import deque
from datetime import datetime, timedelta

import re
import xml.etree.ElementTree as ET

import requests
from flask import Flask, jsonify, request, render_template_string, send_file
from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.graphics.shapes import Drawing, String
from reportlab.graphics.charts.lineplots import LinePlot

# ============================ CONFIGURACIÓN ============================
def carpeta_base():
    """Carpeta donde viven config.json, viento.db y estacion.log.
    Compilado con PyInstaller (--onefile): junto al .exe (no a la carpeta temporal).
    Ejecutado con 'python': junto a este script."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

BASE_DIR = carpeta_base()
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")
URL_BASE_POR_DEFECTO = "http://10.122.120.12/tables.html"

def cargar_config():
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                datos = json.load(f)
            if datos.get("datalogger_url"):
                return {"datalogger_url": str(datos["datalogger_url"]).strip()}
            if datos.get("datalogger_ip"):   # config.json de una versión anterior (solo IP)
                ip = str(datos["datalogger_ip"]).strip()
                return {"datalogger_url": f"http://{ip}/tables.html"}
        except (OSError, ValueError, json.JSONDecodeError):
            pass
    return {"datalogger_url": URL_BASE_POR_DEFECTO}

def guardar_config(nueva_url):
    config["datalogger_url"] = nueva_url
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)

config = cargar_config()

def _base():
    return config["datalogger_url"].rstrip("/")

def url_vivo():
    return f"{_base()}?command=DataQuery&mode=most-recent&format=xml&uri=dl:Datos&p1=1"

def url_backfill(p1):
    return f"{_base()}?command=DataQuery&mode=backfill&format=xml&uri=dl:Datos&p1={p1}"

def url_since_record(p1, p3):
    return (f"{_base()}?command=DataQuery&mode=since-record"
            f"&format=xml&uri=dl:Datos&p1={p1}&p3={p3}")

SEGUNDOS_HISTORICO = 94608000
REFRESCO_HISTORICO_SEG = 60      # el histórico se completa cada minuto (incremental), alineado al minuto
INTERVALO_SEGUNDOS = 1           # el dashboard consulta el datalogger cada 1 s
BUFFER_VIVO = 3600               # puntos en memoria para las gráficas (1 hora a 1 s)
MAX_INCREMENTAL_SEGUNDOS = 259200  # la carga de cada minuto pide como máximo los últimos 3 días
ANTIGUEDAD_MAX_MIN = 5           # si el último registro es más viejo, el dashboard avisa
DB_PATH = os.path.join(BASE_DIR, "viento.db")   # base de datos SQLite (solo histórico)
HOST = "0.0.0.0"                 # 0.0.0.0 = accesible desde la red local
PUERTO = 5000
LOG_FILE = os.path.join(BASE_DIR, "estacion.log")
NOMBRE_ESTACION = "Estación Anemométrica Loyola"
MAX_FILAS_EXCEL = 1_048_575      # límite de filas de una hoja de Excel (menos la cabecera)

NAVY = "16406B"                  # color corporativo principal

logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)

CARDINALES = ['N', 'NNE', 'NE', 'ENE', 'E', 'ESE', 'SE', 'SSE',
              'S', 'SSO', 'SO', 'OSO', 'O', 'ONO', 'NO', 'NNO']

# ============================ BASE DE DATOS ============================
def crear_tabla():
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cur.execute("PRAGMA journal_mode=WAL")   # permite leer la tabla mientras se carga el histórico
    cur.execute("""
        CREATE TABLE IF NOT EXISTS Datos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            Fecha TEXT,
            Record INTEGER,
            Velviento_S_WVT REAL,
            Direccion_D1_WVT REAL,
            Direccion_SD1_WVT REAL,
            Velviento_Max REAL,
            Velviento_Std REAL
        )
    """)
    indice_ok = True
    try:
        cur.execute("CREATE UNIQUE INDEX IF NOT EXISTS ux_datos_fecha ON Datos(Fecha)")
    except sqlite3.Error:
        indice_ok = False          # hay filas repetidas de versiones anteriores
    con.commit()
    depurar_base(con, indice_ok)
    con.close()

COLS_DATOS = ("Fecha, Record, Velviento_S_WVT, Direccion_D1_WVT, "
              "Direccion_SD1_WVT, Velviento_Max, Velviento_Std")

def depurar_base(con, indice_ok):
    """Repara un viento.db antiguo: fechas con 'T'/segundos/fracciones y filas repetidas.
    Deja una sola fila por minuto (si hay varias, gana la última guardada)."""
    cur = con.cursor()
    sucio = cur.execute(
        "SELECT 1 FROM Datos WHERE Fecha NOT LIKE '____-__-__ __:__:00' LIMIT 1").fetchone()
    if indice_ok and not sucio:
        return
    n_antes = cur.execute("SELECT COUNT(*) FROM Datos").fetchone()[0]
    logging.info(f"Depurando base de datos ({n_antes} filas): fechas y repetidos...")
    cur.execute("DROP TABLE IF EXISTS Datos_limpia")
    cur.execute("""CREATE TABLE Datos_limpia (
        id INTEGER PRIMARY KEY AUTOINCREMENT, Fecha TEXT, Record INTEGER,
        Velviento_S_WVT REAL, Direccion_D1_WVT REAL, Direccion_SD1_WVT REAL,
        Velviento_Max REAL, Velviento_Std REAL)""")
    cur.execute("CREATE UNIQUE INDEX ux_limpia_fecha ON Datos_limpia(Fecha)")
    cur.execute(f"""INSERT OR REPLACE INTO Datos_limpia ({COLS_DATOS})
        SELECT substr(replace(Fecha, 'T', ' '), 1, 16) || ':00', Record, Velviento_S_WVT,
               Direccion_D1_WVT, Direccion_SD1_WVT, Velviento_Max, Velviento_Std
        FROM Datos ORDER BY id""")
    cur.execute("DROP TABLE Datos")
    cur.execute("ALTER TABLE Datos_limpia RENAME TO Datos")
    cur.execute("DROP INDEX IF EXISTS ux_limpia_fecha")
    cur.execute("CREATE UNIQUE INDEX ux_datos_fecha ON Datos(Fecha)")
    con.commit()
    n_despues = cur.execute("SELECT COUNT(*) FROM Datos").fetchone()[0]
    logging.info(f"Base depurada: {n_antes} -> {n_despues} filas (una por minuto)")

crear_tabla()

def clausula_periodo(horas):
    """WHERE sobre Fecha (texto ISO, comparable y con índice) para las últimas N horas."""
    if horas and horas > 0:
        corte = (datetime.now() - timedelta(hours=horas)).strftime("%Y-%m-%d %H:%M:%S")
        return "WHERE Fecha >= ?", [corte]
    return "", []

def iterar_filas(horas=0, limite=None):
    """Genera los registros del periodo en orden cronológico, sin cargarlos todos en memoria.
    Si hay límite, devuelve los `limite` más recientes (igualmente en orden ASC)."""
    where, params = clausula_periodo(horas)
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    cur = con.cursor()
    if limite:
        cur.execute(f"""SELECT * FROM (SELECT * FROM Datos {where}
                                       ORDER BY Fecha DESC LIMIT ?)
                        ORDER BY Fecha ASC""", params + [limite])
    else:
        cur.execute(f"SELECT * FROM Datos {where} ORDER BY Fecha ASC", params)
    try:
        for r in cur:
            yield dict(r)
    finally:
        con.close()

def punto_cardinal(deg):
    return CARDINALES[int(round(deg / 22.5)) % 16]

def media_circular_desde_sumas(s, c):
    """Media circular a partir de las sumas de senos y cosenos."""
    return (math.degrees(math.atan2(s, c)) + 360) % 360

# ================== HISTÓRICO COMPLETO (BACKFILL XML, STREAMING) ==================
estado_hist = {"cargando": False, "procesados": 0, "nuevos": 0, "error": None}
hist_lock = threading.Lock()   # evita dos cargas del histórico a la vez

def _local(tag):
    """Nombre de etiqueta sin namespace: '{ns}record' -> 'record'."""
    return tag.split('}')[-1]

def minuto_exacto(txt):
    """'2023-10-30T09:59:00.25' -> '2023-10-30 09:59:00' (una fila por minuto, sin duplicados)."""
    return str(txt).strip().replace('T', ' ')[:16] + ":00"

def registro_desde_elemento(rec):
    """Convierte un <r time=".." no=".."> con hijos <v1>..<v5> en una tupla lista para SQLite.
    v1=Velviento_S_WVT, v2=Direccion_D1_WVT, v3=Direccion_SD1_WVT, v4=Velviento_Max, v5=Velviento_Std.
    Devuelve None si falta algún dato o algún valor es NAN."""
    fecha_txt = rec.get('time')
    if not fecha_txt:
        return None
    valores = {}
    for hijo in rec:
        tag = _local(hijo.tag)
        if tag in ('v1', 'v2', 'v3', 'v4', 'v5') and hijo.text:
            try:
                valores[tag] = float(hijo.text)
            except ValueError:
                return None
    if len(valores) < 5 or any(math.isnan(x) for x in valores.values()):
        return None
    try:
        no = int(float(rec.get('no', 0)))
    except ValueError:
        no = 0
    return (
        minuto_exacto(fecha_txt),
        no,
        valores['v1'], valores['v2'], valores['v3'], valores['v4'], valores['v5'],
    )

SQL_INSERT = """INSERT INTO Datos (Fecha, Record, Velviento_S_WVT, Direccion_D1_WVT,
                    Direccion_SD1_WVT, Velviento_Max, Velviento_Std)
                VALUES (?,?,?,?,?,?,?)
                ON CONFLICT(Fecha) DO UPDATE SET
                    Record = excluded.Record,
                    Velviento_S_WVT = excluded.Velviento_S_WVT,
                    Direccion_D1_WVT = excluded.Direccion_D1_WVT,
                    Direccion_SD1_WVT = excluded.Direccion_SD1_WVT,
                    Velviento_Max = excluded.Velviento_Max,
                    Velviento_Std = excluded.Velviento_Std
                WHERE Datos.Record IS NOT excluded.Record
                   OR Datos.Velviento_S_WVT IS NOT excluded.Velviento_S_WVT
                   OR Datos.Direccion_D1_WVT IS NOT excluded.Direccion_D1_WVT
                   OR Datos.Direccion_SD1_WVT IS NOT excluded.Direccion_SD1_WVT
                   OR Datos.Velviento_Max IS NOT excluded.Velviento_Max
                   OR Datos.Velviento_Std IS NOT excluded.Velviento_Std"""


CARGA_PAGINA = 5000

def registros_desde_xml_stream(fuente):
    """Recorre un XML del datalogger en streaming (sin cargarlo entero en memoria) y
    va devolviendo las tuplas (Fecha, Record, v1..v5) de cada <r>. `fuente` es un
    objeto tipo archivo (p. ej. r.raw de requests)."""
    pila = []
    for evento, el in ET.iterparse(fuente, events=("start", "end")):
        if evento == "start":
            pila.append(el)
            continue
        pila.pop()
        if _local(el.tag) != 'r':
            continue
        reg = registro_desde_elemento(el)
        el.clear()                       # libera memoria del registro ya leído
        if pila:
            pila[-1].remove(el)
        if reg is not None:
            yield reg

def completar_por_registro(con, desde_no):
    """Continúa la carga del histórico por número de registro (XML) hasta agotar el datalogger."""
    siguiente = desde_no
    for _ in range(5000):                       # tope de seguridad
        try:
            with requests.get(url_since_record(siguiente, CARGA_PAGINA),
                              stream=True, timeout=180) as r:
                if r.status_code != 200:
                    break
                r.raw.decode_content = True
                regs = list(registros_desde_xml_stream(r.raw))
        except Exception as e:
            logging.error(f"Completar histórico: {e}")
            break
        if not regs:
            break
        con.executemany(SQL_INSERT, regs)
        con.commit()
        estado_hist["procesados"] += len(regs)
        estado_hist["nuevos"] = con.total_changes
        siguiente = regs[-1][1] + 1
        if len(regs) < CARGA_PAGINA:
            break

def cargar_historico(incremental=False):
    """Descarga el histórico del datalogger y lo guarda en SQLite.
    Se procesa en streaming (no se carga el XML entero en memoria) y se inserta por lotes.
    incremental=True pide solo desde el último registro guardado (menos tráfico)."""
    if not hist_lock.acquire(blocking=False):
        return
    # Solo la carga completa se anuncia en la tabla; la incremental (cada minuto) es silenciosa
    estado_hist.update(cargando=not incremental, procesados=0, nuevos=0, error=None)
    p1 = SEGUNDOS_HISTORICO
    con = None
    try:
        con = sqlite3.connect(DB_PATH)
        if incremental:
            ultima = con.execute("SELECT MAX(Fecha) FROM Datos").fetchone()[0]
            if ultima:
                try:
                    dt = datetime.strptime(ultima[:19], "%Y-%m-%d %H:%M:%S")
                    p1 = min(MAX_INCREMENTAL_SEGUNDOS,
                             int((datetime.now() - dt).total_seconds()) + 180)
                except ValueError:
                    pass
        if not incremental:
            logging.info(f"Carga completa del histórico, p1={p1}")

        with requests.get(url_backfill(p1),
                          stream=True, timeout=(15, 600)) as r:
            if r.status_code != 200:
                raise RuntimeError(f"el datalogger respondió HTTP {r.status_code}")
            r.raw.decode_content = True
            lote = []
            ultimo_no, ultima_fecha = None, None
            cambios0 = con.total_changes
            for reg in registros_desde_xml_stream(r.raw):
                lote.append(reg)
                ultima_fecha, ultimo_no = reg[0], reg[1]
                if len(lote) >= 5000:
                    con.executemany(SQL_INSERT, lote)
                    con.commit()
                    estado_hist["procesados"] += len(lote)
                    estado_hist["nuevos"] = con.total_changes - cambios0
                    lote.clear()
            if lote:
                con.executemany(SQL_INSERT, lote)
                con.commit()
                estado_hist["procesados"] += len(lote)
            estado_hist["nuevos"] = con.total_changes - cambios0
            if not incremental:
                logging.info(f"Backfill XML: último registro leído no={ultimo_no} fecha={ultima_fecha}")
            # El datalogger puede truncar la respuesta al límite de registros por consulta.
            # Si el último dato leído está claramente atrasado, se continúa por número de registro.
            if ultimo_no is not None and ultima_fecha < (
                    datetime.now() - timedelta(hours=2)).strftime("%Y-%m-%d %H:%M:%S"):
                completar_por_registro(con, ultimo_no + 1)
            estado_hist["nuevos"] = con.total_changes - cambios0
        if not incremental or estado_hist["nuevos"] > 0:
            logging.info(f"Histórico {'incremental' if incremental else 'completo'}: "
                         f"{estado_hist['procesados']} recibidos, {estado_hist['nuevos']} nuevos")
    except Exception as e:
        estado_hist["error"] = str(e)
        logging.error(f"Error en la carga del histórico: {e}")
    finally:
        if con:
            con.close()
        estado_hist["cargando"] = False
        hist_lock.release()

def ciclo_historico(stop_event):
    cargar_historico(incremental=False)
    # Cada minuto, 2 s después de que el datalogger cierra el registro del minuto
    while not stop_event.wait(REFRESCO_HISTORICO_SEG - (time.time() % REFRESCO_HISTORICO_SEG) + 2):
        cargar_historico(incremental=True)

# ================== DATOS EN VIVO (solo memoria, NO se guarda) ==================
vivo_lock = threading.Lock()
vivo_ultimo = None                       # última lectura del datalogger
vivo_ultimo_ok = 0.0                     # instante (time.time) de la última consulta correcta
vivo_buffer = deque(maxlen=BUFFER_VIVO)  # ventana deslizante para las gráficas

def obtener_lectura_vivo():
    try:
        r = requests.get(url_vivo(), timeout=5)
        if r.status_code != 200:
            logging.error(f"Error HTTP del datalogger (en vivo): {r.status_code}")
            return None
        raiz = ET.fromstring(r.content)
        for el in raiz.iter():
            if _local(el.tag) != 'r':
                continue
            reg = registro_desde_elemento(el)
            if reg is None:
                return None
            fecha, no, v1, v2, v3, v4, v5 = reg
            return {
                'Fecha': fecha,   # hora del registro (sin redondear a minuto exacto)
                'Record': no,
                'Velviento_S_WVT': v1, 'Direccion_D1_WVT': v2,
                'Direccion_SD1_WVT': v3, 'Velviento_Max': v4, 'Velviento_Std': v5,
            }
        return None
    except Exception as e:
        logging.error(f"Error conectando con datalogger (en vivo): {e}")
        return None

def ciclo_vivo(stop_event):
    global vivo_ultimo, vivo_ultimo_ok
    while not stop_event.is_set():
        lect = obtener_lectura_vivo()
        if lect:
            with vivo_lock:
                vivo_ultimo = lect
                vivo_ultimo_ok = time.time()
                punto = dict(lect)
                punto['Fecha'] = f"{datetime.now():%Y-%m-%d %H:%M:%S}"   # instante de la muestra
                vivo_buffer.append(punto)
        stop_event.wait(INTERVALO_SEGUNDOS)

stop_event = threading.Event()
hilo_vivo = threading.Thread(target=ciclo_vivo, args=(stop_event,), daemon=True)
hilo_hist = threading.Thread(target=ciclo_historico, args=(stop_event,), daemon=True)

# =========================== EXPORTAR A EXCEL ===========================
COLUMNAS = [
    ("Fecha", 20), ("Record", 10),
    ("Velviento_S_WVT", 18), ("Direccion_D1_WVT", 20),
    ("Direccion_SD1_WVT", 22), ("Velviento_Max", 16), ("Velviento_Std", 16),
]

def construir_excel(horas):
    """Modo write_only: escribe fila a fila sin cargar todo en memoria (millones de filas)."""
    wb = Workbook(write_only=True)
    ws = wb.create_sheet("Datos")
    for i, (_, ancho) in enumerate(COLUMNAS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = ancho
    ws.freeze_panes = "A2"
    cabecera = []
    for nombre, _ in COLUMNAS:
        c = WriteOnlyCell(ws, value=nombre)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=NAVY)
        c.alignment = Alignment(horizontal="center")
        cabecera.append(c)
    ws.append(cabecera)
    for f in iterar_filas(horas, limite=MAX_FILAS_EXCEL):
        ws.append([f[c[0]] for c in COLUMNAS])
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio

# ============================ INFORME PDF ============================
def calcular_resumen(horas):
    """Una sola pasada por la base de datos: estadísticas y agrupación para el gráfico."""
    where, params = clausula_periodo(horas)
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cur.execute(f"""SELECT Fecha, Velviento_S_WVT, Direccion_D1_WVT, Direccion_SD1_WVT,
                           Velviento_Max, Velviento_Std
                    FROM Datos {where} ORDER BY Fecha ASC""", params)
    n = 0
    s_vel = s_sd = s_std = s_sin = s_cos = 0.0
    v_max = float("-inf")
    primera = ultima = None
    horarios = {}                                   # 'YYYY-MM-DD HH' -> [suma, cuenta, max]
    for fecha, v, d, sd, vm, vs in cur:
        n += 1
        s_vel += v
        s_sd += sd
        s_std += vs
        rad = math.radians(d)
        s_sin += math.sin(rad)
        s_cos += math.cos(rad)
        if vm > v_max:
            v_max = vm
        if primera is None:
            primera = fecha
        ultima = fecha
        g = horarios.setdefault(fecha[:13], [0.0, 0, float("-inf")])
        g[0] += v
        g[1] += 1
        if vm > g[2]:
            g[2] = vm
    con.close()
    if n == 0:
        return None
    return {
        "n": n, "prom_vel": s_vel / n, "vel_max": v_max,
        "prom_dir": media_circular_desde_sumas(s_sin, s_cos),
        "prom_sd_dir": s_sd / n, "prom_std_vel": s_std / n,
        "primera": primera, "ultima": ultima, "horarios": horarios,
    }

def grafico_agrupado(horarios):
    """Velocidad media y máxima por hora (o por día si el periodo es largo)."""
    claves = sorted(horarios)
    por_dia = len(claves) > 240
    if por_dia:
        dias = {}
        for k in claves:
            suma, cnt, mx = horarios[k]
            d = dias.setdefault(k[:10], [0.0, 0, float("-inf")])
            d[0] += suma
            d[1] += cnt
            d[2] = max(d[2], mx)
        horarios, claves = dias, sorted(dias)
    if len(claves) < 2:
        return None
    datos_prom = [(i, horarios[k][0] / horarios[k][1]) for i, k in enumerate(claves)]
    datos_max = [(i, horarios[k][2]) for i, k in enumerate(claves)]
    titulo = "Velocidad media y máxima por día (m/s)" if por_dia \
        else "Velocidad media y máxima por hora (m/s)"

    d = Drawing(480, 195)
    d.add(String(35, 178, titulo, fontSize=10, fillColor=colors.HexColor("#" + NAVY)))
    lp = LinePlot()
    lp.x, lp.y = 35, 30
    lp.width, lp.height = 425, 140
    lp.data = [datos_prom, datos_max]
    lp.lines[0].strokeColor = colors.HexColor("#" + NAVY)
    lp.lines[0].strokeWidth = 1.6
    lp.lines[1].strokeColor = colors.HexColor("#B04A46")
    lp.lines[1].strokeWidth = 1.6
    lp.xValueAxis.valueMin = 0
    lp.xValueAxis.valueMax = len(claves) - 1
    lp.xValueAxis.valueStep = max(1, len(claves) // 8)
    if por_dia:
        lp.xValueAxis.labelTextFormat = (
            lambda v: claves[int(v)] if 0 <= int(v) < len(claves) else "")
    else:
        lp.xValueAxis.labelTextFormat = (
            lambda v: claves[int(v)][5:10] + " " + claves[int(v)][11:] + "h"
            if 0 <= int(v) < len(claves) else "")
    lp.yValueAxis.valueMin = 0
    for ax in (lp.xValueAxis, lp.yValueAxis):
        ax.strokeColor = colors.HexColor("#9AA5B1")
        ax.labels.fontName = "Helvetica"
        ax.labels.fontSize = 7
        ax.labels.fillColor = colors.HexColor("#4B5563")
    d.add(lp)
    d.add(String(70, 8, "— Media (m/s)", fontSize=8, fillColor=colors.HexColor("#" + NAVY)))
    d.add(String(170, 8, "— Máxima (m/s)", fontSize=8, fillColor=colors.HexColor("#B04A46")))
    return d

def construir_pdf(res, horas):
    rango_txt = (f"Últimas {horas} horas" if horas and horas > 0
                 else "Todo el histórico")

    estilo_titulo = ParagraphStyle("titulo", fontName="Helvetica-Bold",
                                   fontSize=17, textColor=colors.HexColor("#" + NAVY),
                                   spaceAfter=4)
    estilo_sub = ParagraphStyle("sub", fontName="Helvetica", fontSize=10,
                                textColor=colors.HexColor("#4B5563"), spaceAfter=14)

    bio = io.BytesIO()
    doc = SimpleDocTemplate(bio, pagesize=A4,
                            leftMargin=2 * cm, rightMargin=2 * cm,
                            topMargin=2 * cm, bottomMargin=2.2 * cm,
                            title="Informe de viento")

    def pie(canvas, d):
        canvas.saveState()
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(colors.HexColor("#6B7280"))
        canvas.drawString(2 * cm, 1.2 * cm,
                          f"{NOMBRE_ESTACION}  ·  Generado: {datetime.now():%d/%m/%Y %H:%M}")
        canvas.drawRightString(19 * cm, 1.2 * cm, f"Página {d.page}")
        canvas.restoreState()

    elementos = [
        Paragraph("Informe de Viento", estilo_titulo),
        Paragraph(f"{NOMBRE_ESTACION}  ·  Periodo: {rango_txt}  ·  "
                  f"Del {res['primera']} al {res['ultima']}", estilo_sub),
    ]

    datos_tabla = [
        ["Métrica", "Valor"],
        ["Registros analizados", f"{res['n']:,}"],
        ["Promedio velocidad del viento", f"{res['prom_vel']:.2f} m/s"],
        ["Velocidad máxima registrada", f"{res['vel_max']:.2f} m/s"],
        ["Promedio dirección del viento",
         f"{res['prom_dir']:.1f}° ({punto_cardinal(res['prom_dir'])})"],
        ["Desviación estándar media (dirección)", f"{res['prom_sd_dir']:.1f}°"],
        ["Desviación estándar media (velocidad)", f"{res['prom_std_vel']:.2f} m/s"],
    ]
    tabla = Table(datos_tabla, colWidths=[8.5 * cm, 8.5 * cm])
    tabla.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#" + NAVY)),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1),
         [colors.white, colors.HexColor("#F3F6FA")]),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D0D7E2")),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
    ]))
    elementos.append(tabla)
    elementos.append(Spacer(1, 18))

    grafico = grafico_agrupado(res["horarios"])
    if grafico:
        elementos.append(grafico)

    doc.build(elementos, onFirstPage=pie, onLaterPages=pie)
    bio.seek(0)
    return bio

# ============================ SERVIDOR WEB ============================
app = Flask(__name__)

PATRON_URL_BASE = re.compile(r"^https?://[^\s?#]+$", re.IGNORECASE)

def url_base_valida(valor):
    if not valor or len(valor) > 500:
        return False
    return bool(PATRON_URL_BASE.match(valor))

@app.route("/")
def index():
    return render_template_string(PAGINA_HTML)

@app.route("/api/config", methods=["GET", "POST"])
def api_config():
    """Enlace base del datalogger (protocolo + host + ruta, sin parámetros): se puede leer
    y cambiar desde el propio dashboard, sin recompilar ni reiniciar. Se guarda en config.json."""
    if request.method == "GET":
        return jsonify({"datalogger_url": config["datalogger_url"]})
    datos = request.get_json(silent=True) or {}
    nueva_url = str(datos.get("datalogger_url", "")).strip()
    if not url_base_valida(nueva_url):
        return jsonify({"error": "El enlace no es válido. Debe empezar con http:// o https:// "
                                  "y no debe incluir los parámetros (?command=...)."}), 400
    guardar_config(nueva_url)
    logging.info(f"Enlace del datalogger actualizado a {nueva_url}")
    return jsonify({"datalogger_url": config["datalogger_url"]})

@app.route("/api/latest")
def api_latest():
    """Última lectura EN VIVO (memoria). No toca la base de datos."""
    with vivo_lock:
        if vivo_ultimo is None:
            return jsonify({})
        d = dict(vivo_ultimo)
        d['enlace'] = (time.time() - vivo_ultimo_ok) < 10   # el datalogger responde
    try:
        dt = datetime.strptime(d['Fecha'][:19], "%Y-%m-%d %H:%M:%S")
        d['antiguedad_min'] = round((datetime.now() - dt).total_seconds() / 60, 1)
    except ValueError:
        d['antiguedad_min'] = None
    d['reciente'] = d['antiguedad_min'] is not None and d['antiguedad_min'] <= ANTIGUEDAD_MAX_MIN
    return jsonify(d)

@app.route("/api/history")
def api_history():
    """Últimos n puntos EN VIVO (memoria) para las gráficas."""
    n = max(1, min(request.args.get("n", default=300, type=int), BUFFER_VIVO))
    with vivo_lock:
        filas = list(vivo_buffer)[-n:]
    return jsonify([{'Fecha': f['Fecha'],
                     'Velviento_S_WVT': f['Velviento_S_WVT'],
                     'Velviento_Max': f['Velviento_Max'],
                     'Direccion_D1_WVT': f['Direccion_D1_WVT']} for f in filas])

@app.route("/api/tabla")
def api_tabla():
    """Tabla paginada desde el HISTÓRICO guardado (backfill)."""
    pagina = max(1, request.args.get("pagina", default=1, type=int))
    por_pagina = max(1, min(request.args.get("por_pagina", default=20, type=int), 200))
    horas = request.args.get("horas", default=0, type=int)
    where, params = clausula_periodo(horas)
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    cur = con.cursor()
    cur.execute(f"SELECT COUNT(*) AS n, MIN(Fecha) AS d, MAX(Fecha) AS h FROM Datos {where}", params)
    agg = cur.fetchone()
    total = agg['n']
    offset = (pagina - 1) * por_pagina
    cur.execute(f"""SELECT * FROM Datos {where}
                    ORDER BY Fecha DESC LIMIT ? OFFSET ?""",
                params + [por_pagina, offset])
    filas = [dict(r) for r in cur.fetchall()]
    con.close()
    faltantes = 0
    if total and agg['d'] and agg['h']:
        try:
            t0 = datetime.strptime(agg['d'][:19], "%Y-%m-%d %H:%M:%S")
            t1 = datetime.strptime(agg['h'][:19], "%Y-%m-%d %H:%M:%S")
            faltantes = max(0, int((t1 - t0).total_seconds() // 60) + 1 - total)
        except ValueError:
            pass
    return jsonify({'filas': filas, 'total': total, 'faltantes': faltantes,
                    'desde': agg['d'], 'hasta': agg['h'],
                    'pagina': pagina, 'por_pagina': por_pagina,
                    'cargando': estado_hist['cargando'],
                    'procesados': estado_hist['procesados'],
                    'error': estado_hist['error']})

@app.route("/export/excel")
def export_excel():
    horas = request.args.get("horas", default=0, type=int)
    bio = construir_excel(horas)
    nombre = f"datos_viento_{datetime.now():%Y%m%d_%H%M}.xlsx"
    return send_file(
        bio,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True, download_name=nombre)

@app.route("/export/pdf")
def export_pdf():
    horas = request.args.get("horas", default=24, type=int)
    res = calcular_resumen(horas)
    if not res:
        return "No hay datos en el periodo seleccionado", 400
    bio = construir_pdf(res, horas)
    nombre = f"informe_viento_{datetime.now():%Y%m%d_%H%M}.pdf"
    return send_file(bio, mimetype="application/pdf",
                     as_attachment=True, download_name=nombre)

# ============================ PÁGINA WEB ============================
PAGINA_HTML = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Estación Anemométrica Loyola</title>
<script>
    // Se aplica antes de pintar la página, para que no haya parpadeo de claro -> oscuro
    (function () {
        try {
            const t = localStorage.getItem('tema') || 'claro';
            document.documentElement.setAttribute('data-tema', t);
        } catch (e) {}
    })();
</script>
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
<style>
    :root {
        --navy: #16406B; --acento: #2F6FB2;
        --fondo: #F4F6F9; --tarjeta: #FFFFFF; --borde: #E2E8F0;
        --texto: #1F2937; --muted: #6B7280; --banda: #F8FAFC;
        --grid: #EEF2F7; --tick: #6B7280; --code-bg: #EDF3F9;
        --ok-bg: #E8F5EC; --ok-color: #1E7E46; --ok-borde: #BFE3CC;
        --off-bg: #FBECEC; --off-color: #B3261E; --off-borde: #F2C6C4;
        --sombra: rgba(16, 42, 67, .06);
        --alto-header: 52px; --alto-tabs: 40px;
    }
    :root[data-tema="oscuro"] {
        --navy: #6FA8DC; --acento: #6FA8DC;
        --fondo: #0F172A; --tarjeta: #1B2333; --borde: #2E3A50;
        --texto: #E5E9F0; --muted: #96A2B8; --banda: #202A3D;
        --grid: #26314A; --tick: #96A2B8; --code-bg: #26314A;
        --ok-bg: rgba(46, 160, 90, .18); --ok-color: #6FE0A0; --ok-borde: rgba(111, 224, 160, .35);
        --off-bg: rgba(179, 38, 30, .2); --off-color: #FF9C93; --off-borde: rgba(255, 156, 147, .35);
        --sombra: rgba(0, 0, 0, .3);
    }
    * { margin: 0; padding: 0; box-sizing: border-box; }
    html, body { height: 100%; }
    body {
        font-family: 'Segoe UI', -apple-system, 'Helvetica Neue', Arial, sans-serif;
        background: var(--fondo); color: var(--texto);
        display: flex; flex-direction: column; height: 100vh; overflow: hidden;
    }
    header {
        height: var(--alto-header); flex: 0 0 auto;
        background: var(--tarjeta); border-bottom: 1px solid var(--borde);
        padding: 0 24px;
        display: flex; justify-content: space-between; align-items: center;
    }
    .marca { display: flex; align-items: center; gap: 10px; }
    .marca .logo {
        width: 30px; height: 30px; border-radius: 7px;
        background: var(--navy); color: #fff;
        display: flex; align-items: center; justify-content: center;
        font-size: .95rem;
    }
    .marca h1 { font-size: .92rem; font-weight: 600; color: var(--navy); line-height: 1.2; }
    .marca small { display: block; font-weight: 400; color: var(--muted); font-size: .68rem; }
    .estado {
        font-size: .74rem; padding: 4px 12px; border-radius: 20px;
        background: var(--ok-bg); color: var(--ok-color); font-weight: 600;
        border: 1px solid var(--ok-borde); white-space: nowrap;
    }
    .estado.off { background: var(--off-bg); color: var(--off-color); border-color: var(--off-borde); }

    /* ---------- Pestañas ---------- */
    nav.tabs {
        height: var(--alto-tabs); flex: 0 0 auto;
        display: flex; gap: 4px; width: 100%; max-width: 1320px; margin: 0 auto;
        padding: 0 24px; background: var(--tarjeta); border-bottom: 1px solid var(--borde);
    }
    nav.tabs button {
        font-family: inherit; font-size: .82rem; font-weight: 600; color: var(--muted);
        background: none; border: none; border-bottom: 3px solid transparent;
        padding: 0 16px; cursor: pointer; display: flex; align-items: center;
    }
    nav.tabs button:hover { color: var(--navy); }
    nav.tabs button.activa { color: var(--navy); border-bottom-color: var(--navy); }

    main {
        flex: 1 1 auto; min-height: 0; width: 100%; max-width: 1320px; margin: 0 auto;
        padding: 14px 24px; overflow: hidden;
    }
    .tab-panel { display: none; height: 100%; }
    .tab-panel.activa { display: flex; flex-direction: column; }
    #tab-tabla.activa, #tab-config.activa, #tab-acerca.activa {
        display: block; height: 100%; overflow-y: auto;
    }

    h2 {
        font-size: .82rem; font-weight: 600; color: var(--navy);
        text-transform: uppercase; letter-spacing: .5px;
    }
    .grid-kpi { display: grid; grid-template-columns: repeat(5, 1fr); gap: 10px; flex: 0 0 auto; }
    .kpi {
        background: var(--tarjeta); border: 1px solid var(--borde);
        border-radius: 9px; padding: 10px 14px;
        box-shadow: 0 1px 3px var(--sombra);
    }
    .kpi .etq { font-size: .66rem; text-transform: uppercase; letter-spacing: .6px; color: var(--muted); margin-bottom: 4px; }
    .kpi .valor { font-size: 1.25rem; font-weight: 700; color: var(--navy); line-height: 1.15; }
    .kpi .unidad { font-size: .68rem; color: var(--muted); }
    .kpi-brujula { display: flex; align-items: center; gap: 10px; }
    .kpi-brujula .info .valor { font-size: 1.05rem; }
    .kpi-brujula svg { flex: 0 0 auto; }

    .fila-graficos {
        flex: 1 1 auto; min-height: 0; display: grid;
        grid-template-columns: 1fr 1fr; gap: 12px; margin-top: 10px;
    }
    .panel {
        background: var(--tarjeta); border: 1px solid var(--borde);
        border-radius: 10px; padding: 12px 14px;
        box-shadow: 0 1px 3px var(--sombra);
        display: flex; flex-direction: column; min-height: 0;
    }
    .panel-head {
        display: flex; justify-content: space-between; align-items: center;
        flex-wrap: wrap; gap: 8px; margin-bottom: 8px; flex: 0 0 auto;
    }
    .panel-head h2 { margin: 0; }
    .lienzo { position: relative; flex: 1 1 auto; min-height: 0; }
    .controles { display: flex; gap: 8px; align-items: center; flex-wrap: wrap; }
    select, .btn, input[type=text] {
        font-family: inherit; font-size: .82rem; border-radius: 7px;
        padding: 6px 12px; cursor: pointer;
    }
    select, input[type=text] {
        background: var(--tarjeta); color: var(--texto); border: 1px solid var(--borde);
    }
    input[type=text] { cursor: text; min-width: 220px; }
    .btn { text-decoration: none; font-weight: 600; border: 1px solid transparent; display: inline-block; }
    .btn-excel { background: #217346; color: #fff; }
    .btn-excel:hover { background: #1B5E3A; }
    .btn-pdf { background: #B3261E; color: #fff; }
    .btn-pdf:hover { background: #8F1D17; }
    table { width: 100%; border-collapse: collapse; font-size: .84rem; }
    thead th {
        background: var(--navy); color: #fff; font-weight: 600;
        padding: 10px 12px; text-align: left; font-size: .74rem;
        text-transform: uppercase; letter-spacing: .5px; white-space: nowrap;
        position: sticky; top: 0;
    }
    thead th:first-child { border-radius: 7px 0 0 0; }
    thead th:last-child { border-radius: 0 7px 0 0; }
    tbody td { padding: 8px 12px; border-bottom: 1px solid var(--borde); }
    tbody tr:nth-child(even) { background: var(--banda); }
    tbody tr:hover { background: var(--grid); }
    td.num { text-align: right; font-variant-numeric: tabular-nums; }
    .paginacion {
        display: flex; justify-content: space-between; align-items: center;
        margin-top: 12px; font-size: .82rem; color: var(--muted);
    }
    .paginacion .botones { display: flex; gap: 8px; }
    .pag-btn {
        background: var(--tarjeta); border: 1px solid var(--borde); color: var(--texto);
        border-radius: 7px; padding: 6px 14px; font-size: .8rem; cursor: pointer;
        font-family: inherit;
    }
    .pag-btn:disabled { opacity: .4; cursor: not-allowed; }
    .pag-btn:not(:disabled):hover { background: var(--grid); }
    .flecha { transition: transform .8s ease; transform-origin: 45px 45px; }
    .pie-pagina {
        text-align: center; color: var(--muted); font-size: .72rem;
        margin-top: 8px; flex: 0 0 auto;
    }
    code { background: var(--code-bg); padding: 1px 6px; border-radius: 4px; }
    .campo label { display: block; font-size: .78rem; color: var(--muted); margin-bottom: 5px; }

    /* ---------- Interruptor claro/oscuro ---------- */
    .interruptor { display: flex; align-items: center; gap: 12px; }
    .interruptor button {
        width: 48px; height: 26px; border-radius: 20px; border: 1px solid var(--borde);
        background: var(--fondo); position: relative; cursor: pointer; padding: 0;
    }
    .interruptor button .bola {
        position: absolute; top: 2px; left: 2px; width: 20px; height: 20px; border-radius: 50%;
        background: var(--navy); transition: transform .2s ease; display: flex;
        align-items: center; justify-content: center; font-size: .62rem;
    }
    :root[data-tema="oscuro"] .interruptor button .bola { transform: translateX(22px); }

    /* ---------- Pestaña "Acerca de" ---------- */
    .acerca { max-width: 760px; }
    .acerca .panel { display: block; margin-bottom: 14px; }
    .acerca h2 { margin-bottom: 8px; }
    .acerca p { font-size: .88rem; line-height: 1.55; color: var(--texto); margin-bottom: 8px; }
    .acerca ul { list-style: none; }
    .acerca li {
        font-size: .88rem; padding: 6px 0; border-bottom: 1px solid var(--borde);
    }
    .acerca li:last-child { border-bottom: none; }

    @media (max-width: 900px) {
        .grid-kpi { grid-template-columns: repeat(3, 1fr); }
        .fila-graficos { grid-template-columns: 1fr; }
    }
</style>
</head>
<body>
<header>
    <div class="marca">
        <div class="logo">≈</div>
        <h1>Estación Anemométrica Loyola<small>Monitoreo de velocidad y dirección del viento</small></h1>
    </div>
    <span class="estado" id="estado">● Conectado</span>
</header>

<nav class="tabs">
    <button class="tab-btn activa" data-tab="dashboard">Dashboard</button>
    <button class="tab-btn" data-tab="tabla">Tabla</button>
    <button class="tab-btn" data-tab="config">Configuración</button>
    <button class="tab-btn" data-tab="acerca">Acerca de</button>
</nav>

<main>
    <!-- ============================ PESTAÑA 1: DASHBOARD ============================ -->
    <section class="tab-panel activa" id="tab-dashboard">
        <div class="grid-kpi">
            <div class="kpi"><div class="etq">Velocidad media</div>
                <div class="valor" id="vel_media">--</div><div class="unidad">m/s</div></div>
            <div class="kpi"><div class="etq">Velocidad máxima</div>
                <div class="valor" id="vel_max">--</div><div class="unidad">m/s</div></div>
            <div class="kpi kpi-brujula">
                <svg width="90" height="90" viewBox="0 0 90 90">
                    <circle cx="45" cy="45" r="39" fill="none" stroke="var(--borde)" stroke-width="1.5"/>
                    <text x="45" y="15" text-anchor="middle" font-size="10" font-weight="bold" fill="var(--navy)">N</text>
                    <text x="45" y="84" text-anchor="middle" font-size="10" fill="var(--muted)">S</text>
                    <text x="80" y="49" text-anchor="middle" font-size="10" fill="var(--muted)">E</text>
                    <text x="10" y="49" text-anchor="middle" font-size="10" fill="var(--muted)">O</text>
                    <g class="flecha" id="flechaDir">
                        <polygon points="45,12 50,47 45,42 40,47" fill="#B3261E"/>
                        <line x1="45" y1="42" x2="45" y2="73" stroke="var(--muted)" stroke-width="2"/>
                    </g>
                    <circle cx="45" cy="45" r="4" fill="var(--navy)"/>
                </svg>
                <div class="info">
                    <div class="etq">Dirección</div>
                    <div class="valor" id="dir_txt">--</div>
                    <div class="unidad" id="dir_grad">--</div>
                </div>
            </div>
            <div class="kpi"><div class="etq">Desv. estándar dir.</div>
                <div class="valor" id="dir_sd">--</div><div class="unidad">grados</div></div>
            <div class="kpi"><div class="etq">Desv. estándar vel.</div>
                <div class="valor" id="vel_std">--</div><div class="unidad">m/s</div></div>
        </div>

        <div class="fila-graficos">
            <div class="panel">
                <div class="panel-head">
                    <h2>Velocidad del viento</h2>
                    <div class="controles">
                        <label for="ventana" style="font-size:.78rem;color:var(--muted)">Ventana:</label>
                        <select id="ventana">
                            <option value="60">1 min</option>
                            <option value="300" selected>5 min</option>
                            <option value="900">15 min</option>
                            <option value="1800">30 min</option>
                            <option value="3600">1 hora</option>
                        </select>
                    </div>
                </div>
                <div class="lienzo"><canvas id="chart_vel"></canvas></div>
            </div>
            <div class="panel">
                <div class="panel-head"><h2>Dirección del viento</h2></div>
                <div class="lienzo"><canvas id="chart_dir"></canvas></div>
            </div>
        </div>

        <div class="pie-pagina">
            Último dato: <span id="ultima">--</span> · Hora actual: <span id="reloj">--</span>
            · Los datos mostrados pueden tener hasta 3 minutos de latencia respecto al registro real de la estación.
        </div>
    </section>

    <!-- ============================ PESTAÑA 2: TABLA ============================ -->
    <section class="tab-panel" id="tab-tabla">
        <div class="panel" style="height:auto">
            <div class="panel-head">
                <h2>Histórico minuto a minuto</h2>
                <div class="controles">
                    <label for="rango" style="font-size:.78rem;color:var(--muted)">Periodo:</label>
                    <select id="rango">
                        <option value="1">Última hora</option>
                        <option value="6">Últimas 6 horas</option>
                        <option value="24">Últimas 24 horas</option>
                        <option value="168">Última semana</option>
                        <option value="720">Último mes</option>
                        <option value="8760">Último año</option>
                        <option value="0" selected>Todo el histórico</option>
                    </select>
                    <a class="btn btn-excel" id="btnExcel" href="#">Exportar a Excel</a>
                    <a class="btn btn-pdf" id="btnPdf" href="#">Exportar a PDF</a>
                </div>
            </div>
            <div style="overflow-x:auto">
            <table>
                <thead>
                    <tr><th>Fecha</th><th>Hora</th><th>Record</th>
                        <th>Vel. media (m/s)</th><th>Vel. máx (m/s)</th>
                        <th>Dirección (°)</th><th>σ dir. (°)</th><th>σ vel. (m/s)</th></tr>
                </thead>
                <tbody id="cuerpoTabla"></tbody>
            </table>
            </div>
            <div class="paginacion">
                <span id="infoTabla">--</span>
                <div class="botones">
                    <button class="pag-btn" id="btnPrev">Anterior</button>
                    <button class="pag-btn" id="btnNext">Siguiente</button>
                </div>
            </div>
        </div>
    </section>

    <!-- ============================ PESTAÑA 3: CONFIGURACIÓN ============================ -->
    <section class="tab-panel" id="tab-config">
        <div class="panel" style="height:auto">
            <div class="panel-head"><h2>Datalogger</h2></div>
            <div class="controles" style="align-items:flex-end">
                <div class="campo" style="flex:1">
                    <label for="inputUrl">Enlace base del datalogger</label>
                    <input type="text" id="inputUrl" placeholder="http://10.122.120.12/tables.html"
                           style="width:100%; min-width:340px">
                </div>
                <button class="btn btn-excel" id="btnGuardarUrl" style="border:none">Guardar</button>
                <span id="msgConfig" style="font-size:.82rem;color:var(--muted)"></span>
            </div>
            <p style="font-size:.76rem;color:var(--muted);margin-top:10px">
                Es la parte fija del enlace del datalogger, sin los parámetros de la consulta
                (todo lo que va después de <code>?</code>). Por ejemplo:
                <code>http://10.122.120.12/tables.html</code>. Se guarda en <code>config.json</code>,
                junto al programa. No hace falta reiniciar ni recompilar: la próxima consulta al
                datalogger ya usa el nuevo enlace.</p>
        </div>

        <div class="panel" style="height:auto">
            <div class="panel-head"><h2>Apariencia</h2></div>
            <div class="interruptor">
                <button id="btnTema" title="Cambiar entre modo claro y oscuro">
                    <span class="bola" id="bolaTema">○</span>
                </button>
                <span id="txtTema" style="font-size:.86rem;color:var(--muted)">Modo claro</span>
            </div>
        </div>
    </section>

    <!-- ============================ PESTAÑA 4: ACERCA DE ============================ -->
    <section class="tab-panel" id="tab-acerca">
        <div class="acerca">
            <div class="panel">
                <h2>Propósito del proyecto</h2>
                <p>
                    La Estación Anemométrica Loyola registra de forma continua la velocidad y la
                    dirección del viento con fines de investigación del recurso eólico. Los datos
                    recolectados a lo largo de periodos extendidos permiten caracterizar el régimen
                    de viento de la zona: su variabilidad diaria y estacional, la distribución de
                    velocidades y la rosa de los vientos, información básica para evaluar el
                    potencial eólico de un sitio.
                </p>
                <p>
                    Este tipo de caracterización es un insumo típico en los estudios de factibilidad
                    para proyectos de energía eólica, ya que permite estimar la disponibilidad del
                    recurso antes de considerar la instalación de aerogeneradores o sistemas de
                    generación eólica de menor escala.
                </p>
                <p>
                    El sistema centraliza y organiza los datos del datalogger de la estación en un
                    solo lugar, con visualización en vivo, un histórico consultable minuto a minuto
                    y exportación a Excel y PDF para su posterior análisis estadístico.
                </p>
            </div>
            <div class="panel">
                <h2>Créditos</h2>
                <ul>
                    <li>Luis David Lorenzo Sánchez</li>
                    <li>Bryant Rosario</li>
                    <li>Antonio Carvajal Rodríguez</li>
                </ul>
            </div>
        </div>
    </section>
</main>

<script>
const CARDINALES = ['N','NNE','NE','ENE','E','ESE','SE','SSE','S','SSO','SO','OSO','O','ONO','NO','NNO'];
const cardinal = deg => CARDINALES[Math.round(deg / 22.5) % 16];
const fmt = n => Number(n).toFixed(2);

// ---------- Pestañas ----------
const botonesTab = document.querySelectorAll('.tab-btn');
const panelesTab = document.querySelectorAll('.tab-panel');
botonesTab.forEach(btn => btn.addEventListener('click', () => {
    botonesTab.forEach(b => b.classList.remove('activa'));
    panelesTab.forEach(p => p.classList.remove('activa'));
    btn.classList.add('activa');
    document.getElementById('tab-' + btn.dataset.tab).classList.add('activa');
    if (btn.dataset.tab === 'dashboard') {
        // el canvas pudo quedar con un tamaño incorrecto mientras estaba oculto
        chartVel.resize(); chartDir.resize();
    }
}));

// ---------- Modo claro / oscuro ----------
const btnTema = document.getElementById('btnTema');
const bolaTema = document.getElementById('bolaTema');
const txtTema = document.getElementById('txtTema');

function coloresGrafico() {
    const st = getComputedStyle(document.documentElement);
    return { texto: st.getPropertyValue('--tick').trim(), grid: st.getPropertyValue('--grid').trim() };
}

function aplicarColoresGraficos() {
    const c = coloresGrafico();
    [chartVel, chartDir].forEach(ch => {
        ch.options.plugins.legend.labels.color = c.texto;
        ch.options.scales.x.ticks.color = c.texto;
        ch.options.scales.x.grid.color = c.grid;
        ch.options.scales.y.ticks.color = c.texto;
        ch.options.scales.y.grid.color = c.grid;
        ch.update('none');
    });
}

function fijarTema(tema, guardar) {
    document.documentElement.setAttribute('data-tema', tema);
    bolaTema.textContent = tema === 'oscuro' ? '●' : '○';
    txtTema.textContent = tema === 'oscuro' ? 'Modo oscuro' : 'Modo claro';
    if (guardar) { try { localStorage.setItem('tema', tema); } catch (e) {} }
    if (typeof chartVel !== 'undefined') aplicarColoresGraficos();
}

btnTema.addEventListener('click', () => {
    const actual = document.documentElement.getAttribute('data-tema') === 'oscuro' ? 'claro' : 'oscuro';
    fijarTema(actual, true);
});

// ---------- Gráficos (colores iniciales según el tema ya aplicado en <head>) ----------
const _c0 = coloresGrafico();
const baseOpts = {
    animation: false,
    maintainAspectRatio: false,
    plugins: { legend: { labels: { color: _c0.texto, font: { size: 10 }, boxWidth: 14 } } },
    scales: {
        x: { ticks: { color: _c0.texto, maxTicksLimit: 8, font: { size: 9 } },
             grid: { color: _c0.grid } },
        y: { ticks: { color: _c0.texto, font: { size: 9 } }, grid: { color: _c0.grid } }
    }
};
const chartVel = new Chart(document.getElementById('chart_vel'), {
    type: 'line',
    data: { labels: [], datasets: [
        { label: 'Media (m/s)', data: [], borderColor: '#2F6FB2',
          backgroundColor: 'rgba(47,111,178,.14)', fill: true, tension: .3, pointRadius: 0, borderWidth: 2 },
        { label: 'Máxima (m/s)', data: [], borderColor: '#B04A46',
          tension: .3, pointRadius: 0, borderWidth: 1.5 }
    ]},
    options: { ...baseOpts, scales: { ...baseOpts.scales, y: { ...baseOpts.scales.y, beginAtZero: true } } }
});
const chartDir = new Chart(document.getElementById('chart_dir'), {
    type: 'line',
    data: { labels: [], datasets: [
        { label: 'Dirección (°)', data: [], borderColor: '#3D9970',
          tension: .3, pointRadius: 0, borderWidth: 1.5 }
    ]},
    options: { ...baseOpts, scales: { ...baseOpts.scales,
        y: { ...baseOpts.scales.y, min: 0, max: 360, ticks: { ...baseOpts.scales.y.ticks, stepSize: 90 } } } }
});
// Estado inicial del interruptor, ya con las gráficas creadas
fijarTema(document.documentElement.getAttribute('data-tema') || 'claro', false);
window.addEventListener('resize', () => { chartVel.resize(); chartDir.resize(); });

// ---------- Tabla histórica con paginación ----------
let paginaActual = 1;
const porPagina = 20;
let totalPaginas = 1;

async function cargarTabla() {
    const horas = document.getElementById('rango').value;
    try {
        const r = await fetch(`/api/tabla?pagina=${paginaActual}&por_pagina=${porPagina}&horas=${horas}`);
        const d = await r.json();
        totalPaginas = Math.max(1, Math.ceil(d.total / porPagina));
        const tb = document.getElementById('cuerpoTabla');
        tb.innerHTML = d.filas.map(f => `
            <tr>
                <td>${f.Fecha.slice(0, 10)}</td>
                <td>${f.Fecha.slice(11, 19)}</td>
                <td class="num">${f.Record}</td>
                <td class="num">${fmt(f.Velviento_S_WVT)}</td>
                <td class="num">${fmt(f.Velviento_Max)}</td>
                <td class="num">${f.Direccion_D1_WVT.toFixed(1)} (${cardinal(f.Direccion_D1_WVT)})</td>
                <td class="num">${f.Direccion_SD1_WVT.toFixed(1)}</td>
                <td class="num">${fmt(f.Velviento_Std)}</td>
            </tr>`).join('') ||
            '<tr><td colspan="8" style="text-align:center;color:var(--muted);padding:20px">' +
            (d.cargando ? 'Cargando histórico del datalogger…' : 'Sin datos en este periodo') + '</td></tr>';
        const rangoTxt = (d.desde && d.hasta) ? ` · ${d.desde.slice(0, 10)} → ${d.hasta.slice(0, 10)}` : '';
        let info = `${d.total.toLocaleString('es')} registros${rangoTxt} · Página ${paginaActual} de ${totalPaginas}`;
        if (d.faltantes > 0) info += ` · ${d.faltantes.toLocaleString('es')} minutos sin dato en el datalogger`;
        if (d.cargando) info += ` · Cargando histórico (${d.procesados.toLocaleString('es')} procesados)…`;
        else if (d.error) info += ` · Error al cargar histórico: ${d.error}`;
        document.getElementById('infoTabla').textContent = info;
        document.getElementById('btnPrev').disabled = paginaActual <= 1;
        document.getElementById('btnNext').disabled = paginaActual >= totalPaginas;
        return d.cargando;
    } catch (e) { return false; }
}

// Mientras se descarga el histórico la tabla se refresca cada 5 s; después, cada minuto
async function cicloTabla() {
    const cargando = await cargarTabla();
    // Con el histórico ya cargado, la tabla se refresca justo después de cada minuto nuevo
    const alMinuto = (60 - (Date.now() / 1000) % 60 + 5) * 1000;
    setTimeout(cicloTabla, cargando ? 5000 : alMinuto);
}

document.getElementById('btnPrev').onclick = () => { if (paginaActual > 1) { paginaActual--; cargarTabla(); } };
document.getElementById('btnNext').onclick = () => { if (paginaActual < totalPaginas) { paginaActual++; cargarTabla(); } };

// ---------- Enlaces de exportación (siguen el periodo seleccionado) ----------
function actualizarEnlaces() {
    const horas = document.getElementById('rango').value;
    document.getElementById('btnExcel').href = `/export/excel?horas=${horas}`;
    document.getElementById('btnPdf').href = `/export/pdf?horas=${horas}`;
}

document.getElementById('rango').addEventListener('change', () => {
    paginaActual = 1;
    actualizarEnlaces();
    cargarTabla();
});
actualizarEnlaces();
document.getElementById('ventana').addEventListener('change', actualizar);

// ---------- Configuración del enlace del datalogger ----------
const inputUrl = document.getElementById('inputUrl');
const msgConfig = document.getElementById('msgConfig');

async function cargarConfig() {
    try {
        const d = await (await fetch('/api/config')).json();
        inputUrl.value = d.datalogger_url || '';
    } catch (e) { /* se reintenta la próxima vez que se abra la pestaña */ }
}

document.getElementById('btnGuardarUrl').onclick = async () => {
    const url = inputUrl.value.trim();
    msgConfig.textContent = 'Guardando...';
    msgConfig.style.color = 'var(--muted)';
    try {
        const r = await fetch('/api/config', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ datalogger_url: url })
        });
        const d = await r.json();
        if (!r.ok) throw new Error(d.error || 'Error al guardar');
        msgConfig.textContent = `Guardado: ${d.datalogger_url}`;
        msgConfig.style.color = 'var(--ok-color)';
    } catch (e) {
        msgConfig.textContent = e.message;
        msgConfig.style.color = 'var(--off-color)';
    }
};

cargarConfig();

// ---------- Dashboard en vivo: cada segundo, solo muestra (no guarda nada) ----------
let actualizando = false;
async function actualizar() {
    if (actualizando) return;          // evita solapar peticiones si una tarda
    actualizando = true;
    const est = document.getElementById('estado');
    try {
        document.getElementById('reloj').textContent = new Date().toLocaleTimeString('es-ES');
        try {
            const r = await fetch('/api/latest');
            const d = await r.json();
            if (!d.Fecha) throw 0;
            document.getElementById('vel_media').textContent = fmt(d.Velviento_S_WVT);
            document.getElementById('vel_max').textContent = fmt(d.Velviento_Max);
            document.getElementById('dir_txt').textContent = cardinal(d.Direccion_D1_WVT);
            document.getElementById('dir_grad').textContent = d.Direccion_D1_WVT.toFixed(1) + '°';
            document.getElementById('dir_sd').textContent = d.Direccion_SD1_WVT.toFixed(1);
            document.getElementById('vel_std').textContent = fmt(d.Velviento_Std);
            document.getElementById('flechaDir').style.transform = `rotate(${d.Direccion_D1_WVT}deg)`;
            document.getElementById('ultima').textContent = d.Fecha;
            if (!d.enlace) { est.textContent = '● Sin conexión con la estación'; est.classList.add('off'); }
            else if (!d.reciente) { est.textContent = `● Estación sin datos nuevos (último: ${d.Fecha})`; est.classList.add('off'); }
            else { est.textContent = '● Conectado'; est.classList.remove('off'); }
        } catch (e) {
            est.textContent = '● Sin conexión con la estación'; est.classList.add('off');
        }
        const n = document.getElementById('ventana').value;
        const h = await (await fetch('/api/history?n=' + n)).json();
        const etq = h.map(p => p.Fecha.slice(11, 19));
        chartVel.data.labels = etq;
        chartVel.data.datasets[0].data = h.map(p => p.Velviento_S_WVT);
        chartVel.data.datasets[1].data = h.map(p => p.Velviento_Max);
        chartVel.update('none');
        chartDir.data.labels = etq;
        chartDir.data.datasets[0].data = h.map(p => p.Direccion_D1_WVT);
        chartDir.update('none');
    } catch (e) {
        est.textContent = '● Sin conexión con la estación'; est.classList.add('off');
    }
    finally { actualizando = false; }
}

actualizar();
setInterval(actualizar, 1000);      // KPIs y gráficas: cada segundo
cicloTabla();                       // tabla del histórico: cada minuto (justo tras el minuto nuevo)
</script>
</body>
</html>
"""

# ============================ ARRANQUE ============================
def iniciar_servidor_flask():
    # use_reloader=False: con el reloader, Flask se relanzaría a sí mismo y
    # arrancaría los hilos (adquisición, histórico) por duplicado.
    app.run(host=HOST, port=PUERTO, threaded=True, use_reloader=False)

def esperar_servidor(url, intentos=40, espera=0.25):
    """Espera a que Flask ya esté escuchando antes de abrir la ventana/el navegador."""
    for _ in range(intentos):
        try:
            if requests.get(url, timeout=1).status_code == 200:
                return True
        except Exception:
            pass
        time.sleep(espera)
    return False

if __name__ == "__main__":
    hilo_vivo.start()      # lectura en vivo (solo memoria)
    hilo_hist.start()      # histórico completo al arrancar y luego cada minuto -> SQLite

    hilo_servidor = threading.Thread(target=iniciar_servidor_flask, daemon=True)
    hilo_servidor.start()
    logging.info(f"Servidor iniciado en http://{HOST}:{PUERTO}")

    url_local = f"http://127.0.0.1:{PUERTO}"
    esperar_servidor(url_local)

    try:
        import webview   # pywebview: ventana nativa de escritorio, sin CMD ni navegador
        webview.create_window(
            "Estación Anemométrica Loyola",
            url_local,
            width=1300, height=880, min_size=(1000, 650),
        )
        logging.info("Abriendo ventana de escritorio (pywebview)")
        webview.start()          # bloquea hasta que se cierre la ventana
    except Exception as e:
        # Sin pywebview instalado (p. ej. ejecutando servidor_estacion.py suelto con
        # 'python'): se abre en el navegador y el programa sigue corriendo en la consola.
        logging.warning(f"pywebview no disponible ({e}); se abre en el navegador")
        import webbrowser
        webbrowser.open(url_local)
        try:
            while hilo_servidor.is_alive():
                time.sleep(1)
        except KeyboardInterrupt:
            pass
    finally:
        stop_event.set()
