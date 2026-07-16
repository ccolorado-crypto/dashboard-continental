import pandas as pd
import os
import glob
import base64
import json
from datetime import datetime, timedelta

print("Iniciando generación del Dashboard Corporativo ÁRTIMO Completo con UX/UI Avanzado...")

carpeta_actual = os.getcwd()
carpeta_data = os.path.join(carpeta_actual, 'data')
carpeta_public = os.path.join(carpeta_actual, 'public')

os.makedirs(carpeta_public, exist_ok=True)

# --- FECHA DE ACTUALIZACIÓN (AJUSTADA A HORA COLOMBIA UTC-5) ---
fecha_servidor = datetime.now()
fecha_colombia = fecha_servidor - timedelta(hours=5)
fecha_actualizacion = fecha_colombia.strftime("%d/%m/%Y %I:%M %p")

# 1. LOGO OFICIAL DE ÁRTIMO
ruta_logo_oficial = os.path.join(carpeta_data, 'logoartimogrande.jpg')
logo_base64_src = ""

if os.path.exists(ruta_logo_oficial):
    with open(ruta_logo_oficial, "rb") as image_file:
        encoded_string = base64.b64encode(image_file.read()).decode('utf-8')
        logo_base64_src = f"data:image/jpeg;base64,{encoded_string}"
else:
    ruta_logo_alternativo = glob.glob(os.path.join(carpeta_data, 'logo.*'))
    if ruta_logo_alternativo:
        with open(ruta_logo_alternativo[0], "rb") as image_file:
            encoded_string = base64.b64encode(image_file.read()).decode('utf-8')
            mime_type = "image/png" if ruta_logo_alternativo[0].lower().endswith('png') else "image/jpeg"
            logo_base64_src = f"data:{mime_type};base64,{encoded_string}"

html_logo = f'<img src="{logo_base64_src}" alt="ÁRTIMO" class="brand-logo">' if logo_base64_src else '<h1 style="color: #C8102E; font-size: 28px; margin: 0; font-weight: 800; letter-spacing: 2px;">ÁRTIMO</h1>'

# 2. LECTURA DE DATOS
archivos_datos = glob.glob(os.path.join(carpeta_data, '*.xlsx')) + glob.glob(os.path.join(carpeta_data, '*.csv'))
if not archivos_datos:
    print("Error: No se encontró NINGÚN archivo de datos en la carpeta data.")
    exit(1)

ruta_archivo = archivos_datos[0]
if ruta_archivo.lower().endswith('.csv'):
    df = pd.read_csv(ruta_archivo, on_bad_lines='skip', engine='python')
else:
    import warnings
    warnings.filterwarnings('ignore') # Ignora alertas visuales de openpyxl
    df = pd.read_excel(ruta_archivo)

# Limpiar fila de cabecera duplicada si existe
if not df.empty and str(df.iloc[0, 0]).strip() == 'Estado en línea':
    df = df.drop(index=0)
df.columns = df.columns.str.strip()

# --- EXTRACCIÓN DE COLUMNAS (Incluyendo la Flota) ---
if 'Flota asignada' not in df.columns:
    col_flota_fallback = df.columns[2] if len(df.columns) > 2 else df.columns[1]
    df['Flota asignada'] = df[col_flota_fallback]

columnas_base = ['Número de matrícula', 'Flota asignada', 'Tiempo fuera de línea', 'Estado de salud del almacenamiento 1']
dashboard_data = df[columnas_base].copy()
dashboard_data.columns = ['Máquina', 'Flota', 'Última transmisión', 'Estado del Disco 1']
total_maquinas = len(dashboard_data)

dashboard_data['Flota'] = dashboard_data['Flota'].fillna('Sin Flota asignada')
flotas_unicas = sorted([str(f).strip() for f in dashboard_data['Flota'].unique() if str(f).strip() != ''])
opciones_flota = '\n'.join([f'<option value="{f}">{f}</option>' for f in flotas_unicas])

# --- PROCESAMIENTO AVANZADO ---
ahora = datetime.now()

def calcular_dias_offline(val):
    val_str = str(val).strip().lower()
    if val_str == 'en línea': 
        return 0
    try:
        fecha_trans = pd.to_datetime(val)
        diff = (ahora - fecha_trans).days
        return max(0, diff)
    except:
        return 30

dashboard_data['Días Offline'] = dashboard_data['Última transmisión'].fillna('En línea').apply(calcular_dias_offline)
equipos_criticos_antiguedad = int((dashboard_data['Días Offline'] > 15).sum())

def mapear_estado_disco(estado):
    est = str(estado).strip().lower()
    if est == 'normal': return 'Normal'
    elif est in ['daños leves', 'falla']: return 'Falla'
    else: return 'No Detectado'

dashboard_data['Status_Disco'] = dashboard_data['Estado del Disco 1'].apply(mapear_estado_disco)
conteo_discos = dashboard_data['Status_Disco'].value_counts()
disco_normal_cnt = int(conteo_discos.get('Normal', 0))
disco_falla_cnt = int(conteo_discos.get('Falla', 0))
disco_nodet_cnt = int(conteo_discos.get('No Detectado', 0))

dashboard_data['Status_Transmision'] = dashboard_data['Días Offline'].apply(lambda x: 'Falla' if x >= 5 else 'Operando')
conteo_transmisiones = dashboard_data['Status_Transmision'].value_counts()
dashboard_data['Última transmisión'] = dashboard_data['Última transmisión'].fillna('En línea')
operando_cnt = int(conteo_transmisiones.get('Operando', 0))
falla_trans_cnt = int(conteo_transmisiones.get('Falla', 0))

# --- MAPEO DE CÁMARAS PERMITIDAS Y NO PERMITIDAS ---
nombres_camaras_personalizados = { 1: 'READ', 5: 'DMS', 6: 'ADAS', 7: 'LEFTDOWN', 8: 'LEFTREAR', 10: 'RIGHTDOWN', 11: 'RIGHTREAR', 12: 'FRONT' }
camaras_encontradas = []
total_cam_ok = 0
total_cam_falla = 0

for i in range(1, 13):
    col_hab = f'Cámara {i} habilitada'
    col_est = f'Estado de la cámara {i}'
    if col_hab in df.columns and col_est in df.columns:
        if i in nombres_camaras_personalizados:
            cam_name = nombres_camaras_personalizados[i]
            camaras_encontradas.append((col_hab, col_est, cam_name))
            
            def evaluar_camara(row, h=col_hab, e=col_est):
                hab = str(row.get(h, '')).strip().lower()
                est = str(row.get(e, '')).strip().lower()
                if hab == 'abrir':
                    return 'Normal' if est == 'normal' else 'Falla'
                return 'N/A'
                
            dashboard_data[cam_name] = df.apply(evaluar_camara, axis=1)
            total_cam_ok += (dashboard_data[cam_name] == 'Normal').sum()
            total_cam_falla += (dashboard_data[cam_name] == 'Falla').sum()

canales_no_permitidos = [i for i in range(1, 13) if i not in nombres_camaras_personalizados]

def generar_comentario(row):
    canales_activos_erroneos = []
    for i in canales_no_permitidos:
        col_hab = f'Cámara {i} habilitada'
        if col_hab in df.columns and str(row.get(col_hab, '')).strip().lower() == 'abrir':
            canales_activos_erroneos.append(f"CAM {i}")
    
    if canales_activos_erroneos:
        return f"⚠️ Canales activos no permitidos: {', '.join(canales_activos_erroneos)}"
    return "Sin observaciones"

dashboard_data['Comentario'] = df.apply(generar_comentario, axis=1)
alertas_hardware = disco_falla_cnt + disco_nodet_cnt + total_cam_falla

def evaluar_gravedad(row):
    if row['Status_Transmision'] == 'Falla' or row['Status_Disco'] in ['Falla', 'No Detectado']:
        return '🔴 Crítico'
    tiene_camaras_danadas = any(row[c[2]] == 'Falla' for c in camaras_encontradas)
    if tiene_camaras_danadas:
        return '🟡 Advertencia'
    return '🟢 Excelente'

dashboard_data['Gravedad'] = dashboard_data.apply(evaluar_gravedad, axis=1)
dashboard_data = dashboard_data.sort_values(by='Días Offline', ascending=False)

# --- CÁLCULO DE HEALTH SCORE Y RESUMEN EJECUTIVO (NUEVO) ---
equipos_excelentes = int((dashboard_data['Gravedad'] == '🟢 Excelente').sum())
health_score = round((equipos_excelentes / total_maquinas) * 100) if total_maquinas > 0 else 100

if health_score >= 90:
    hs_color, hs_icono = "#2A9D8F", "🌟"
elif health_score >= 70:
    hs_color, hs_icono = "#E9C46A", "⚠️"
else:
    hs_color, hs_icono = "#C8102E", "🚨"

equipos_mal_configurados = int((dashboard_data['Comentario'] != 'Sin observaciones').sum())

# Generación inteligente del texto
resumen_insights = f"El índice de salud operativo de la flota es del <strong>{health_score}%</strong>. "
if equipos_criticos_antiguedad > 0:
    resumen_insights += f"Se registran <strong style='color: var(--artimo-red);'>{equipos_criticos_antiguedad}</strong> equipos en estado crítico de transmisión (>15 días). "
if equipos_mal_configurados > 0:
    resumen_insights += f"Se detectaron <strong>{equipos_mal_configurados}</strong> vehículos con canales de cámara no permitidos activos. "
if alertas_hardware > 0:
    resumen_insights += f"Hay <strong>{alertas_hardware}</strong> incidencias activas de hardware (discos o cámaras)."
if equipos_criticos_antiguedad == 0 and equipos_mal_configurados == 0 and alertas_hardware == 0:
    resumen_insights += "Todos los sistemas operan dentro de los parámetros normales."

# --- DATA CRUDA EN FORMATO JSON ---
raw_export_data = dashboard_data.to_dict(orient='records')
json_raw_data = json.dumps(raw_export_data, default=str)

# --- CONSTRUCCIÓN DE LA TABLA HTML ---
def style_gravedad(grav):
    if 'Crítico' in grav: return '<b style="color: #C8102E;">🔴 Crítico</b>'
    if 'Advertencia' in grav: return '<b style="color: #E9C46A;">🟡 Advertencia</b>'
    return '<b style="color: #2A9D8F;">🟢 Excelente</b>'

def style_disk_status(st):
    if st == 'Normal': return '<span class="disk-status Normal">Normal</span>'
    if st == 'Falla': return '<span class="disk-status Falla">Falla</span>'
    return '<span class="disk-status SinDisco">No Detectado</span>'

def style_cam_status(st):
    if st == 'Normal': return '<span class="disk-status Normal">✓ OK</span>'
    if st == 'Falla': return '<span class="disk-status Falla">✗ Falla</span>'
    return '<span style="color: #5A5A59; font-weight: bold;">-</span>'

columnas_mostrar = ['Gravedad', 'Máquina', 'Flota', 'Días Offline', 'Última transmisión', 'Estado del Disco 1'] + [c[2] for c in camaras_encontradas] + ['Comentario']

html_table = '<table class="tabla-maquinas" id="tablaPrincipal">\n<thead>\n<tr>'
for col in columnas_mostrar:
    html_table += f'<th>{col}</th>'
html_table += '</tr>\n</thead>\n<tbody>\n'

for idx, row in dashboard_data.iterrows():
    t_stat = row['Status_Transmision']
    d_stat = row['Status_Disco']
    c_stat = 'Falla' if any(row[c[2]] == 'Falla' for c in camaras_encontradas) else 'Normal'
    
    flota_escaped = str(row["Flota"]).replace('"', '&quot;')
    maquina_str = str(row["Máquina"]).replace("'", "\\'")
    # NUEVO: Agregamos data-maquina para hacer el filtrado por texto más eficiente
    maquina_escaped = str(row["Máquina"]).replace('"', '&quot;').lower()
    
    # Se añade la función onclick a la fila para abrir el Modal y data-maquina
    html_table += f'<tr class="data-row" data-trans="{t_stat}" data-disk="{d_stat}" data-cam="{c_stat}" data-flota="{flota_escaped}" data-maquina="{maquina_escaped}" onclick="abrirModal(\'{maquina_str}\')">'
    html_table += f'<td>{style_gravedad(row["Gravedad"])}</td>'
    html_table += f'<td>{row["Máquina"]}</td>'
    html_table += f'<td style="font-size: 11px; max-width: 150px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="{flota_escaped}">{row["Flota"]}</td>'
    
    # Evitamos que el fondo rojo fije el color en modo oscuro, usamos una clase
    alerta_class = "cell-alert" if row["Días Offline"] >= 5 else ""
    html_table += f'<td class="{alerta_class}" style="font-weight: 700;">{row["Días Offline"]}</td>'
    
    html_table += f'<td>{row["Última transmisión"]}</td>'
    html_table += f'<td>{style_disk_status(row["Status_Disco"])}</td>'
    
    for _, _, cam_name in camaras_encontradas:
        html_table += f'<td>{style_cam_status(row[cam_name])}</td>'
    
    comentario_texto = row['Comentario']
    if "⚠️" in comentario_texto:
        style_com = 'class="comentario-alerta"'
    else:
        style_com = 'class="comentario-normal"'
        
    html_table += f'<td {style_com}>{comentario_texto}</td>'
    html_table += '</tr>\n'
html_table += '</tbody>\n</table>'

# 3. PLANTILLA HTML CON NUEVAS VARIABLES CSS PARA MODO OSCURO
plantilla_base = f'''
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>ÁRTIMO | DIAGNÓSTICO CONTINENTAL GOLD</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        :root {{ 
            --artimo-red: #C8102E; 
            --artimo-dark: #333333; 
            --artimo-grey: #5A5A59; 
            --artimo-light: #F4F4F4;
            --blanco: #FFFFFF; 
            --border-color: #E5E7EB;
            --alert-bg: #FFF5F5;
            --font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
        }}
        
        /* VARIABLES PARA MODO OSCURO */
        body.dark-mode {{
            --artimo-dark: #F3F4F6;
            --artimo-light: #111827;
            --blanco: #1F2937;
            --border-color: #374151;
            --artimo-grey: #9CA3AF;
            --alert-bg: #451a1e;
        }}

        body {{ font-family: var(--font-family); margin: 0; padding: 0; background-color: var(--artimo-light); color: var(--artimo-dark); transition: background-color 0.3s, color 0.3s; }}
        
        .top-navbar {{ 
            height: 56px; 
            background-color: var(--blanco); 
            border-bottom: 1px solid var(--border-color); 
            box-shadow: 0 2px 8px rgba(0,0,0,0.04);
            display: flex; 
            justify-content: space-between; 
            align-items: center; 
            padding: 0 24px;
            position: sticky;
            top: 0;
            z-index: 1000;
        }}
        .brand-container {{ display: flex; align-items: center; gap: 12px; }}
        .brand-logo {{ max-height: 36px; object-fit: contain; }}
        .navbar-title {{ font-size: 14px; font-weight: 700; color: var(--artimo-dark); letter-spacing: 0.5px; text-transform: uppercase; }}
        
        .navbar-actions {{ display: flex; align-items: center; gap: 10px; }}
        .btn-action {{ 
            background-color: var(--blanco); color: var(--artimo-dark); border: 1px solid var(--border-color); padding: 6px 14px; border-radius: 8px; font-size: 12px; font-weight: 600; cursor: pointer; display: flex; align-items: center; gap: 6px; transition: all 0.15s ease;
        }}
        .btn-action:hover {{ background-color: var(--artimo-light); }}
        .btn-action.primary {{ background-color: var(--artimo-red); color: #FFF; border-color: var(--artimo-red); }}
        .btn-action.primary:hover {{ background-color: #A50D24; }}
        .timestamp {{ font-size: 11px; color: var(--artimo-grey); background: var(--artimo-light); padding: 6px 12px; border-radius: 12px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px; }}
        
        .dashboard-container {{ display: flex; flex-direction: column; align-items: center; gap: 20px; padding: 24px; max-width: 1400px; margin: 0 auto; }}
        
        /* BANNER EJECUTIVO */
        .executive-summary {{ width: 100%; background-color: var(--blanco); border-left: 4px solid var(--artimo-red); padding: 14px 20px; border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.04); font-size: 14px; color: var(--artimo-dark); display: flex; align-items: center; gap: 12px; box-sizing: border-box; border-right: 1px solid var(--border-color); border-top: 1px solid var(--border-color); border-bottom: 1px solid var(--border-color); }}
        
        .kpi-section {{ display: flex; flex-wrap: wrap; justify-content: space-between; width: 100%; gap: 16px; }}
        .kpi-card {{ flex: 1; min-width: 200px; background: var(--blanco); border-radius: 12px; padding: 20px; text-align: center; box-shadow: 0 2px 8px rgba(0,0,0,0.04); border: 1px solid var(--border-color); position: relative; }}
        .kpi-card h3 {{ margin: 0 0 8px 0; color: var(--artimo-grey); font-size: 11px; text-transform: uppercase; letter-spacing: 0.5px; font-weight: 600; }}
        .kpi-card .number {{ margin: 0; font-size: 38px; font-weight: 800; color: var(--artimo-dark); }}
        
        /* TARJETA DE SALUD DESTACADA */
        .kpi-card.health-score {{ background: {hs_color}15; border: 2px solid {hs_color}; }}
        .kpi-card.health-score .number {{ color: {hs_color}; }}

        .charts-section {{ display: flex; flex-wrap: wrap; justify-content: center; gap: 16px; width: 100%; }}
        .card {{ background-color: var(--blanco); border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.04); border: 1px solid var(--border-color); padding: 20px; box-sizing: border-box; display: flex; flex-direction: column; }}
        .chart-card {{ flex: 1; min-width: 300px; max-width: 32%; min-height: 430px; transition: transform 0.15s ease; }}
        .chart-container {{ position: relative; width: 100%; height: 260px; flex-grow: 1; }}
        .chart-info-text {{ font-size: 11px; color: var(--artimo-grey); background-color: var(--artimo-light); padding: 10px; margin-top: 12px; border-radius: 8px; border-left: 3px solid var(--artimo-red); line-height: 1.5; }}
        
        .filter-bar {{ width: 100%; display: flex; justify-content: space-between; align-items: center; padding: 12px 20px; background: var(--blanco); border-radius: 12px; border: 1px solid var(--border-color); box-sizing: border-box; border-left: 4px solid var(--artimo-dark); flex-wrap: wrap; gap: 15px; }}
        .filter-group {{ display: flex; align-items: center; gap: 10px; }}
        .filter-label {{ font-weight: 700; font-size: 12px; color: var(--artimo-dark); text-transform: uppercase; }}
        .filter-select, .filter-input {{ padding: 6px 12px; border-radius: 6px; border: 1px solid var(--border-color); font-family: var(--font-family); font-size: 12px; outline: none; background: var(--blanco); color: var(--artimo-dark); }}
        .filter-select {{ cursor: pointer; max-width: 250px; }}
        .filter-input {{ max-width: 200px; }}
        .filter-msg {{ font-weight: 600; font-size: 13px; background: var(--alert-bg); padding: 6px 12px; border-radius: 6px; color: var(--artimo-red); }}
        
        /* Contenedor flexible para los filtros (nuevo) */
        .filters-container {{ display: flex; gap: 15px; flex-wrap: wrap; }}

        .table-card {{ width: 100%; padding: 0; overflow-x: auto; border-radius: 12px; border: 1px solid var(--border-color); }}
        table.tabla-maquinas {{ width: 100%; border-collapse: collapse; white-space: nowrap; font-size: 13px; }}
        th, td {{ padding: 12px 14px; text-align: center; border-bottom: 1px solid var(--border-color); }}
        th {{ background-color: var(--artimo-light); color: var(--artimo-dark); font-weight: 700; text-transform: uppercase; font-size: 11px; letter-spacing: 0.5px; border-top: none; }}
        tr.data-row {{ cursor: pointer; transition: background-color 0.2s; }}
        tr.data-row:hover {{ background-color: var(--artimo-light); }}
        
        td.cell-alert {{ background-color: var(--alert-bg); }}
        .comentario-alerta {{ color: #C8102E; font-weight: bold; font-size: 11px; text-align: left; }}
        .comentario-normal {{ color: var(--artimo-grey); font-size: 11px; text-align: left; }}
        
        .disk-status {{ padding: 3px 10px; border-radius: 12px; font-size: 11px; font-weight: 600; display: inline-block; }}
        .disk-status.Normal {{ background-color: #E8F8F5; color: #2A9D8F; }}
        .disk-status.Falla {{ background-color: #FDEDEC; color: #C8102E; }}
        .disk-status.SinDisco {{ background-color: var(--artimo-light); color: var(--artimo-grey); }}
        
        /* DISEÑO DEL MODAL */
        .modal-overlay {{ display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.6); z-index: 2000; justify-content: center; align-items: center; backdrop-filter: blur(3px); }}
        .modal-content {{ background: var(--blanco); padding: 24px; border-radius: 12px; max-width: 600px; width: 90%; color: var(--artimo-dark); border: 1px solid var(--border-color); box-shadow: 0 10px 25px rgba(0,0,0,0.2); max-height: 85vh; overflow-y: auto; }}
        .modal-header {{ display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid var(--border-color); padding-bottom: 14px; margin-bottom: 20px; }}
        .modal-header h2 {{ margin: 0; font-size: 20px; display: flex; align-items: center; gap: 10px; }}
        .modal-close {{ background: none; border: none; font-size: 24px; cursor: pointer; color: var(--artimo-grey); padding: 0; line-height: 1; }}
        .modal-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }}
        .modal-box {{ background: var(--artimo-light); padding: 12px; border-radius: 8px; border: 1px solid var(--border-color); }}
        .modal-box span {{ display: block; font-size: 11px; color: var(--artimo-grey); text-transform: uppercase; font-weight: 700; margin-bottom: 4px; }}
        .modal-box strong {{ font-size: 14px; color: var(--artimo-dark); }}
        .modal-full-width {{ grid-column: span 2; }}

        .footer-firma {{ text-align: center; padding: 30px 20px; margin-top: 20px; color: var(--artimo-grey); font-size: 12px; letter-spacing: 0.5px; font-weight: 300; }}
    </style>
</head>
<body>
    <div class="top-navbar">
        <div class="brand-container">
            {html_logo}
            <div class="navbar-title">Diagnóstico Corporativo</div>
        </div>
        <div class="navbar-actions">
            <span class="timestamp">🔄 {fecha_actualizacion}</span>
            <button class="btn-action" onclick="toggleDarkMode()" id="btnDarkMode">
                🌙 Modo Oscuro
            </button>
            <button class="btn-action" onclick="exportCSV()">
                📥 Exportar Data
            </button>
            <button class="btn-action primary" onclick="window.print()">
                📄 Guardar PDF
            </button>
        </div>
    </div>
    
    <div class="dashboard-container">
        <div class="executive-summary">
            <span style="font-size: 18px;">🤖</span>
            <div><strong>Análisis de IA:</strong> {resumen_insights}</div>
        </div>

        <div class="kpi-section">
            <div class="kpi-card health-score">
                <h3>Índice de Salud Global</h3>
                <p class="number">{hs_icono} {health_score}%</p>
            </div>
            <div class="kpi-card"><h3>Total de Máquinas</h3><p class="number">{total_maquinas}</p></div>
            <div class="kpi-card"><h3>Equipos Operando</h3><p class="number" style="color: #2A9D8F;">{operando_cnt}</p></div>
            <div class="kpi-card"><h3>Críticos (>15 Días)</h3><p class="number" style="color: #E9C46A;">{equipos_criticos_antiguedad}</p></div>
            <div class="kpi-card"><h3>Alertas Hardware</h3><p class="number" style="color: #C8102E;">{alertas_hardware}</p></div>
        </div>
        
        <div class="charts-section">
            <div class="card chart-card">
                <div class="chart-container"><canvas id="graficaTransmision"></canvas></div>
                <div class="chart-info-text">⚫ <strong>Operando:</strong> Transmitió hace menos de 5 días.<br>🔴 <strong>Falla:</strong> 5 días o más sin reportar.</div>
            </div>
            <div class="card chart-card">
                <div class="chart-container"><canvas id="graficaDisco"></canvas></div>
                <div class="chart-info-text">💡 Filtra los equipos interactuando con la gráfica.</div>
            </div>
            <div class="card chart-card">
                <div class="chart-container"><canvas id="graficaCamaras"></canvas></div>
                <div class="chart-info-text">💡 Filtra por equipos con cámaras Dañadas u OK.</div>
            </div>
        </div>
        
        <div class="filter-bar" id="filterBar">
            <div class="filters-container">
                <div class="filter-group">
                    <label for="flotaFilter" class="filter-label">🏢 Flota:</label>
                    <select id="flotaFilter" class="filter-select" onchange="applyFilters()">
                        <option value="ALL">Mostrar Todas las Flotas</option>
                        {opciones_flota}
                    </select>
                </div>
                <div class="filter-group">
                    <label for="maquinaFilter" class="filter-label">🚜 Máquina:</label>
                    <input type="text" id="maquinaFilter" class="filter-input" placeholder="Ej. MAQ-001..." onkeyup="applyFilters()">
                </div>
            </div>
            <div class="filter-msg" id="filterMessage" style="display: none;"></div>
            <button class="btn-action" onclick="resetFilters()" id="btnReset" style="display: none; background: var(--artimo-dark); color: white;">Limpiar Filtros</button>
        </div>
        
        <div class="card table-card">
            <div style="padding: 10px 15px; font-size: 11px; color: var(--artimo-grey); border-bottom: 1px solid var(--border-color);">
                <i>* Haz clic en cualquier fila para ver el detalle técnico ampliado del vehículo.</i>
            </div>
            {html_table}
        </div>
        <div class="footer-firma">Dashboard Analytics generado por Ártimo</div>
    </div>

    <div class="modal-overlay" id="detalleModal" onclick="cerrarModalFuera(event)">
        <div class="modal-content">
            <div class="modal-header">
                <h2 id="modalTitle">🚜 Máquina: --</h2>
                <button class="modal-close" onclick="cerrarModal()">&times;</button>
            </div>
            <div class="modal-grid" id="modalGrid">
                </div>
        </div>
    </div>
    
    <script>
        const rawJsonData = {json_raw_data};
        let chartsInstanced = [];

        // --- FUNCIONALIDAD MODO OSCURO ---
        function toggleDarkMode() {{
            const body = document.body;
            body.classList.toggle('dark-mode');
            const isDark = body.classList.contains('dark-mode');
            
            document.getElementById('btnDarkMode').innerHTML = isDark ? '☀️ Modo Claro' : '🌙 Modo Oscuro';
            
            // Actualizar color de fuente en Chart.js para que se lea en fondo oscuro
            Chart.defaults.color = isDark ? '#9CA3AF' : '#5A5A59';
            chartsInstanced.forEach(chart => chart.update());
        }}

        // --- FUNCIONALIDAD MODAL (VISTA DE DETALLE) ---
        function abrirModal(maquinaId) {{
            const dataRow = rawJsonData.find(d => String(d['Máquina']) === String(maquinaId));
            if(!dataRow) return;

            document.getElementById('modalTitle').innerHTML = `🚜 Máquina: ${{dataRow['Máquina']}} &nbsp; <span style="font-size: 14px;">(${{dataRow['Gravedad']}})</span>`;
            
            const grid = document.getElementById('modalGrid');
            grid.innerHTML = `
                <div class="modal-box modal-full-width"><span>🏢 Flota Asignada</span><strong>${{dataRow['Flota']}}</strong></div>
                <div class="modal-box"><span>⏱️ Días Offline</span><strong>${{dataRow['Días Offline']}} días</strong></div>
                <div class="modal-box"><span>📡 Última Transmisión</span><strong>${{dataRow['Última transmisión']}}</strong></div>
                <div class="modal-box"><span>💾 Estado de Disco</span><strong>${{dataRow['Status_Disco']}}</strong></div>
                <div class="modal-box"><span>📹 Status Red Cámaras</span><strong>${{dataRow['Gravedad'].includes('Advertencia') ? 'Fallas Detectadas' : 'Operando OK'}}</strong></div>
                <div class="modal-box modal-full-width">
                    <span>⚠️ Comentarios del Sistema</span>
                    <strong style="color: ${{dataRow['Comentario'].includes('⚠️') ? '#C8102E' : 'inherit'}}">${{dataRow['Comentario']}}</strong>
                </div>
            `;
            
            document.getElementById('detalleModal').style.display = 'flex';
        }}

        function cerrarModal() {{ document.getElementById('detalleModal').style.display = 'none'; }}
        function cerrarModalFuera(event) {{ if(event.target.id === 'detalleModal') cerrarModal(); }}

        // --- FUNCIONES CLÁSICAS (EXPORTAR, FILTROS, GRÁFICAS) ---
        function exportCSV() {{
            const headers = Object.keys(rawJsonData[0]);
            const csvRows = [headers.join(',')];
            for (const row of rawJsonData) {{
                const values = headers.map(header => {{
                    const val = row[header];
                    const valEscaped = ('' + (val !== null ? val : '')).replace(/"/g, '\\"');
                    return `"${{valEscaped}}"`;
                }});
                csvRows.push(values.join(','));
            }}
            const link = document.createElement("a");
            link.setAttribute("href", encodeURI("data:text/csv;charset=utf-8,\\uFEFF" + csvRows.join('\\n')));
            link.setAttribute("download", `Data_Cruda_${{new Date().toISOString().split('T')[0]}}.csv`);
            document.body.appendChild(link); link.click(); document.body.removeChild(link);
        }}

        let currentChartFilter = null;
        function filterTable(category, value, labelName) {{
            currentChartFilter = {{ category: category, value: value, label: labelName }};
            applyFilters();
            document.getElementById('filterBar').scrollIntoView({{ behavior: 'smooth', block: 'center' }});
        }}

        function applyFilters() {{
            const flotaValue = document.getElementById('flotaFilter').value;
            // NUEVO: Obtenemos lo que se escriba en el input y lo pasamos a minúsculas
            const maquinaValue = document.getElementById('maquinaFilter').value.trim().toLowerCase();
            const rows = document.querySelectorAll('.data-row');
            let count = 0;
            
            rows.forEach(row => {{
                let show = true;
                
                // Filtro por select de flota
                if (flotaValue !== 'ALL' && row.getAttribute('data-flota') !== flotaValue) show = false;
                
                // NUEVO: Filtro de texto por máquina (búsqueda parcial)
                if (maquinaValue && !row.getAttribute('data-maquina').includes(maquinaValue)) show = false;
                
                // Filtro por clicks en gráficas
                if (currentChartFilter) {{
                    let rVal = '';
                    if (currentChartFilter.category === 'trans') rVal = row.getAttribute('data-trans');
                    if (currentChartFilter.category === 'disk') rVal = row.getAttribute('data-disk');
                    if (currentChartFilter.category === 'cam') rVal = row.getAttribute('data-cam');
                    if (rVal !== currentChartFilter.value) show = false;
                }}
                
                row.style.display = show ? '' : 'none';
                if (show) count++;
            }});
            
            const msgEl = document.getElementById('filterMessage');
            const resetBtn = document.getElementById('btnReset');
            let activeFilters = [];
            
            if (currentChartFilter) activeFilters.push(`Gráfica: ${{currentChartFilter.label}}`);
            if (flotaValue !== 'ALL') activeFilters.push(`Flota`);
            if (maquinaValue !== '') activeFilters.push(`Búsqueda: "${{maquinaValue}}"`);
            
            if (activeFilters.length > 0) {{
                msgEl.style.display = 'inline-block'; resetBtn.style.display = 'inline-block';
                msgEl.innerText = `🔎 Filtrado por: ${{activeFilters.join(' + ')}} (${{count}} equipos encontrados)`;
            }} else {{
                msgEl.style.display = 'none'; resetBtn.style.display = 'none';
            }}
        }}

        function resetFilters() {{
            document.getElementById('flotaFilter').value = 'ALL';
            document.getElementById('maquinaFilter').value = ''; // Limpiar el input de máquina
            currentChartFilter = null;
            applyFilters();
        }}

        const titleOptions = (text) => ({{ display: true, text: text, font: {{ size: 14, weight: '700', family: 'var(--font-family)' }}, padding: {{bottom: 12}} }});
        const onChartClick = (category) => (evt, elements, chart) => {{
            if (elements.length === 0) return;
            const labelFull = chart.data.labels[elements[0].index];
            let value = '';
            if (category === 'trans') value = labelFull.includes('Operando') ? 'Operando' : 'Falla';
            if (category === 'disk') value = labelFull.includes('Normal') ? 'Normal' : (labelFull.includes('Falla') ? 'Falla' : 'No Detectado');
            if (category === 'cam') value = labelFull.includes('OK') ? 'Normal' : 'Falla';
            filterTable(category, value, labelFull);
        }};

        // Inicialización de Gráficas y guardado en arreglo para Modo Oscuro
        Chart.defaults.font.family = 'var(--font-family)';
        
        chartsInstanced.push(new Chart(document.getElementById('graficaTransmision').getContext('2d'), {{ 
            type: 'doughnut', 
            data: {{ labels: ['Operando', 'Falla de Transmisión'], datasets: [{{ data: [{operando_cnt}, {falla_trans_cnt}], backgroundColor: ['#2A9D8F', '#C8102E'], borderWidth: 0 }}] }}, 
            options: {{ responsive: true, maintainAspectRatio: false, plugins: {{ legend: {{ position: 'bottom' }}, title: titleOptions('Estado de Transmisión') }}, onClick: onChartClick('trans') }} 
        }}));
        
        chartsInstanced.push(new Chart(document.getElementById('graficaDisco').getContext('2d'), {{ 
            type: 'doughnut', 
            data: {{ labels: ['Normal', 'Falla', 'No Detectado'], datasets: [{{ data: [{disco_normal_cnt}, {disco_falla_cnt}, {disco_nodet_cnt}], backgroundColor: ['#2A9D8F', '#C8102E', '#5A5A59'], borderWidth: 0 }}] }}, 
            options: {{ responsive: true, maintainAspectRatio: false, plugins: {{ legend: {{ position: 'bottom' }}, title: titleOptions('Estado del Disco Duro') }}, onClick: onChartClick('disk') }} 
        }}));
        
        chartsInstanced.push(new Chart(document.getElementById('graficaCamaras').getContext('2d'), {{ 
            type: 'pie', 
            data: {{ labels: ['Equipos 100% OK', 'Equipos con Cámara Dañada'], datasets: [{{ label: 'Unidad', data: [{(dashboard_data['Status_Transmision'] != 'N/A').sum() - total_cam_falla}, {total_cam_falla}], backgroundColor: ['#2A9D8F', '#E9C46A'], borderWidth: 0 }}] }}, 
            options: {{ responsive: true, maintainAspectRatio: false, plugins: {{ legend: {{ position: 'bottom' }}, title: titleOptions('Salud de Cámaras') }}, onClick: onChartClick('cam') }} 
        }}));
    </script>
</body>
</html>
'''

ruta_guardado = os.path.join(carpeta_public, 'index.html')
with open(ruta_guardado, 'w', encoding='utf-8') as f:
    f.write(plantilla_base)

print(f"¡Dashboard avanzado (Dark Mode + IA Summary + Búsqueda por Máquina) generado con éxito!")
