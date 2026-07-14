import pandas as pd
import os
import glob
import base64
import json
from datetime import datetime, timedelta

print("Iniciando generación del Dashboard Corporativo ÁRTIMO Completo...")

carpeta_actual = os.getcwd()
carpeta_data = os.path.join(carpeta_actual, 'data')
carpeta_public = os.path.join(carpeta_actual, 'public')

os.makedirs(carpeta_public, exist_ok=True)

# --- FECHA DE ACTUALIZACIÓN (AJUSTADA A HORA COLOMBIA UTC-5) ---
fecha_servidor = datetime.now()
fecha_colombia = fecha_servidor - timedelta(hours=5)
fecha_actualizacion = fecha_colombia.strftime("%d/%m/%Y %I:%M %p")

# 1. LOGO OFICIAL DE ÁRTIMO (logoartimogrande.jpg)
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
    df = pd.read_excel(ruta_archivo)

if not df.empty and str(df.iloc[0, 0]).strip() == 'Estado en línea':
    df = df.drop(index=0)
df.columns = df.columns.str.strip()

columnas_base = ['Número de matrícula', 'Tiempo fuera de línea', 'Estado de salud del almacenamiento 1']
dashboard_data = df[columnas_base].copy()
dashboard_data.columns = ['Máquina', 'Última transmisión', 'Estado del Disco 1']
total_maquinas = len(dashboard_data)

# --- PROCESAMIENTO AVANZADO ---
ahora = datetime.now()

# Calcular Días Exactos Offline
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

# Filtrado por Antigüedad Crítica (>15 días sin transmitir)
equipos_criticos_antiguedad = int((dashboard_data['Días Offline'] > 15).sum())

# Mapeo de Discos
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

# Evaluación de Transmisión
dashboard_data['Status_Transmision'] = dashboard_data['Días Offline'].apply(lambda x: 'Falla' if x >= 5 else 'Operando')
conteo_transmisiones = dashboard_data['Status_Transmision'].value_counts()
dashboard_data['Última transmisión'] = dashboard_data['Última transmisión'].fillna('En línea')
operando_cnt = int(conteo_transmisiones.get('Operando', 0))
falla_trans_cnt = int(conteo_transmisiones.get('Falla', 0))


# --- CANALES DE CÁMARAS ACTUALIZADOS ---
mapeo_camaras = {
    1: 'REAR',
    5: 'DMS',
    6: 'ADAS',
    7: 'LEFTDOWN',
    8: 'LEFTREAR',
    10: 'RIGHTDOWN',
    11: 'RIGHTREAR',
    12: 'FRONT'
}

camaras_encontradas = []
total_cam_ok = 0
total_cam_falla = 0

for i in range(1, 13):
    col_hab = f'Cámara {i} habilitada'
    col_est = f'Estado de la cámara {i}'
    if col_hab in df.columns and col_est in df.columns:
        # Se asigna el nombre del mapeo si existe, si no, se deja como 'CAM X' por defecto
        cam_name = mapeo_camaras.get(i, f'CAM {i}')
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

alertas_hardware = disco_falla_cnt + disco_nodet_cnt + total_cam_falla

# --- CALCULAR INDICADOR DE GRAVEDAD POR VEHÍCULO ---
def evaluar_gravedad(row):
    if row['Status_Transmision'] == 'Falla' or row['Status_Disco'] in ['Falla', 'No Detectado']:
        return '🔴 Crítico'
    
    tiene_camaras_danadas = any(row[c[2]] == 'Falla' for c in camaras_encontradas)
    if tiene_camaras_danadas:
        return '🟡 Advertencia'
        
    return '  Excelente'

dashboard_data['Gravedad'] = dashboard_data.apply(evaluar_gravedad, axis=1)

# ORDENAR LA TABLA DE MAYOR A MENOR SEGÚN DÍAS OFFLINE
dashboard_data = dashboard_data.sort_values(by='Días Offline', ascending=False)

# --- DATA CRUDA EN FORMATO JSON PARA JS DOWNLOAD ---
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

columnas_mostrar = ['Gravedad', 'Máquina', 'Días Offline', 'Última transmisión', 'Estado del Disco 1'] + [c[2] for c in camaras_encontradas]

html_table = '<table class="tabla-maquinas">\n<thead>\n<tr>'
for col in columnas_mostrar:
    html_table += f'<th>{col}</th>'
html_table += '</tr>\n</thead>\n<tbody>\n'

for idx, row in dashboard_data.iterrows():
    t_stat = row['Status_Transmision']
    d_stat = row['Status_Disco']
    c_stat = 'Falla' if any(row[c[2]] == 'Falla' for c in camaras_encontradas) else 'Normal'
    
    html_table += f'<tr class="data-row" data-trans="{t_stat}" data-disk="{d_stat}" data-cam="{c_stat}">'
    html_table += f'<td>{style_gravedad(row["Gravedad"])}</td>'
    html_table += f'<td>{row["Máquina"]}</td>'
    html_table += f'<td style="font-weight: 700; background-color: {"#FFF5F5" if row["Días Offline"] >= 5 else "inherit"}">{row["Días Offline"]}</td>'
    html_table += f'<td>{row["Última transmisión"]}</td>'
    html_table += f'<td>{style_disk_status(row["Status_Disco"])}</td>'
    
    for _, _, cam_name in camaras_encontradas:
        html_table += f'<td>{style_cam_status(row[cam_name])}</td>'
        
    html_table += '</tr>\n'
html_table += '</tbody>\n</table>'

# 3. PLANTILLA HTML CON IDENTIDAD CORPORATIVA Y BOTONES DE DESCARGA
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
            --font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
        }}
        body {{ font-family: var(--font-family); margin: 0; padding: 0; background-color: var(--artimo-light); color: var(--artimo-dark); }}
        
        /* BARRA SUPERIOR FIJA 56PX */
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
        
        /* PANEL DE CONTROLES Y DESCARGAS */
        .navbar-actions {{ display: flex; align-items: center; gap: 10px; }}
        .btn-download {{ 
            background-color: var(--blanco);
            color: var(--artimo-dark);
            border: 1px solid var(--border-color);
            padding: 6px 14px;
            border-radius: 8px;
            font-size: 12px;
            font-weight: 600;
            cursor: pointer;
            display: flex;
            align-items: center;
            gap: 6px;
            transition: all 0.15s ease;
        }}
        .btn-download:hover {{ background-color: var(--artimo-light); border-color: var(--artimo-grey); }}
        .btn-download.primary {{ background-color: var(--artimo-red); color: var(--blanco); border-color: var(--artimo-red); }}
        .btn-download.primary:hover {{ background-color: #A50D24; }}
        
        .timestamp {{ font-size: 11px; color: var(--artimo-grey); background: var(--artimo-light); padding: 6px 12px; border-radius: 12px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px; }}
        
        .dashboard-container {{ display: flex; flex-direction: column; align-items: center; gap: 24px; padding: 24px; max-width: 1400px; margin: 0 auto; }}
        
        .kpi-section {{ display: flex; flex-wrap: wrap; justify-content: space-between; width: 100%; gap: 16px; }}
        .kpi-card {{ 
            flex: 1; 
            min-width: 220px; 
            background: var(--blanco); 
            border-radius: 12px; 
            padding: 20px; 
            text-align: center; 
            box-shadow: 0 2px 8px rgba(0,0,0,0.08); 
            border: 1px solid var(--border-color); 
        }}
        .kpi-card h3 {{ margin: 0 0 8px 0; color: var(--artimo-grey); font-size: 11px; text-transform: uppercase; letter-spacing: 0.5px; font-weight: 600; }}
        .kpi-card .number {{ margin: 0; font-size: 38px; font-weight: 800; color: var(--artimo-dark); }}
        .kpi-card.highlight {{ border-top: 4px solid var(--artimo-red); }}
        
        .charts-section {{ display: flex; flex-wrap: wrap; justify-content: center; gap: 16px; width: 100%; }}
        .card {{ background-color: var(--blanco); border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); border: 1px solid var(--border-color); padding: 20px; box-sizing: border-box; display: flex; flex-direction: column; }}
        .chart-card {{ flex: 1; min-width: 300px; max-width: 32%; min-height: 430px; transition: transform 0.15s ease; }}
        .chart-card:hover {{ transform: translateY(-3px); cursor: pointer; border-color: var(--artimo-grey); }}
        .chart-container {{ position: relative; width: 100%; height: 260px; flex-grow: 1; }}
        
        .chart-info-text {{ font-size: 11px; color: var(--artimo-grey); background-color: var(--artimo-light); padding: 10px; margin-top: 12px; border-radius: 8px; border-left: 3px solid var(--artimo-red); line-height: 1.5; }}
        
        .filter-bar {{ width: 100%; display: flex; justify-content: space-between; align-items: center; padding: 12px 20px; background: #fff; border-radius: 12px; border: 1px solid var(--border-color); box-sizing: border-box; border-left: 4px solid var(--artimo-red); }}
        .filter-msg {{ font-weight: 600; color: var(--artimo-dark); font-size: 13px; text-transform: uppercase; letter-spacing: 0.5px; }}
        .btn-reset {{ background: var(--artimo-dark); color: white; border: none; padding: 8px 15px; border-radius: 8px; cursor: pointer; font-weight: 600; font-size: 11px; text-transform: uppercase; letter-spacing: 0.5px; }}
        .btn-reset:hover {{ background: var(--artimo-red); }}
        
        .table-card {{ width: 100%; padding: 0; overflow-x: auto; border-radius: 12px; border: 1px solid var(--border-color); }}
        table.tabla-maquinas {{ width: 100%; border-collapse: collapse; white-space: nowrap; font-size: 13px; }}
        th, td {{ padding: 12px 14px; text-align: center; border-bottom: 1px solid var(--border-color); }}
        th {{ background-color: var(--artimo-dark); color: var(--blanco); font-weight: 600; text-transform: uppercase; font-size: 11px; letter-spacing: 0.5px; }}
        tr.data-row:hover {{ background-color: var(--artimo-light); }}
        
        .disk-status {{ padding: 3px 10px; border-radius: 12px; font-size: 11px; font-weight: 600; display: inline-block; }}
        .disk-status.Normal {{ background-color: #E8F8F5; color: #2A9D8F; }}
        .disk-status.Falla {{ background-color: #FDEDEC; color: #C8102E; }}
        .disk-status.SinDisco {{ background-color: var(--artimo-light); color: var(--artimo-grey); }}
        
        .footer-firma {{ text-align: center; padding: 30px 20px; margin-top: 20px; color: var(--artimo-grey); font-size: 12px; letter-spacing: 0.5px; border-top: 1px solid var(--border-color); font-weight: 300; }}
        .footer-firma span {{ font-size: 13px; font-weight: 700; color: var(--artimo-dark); }}

        /* ESTILOS EXCLUSIVOS PARA IMPRESIÓN (PDF) */
        @media print {{
            body {{ background-color: #ffffff; color: #000000; font-size: 10px; }}
            .top-navbar, .navbar-actions, .filter-bar, .chart-info-text, .footer-firma {{ display: none !important; }}
            .dashboard-container {{ padding: 0; gap: 15px; }}
            .kpi-card, .card {{ box-shadow: none !important; border: 1px solid #ccc !important; page-break-inside: avoid; }}
            .chart-card {{ max-width: 31% !important; min-height: auto !important; height: 300px !important; }}
            .table-card {{ border: none !important; margin-top: 15px; }}
            table.tabla-maquinas {{ font-size: 10px !important; width: 100% !important; }}
            th {{ background-color: #333 !important; color: #fff !important; }}
            td, th {{ padding: 6px 8px !important; }}
        }}
    </style>
</head>
<body>
    <div class="top-navbar">
        <div class="brand-container">
            {html_logo}
            <div class="navbar-title">Diagnóstico Continental Gold</div>
        </div>
        <div class="navbar-actions">
            <span class="timestamp">🔄 {fecha_actualizacion}</span>
            <button class="btn-download" onclick="exportCSV()">
                📥 Data Cruda (.csv)
            </button>
            <button class="btn-download primary" onclick="exportPDF()">
                📄 Exportar PDF
            </button>
        </div>
    </div>
    
    <div class="dashboard-container">
        <div class="kpi-section">
            <div class="kpi-card"><h3>Total de Máquinas</h3><p class="number">{total_maquinas}</p></div>
            <div class="kpi-card highlight"><h3>Equipos Operando</h3><p class="number" style="color: #2A9D8F;">{operando_cnt}</p></div>
            <div class="kpi-card"><h3>Offline >15 Días (Crítico)</h3><p class="number" style="color: #E9C46A;">{equipos_criticos_antiguedad}</p></div>
            <div class="kpi-card highlight"><h3>Alertas Hardware</h3><p class="number" style="color: #C8102E;">{alertas_hardware}</p></div>
        </div>
        
        <div class="charts-section">
            <div class="card chart-card">
                <div class="chart-container"><canvas id="graficaTransmision"></canvas></div>
                <div class="chart-info-text">⚫ <strong>Operando:</strong> Transmitió hace menos de 5 days.<br>🔴 <strong>Falla:</strong> 5 días o más sin reportar datos.</div>
            </div>
            <div class="card chart-card">
                <div class="chart-container"><canvas id="graficaDisco"></canvas></div>
                <div class="chart-info-text">💡 Filtra la tabla por estado de disco duro.</div>
            </div>
            <div class="card chart-card">
                <div class="chart-container"><canvas id="graficaCamaras"></canvas></div>
                <div class="chart-info-text">💡 Filtra por equipos con cámaras Dañadas u OK.</div>
            </div>
        </div>
        
        <div class="filter-bar" id="filterBar" style="display: none;">
            <div class="filter-msg" id="filterMessage">Mostrando resultados filtrados</div>
            <button class="btn-reset" onclick="resetFilters()">Mostrar Todos</button>
        </div>
        
        <div class="card table-card">
            {html_table}
        </div>
        <div class="footer-firma">Dashboard realizado por <span>Carlos Colorado</span></div>
    </div>
    
    <script>
        // Inyectamos la data cruda directamente en formato JSON
        const rawJsonData = {json_raw_data};

        function exportCSV() {{
            if (!rawJsonData || rawJsonData.length === 0) return;
            
            // Extraer cabeceras del JSON
            const headers = Object.keys(rawJsonData[0]);
            const csvRows = [headers.join(',')];
            
            for (const row of rawJsonData) {{
                const values = headers.map(header => {{
                    const val = row[header];
                    // Escapar strings que contengan comas
                    const valEscaped = ('' + (val !== null ? val : '')).replace(/"/g, '\\"');
                    return `"${{valEscaped}}"`;
                }});
                csvRows.push(values.join(','));
            }}
            
            const csvContent = "data:text/csv;charset=utf-8,\\uFEFF" + csvRows.join('\\n');
            const encodedUri = encodeURI(csvContent);
            const link = document.createElement("a");
            link.setAttribute("href", encodedUri);
            link.setAttribute("download", `Data_Cruda_Artimo_${{new Date().toISOString().split('T')[0]}}.csv`);
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }}

        function exportPDF() {{
            window.print();
        }}

        function filterTable(category, value, labelName) {{
            const rows = document.querySelectorAll('.data-row');
            let count = 0;
            rows.forEach(row => {{
                let rowValue = '';
                if (category === 'trans') rowValue = row.getAttribute('data-trans');
                if (category === 'disk') rowValue = row.getAttribute('data-disk');
                if (category === 'cam') rowValue = row.getAttribute('data-cam');
                
                if (rowValue === value) {{
                    row.style.display = '';
                    count++;
                } else {{
                    row.style.display = 'none';
                }}
            }});
            document.getElementById('filterBar').style.display = 'flex';
            document.getElementById('filterMessage').innerText = `🔎 Filtrado por: ${{labelName}} (${{count}} equipos)`;
            document.getElementById('filterBar').scrollIntoView({{ behavior: 'smooth', block: 'center' }});
        }}

        function resetFilters() {{
            document.querySelectorAll('.data-row').forEach(row => row.style.display = '');
            document.getElementById('filterBar').style.display = 'none';
        }}

        const titleOptions = (titleText) => ({{ display: true, text: titleText, font: {{ size: 14, weight: '700', family: 'var(--font-family)' }}, color: 'var(--artimo-dark)', padding: {{bottom: 12}} }});
        
        const onChartClick = (category) => (evt, elements, chart) => {{
            if (elements.length === 0) return;
            const index = elements[0].index;
            const labelFull = chart.data.labels[index];
            let value = '';
            if (category === 'trans') value = labelFull.includes('Operando') ? 'Operando' : 'Falla';
            if (category === 'disk') value = labelFull.includes('Normal') ? 'Normal' : (labelFull.includes('Falla') ? 'Falla' : 'No Detectado');
            if (category === 'cam') value = labelFull.includes('OK') ? 'Normal' : 'Falla';
            filterTable(category, value, labelFull);
        }};

        new Chart(document.getElementById('graficaTransmision').getContext('2d'), {{ 
            type: 'doughnut', 
            data: {{ labels: ['Operando', 'Falla de Transmisión'], datasets: [{{ data: [{operando_cnt}, {falla_trans_cnt}], backgroundColor: ['#2A9D8F', '#C8102E'], borderWidth: 1 }}] }}, 
            options: {{ responsive: true, maintainAspectRatio: false, plugins: {{ legend: {{ position: 'bottom', labels: {{ boxWidth: 10, font: {{ size: 11 }} }} }}, title: titleOptions('Estado de Transmisión') }}, onClick: onChartClick('trans') }} 
        }});
        
        new Chart(document.getElementById('graficaDisco').getContext('2d'), {{ 
            type: 'doughnut', 
            data: {{ labels: ['Normal', 'Falla', 'No Detectado'], datasets: [{{ data: [{disco_normal_cnt}, {disco_falla_cnt}, {disco_nodet_cnt}], backgroundColor: ['#2A9D8F', '#C8102E', '#5A5A59'], borderWidth: 1 }}] }}, 
            options: {{ responsive: true, maintainAspectRatio: false, plugins: {{ legend: {{ position: 'bottom', labels: {{ boxWidth: 10, font: {{ size: 11 }} }} }}, title: titleOptions('Estado del Disco Duro') }}, onClick: onChartClick('disk') }} 
        }});
        
        new Chart(document.getElementById('graficaCamaras').getContext('2d'), {{ 
            type: 'pie', 
            data: {{ labels: ['Equipos 100% OK', 'Equipos con Cámara Dañada'], datasets: [{{ label: 'Unidad', data: [{(dashboard_data['Status_Transmision'] != 'N/A').sum() - total_cam_falla}, {total_cam_falla}], backgroundColor: ['#2A9D8F', '#E9C46A'], borderWidth: 1 }}] }}, 
            options: {{ responsive: true, maintainAspectRatio: false, plugins: {{ legend: {{ position: 'bottom', labels: {{ boxWidth: 10, font: {{ size: 11 }} }} }}, title: titleOptions('Salud de Cámaras') }}, onClick: onChartClick('cam') }} 
        }});
    </script>
</body>
</html>
'''

ruta_guardado = os.path.join(carpeta_public, 'index.html')
with open(ruta_guardado, 'w', encoding='utf-8') as f:
    f.write(plantilla_base)

print(f"¡Dashboard analítico definitivo generado con éxito!")
