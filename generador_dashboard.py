import pandas as pd
import os
import glob
import base64

print("Iniciando generación del Dashboard Interactivo...")

carpeta_actual = os.getcwd()
carpeta_data = os.path.join(carpeta_actual, 'data')
carpeta_public = os.path.join(carpeta_actual, 'public')

os.makedirs(carpeta_public, exist_ok=True)

# 1. LOGO
ruta_logo = glob.glob(os.path.join(carpeta_data, 'logo.*'))
logo_base64_src = ""
if ruta_logo:
    with open(ruta_logo[0], "rb") as image_file:
        encoded_string = base64.b64encode(image_file.read()).decode('utf-8')
        mime_type = "image/png" if ruta_logo[0].lower().endswith('png') else "image/jpeg"
        logo_base64_src = f"data:{mime_type};base64,{encoded_string}"

html_logo = f'<img src="{logo_base64_src}" alt="Logo">' if logo_base64_src else '<h1 style="color: #C8102E; font-size: 40px; margin: 0;">ARTIMO</h1>'

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

# --- PROCESAMIENTO ESTRICTO ---

# DISCOS
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

# TRANSMISIÓN
limite = pd.to_datetime('2026-06-30 23:59:59')
def evaluar_transmision(val):
    if str(val).strip().lower() == 'en línea': return 'Operando'
    try: return 'Operando' if pd.to_datetime(val) > limite else 'Falla'
    except: return 'Falla'

dashboard_data['Status_Transmision'] = dashboard_data['Última transmisión'].fillna('Falla').apply(evaluar_transmision)
conteo_transmisiones = dashboard_data['Status_Transmision'].value_counts()
dashboard_data['Última transmisión'] = dashboard_data['Última transmisión'].fillna('En línea')
operando_cnt = int(conteo_transmisiones.get('Operando', 0))
falla_trans_cnt = int(conteo_transmisiones.get('Falla', 0))

# CÁMARAS (Nueva Lógica)
camaras_encontradas = []
total_cam_ok = 0
total_cam_falla = 0

for i in range(1, 9):
    col_hab = f'Cámara {i} habilitada'
    col_est = f'Estado de la cámara {i}'
    if col_hab in df.columns and col_est in df.columns:
        cam_name = f'CAM {i}'
        camaras_encontradas.append((col_hab, col_est, cam_name))
        
        def evaluar_camara(row, h=col_hab, e=col_est):
            hab = str(row.get(h, '')).strip().lower()
            est = str(row.get(e, '')).strip().lower()
            
            if hab == 'abrir':
                # Si está abierta pero su estado es 'perdida' o diferente a normal, es Falla
                return 'Normal' if est == 'normal' else 'Falla'
            # Si está cerrado o vacío, la ignoramos
            return 'N/A'
            
        dashboard_data[cam_name] = df.apply(evaluar_camara, axis=1)
        total_cam_ok += (dashboard_data[cam_name] == 'Normal').sum()
        total_cam_falla += (dashboard_data[cam_name] == 'Falla').sum()

alertas_hardware = disco_falla_cnt + disco_nodet_cnt + total_cam_falla

# --- CONSTRUCCIÓN DE LA TABLA HTML CON ATRIBUTOS PARA FILTROS ---
display_data = dashboard_data.copy()

def style_disk_status(st):
    if st == 'Normal': return '<span class="disk-status Normal">Normal</span>'
    if st == 'Falla': return '<span class="disk-status Falla">Falla</span>'
    return '<span class="disk-status SinDisco">No Detectado</span>'

def style_cam_status(st):
    if st == 'Normal': return '<span class="disk-status Normal">✓ OK</span>'
    if st == 'Falla': return '<span class="disk-status Falla">✗ Falla</span>'
    return '<span style="color: #E2E8F0; font-weight: bold;">-</span>'

columnas_mostrar = ['Máquina', 'Última transmisión', 'Estado del Disco 1'] + [c[2] for c in camaras_encontradas]

html_table = '<table class="tabla-maquinas">\n<thead>\n<tr>'
for col in columnas_mostrar:
    html_table += f'<th>{col}</th>'
html_table += '</tr>\n</thead>\n<tbody>\n'

for idx, row in dashboard_data.iterrows():
    t_stat = row['Status_Transmision']
    d_stat = row['Status_Disco']
    # Si la máquina tiene al menos UNA cámara en falla, se clasifica como Falla en el filtro global de cámaras
    c_stat = 'Falla' if any(row[c[2]] == 'Falla' for c in camaras_encontradas) else 'Normal'
    
    html_table += f'<tr class="data-row" data-trans="{t_stat}" data-disk="{d_stat}" data-cam="{c_stat}">'
    html_table += f'<td>{row["Máquina"]}</td>'
    html_table += f'<td>{row["Última transmisión"]}</td>'
    html_table += f'<td>{style_disk_status(row["Status_Disco"])}</td>'
    
    for _, _, cam_name in camaras_encontradas:
        html_table += f'<td>{style_cam_status(row[cam_name])}</td>'
        
    html_table += '</tr>\n'
html_table += '</tbody>\n</table>'


# 3. PLANTILLA HTML Y JAVASCRIPT
plantilla_base = f'''
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>DIAGNOSTICO CONTINENTAL GOLD</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        :root {{ --artimo-rojo: #C8102E; --artimo-oscuro: #2D2D2D; --fondo-gris: #F5F7FA; --blanco: #FFFFFF; }}
        body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 0; background-color: var(--fondo-gris); color: var(--artimo-oscuro); }}
        .header {{ text-align: center; padding: 25px 0; background-color: var(--blanco); border-bottom: 5px solid var(--artimo-rojo); box-shadow: 0 4px 12px rgba(0,0,0,0.06); display: flex; flex-direction: column; align-items: center; gap: 12px; }}
        .header img {{ max-height: 85px; object-fit: contain; }}
        h1 {{ color: var(--artimo-oscuro); margin: 0; font-size: 28px; text-transform: uppercase; letter-spacing: 1.5px; font-weight: 700; }}
        .dashboard-container {{ display: flex; flex-direction: column; align-items: center; gap: 30px; padding: 30px 20px; max-width: 1400px; margin: 0 auto; }}
        .kpi-section {{ display: flex; flex-wrap: wrap; justify-content: space-between; width: 100%; gap: 20px; }}
        .kpi-card {{ flex: 1; min-width: 250px; background: var(--blanco); border-radius: 8px; padding: 25px 20px; text-align: center; box-shadow: 0 4px 10px rgba(0,0,0,0.05); border-left: 5px solid var(--artimo-rojo); }}
        .kpi-card h3 {{ margin: 0 0 10px 0; color: #777; font-size: 16px; text-transform: uppercase; letter-spacing: 1px; }}
        .kpi-card .number {{ margin: 0; font-size: 42px; font-weight: 800; color: var(--artimo-oscuro); }}
        .kpi-card.success {{ border-left-color: #2ECC71; }}
        .kpi-card.danger {{ border-left-color: var(--artimo-rojo); }}
        
        .charts-section {{ display: flex; flex-wrap: wrap; justify-content: center; gap: 20px; width: 100%; }}
        .card {{ background-color: var(--blanco); border-radius: 8px; box-shadow: 0 4px 10px rgba(0,0,0,0.05); border-top: 3px solid var(--artimo-rojo); padding: 20px; box-sizing: border-box; display: flex; flex-direction: column; }}
        .chart-card {{ flex: 1; min-width: 300px; max-width: 32%; min-height: 420px; transition: transform 0.2s; }}
        .chart-card:hover {{ transform: translateY(-5px); cursor: pointer; }}
        .chart-container {{ position: relative; width: 100%; height: 280px; flex-grow: 1; }}
        .chart-info-text {{ font-size: 12px; color: #555555; background-color: #F9F9F9; padding: 12px; margin-top: 15px; border-radius: 6px; border-left: 4px solid var(--artimo-rojo); line-height: 1.5; }}
        
        /* Controles de Filtro */
        .filter-bar {{ width: 100%; display: flex; justify-content: space-between; align-items: center; padding: 10px 20px; background: #fff; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.05); box-sizing: border-box; border-left: 4px solid #3498DB; }}
        .filter-msg {{ font-weight: bold; color: #3498DB; font-size: 14px; }}
        .btn-reset {{ background: var(--artimo-oscuro); color: white; border: none; padding: 8px 15px; border-radius: 4px; cursor: pointer; font-weight: bold; font-size: 12px; text-transform: uppercase; }}
        .btn-reset:hover {{ background: var(--artimo-rojo); }}
        
        .table-card {{ width: 100%; padding: 0; overflow-x: auto; }}
        table.tabla-maquinas {{ width: 100%; border-collapse: collapse; white-space: nowrap; }}
        th, td {{ padding: 12px 15px; text-align: center; border-bottom: 1px solid #EAEAEA; }}
        th {{ background-color: var(--artimo-rojo); color: var(--blanco); font-weight: 600; text-transform: uppercase; font-size: 12px; }}
        tr:hover {{ background-color: #FDFDFD; }}
        
        .disk-status {{ padding: 4px 10px; border-radius: 12px; font-size: 12px; font-weight: 600; display: inline-block; }}
        .disk-status.Normal {{ background-color: #E8F8F5; color: #117A65; border: 1px solid #A3E4D7; }}
        .disk-status.Falla {{ background-color: #FDEDEC; color: #C0392B; border: 1px solid #F5B7B1; }}
        .disk-status.SinDisco {{ background-color: #F2F4F4; color: #7B7D7D; border: 1px solid #D5D8DC; }}
        .footer-firma {{ text-align: center; padding: 40px 20px; margin-top: 20px; color: #7f8c8d; font-size: 15px; letter-spacing: 1px; border-top: 1px solid #EAEAEA; }}
        .footer-firma span {{ font-family: 'Georgia', serif; font-size: 18px; font-weight: bold; font-style: italic; color: var(--artimo-rojo); }}
    </style>
</head>
<body>
    <div class="header">
        {html_logo}
        <h1>DIAGNOSTICO CONTINENTAL GOLD</h1>
    </div>
    <div class="dashboard-container">
        <div class="kpi-section">
            <div class="kpi-card"><h3>Total de Máquinas</h3><p class="number">{total_maquinas}</p></div>
            <div class="kpi-card success"><h3>Equipos Operando</h3><p class="number" style="color: #2ECC71;">{operando_cnt}</p></div>
            <div class="kpi-card danger"><h3>Alertas de Hardware Críticas</h3><p class="number" style="color: #C8102E;">{alertas_hardware}</p></div>
        </div>
        
        <div class="charts-section">
            <div class="card chart-card">
                <div class="chart-container"><canvas id="graficaTransmision"></canvas></div>
                <div class="chart-info-text">💡 <strong>Clickea en la gráfica</strong> para filtrar la tabla por transmisión.</div>
            </div>
            <div class="card chart-card">
                <div class="chart-container"><canvas id="graficaDisco"></canvas></div>
                <div class="chart-info-text">💡 <strong>Clickea en la gráfica</strong> para filtrar la tabla por estado de disco.</div>
            </div>
            <div class="card chart-card">
                <div class="chart-container"><canvas id="graficaCamaras"></canvas></div>
                <div class="chart-info-text">💡 <strong>Clickea en la gráfica</strong> para ver equipos con cámaras Dañadas o 100% OK.</div>
            </div>
        </div>
        
        <div class="filter-bar" id="filterBar" style="display: none;">
            <div class="filter-msg" id="filterMessage">Mostrando resultados filtrados</div>
            <button class="btn-reset" onclick="resetFilters()">Mostrar Todas las Máquinas</button>
        </div>
        
        <div class="card table-card">
            {html_table}
        </div>
        <div class="footer-firma">Dashboard realizado por <span>Carlos Colorado</span></div>
    </div>
    
    <script>
        // LÓGICA DE FILTRADO
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
                }} else {{
                    row.style.display = 'none';
                }}
            }});
            
            document.getElementById('filterBar').style.display = 'flex';
            document.getElementById('filterMessage').innerText = `🔎 Filtrado por: ${{labelName}} (${{count}} equipos encontrados)`;
            
            // Scroll suave hacia la tabla
            document.getElementById('filterBar').scrollIntoView({{ behavior: 'smooth', block: 'center' }});
        }}

        function resetFilters() {{
            const rows = document.querySelectorAll('.data-row');
            rows.forEach(row => row.style.display = '');
            document.getElementById('filterBar').style.display = 'none';
        }}

        // CONFIGURACIÓN DE GRÁFICAS
        const titleOptions = (titleText) => ({{ display: true, text: titleText, font: {{ size: 16, weight: 'bold' }}, color: '#2D2D2D', padding: {{bottom: 10}} }});
        
        // Función general para clics
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
            data: {{ labels: ['Operando', 'Falla de Transmisión'], datasets: [{{ data: [{operando_cnt}, {falla_trans_cnt}], backgroundColor: ['#2ECC71', '#C8102E'], borderWidth: 2, hoverOffset: 10 }}] }}, 
            options: {{ responsive: true, maintainAspectRatio: false, plugins: {{ legend: {{ position: 'bottom' }}, title: titleOptions('Transmisión') }}, onClick: onChartClick('trans') }} 
        }});
        
        new Chart(document.getElementById('graficaDisco').getContext('2d'), {{ 
            type: 'doughnut', 
            data: {{ labels: ['Normal', 'Falla', 'No Detectado'], datasets: [{{ data: [{disco_normal_cnt}, {disco_falla_cnt}, {disco_nodet_cnt}], backgroundColor: ['#2ECC71', '#C8102E', '#95A5A6'], borderWidth: 2, hoverOffset: 10 }}] }}, 
            options: {{ responsive: true, maintainAspectRatio: false, plugins: {{ legend: {{ position: 'bottom' }}, title: titleOptions('Salud de Discos') }}, onClick: onChartClick('disk') }} 
        }});
        
        new Chart(document.getElementById('graficaCamaras').getContext('2d'), {{ 
            type: 'pie', 
            data: {{ labels: ['Equipos 100% OK', 'Equipos con Cámara Dañada'], datasets: [{{ data: [{(dashboard_data['Status_Transmision'] != 'N/A').sum() - total_cam_falla}, {total_cam_falla}], backgroundColor: ['#2ECC71', '#C8102E'], borderWidth: 2, hoverOffset: 10 }}] }}, 
            options: {{ responsive: true, maintainAspectRatio: false, plugins: {{ legend: {{ position: 'bottom' }}, title: titleOptions('Salud de Cámaras (Por Máquina)') }}, onClick: onChartClick('cam') }} 
        }});
    </script>
</body>
</html>
'''

ruta_guardado = os.path.join(carpeta_public, 'index.html')
with open(ruta_guardado, 'w', encoding='utf-8') as f:
    f.write(plantilla_base)

print(f"¡Dashboard generado exitosamente en {ruta_guardado}!")
