import pandas as pd
import os
import glob
import base64

print("Iniciando generación del Dashboard Corporativo...")

carpeta_actual = os.getcwd()
carpeta_data = os.path.join(carpeta_actual, 'data')
carpeta_public = os.path.join(carpeta_actual, 'public')

# Crear carpeta public automáticamente
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

# 2. DATOS
archivos_datos = glob.glob(os.path.join(carpeta_data, '*.xlsx')) + glob.glob(os.path.join(carpeta_data, '*.csv'))
if not archivos_datos:
    print("Error: No se encontró NINGÚN archivo de datos en la carpeta data.")
    exit(1)

ruta_archivo = archivos_datos[0]
if ruta_archivo.lower().endswith('.csv'):
    df = pd.read_csv(ruta_archivo, on_bad_lines='skip', engine='python')
else:
    df = pd.read_excel(ruta_archivo)

if not df.empty and df.iloc[0, 0] == 'Estado en línea':
    df = df.drop(index=0)
df.columns = df.columns.str.strip()

columnas_base = ['Número de matrícula', 'Tiempo fuera de línea', 'Estado de salud del almacenamiento 1']
dashboard_data = df[columnas_base].copy()
dashboard_data.columns = ['Máquina', 'Última transmisión', 'Estado del Disco 1']
total_maquinas = len(dashboard_data)

def mapear_estado_disco(estado):
    if pd.isna(estado) or estado == 'Ninguna descripción' or str(estado).strip() == '': return 'No se detecta disco duro'
    elif estado == 'Normal': return 'Normal'
    elif estado == 'Daños leves': return 'Falla'
    else: return 'No se detecta disco duro'

dashboard_data['Estado del Disco 1'] = dashboard_data['Estado del Disco 1'].apply(mapear_estado_disco)
conteo_discos = dashboard_data['Estado del Disco 1'].value_counts()
disco_normal_cnt = int(conteo_discos.get('Normal', 0))
disco_falla_cnt = int(conteo_discos.get('Falla', 0))
disco_nodet_cnt = int(conteo_discos.get('No se detecta disco duro', 0))

limite = pd.to_datetime('2026-06-30 23:59:59')
def evaluar_transmision(val):
    if str(val).strip().lower() == 'en línea': return 'Operando'
    try: return 'Operando' if pd.to_datetime(val) > limite else 'Falla de Transmisión'
    except: return 'Falla de Transmisión'

dashboard_data['Temporal_Status'] = dashboard_data['Última transmisión'].fillna('Falla_Temp').apply(evaluar_transmision)
conteo_transmisiones = dashboard_data['Temporal_Status'].value_counts()
dashboard_data['Última transmisión'] = dashboard_data['Última transmisión'].fillna('En línea')
dashboard_data = dashboard_data.drop(columns=['Temporal_Status'])

operando_cnt = int(conteo_transmisiones.get('Operando', 0))
falla_trans_cnt = int(conteo_transmisiones.get('Falla de Transmisión', 0))

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
            hab = str(row.get(h, '')).strip()
            est = str(row.get(e, '')).strip()
            if hab == 'Abrir': return 'Normal' if est == 'Normal' else 'Falla'
            return 'N/A'
        dashboard_data[cam_name] = df.apply(evaluar_camara, axis=1)
        total_cam_ok += (dashboard_data[cam_name] == 'Normal').sum()
        total_cam_falla += (dashboard_data[cam_name] == 'Falla').sum()

alertas_hardware = disco_falla_cnt + disco_nodet_cnt + total_cam_falla
display_data = dashboard_data.copy()

def style_disk_status(st):
    if st == 'Normal': return '<span class="disk-status Normal">Normal</span>'
    if st == 'Falla': return '<span class="disk-status Falla">Falla</span>'
    return '<span class="disk-status SinDisco">Sin Disco</span>'
display_data['Estado del Disco 1'] = display_data['Estado del Disco 1'].apply(style_disk_status)

def style_cam_status(st):
    if st == 'Normal': return '<span class="disk-status Normal">✓ OK</span>'
    if st == 'Falla': return '<span class="disk-status Falla">✗ Falla</span>'
    return '<span style="color: #A0AEC0; font-weight: bold;">-</span>'

for _, _, cam_name in camaras_encontradas:
    display_data[cam_name] = display_data[cam_name].apply(style_cam_status)

html_table = display_data.to_html(index=False, classes='tabla-maquinas', escape=False)

# 3. PLANTILLA HTML
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
        .chart-card {{ flex: 1; min-width: 300px; max-width: 32%; min-height: 420px; }}
        .chart-container {{ position: relative; width: 100%; height: 280px; flex-grow: 1; }}
        .chart-info-text {{ font-size: 12px; color: #555555; background-color: #F9F9F9; padding: 12px; margin-top: 15px; border-radius: 6px; border-left: 4px solid var(--artimo-rojo); line-height: 1.5; }}
        .table-card {{ width: 100%; padding: 0; overflow-x: auto; }}
        table.tabla-maquinas {{ width: 100%; border-collapse: collapse; white-space: nowrap; }}
        th, td {{ padding: 12px 15px; text-align: center; border-bottom: 1px solid #EAEAEA; }}
        th {{ background-color: var(--artimo-rojo); color: var(--blanco); font-weight: 600; text-transform: uppercase; font-size: 12px; }}
        .disk-status {{ padding: 4px 10px; border-radius: 12px; font-size: 12px; font-weight: 600; display: inline-block; }}
        .disk-status
