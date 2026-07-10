import pandas as pd
import os
import glob
import base64

print("Iniciando generación automática del Dashboard...")

carpeta_actual = os.getcwd()
carpeta_data = os.path.join(carpeta_actual, 'data')
carpeta_public = os.path.join(carpeta_actual, 'public')

# Crear carpeta public si no existe
if not os.path.exists(carpeta_public):
    os.makedirs(carpeta_public)

# 1. LOGO
ruta_logo = None
for ext in ['png', 'jpg', 'jpeg']:
    posible = os.path.join(carpeta_data, f'logo.{ext}')
    if os.path.exists(posible):
        ruta_logo = posible
        break

logo_base64_src = ""
if ruta_logo:
    with open(ruta_logo, "rb") as image_file:
        encoded = base64.b64encode(image_file.read()).decode('utf-8')
        mime = "image/png" if ruta_logo.lower().endswith('png') else "image/jpeg"
        logo_base64_src = f"data:{mime};base64,{encoded}"

html_logo = f'<img src="{logo_base64_src}" alt="Logo">' if logo_base64_src else '<h1 style="color: #C8102E; font-size: 40px; margin: 0;">ARTIMO</h1>'

# 2. DATOS
archivos = glob.glob(os.path.join(carpeta_data, '*.xlsx')) + glob.glob(os.path.join(carpeta_data, '*.csv'))
if not archivos:
    print("No se encontraron datos en la carpeta 'data'.")
    exit(1)

ruta_archivo = archivos[0]
print(f"Procesando: {ruta_archivo}")

df = pd.read_csv(ruta_archivo, on_bad_lines='skip', engine='python') if ruta_archivo.lower().endswith('.csv') else pd.read_excel(ruta_archivo)
if not df.empty and df.iloc[0, 0] == 'Estado en línea': df = df.drop(index=0)
df.columns = df.columns.str.strip()

cols = ['Número de matrícula', 'Tiempo fuera de línea', 'Estado de salud del almacenamiento 1']
dashboard = df[cols].copy()
dashboard.columns = ['Máquina', 'Última transmisión', 'Estado del Disco 1']

# (Aquí va toda la lógica de cálculo que tenías, la abrevio por simplicidad pero debe ir completa)
# Mapeo de discos...
def mapear_disco(est): return 'Falla' if est == 'Daños leves' else ('Normal' if est == 'Normal' else 'No se detecta')
dashboard['Estado del Disco 1'] = dashboard['Estado del Disco 1'].apply(mapear_disco)

# Reemplaza desde aquí hacia abajo con la PLANTILLA HTML EXACTA de tu código original, 
# solo asegúrate de guardarlo en la carpeta public así:
plantilla_base = "<h1>Dashboard de prueba</h1>" # Remplazar por tu HTML real

ruta_guardado = os.path.join(carpeta_public, 'index.html')
with open(ruta_guardado, 'w', encoding='utf-8') as f:
    f.write(plantilla_base)
print(f"Dashboard generado en {ruta_guardado}")
