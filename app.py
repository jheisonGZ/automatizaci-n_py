"""
================================================================================
AUTOMATIZACIÓN DE REGISTRO MASIVO - ALTERNATIVA LIBERAL POPULAR
Backend Flask — Código original + interfaz web
================================================================================
"""

from flask import Flask, render_template, request, jsonify, send_file
from flask_socketio import SocketIO
import pandas as pd
import unicodedata
import threading
import tempfile
import time
import os
import json
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
import undetected_chromedriver as uc

app = Flask(__name__)
app.config['SECRET_KEY'] = 'alp-secret-2024'
socketio = SocketIO(app, cors_allowed_origins="*", async_mode='threading')

# ── Archivos temporales compatibles Windows/Linux ──────────
TEMP_DIR    = tempfile.gettempdir()
DATOS_JSON  = os.path.join(TEMP_DIR, 'datos_proceso.json')
ERRORES_XLS = os.path.join(TEMP_DIR, 'NO_SUBIDOS.xlsx')

# ── Estado global ──────────────────────────────────────────
proceso_activo    = False
proceso_cancelado = False

COLUMNAS_REQUERIDAS = {
    "CEDULA", "NOMBRES", "APELLIDOS",
    "TELEFONO", "DIRECCION", "DEPARTAMENTO", "MUNICIPIO"
}


# ==============================================================================
# FUNCIÓN AUXILIAR — igual que en automa.py original
# ==============================================================================
def normalizar(texto):
    return ''.join(
        c for c in unicodedata.normalize('NFD', str(texto))
        if unicodedata.category(c) != 'Mn'
    ).strip().upper()


def log(msg, tipo="info"):
    socketio.emit('log', {'msg': msg, 'tipo': tipo})


def progreso(actual, total, cedula="", depto="", municipio=""):
    socketio.emit('progreso', {
        'actual':      actual,
        'total':       total,
        'cedula':      cedula,
        'depto':       depto,
        'municipio':   municipio,
        'porcentaje':  round((actual / total) * 100) if total > 0 else 0
    })


# ==============================================================================
# RUTAS FLASK
# ==============================================================================

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/validar-excel', methods=['POST'])
def validar_excel():
    if 'file' not in request.files:
        return jsonify({'ok': False, 'error': 'No se recibió archivo'})

    file = request.files['file']
    if file.filename == '':
        return jsonify({'ok': False, 'error': 'Nombre de archivo vacío'})

    try:
        df = pd.read_excel(file)
        df.columns = [normalizar(col) for col in df.columns]

        columnas_presentes = set(df.columns)
        columnas_faltantes = COLUMNAS_REQUERIDAS - columnas_presentes

        if columnas_faltantes:
            return jsonify({
                'ok':          False,
                'faltantes':   sorted(list(columnas_faltantes)),
                'encontradas': list(df.columns),
                'requeridas':  sorted(list(COLUMNAS_REQUERIDAS))
            })

        if df.empty:
            return jsonify({'ok': False, 'error': 'El archivo no contiene filas de datos'})

        df.to_json(DATOS_JSON, orient='records', force_ascii=False)

        return jsonify({
            'ok':      True,
            'total':   len(df),
            'columnas': list(df.columns),
            'preview':  df.head(3).to_dict(orient='records')
        })

    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)})


@app.route('/iniciar', methods=['POST'])
def iniciar():
    global proceso_activo, proceso_cancelado

    if proceso_activo:
        return jsonify({'ok': False, 'error': 'Ya hay un proceso en curso'})

    data = request.json
    codigo_referido = data.get('codigo_referido', '').strip()

    if not codigo_referido:
        return jsonify({'ok': False, 'error': 'Código de referido requerido'})

    if not os.path.exists(DATOS_JSON):
        return jsonify({'ok': False, 'error': 'No hay datos cargados. Sube el Excel primero.'})

    proceso_activo    = True
    proceso_cancelado = False

    thread = threading.Thread(target=ejecutar_proceso, args=(codigo_referido,))
    thread.daemon = True
    thread.start()

    return jsonify({'ok': True})


@app.route('/cancelar', methods=['POST'])
def cancelar():
    global proceso_cancelado
    proceso_cancelado = True
    log('⛔ Cancelación solicitada por el usuario...', 'warning')
    return jsonify({'ok': True})


@app.route('/descargar-errores')
def descargar_errores():
    if os.path.exists(ERRORES_XLS):
        return send_file(ERRORES_XLS, as_attachment=True, download_name='NO_SUBIDOS.xlsx')
    return jsonify({'error': 'No hay archivo de errores'}), 404


# ==============================================================================
# PROCESO SELENIUM — lógica 100% original de automa.py
# ==============================================================================
def ejecutar_proceso(codigo_referido):
    global proceso_activo, proceso_cancelado

    errores_definitivos = []

    try:
        with open(DATOS_JSON, 'r', encoding='utf-8') as f:
            registros = json.load(f)

        df    = pd.DataFrame(registros)
        total = len(df)

        log(f'🚀 Iniciando proceso con {total} registros...', 'success')
        log(f'🔑 Código de referido: {codigo_referido}', 'info')

        # ── Mismas opciones que automa.py original ─────────────────────────
        options = uc.ChromeOptions()
        options.add_argument("--start-maximized")

        driver = uc.Chrome(options=options, version_main=145)
        driver.set_window_position(0, 0)

        wait   = WebDriverWait(driver, 20)
        errores = []
        exitos  = []

        try:
            driver.get("https://dos.alternativaliberalpopular.org/registro.php")
            time.sleep(3)

            # ── PRIMERA VUELTA ─────────────────────────────────────────────
            for index, row in df.iterrows():
                if proceso_cancelado:
                    log('⛔ Proceso cancelado.', 'warning')
                    break

                cedula = str(row["CEDULA"]).replace(".", "").strip()

                try:
                    nombres      = str(row["NOMBRES"]).strip()
                    apellidos    = str(row["APELLIDOS"]).strip()
                    telefono     = str(row["TELEFONO"]).replace(".0", "").strip()
                    direccion    = str(row["DIRECCION"]).strip()
                    departamento = str(row["DEPARTAMENTO"]).strip().upper()
                    municipio    = str(row["MUNICIPIO"]).strip().upper()

                    wait.until(EC.presence_of_element_located((By.NAME, "identification")))

                    driver.find_element(By.NAME, "identification").clear()
                    driver.find_element(By.NAME, "identification").send_keys(cedula)
                    time.sleep(1)

                    driver.find_element(By.NAME, "identification_confirm").clear()
                    driver.find_element(By.NAME, "identification_confirm").send_keys(cedula)
                    time.sleep(0.5)

                    driver.find_element(By.NAME, "name").send_keys(nombres)
                    driver.find_element(By.NAME, "lastName").send_keys(apellidos)

                    telefono = telefono[:10]
                    driver.find_element(By.NAME, "cellPhone").send_keys(telefono)

                    Select(driver.find_element(By.NAME, "department")).select_by_visible_text(departamento)
                    wait.until(lambda d: len(Select(d.find_element(By.NAME, "municipalityid")).options) > 1)

                    municipio_select = Select(driver.find_element(By.NAME, "municipalityid"))
                    municipio_select.select_by_visible_text(municipio)
                    time.sleep(0.5)

                    driver.find_element(By.NAME, "neighborhood").send_keys("Centro")
                    driver.find_element(By.NAME, "direction").send_keys(direccion)
                    driver.find_element(By.NAME, "whoInvited").send_keys(codigo_referido)

                    checkbox = wait.until(EC.element_to_be_clickable((By.ID, "terminos")))
                    driver.execute_script("arguments[0].click();", checkbox)
                    time.sleep(0.5)

                    boton = wait.until(EC.element_to_be_clickable((By.ID, "btnRegistrar")))
                    driver.execute_script("arguments[0].click();", boton)

                    log(f'✔ [{index+1}/{total}] Enviado: {cedula} — {departamento} / {municipio}', 'success')
                    exitos.append(cedula)
                    progreso(len(exitos), total, cedula, departamento, municipio)

                    time.sleep(6)
                    driver.get("https://dos.alternativaliberalpopular.org/registro.php")
                    time.sleep(3)

                except Exception as e:
                    log(f'❌ [{index+1}/{total}] Error con {cedula}: {e}', 'error')
                    errores.append(row)
                    driver.get("https://dos.alternativaliberalpopular.org/registro.php")
                    time.sleep(3)

            # ── SEGUNDA VUELTA ─────────────────────────────────────────────
            log('🔁 Intentando nuevamente los fallidos...', 'warning')

            for i, row in enumerate(errores):
                if proceso_cancelado:
                    break

                cedula       = str(row["CEDULA"]).replace(".", "").strip()
                departamento = str(row["DEPARTAMENTO"]).strip().upper()
                municipio    = str(row["MUNICIPIO"]).strip().upper()

                try:
                    wait.until(EC.presence_of_element_located((By.NAME, "identification")))

                    driver.find_element(By.NAME, "identification").clear()
                    driver.find_element(By.NAME, "identification").send_keys(cedula)
                    time.sleep(1)

                    driver.find_element(By.NAME, "identification_confirm").clear()
                    driver.find_element(By.NAME, "identification_confirm").send_keys(cedula)
                    time.sleep(0.5)

                    driver.find_element(By.NAME, "name").send_keys(str(row["NOMBRES"]))
                    driver.find_element(By.NAME, "lastName").send_keys(str(row["APELLIDOS"]))

                    telefono = str(row["TELEFONO"]).replace(".0", "")[:10]
                    driver.find_element(By.NAME, "cellPhone").send_keys(telefono)

                    Select(driver.find_element(By.NAME, "department")).select_by_visible_text(departamento)
                    wait.until(lambda d: len(Select(d.find_element(By.NAME, "municipalityid")).options) > 1)

                    municipio_select = Select(driver.find_element(By.NAME, "municipalityid"))
                    municipio_select.select_by_visible_text(municipio)
                    time.sleep(0.5)

                    driver.find_element(By.NAME, "neighborhood").send_keys("Centro")
                    driver.find_element(By.NAME, "direction").send_keys(str(row["DIRECCION"]))
                    driver.find_element(By.NAME, "whoInvited").send_keys(codigo_referido)

                    checkbox = wait.until(EC.element_to_be_clickable((By.ID, "terminos")))
                    driver.execute_script("arguments[0].click();", checkbox)
                    time.sleep(0.5)

                    boton = wait.until(EC.element_to_be_clickable((By.ID, "btnRegistrar")))
                    driver.execute_script("arguments[0].click();", boton)

                    log(f'✅ [{i+1}/{len(errores)}] Recuperado: {cedula} — {departamento} / {municipio}', 'success')

                    time.sleep(6)
                    driver.get("https://dos.alternativaliberalpopular.org/registro.php")
                    time.sleep(3)

                except Exception as e:
                    log(f'⛔ [{i+1}/{len(errores)}] No se pudo subir: {cedula} — {e}', 'error')
                    errores_definitivos.append(row)
                    driver.get("https://dos.alternativaliberalpopular.org/registro.php")
                    time.sleep(3)

            # ── EXPORTAR FALLIDOS ──────────────────────────────────────────
            if errores_definitivos:
                pd.DataFrame(errores_definitivos).to_excel(ERRORES_XLS, index=False)
                log('📁 Archivo NO_SUBIDOS.xlsx listo para descargar.', 'warning')

        finally:
            driver.quit()  # ← igual que automa.py original, evita WinError 6

        # ── RESUMEN FINAL ──────────────────────────────────────────────────
        log('========== RESUMEN FINAL ==========', 'info')
        log(f'📥 Total procesados  : {total}', 'info')
        log(f'✔  Exitosos          : {len(exitos)}', 'success')
        log(f'🔁 Reintentados      : {len(errores)}', 'warning')
        log(f'❌ Fallidos definitiv: {len(errores_definitivos)}', 'error')
        log(f'🔑 Código referido   : {codigo_referido}', 'info')
        log('====================================', 'info')
        log('🏁 Proceso terminado.', 'success')

        socketio.emit('resumen', {
            'total':              total,
            'exitos':             len(exitos),
            'errores':            len(errores),
            'definitivos':        len(errores_definitivos),
            'tiene_errores_xlsx': os.path.exists(ERRORES_XLS)
        })

    except Exception as e:
        log(f'💥 Error crítico: {e}', 'error')

    finally:
        proceso_activo = False


# ==============================================================================
if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    socketio.run(app, host='0.0.0.0', port=port, debug=False)