"""
AUTOMATIZACIÓN DE REGISTRO MASIVO - ALTERNATIVA LIBERAL POPULAR
Versión optimizada para Render / Linux
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


TEMP_DIR = tempfile.gettempdir()
DATOS_JSON = os.path.join(TEMP_DIR, 'datos_proceso.json')
ERRORES_XLS = os.path.join(TEMP_DIR, 'NO_SUBIDOS.xlsx')

proceso_activo = False
proceso_cancelado = False


COLUMNAS_REQUERIDAS = {
    "CEDULA",
    "NOMBRES",
    "APELLIDOS",
    "TELEFONO",
    "DIRECCION",
    "DEPARTAMENTO",
    "MUNICIPIO"
}


def normalizar(texto):
    return ''.join(
        c for c in unicodedata.normalize('NFD', str(texto))
        if unicodedata.category(c) != 'Mn'
    ).strip().upper()


def log(msg, tipo="info"):
    socketio.emit('log', {'msg': msg, 'tipo': tipo})


def progreso(actual, total, cedula="", depto="", municipio=""):
    socketio.emit('progreso', {
        'actual': actual,
        'total': total,
        'cedula': cedula,
        'depto': depto,
        'municipio': municipio,
        'porcentaje': round((actual / total) * 100) if total > 0 else 0
    })


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/validar-excel', methods=['POST'])
def validar_excel():

    if 'file' not in request.files:
        return jsonify({'ok': False, 'error': 'No se recibió archivo'})

    file = request.files['file']

    try:

        df = pd.read_excel(file)
        df.columns = [normalizar(col) for col in df.columns]

        columnas_presentes = set(df.columns)
        columnas_faltantes = COLUMNAS_REQUERIDAS - columnas_presentes

        if columnas_faltantes:
            return jsonify({
                'ok': False,
                'faltantes': sorted(list(columnas_faltantes))
            })

        if df.empty:
            return jsonify({'ok': False, 'error': 'Archivo sin datos'})

        df.to_json(DATOS_JSON, orient='records', force_ascii=False)

        return jsonify({
            'ok': True,
            'total': len(df),
            'preview': df.head(3).to_dict(orient='records')
        })

    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)})


@app.route('/iniciar', methods=['POST'])
def iniciar():

    global proceso_activo, proceso_cancelado

    if proceso_activo:
        return jsonify({'ok': False, 'error': 'Proceso ya activo'})

    data = request.json
    codigo_referido = data.get('codigo_referido', '').strip()

    if not os.path.exists(DATOS_JSON):
        return jsonify({'ok': False, 'error': 'Sube primero el Excel'})

    proceso_activo = True
    proceso_cancelado = False

    thread = threading.Thread(target=ejecutar_proceso, args=(codigo_referido,))
    thread.start()

    return jsonify({'ok': True})


@app.route('/cancelar', methods=['POST'])
def cancelar():
    global proceso_cancelado
    proceso_cancelado = True
    log("Cancelación solicitada", "warning")
    return jsonify({'ok': True})


@app.route('/descargar-errores')
def descargar_errores():
    if os.path.exists(ERRORES_XLS):
        return send_file(ERRORES_XLS, as_attachment=True)
    return jsonify({'error': 'No existe archivo'})


def iniciar_driver():

    options = uc.ChromeOptions()

    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")

    driver = uc.Chrome(options=options)

    return driver


def ejecutar_proceso(codigo_referido):

    global proceso_activo

    errores_definitivos = []

    try:

        with open(DATOS_JSON, 'r', encoding='utf-8') as f:
            registros = json.load(f)

        df = pd.DataFrame(registros)

        total = len(df)

        log(f"Iniciando proceso con {total} registros", "success")

        driver = iniciar_driver()

        wait = WebDriverWait(driver, 20)

        errores = []
        exitos = []

        driver.get("https://dos.alternativaliberalpopular.org/registro.php")

        for index, row in df.iterrows():

            if proceso_cancelado:
                break

            cedula = str(row["CEDULA"]).replace(".", "")

            try:

                nombres = str(row["NOMBRES"])
                apellidos = str(row["APELLIDOS"])
                telefono = str(row["TELEFONO"]).replace(".0", "")[:10]
                direccion = str(row["DIRECCION"])
                departamento = str(row["DEPARTAMENTO"]).upper()
                municipio = str(row["MUNICIPIO"]).upper()

                wait.until(EC.presence_of_element_located((By.NAME, "identification")))

                driver.find_element(By.NAME, "identification").send_keys(cedula)
                driver.find_element(By.NAME, "identification_confirm").send_keys(cedula)

                driver.find_element(By.NAME, "name").send_keys(nombres)
                driver.find_element(By.NAME, "lastName").send_keys(apellidos)

                driver.find_element(By.NAME, "cellPhone").send_keys(telefono)

                Select(driver.find_element(By.NAME, "department")).select_by_visible_text(departamento)

                wait.until(lambda d: len(Select(d.find_element(By.NAME, "municipalityid")).options) > 1)

                Select(driver.find_element(By.NAME, "municipalityid")).select_by_visible_text(municipio)

                driver.find_element(By.NAME, "neighborhood").send_keys("Centro")
                driver.find_element(By.NAME, "direction").send_keys(direccion)
                driver.find_element(By.NAME, "whoInvited").send_keys(codigo_referido)

                checkbox = wait.until(EC.element_to_be_clickable((By.ID, "terminos")))
                driver.execute_script("arguments[0].click();", checkbox)

                boton = wait.until(EC.element_to_be_clickable((By.ID, "btnRegistrar")))
                driver.execute_script("arguments[0].click();", boton)

                exitos.append(cedula)

                progreso(len(exitos), total)

                log(f"Enviado {cedula}", "success")

                time.sleep(6)

                driver.get("https://dos.alternativaliberalpopular.org/registro.php")

            except Exception as e:

                errores.append(row)

                log(f"Error {cedula} {e}", "error")

        if errores:

            pd.DataFrame(errores).to_excel(ERRORES_XLS, index=False)

        driver.quit()

        log("Proceso finalizado", "success")

    except Exception as e:

        log(f"Error crítico {e}", "error")

    finally:

        proceso_activo = False


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    socketio.run(app, host="0.0.0.0", port=port, allow_unsafe_werkzeug=True)