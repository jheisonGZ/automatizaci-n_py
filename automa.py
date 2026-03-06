"""
================================================================================
AUTOMATIZACIÓN DE REGISTRO MASIVO - ALTERNATIVA LIBERAL POPULAR
================================================================================
Descripción : Automatiza el registro de personas en el portal web a partir
              de un archivo Excel seleccionado por el usuario mediante un
              diálogo gráfico. Valida la estructura del archivo antes de
              iniciar el proceso. Incluye segunda vuelta para reintentos.
              Normaliza tildes y mayúsculas en los encabezados del Excel.
Dependencias: selenium, undetected-chromedriver, pandas, openpyxl, tkinter
Autor       : ---
Versión     : 8.0
================================================================================
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import unicodedata
import time
import sys
from selenium.webdriver.chrome.options import Options
import undetected_chromedriver as uc
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog


# ==============================================================================
# FUNCIÓN AUXILIAR — Normalizar texto eliminando tildes y espacios
# ==============================================================================
def normalizar(texto):
    return ''.join(
        c for c in unicodedata.normalize('NFD', str(texto))
        if unicodedata.category(c) != 'Mn'
    ).strip().upper()


# ==============================================================================
# FUNCIÓN AUXILIAR — Seleccionar y validar Excel (con opción de reintentar)
# ==============================================================================
def seleccionar_y_validar_excel(root):
    """
    Muestra el diálogo de selección de archivo y valida las columnas.
    - Si el usuario cancela la selección → termina el proceso.
    - Si el archivo no tiene las columnas correctas → pregunta si quiere
      cargar otro archivo o cancelar. NO queda en un limbo esperando.
    - Retorna (df, ruta_excel) cuando todo es válido.
    """

    COLUMNAS_REQUERIDAS = {
        "CEDULA", "NOMBRES", "APELLIDOS",
        "TELEFONO", "DIRECCION", "DEPARTAMENTO", "MUNICIPIO"
    }

    while True:
        # ── Selección del archivo ──────────────────────────────────────────
        print("📂 Selecciona el archivo Excel con los datos a registrar...")

        ruta_excel = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[
                ("Archivos Excel", "*.xlsx *.xls"),
                ("Todos los archivos", "*.*")
            ]
        )

        # El usuario cerró el diálogo sin elegir archivo
        if not ruta_excel:
            messagebox.showerror(
                "Cancelado",
                "No se seleccionó ningún archivo.\nEl proceso se cancelará."
            )
            print("❌ No se seleccionó archivo. Proceso cancelado.")
            root.destroy()
            sys.exit()

        print(f"✅ Archivo seleccionado: {ruta_excel}")

        # ── Leer y validar columnas ────────────────────────────────────────
        try:
            df = pd.read_excel(ruta_excel)
            df.columns = [normalizar(col) for col in df.columns]

            print(f"\n📋 Columnas detectadas (normalizadas): {list(df.columns)}")

            columnas_presentes  = set(df.columns)
            columnas_faltantes  = COLUMNAS_REQUERIDAS - columnas_presentes

            # ── Archivo con columnas incorrectas ──────────────────────────
            if columnas_faltantes:
                mensaje = (
                    f"El archivo no tiene las columnas necesarias.\n\n"
                    f"❌ Columnas faltantes:\n  {', '.join(sorted(columnas_faltantes))}\n\n"
                    f"📋 Columnas encontradas:\n  {', '.join(df.columns)}\n\n"
                    f"Verifica que el Excel tenga estos encabezados\n"
                    f"(sin importar mayúsculas o tildes):\n"
                    f"  {', '.join(sorted(COLUMNAS_REQUERIDAS))}\n\n"
                    f"¿Deseas seleccionar otro archivo?"
                )
                reintentar = messagebox.askyesno("Columnas inválidas", mensaje)

                if reintentar:
                    # Vuelve al inicio del bucle → abre el diálogo de nuevo
                    print("🔄 El usuario eligió cargar otro archivo...")
                    continue
                else:
                    # El usuario dijo NO → cerrar y terminar
                    print("⛔ Proceso cancelado por el usuario.")
                    root.destroy()
                    sys.exit()

            # ── Archivo vacío ─────────────────────────────────────────────
            if df.empty:
                reintentar = messagebox.askyesno(
                    "Archivo vacío",
                    "El archivo Excel no contiene filas de datos.\n\n"
                    "¿Deseas seleccionar otro archivo?"
                )
                if reintentar:
                    print("🔄 Archivo vacío. El usuario eligió cargar otro...")
                    continue
                else:
                    print("⛔ Proceso cancelado por el usuario.")
                    root.destroy()
                    sys.exit()

            # ── Todo correcto ─────────────────────────────────────────────
            print(f"\n✅ Validación exitosa.")
            print(f"   → Columnas requeridas: presentes")
            print(f"   → Total de registros a procesar: {len(df)}")
            return df, ruta_excel

        except FileNotFoundError:
            messagebox.showerror("Error", "No se encontró el archivo seleccionado.")
            print("❌ Archivo no encontrado.")
            root.destroy()
            sys.exit()

        except Exception as e:
            reintentar = messagebox.askyesno(
                "Error al leer el archivo",
                f"No se pudo leer el Excel:\n{e}\n\n¿Deseas seleccionar otro archivo?"
            )
            if reintentar:
                continue
            else:
                print(f"❌ Error al leer el archivo: {e}")
                root.destroy()
                sys.exit()


# ==============================================================================
# PASO 1 — INICIAR VENTANA TKINTER
# ==============================================================================
root = tk.Tk()
root.withdraw()

# ==============================================================================
# PASO 2 — SELECCIÓN Y VALIDACIÓN DEL EXCEL (con reintentos integrados)
# ==============================================================================
df, ruta_excel = seleccionar_y_validar_excel(root)


# ==============================================================================
# PASO 3 — SOLICITAR CÓDIGO DE REFERIDO
# ==============================================================================
codigo_referido = simpledialog.askstring(
    title="Código de Referido",
    prompt="Ingresa el código de quien invita (whoInvited):"
)

if not codigo_referido or not codigo_referido.strip():
    messagebox.showerror("Error", "Debes ingresar un código de referido para continuar.")
    print("❌ No se ingresó código de referido. Proceso cancelado.")
    root.destroy()
    sys.exit()

codigo_referido = codigo_referido.strip()
print(f"✅ Código de referido ingresado: {codigo_referido}")


# ==============================================================================
# PASO 4 — CONFIRMACIÓN FINAL ANTES DE INICIAR
# ==============================================================================
confirmacion = messagebox.askyesno(
    "Confirmar inicio",
    f"✅ Todo listo para iniciar\n\n"
    f"📄 Archivo     : {ruta_excel.split('/')[-1]}\n"
    f"👥 Registros   : {len(df)}\n"
    f"🔑 Referido    : {codigo_referido}\n\n"
    f"¿Deseas iniciar el proceso de registro?"
)

if not confirmacion:
    print("⛔ Proceso cancelado por el usuario.")
    root.destroy()
    sys.exit()

root.destroy()


# ==============================================================================
# PASO 5 — CONFIGURACIÓN E INICIO DEL NAVEGADOR
# ==============================================================================
options = uc.ChromeOptions()
options.add_argument("--start-maximized")

driver = uc.Chrome(options=options, version_main=145)
driver.set_window_position(0, 0)

try:
    driver.get("https://dos.alternativaliberalpopular.org/registro.php")
    time.sleep(3)

    wait = WebDriverWait(driver, 20)

    errores = []
    exitos  = []

    # ==============================================================================
    # PRIMERA VUELTA — Procesar todos los registros del Excel
    # ==============================================================================
    for index, row in df.iterrows():

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

            print(f"✔  [{index + 1}/{len(df)}] Enviado: {cedula} — {departamento} / {municipio}")
            exitos.append(cedula)

            time.sleep(6)
            driver.get("https://dos.alternativaliberalpopular.org/registro.php")
            time.sleep(3)

        except Exception as e:
            print(f"❌ [{index + 1}/{len(df)}] Error con {cedula}: {e}")
            errores.append(row)

            driver.get("https://dos.alternativaliberalpopular.org/registro.php")
            time.sleep(3)

    # ==============================================================================
    # SEGUNDA VUELTA — Reintento de registros fallidos
    # ==============================================================================
    print("\n🔁 Intentando nuevamente los fallidos...\n")

    errores_definitivos = []

    for i, row in enumerate(errores):

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

            print(f"✅ [{i + 1}/{len(errores)}] Recuperado: {cedula} — {departamento} / {municipio}")

            time.sleep(6)
            driver.get("https://dos.alternativaliberalpopular.org/registro.php")
            time.sleep(3)

        except Exception as e:
            print(f"⛔ [{i + 1}/{len(errores)}] No se pudo subir: {cedula} — {e}")
            errores_definitivos.append(row)

            driver.get("https://dos.alternativaliberalpopular.org/registro.php")
            time.sleep(3)

    # ==============================================================================
    # EXPORTAR REGISTROS FALLIDOS DEFINITIVOS
    # ==============================================================================
    if errores_definitivos:
        df_errores = pd.DataFrame(errores_definitivos)
        df_errores.to_excel("NO_SUBIDOS.xlsx", index=False)
        print("\n📁 Archivo NO_SUBIDOS.xlsx generado con los registros que fallaron.")

    # ==============================================================================
    # RESUMEN FINAL
    # ==============================================================================
    print("\n========== RESUMEN FINAL ==========")
    print(f"📥 Total procesados  : {len(df)}")
    print(f"✔  Exitosos          : {len(exitos)}")
    print(f"🔁 Reintentados      : {len(errores)}")
    print(f"❌ Fallidos definitiv: {len(errores_definitivos)}")
    print(f"🔑 Código referido   : {codigo_referido}")
    print("====================================")
    print("Proceso terminado.")

finally:
    driver.quit()  # ← Siempre se ejecuta, evita el WinError 6