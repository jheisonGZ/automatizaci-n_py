# ALP — Registro Masivo (Web App)

Versión web del automatizador de registros para **Alternativa Liberal Popular**, lista para desplegar en [Render](https://render.com).

## Archivos

```
app.py              → Backend Flask + WebSocket + Selenium
templates/index.html → Interfaz gráfica web
requirements.txt    → Dependencias Python
render.yaml         → Configuración Render
```

## Despliegue en Render

### Opción A — Con render.yaml (automático)
1. Sube todos los archivos a un repositorio GitHub
2. En Render → **New Web Service** → conecta el repo
3. Render detectará el `render.yaml` y configurará todo solo
4. Haz clic en **Deploy**

### Opción B — Manual
1. Sube los archivos a GitHub
2. Render → **New Web Service** → conecta el repo
3. Configura:
   - **Runtime**: Python 3
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `python app.py`
4. **Deploy**

## Variables de entorno
| Variable | Valor     | Descripción          |
|----------|-----------|----------------------|
| PORT     | 10000     | Puerto (Render asigna automáticamente) |

## Uso

1. Abre la URL que asigna Render
2. **Arrastra o selecciona** tu archivo Excel (.xlsx / .xls)
3. El sistema valida automáticamente las columnas requeridas:
   `CEDULA, NOMBRES, APELLIDOS, TELEFONO, DIRECCION, DEPARTAMENTO, MUNICIPIO`
4. Ingresa el **código de referido**
5. Haz clic en **Iniciar registro**
6. Monitorea el progreso en tiempo real en la consola derecha
7. Al terminar, descarga `NO_SUBIDOS.xlsx` si hubo errores

## Columnas requeridas en el Excel
| Columna      | Notas                              |
|--------------|------------------------------------|
| CEDULA       | Sin puntos ni espacios             |
| NOMBRES      | Nombres de la persona              |
| APELLIDOS    | Apellidos de la persona            |
| TELEFONO     | Máximo 10 dígitos                  |
| DIRECCION    | Dirección completa                 |
| DEPARTAMENTO | Nombre exacto como aparece en web  |
| MUNICIPIO    | Nombre exacto como aparece en web  |

> Los encabezados pueden tener tildes o minúsculas, el sistema los normaliza automáticamente.

## Notas técnicas
- Usa `undetected-chromedriver` en modo headless para el servidor
- Los logs se transmiten en tiempo real via **WebSocket (Socket.IO)**
- Los errores definitivos se exportan a `/tmp/NO_SUBIDOS.xlsx` y se pueden descargar desde la interfaz