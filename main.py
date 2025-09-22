import os
import re
import io
import secrets
import openpyxl
from collections import defaultdict
from datetime import datetime, timedelta
from fastapi import FastAPI, UploadFile, File, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse, JSONResponse

# ==========================
# Configuración
# ==========================
CARACTERES_PROHIBIDOS = set("!@#$%&/()=\u00a1\u00a8*[];:_°|\u00ac")
ENCABEZADOS_ESPERADOS = ["Capitulo", "Subcapitulo", "Preguntas"]

DOWNLOADS: dict[str, tuple[bytes, str, str, datetime]] = {}
DOWNLOAD_TTL_SECS = 600  # 10 min
TXT_MEDIA_TYPE = "text/plain"

# ==========================
# Inicializar API
# ==========================
app = FastAPI(title="Validador de Excel Preguntas")

ALLOWED_ORIGINS = [
    "https://www.dipli.ai",
    "https://dipli.ai",
    "https://isagarcivill09.wixsite.com/turop",
    "https://isagarcivill09.wixsite.com/turop/tienda",
    "https://isagarcivill09-wixsite-com.filesusr.com"
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=ALLOWED_ORIGINS,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ==========================
# Funciones de validación
# ==========================
def validar_encabezados(sheet):
    errores = []
    for col, esperado in zip(['A', 'B', 'C'], ENCABEZADOS_ESPERADOS):
        celda = f"{col}1"
        valor = str(sheet[celda].value).strip() if sheet[celda].value else ""
        if valor != esperado:
            errores.append(f"❌ Celda {celda} debería contener '{esperado}', pero tiene '{valor}'")
    return errores

def buscar_preguntas_duplicadas(sheet):
    preguntas = defaultdict(list)
    for row in range(2, sheet.max_row + 1):
        valor = sheet[f"C{row}"].value
        if valor:
            valor = str(valor).strip()
            preguntas[valor].append(row)
    return [f"❌ Pregunta duplicada en filas {v}: '{k}'" for k, v in preguntas.items() if len(v) > 1]

def buscar_caracteres_prohibidos(sheet):
    errores = []
    for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row):
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                for c in cell.value:
                    if c in CARACTERES_PROHIBIDOS:
                        errores.append(
                            f"❌ Celda {cell.coordinate} contiene caracter prohibido '{c}' en: '{cell.value}'"
                        )
                        break
    return errores

def generar_reporte(errores, nombre_archivo: str):
    fecha = datetime.now().strftime("%Y%m%d_%H%M%S")
    nombre = f"reporte_errores_{os.path.splitext(os.path.basename(nombre_archivo))[0]}_{fecha}.txt"
    buffer = io.StringIO()

    if not errores:
        buffer.write("✅ VALIDACIÓN EXITOSA: No se encontraron errores.\n")
    else:
        buffer.write("❌ VALIDACIÓN FALLIDA: Se encontraron errores:\n\n")
        buffer.writelines(f"{err}\n" for err in errores)

    return buffer.getvalue().encode("utf-8"), nombre

# ==========================
# Descargas temporales
# ==========================
def cleanup_downloads():
    now = datetime.utcnow()
    expired = [t for t, (_, _, _, exp) in DOWNLOADS.items() if exp <= now]
    for t in expired:
        DOWNLOADS.pop(t, None)

def register_download(data: bytes, filename: str, media_type: str):
    cleanup_downloads()
    token = secrets.token_urlsafe(16)
    expires_at = datetime.utcnow() + timedelta(seconds=DOWNLOAD_TTL_SECS)
    DOWNLOADS[token] = (data, filename, media_type, expires_at)
    return token

# ==========================
# Endpoints
# ==========================
@app.post("/validar/")
async def validar_excel(request: Request, file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="El archivo debe ser .xlsx")

    try:
        wb = openpyxl.load_workbook(io.BytesIO(await file.read()))
        sheet = wb.active

        errores = []
        errores.extend(validar_encabezados(sheet))
        errores.extend(buscar_preguntas_duplicadas(sheet))
        errores.extend(buscar_caracteres_prohibidos(sheet))

        # Crear reporte
        data, nombre = generar_reporte(errores, file.filename)
        token = register_download(data, nombre, TXT_MEDIA_TYPE)

        base_url = str(request.base_url).rstrip('/')
        download_url = f"{base_url}/download/{token}"

        return {"download_url": download_url, "expires_in_seconds": DOWNLOAD_TTL_SECS, "errores": len(errores)}

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error procesando el archivo: {e}")

@app.get("/download/{token}")
def download_token(token: str):
    cleanup_downloads()
    item = DOWNLOADS.get(token)
    if not item:
        raise HTTPException(status_code=404, detail="Link expirado o inválido")
    data, exp = item
    if exp <= datetime.utcnow():
        DOWNLOADS.pop(token, None)
        raise HTTPException(status_code=410, detail="Link expirado")

    filename = f"reporte_errores_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    headers = {
        "Content-Disposition": f'attachment; filename="{filename}"',
        "Cache-Control": "no-store",
    }
    return StreamingResponse(io.BytesIO(data), media_type="text/plain", headers=headers)

@app.get("/")
def root():
    return {"message": "API de Validación de Excel funcionando", "version": "1.0.0"}

