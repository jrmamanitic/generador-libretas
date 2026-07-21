# ==============================================================
# motor.py — Generador automático de libretas (por lotes)
# Sin IA, sin API keys: 100% determinista, lee Excel -> genera PDF.
# ==============================================================
import io
import re
import subprocess
import unicodedata
from pathlib import Path

import openpyxl

BASE_DIR = Path(__file__).parent
RUTA_PLANTILLA_SECUNDARIA = BASE_DIR / "assets" / "PLANTILLA_SECUNDARIA.xlsx"
RUTA_PLANTILLA_PRIMARIA = BASE_DIR / "assets" / "PLANTILLA_PRIMARIA.xlsx"

_PAT_IDX = re.compile(r"!\$([A-Z]{1,2})\$1(?!:)")  # ya no depende del nombre de la hoja (CONSOLIDADO/PRIMARIA/etc.)
_PAT_GRADO = re.compile(r"([1-5])\s*(?:RO|DO|ER|TO|ERO|°|º)?\s*A[ÑN]O", re.I)
_PAT_BIM = re.compile(r"\b(I{1,3}|IV)\s*BIM", re.I)


def _nfc(s):
    return unicodedata.normalize("NFC", str(s))


def _soffice_cmd():
    import shutil
    exe = shutil.which("soffice") or shutil.which("soffice.exe")
    if exe:
        return exe
    for c in [r"C:\Program Files\LibreOffice\program\soffice.exe",
              r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
              "/Applications/LibreOffice.app/Contents/MacOS/soffice",
              "/usr/bin/soffice", "/usr/local/bin/soffice", "/snap/bin/libreoffice"]:
        if Path(c).exists():
            return c
    raise FileNotFoundError(
        "No se encontró LibreOffice instalado. Descárgalo de https://www.libreoffice.org/download/"
    )


def _buscar_valor_por_etiqueta(ws, etiquetas, filas=range(1, 15), cols=range(1, 10)):
    """Busca una celda cuyo texto contenga alguna de las 'etiquetas' (ej. 'GRADO') y
    devuelve el primer valor no vacío a su derecha, en la misma fila."""
    for r in filas:
        for c in cols:
            v = ws.cell(r, c).value
            if v and isinstance(v, str) and any(e in v.upper() for e in etiquetas):
                for c2 in range(c + 1, c + 8):
                    v2 = ws.cell(r, c2).value
                    if v2 not in (None, ""):
                        return str(v2).strip()
    return None


def detectar_nivel_grado_bimestre(reg_bytes, nombre_archivo=""):
    """Detecta NIVEL (primaria/secundaria), grado y bimestre leyendo el CONTENIDO
    real del archivo (no el nombre, que puede venir mal escrito o incompleto)."""
    wb = openpyxl.load_workbook(io.BytesIO(reg_bytes), data_only=True, read_only=True)
    nivel = "PRIMARIA" if "PERS" in wb.sheetnames else "SECUNDARIA"

    grado, bimestre = None, None
    if "ASISTEN" in wb.sheetnames:
        ws = wb["ASISTEN"]
        grado = _buscar_valor_por_etiqueta(ws, ["GRADO", "AÑO"])
        bimestre = _buscar_valor_por_etiqueta(ws, ["PERIODO", "BIMESTRE"])

    if not grado:
        nfc = _nfc(nombre_archivo)
        mg = _PAT_GRADO.search(nfc)
        sufijo = "GRADO" if nivel == "PRIMARIA" else "AÑO"
        grado = f"{mg.group(1)}° {sufijo}" if mg else "GRADO/AÑO DESCONOCIDO"
    if not bimestre:
        nfc = _nfc(nombre_archivo)
        mb = _PAT_BIM.search(nfc)
        bimestre = (mb.group(1).upper() + " BIM") if mb else "BIM"

    return nivel, grado, bimestre


def detectar_grado_bimestre(nombre_archivo):
    """(Obsoleto, se mantiene por compatibilidad) Extrae grado/bimestre solo del nombre."""
    nfc = _nfc(nombre_archivo)
    mg = _PAT_GRADO.search(nfc)
    mb = _PAT_BIM.search(nfc)
    grado = f"{mg.group(1)}° AÑO" if mg else "GRADO"
    bim = (mb.group(1).upper() + " BIM") if mb else "BIM"
    return grado, bim


def leer_alumnos(reg_bytes):
    """Devuelve {n_lista: nombre} desde la hoja CONSOLIDADO del registro."""
    wb = openpyxl.load_workbook(io.BytesIO(reg_bytes), data_only=True)
    cons = wb["CONSOLIDADO"]
    alumnos = {}
    for r in range(5, cons.max_row + 1):
        n, nom = cons.cell(r, 1).value, cons.cell(r, 2).value
        if n and nom and str(nom) != "0":
            try:
                alumnos[int(float(n))] = str(nom).strip()
            except (ValueError, TypeError):
                pass
    return alumnos


def _leer_tutor(reg_bytes):
    try:
        wb = openpyxl.load_workbook(io.BytesIO(reg_bytes), data_only=True, read_only=True)
        if "ASISTEN" not in wb.sheetnames:
            return None
        ws = wb["ASISTEN"]
        return _buscar_valor_por_etiqueta(ws, ["TUTOR"])
    except Exception:
        return None


def generar_una_libreta(reg_bytes, n_lista, grado, bimestre, plantilla_bytes, out_dir: Path):
    """Genera el PDF de UN alumno. Devuelve (ruta_pdf, nombre)."""
    reg_cons = openpyxl.load_workbook(io.BytesIO(reg_bytes), data_only=True)["CONSOLIDADO"]
    fila = n_lista + 4
    nombre = reg_cons.cell(fila, 2).value
    if not nombre or str(nombre) == "0":
        raise ValueError(f"No existe el alumno N° {n_lista}")
    nombre = str(nombre).strip()

    wb_out = openpyxl.load_workbook(io.BytesIO(plantilla_bytes), data_only=False)
    lib = wb_out["Libreta"]
    lib["C6"] = n_lista
    lib["D5"] = grado
    lib["C8"] = _leer_tutor(reg_bytes) or "-"
    lib["E7"] = bimestre.replace("BIM", "BIMESTRE") if "BIM" in bimestre else bimestre

    for row in lib.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith("="):
                refs = _PAT_IDX.findall(cell.value)
                if not refs:
                    continue
                idx = reg_cons[refs[-1] + "1"].value
                val = reg_cons.cell(fila, int(idx)).value if idx else None
                cell.value = val if val not in (None, "", "0") else "-"

    for nombre_hoja in list(wb_out.sheetnames):
        if nombre_hoja != "Libreta":
            del wb_out[nombre_hoja]

    lib.sheet_properties.pageSetUpPr.fitToPage = True
    lib.page_setup.fitToWidth = 1
    lib.page_setup.fitToHeight = 2  # fuerza que TODO quepa en 2 páginas (ancho 1 x alto 2),
                                     # se autoajusta sin depender de un % fijo ni de la fuente instalada
    lib.page_setup.scale = None

    gsafe = re.sub(r"\W+", "", grado)
    bsafe = re.sub(r"\W+", "", bimestre)
    nsafe = re.sub(r"\W+", "_", nombre)[:40]
    tmp_xlsx = out_dir / f"{n_lista:02d}_{nsafe}.xlsx"
    wb_out.save(tmp_xlsx)

    subprocess.run(
        [_soffice_cmd(), "--headless", "--convert-to", "pdf", "--outdir", str(out_dir), str(tmp_xlsx)],
        check=True, capture_output=True, timeout=120,
    )
    pdf_path = tmp_xlsx.with_suffix(".pdf")
    tmp_xlsx.unlink(missing_ok=True)  # no dejamos el .xlsx intermedio, solo el PDF final
    if not pdf_path.exists():
        raise RuntimeError("LibreOffice no generó el PDF.")
    return str(pdf_path), nombre


def generar_todas(reg_bytes, nombre_archivo_reg, plantilla_bytes_override, out_dir: Path, progreso_callback=None):
    """Genera TODAS las libretas de un archivo REG. Elige automáticamente la plantilla
    de Primaria o Secundaria según el contenido real del archivo (a menos que se
    pase una plantilla específica en plantilla_bytes_override)."""
    nivel, grado, bimestre = detectar_nivel_grado_bimestre(reg_bytes, nombre_archivo_reg)

    if plantilla_bytes_override is not None:
        plantilla_bytes = plantilla_bytes_override
    elif nivel == "PRIMARIA" and RUTA_PLANTILLA_PRIMARIA.exists():
        plantilla_bytes = RUTA_PLANTILLA_PRIMARIA.read_bytes()
    elif RUTA_PLANTILLA_SECUNDARIA.exists():
        plantilla_bytes = RUTA_PLANTILLA_SECUNDARIA.read_bytes()
    else:
        raise FileNotFoundError(f"No hay plantilla disponible para nivel {nivel}.")

    alumnos = leer_alumnos(reg_bytes)
    resultados = []
    total = len(alumnos)
    for i, n_lista in enumerate(sorted(alumnos), start=1):
        nombre_esperado = alumnos[n_lista]
        try:
            ruta, nombre = generar_una_libreta(reg_bytes, n_lista, grado, bimestre, plantilla_bytes, out_dir)
            resultados.append({"n": n_lista, "nombre": nombre, "ok": True, "ruta": ruta, "error": None})
        except Exception as e:
            resultados.append({"n": n_lista, "nombre": nombre_esperado, "ok": False, "ruta": None, "error": str(e)})
        if progreso_callback:
            progreso_callback(i, total, nombre_esperado)
    return nivel, grado, bimestre, resultados
