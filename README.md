# 📄 Generador Automático de Libretas — Steve Jobs College (Tacna)

Sistema web simple: subes el registro de notas de un grado (`REG_SECUNDARIA...xlsx`)
y genera automáticamente la libreta en PDF de **todos** los alumnos, lista para
descargar en un ZIP. Sin IA, sin API keys, sin límites — solo procesamiento directo
del Excel.

## Instalación local

```bash
pip install -r requirements.txt
```

También necesitas **LibreOffice** instalado (para exportar a PDF):
[libreoffice.org/download](https://www.libreoffice.org/download/main/windows/)

## Correr la app

```bash
streamlit run app.py
```

Se abre en `http://localhost:8501`.

## Cómo usarla

1. (Opcional) Si quieres usar una plantilla distinta a la incluida, súbela en
   "Plantilla de libreta".
2. Sube uno o varios archivos `REG_SECUNDARIA - ... AÑO ... BIMESTRE.xlsx`
   (puedes subir varios grados a la vez).
3. Clic en **"Generar todas las libretas"**.
4. Espera a que termine la barra de progreso (una por cada grado subido).
5. Clic en **"Descargar todas las libretas (ZIP)"** — te da un ZIP con una
   carpeta por grado, cada una con el PDF de cada alumno.

## Desplegarlo público y gratis (Streamlit Community Cloud)

1. Sube esta carpeta a un repositorio de GitHub.
2. Ve a [share.streamlit.io](https://share.streamlit.io) → New app → selecciona
   el repo → Main file: `app.py` → Deploy.
3. Streamlit instala automáticamente `requirements.txt` y `packages.txt`
   (LibreOffice) — no necesitas configurar ningún Secret, esta versión no usa
   ninguna API externa.

## Estructura

```
libretas_batch/
├── app.py              # Interfaz web
├── motor.py              # Lógica: lee Excel, genera PDF
├── assets/PLANTILLA_SECUNDARIA.xlsx
├── requirements.txt
├── packages.txt
└── .streamlit/config.toml
```
