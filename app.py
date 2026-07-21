# ==============================================================
# app.py — Generador automático de Libretas de Notas
# Steve Jobs College (Tacna) · Sin IA, 100% determinista
# ==============================================================
import io
import tempfile
import zipfile
from pathlib import Path

import streamlit as st

import motor

st.set_page_config(page_title="Generador de Libretas · SJC", page_icon="📄", layout="centered")

st.markdown("""
<style>
    :root { --sj-orange: #F26B21; --sj-bg: #0f1115; --sj-panel: #161920; --sj-border: #262b36; }
    .stApp { background-color: var(--sj-bg); }
    .sj-header { display:flex; align-items:center; gap:12px; padding-bottom:18px;
                 border-bottom:1px solid var(--sj-border); margin-bottom:22px; }
    .sj-header .logo { width:44px; height:44px; border-radius:12px;
        background:linear-gradient(135deg,#F26B21,#C9531A); display:flex; align-items:center;
        justify-content:center; font-size:22px; font-weight:700; color:white; }
    .sj-header .title { font-size:21px; font-weight:700; color:#f2f2f2; }
    .sj-header .subtitle { font-size:13px; color:#9aa1ad; }
    div[data-testid="stFileUploaderDropzone"] { background-color: var(--sj-panel); }
    #MainMenu, footer { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="sj-header">
    <div class="logo">SJ</div>
    <div>
        <div class="title">📄 Generador automático de Libretas</div>
        <div class="subtitle">Steve Jobs College · Tacna</div>
    </div>
</div>
""", unsafe_allow_html=True)

st.markdown(
    "Sube el **registro de notas (REG_SECUNDARIA o REG_PRIMARIA)** de uno o varios grados y el "
    "sistema detecta automáticamente el nivel y genera la libreta en PDF de **todos** los alumnos, "
    "lista para descargar."
)

# ---------------------------------------------------------------
# Plantillas: usa las incluidas (primaria y secundaria), o permite reemplazarlas
# ---------------------------------------------------------------
with st.expander("⚙️ Plantillas (opcional — ya incluye una de cada nivel por defecto)"):
    st.caption("El sistema detecta solo si el registro es de Primaria o Secundaria y usa la "
               "plantilla correcta automáticamente. Solo súbelas aquí si quieres reemplazarlas.")
    plantilla_sec_subida = st.file_uploader("Reemplazar PLANTILLA_SECUNDARIA.xlsx", type="xlsx", key="psec")
    plantilla_pri_subida = st.file_uploader("Reemplazar PLANTILLA_PRIMARIA.xlsx", type="xlsx", key="ppri")

plantilla_override = None
if plantilla_sec_subida or plantilla_pri_subida:
    st.caption("⚠️ Con plantillas personalizadas, todos los registros subidos usarán la MISMA "
               "plantilla que subas aquí (no se detecta nivel por archivo en ese caso).")
    plantilla_override = (plantilla_sec_subida or plantilla_pri_subida).getvalue()

tiene_default = motor.RUTA_PLANTILLA_SECUNDARIA.exists() or motor.RUTA_PLANTILLA_PRIMARIA.exists()
if not plantilla_override and not tiene_default:
    st.error("No hay ninguna plantilla disponible. Sube al menos una arriba.")

st.divider()

# ---------------------------------------------------------------
# Registro(s) de notas
# ---------------------------------------------------------------
regs_subidos = st.file_uploader(
    "Registro(s) de notas (REG_SECUNDARIA o REG_PRIMARIA - ... GRADO/AÑO ... BIMESTRE.xlsx)",
    type="xlsx", accept_multiple_files=True,
)

if regs_subidos and (plantilla_override or tiene_default):
    st.divider()
    if st.button("🚀 Generar todas las libretas", type="primary", use_container_width=True):
        with tempfile.TemporaryDirectory() as tmp:
            out_dir = Path(tmp)
            resumen_total = []
            for reg_file in regs_subidos:
                st.markdown(f"#### 📘 {reg_file.name}")
                barra = st.progress(0.0)
                estado_txt = st.empty()

                def _cb(i, total, nombre, _barra=barra, _txt=estado_txt):
                    _barra.progress(i / total)
                    _txt.caption(f"Generando {i}/{total}: {nombre}")

                try:
                    nivel, grado, bimestre, resultados = motor.generar_todas(
                        reg_file.getvalue(), reg_file.name, plantilla_override, out_dir, progreso_callback=_cb
                    )
                except FileNotFoundError as e:
                    st.error(str(e))
                    continue

                barra.progress(1.0)
                ok = sum(1 for r in resultados if r["ok"])
                estado_txt.caption(f"✅ {ok}/{len(resultados)} libretas generadas — {nivel} · {grado} ({bimestre})")

                fallidos = [r for r in resultados if not r["ok"]]
                if fallidos:
                    with st.expander(f"⚠️ {len(fallidos)} con error"):
                        for r in fallidos:
                            st.caption(f"N°{r['n']} {r['nombre']}: {r['error']}")

                resumen_total.append((reg_file.name, nivel, grado, bimestre, resultados))

            # Empaquetar todo en un solo ZIP, organizado por nivel y grado
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for nombre_archivo, nivel, grado, bimestre, resultados in resumen_total:
                    carpeta = f"{nivel} - {grado} - {bimestre}".replace("°", "")
                    for r in resultados:
                        if r["ok"]:
                            zf.write(r["ruta"], arcname=f"{carpeta}/{Path(r['ruta']).name}")
            zip_buffer.seek(0)

            if resumen_total:
                st.divider()
                st.success("¡Listo! Todas las libretas fueron generadas.")
                st.download_button(
                    "⬇️ Descargar todas las libretas (ZIP)",
                    data=zip_buffer, file_name="Libretas_SteveJobsCollege.zip",
                    mime="application/zip", use_container_width=True, type="primary",
                )
elif regs_subidos and not (plantilla_override or tiene_default):
    st.warning("Falta la plantilla — sube al menos una en la sección de arriba.")
else:
    st.info("Sube al menos un archivo de registro para comenzar.")
