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
    .sj-alumno { padding:6px 0; border-bottom:1px solid var(--sj-border); }
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
    "organizadas por nivel y grado, listas para descargar una por una o en bloque."
)

# ---------------------------------------------------------------
# Plantillas: usa las incluidas (primaria y secundaria), o permite reemplazarlas
# Cada nivel tiene su PROPIA plantilla — reemplazar una nunca afecta a la otra.
# ---------------------------------------------------------------
with st.expander("⚙️ Plantillas (opcional — ya incluye una de cada nivel por defecto)"):
    st.caption(
        "El sistema detecta solo si el registro es de Primaria o Secundaria y usa la plantilla "
        "correspondiente automáticamente. Si reemplazas solo una (por ejemplo Secundaria), la otra "
        "sigue usando la plantilla incluida por defecto — nunca se mezclan entre sí."
    )
    col_pri, col_sec = st.columns(2)
    with col_pri:
        plantilla_pri_subida = st.file_uploader("Reemplazar PLANTILLA_PRIMARIA.xlsx", type="xlsx", key="ppri")
        if plantilla_pri_subida:
            ok, msg = motor.validar_plantilla(plantilla_pri_subida.getvalue())
            (st.success if ok else st.error)(msg)
    with col_sec:
        plantilla_sec_subida = st.file_uploader("Reemplazar PLANTILLA_SECUNDARIA.xlsx", type="xlsx", key="psec")
        if plantilla_sec_subida:
            ok, msg = motor.validar_plantilla(plantilla_sec_subida.getvalue())
            (st.success if ok else st.error)(msg)

plantilla_pri_bytes = plantilla_pri_subida.getvalue() if plantilla_pri_subida else None
plantilla_sec_bytes = plantilla_sec_subida.getvalue() if plantilla_sec_subida else None

hay_pri = plantilla_pri_bytes is not None or motor.RUTA_PLANTILLA_PRIMARIA.exists()
hay_sec = plantilla_sec_bytes is not None or motor.RUTA_PLANTILLA_SECUNDARIA.exists()
if not hay_pri and not hay_sec:
    st.error("No hay ninguna plantilla disponible (ni Primaria ni Secundaria). Sube al menos una arriba.")

st.divider()

# ---------------------------------------------------------------
# Registro(s) de notas
# ---------------------------------------------------------------
regs_subidos = st.file_uploader(
    "Registro(s) de notas (REG_SECUNDARIA o REG_PRIMARIA - ... GRADO/AÑO ... BIMESTRE.xlsx)",
    type="xlsx", accept_multiple_files=True,
)

if "lote" not in st.session_state:
    st.session_state.lote = None       # resultados ya generados, listos para mostrar/descargar
if "zip_total" not in st.session_state:
    st.session_state.zip_total = None  # ZIP con TODO, precalculado junto con el lote

puede_generar = bool(regs_subidos) and (hay_pri or hay_sec)

if regs_subidos and not puede_generar:
    st.warning("Falta la plantilla del nivel correspondiente — sube al menos una en la sección de arriba.")
elif not regs_subidos:
    st.info("Sube al menos un archivo de registro para comenzar.")

if puede_generar:
    st.divider()
    if st.button("🚀 Generar todas las libretas", type="primary", use_container_width=True):
        grupos = []
        with tempfile.TemporaryDirectory() as tmp:
            out_dir = Path(tmp)
            for reg_file in regs_subidos:
                st.markdown(f"#### 📘 {reg_file.name}")
                barra = st.progress(0.0)
                estado_txt = st.empty()

                def _cb(i, total, nombre, _barra=barra, _txt=estado_txt):
                    _barra.progress(min(i / total, 1.0))
                    _txt.caption(f"Generando {i}/{total}: {nombre}")

                try:
                    nivel, grado, bimestre, resultados = motor.generar_todas(
                        reg_file.getvalue(), reg_file.name, out_dir,
                        plantilla_primaria_bytes=plantilla_pri_bytes,
                        plantilla_secundaria_bytes=plantilla_sec_bytes,
                        progreso_callback=_cb,
                    )
                except FileNotFoundError as e:
                    st.error(str(e))
                    continue

                barra.progress(1.0)
                ok = sum(1 for r in resultados if r["ok"])
                estado_txt.caption(f"✅ {ok}/{len(resultados)} libretas generadas — {nivel} · {grado} ({bimestre})")

                # Cargamos cada PDF a memoria AHORA, mientras la carpeta temporal
                # todavía existe, para poder ofrecer descargas individuales después
                # sin tener que volver a generar nada (eso evita cargarle CPU de más
                # a Streamlit Cloud cada vez que alguien hace clic en "Descargar").
                alumnos = []
                for r in resultados:
                    pdf_bytes, nombre_pdf = None, None
                    if r["ok"]:
                        p = Path(r["ruta"])
                        pdf_bytes = p.read_bytes()
                        nombre_pdf = p.name
                    alumnos.append({
                        "n": r["n"], "nombre": r["nombre"], "ok": r["ok"], "error": r["error"],
                        "pdf_bytes": pdf_bytes, "filename": nombre_pdf,
                    })

                zip_grado = io.BytesIO()
                with zipfile.ZipFile(zip_grado, "w", zipfile.ZIP_DEFLATED) as zf:
                    for a in alumnos:
                        if a["ok"]:
                            zf.writestr(a["filename"], a["pdf_bytes"])

                grupos.append({
                    "archivo": reg_file.name, "nivel": nivel, "grado": grado, "bimestre": bimestre,
                    "alumnos": alumnos, "zip_bytes": zip_grado.getvalue(),
                })

        # ZIP con absolutamente todo, organizado por carpetas (nivel - grado - bimestre)
        zip_total = io.BytesIO()
        with zipfile.ZipFile(zip_total, "w", zipfile.ZIP_DEFLATED) as zf:
            for grupo in grupos:
                carpeta = f"{grupo['nivel']} - {grupo['grado']} - {grupo['bimestre']}".replace("°", "")
                for a in grupo["alumnos"]:
                    if a["ok"]:
                        zf.writestr(f"{carpeta}/{a['filename']}", a["pdf_bytes"])

        st.session_state.lote = grupos
        st.session_state.zip_total = zip_total.getvalue()

# ---------------------------------------------------------------
# Resultados — se leen de session_state, así que NO se regeneran
# (ni gastan CPU de nuevo) cada vez que alguien hace clic en un botón
# de descarga; solo se recalculan al presionar "Generar" de nuevo.
# ---------------------------------------------------------------
if st.session_state.lote:
    st.divider()
    st.success("¡Listo! Tus libretas están organizadas por nivel y grado abajo.")

    ICONOS = {"PRIMARIA": "🧒", "SECUNDARIA": "🎓"}
    for nivel in ["PRIMARIA", "SECUNDARIA"]:
        grupos_nivel = [g for g in st.session_state.lote if g["nivel"] == nivel]
        if not grupos_nivel:
            continue

        st.markdown(f"## {ICONOS[nivel]} {nivel.capitalize()}")

        for gi, grupo in enumerate(grupos_nivel):
            col1, col2 = st.columns([3, 1])
            with col1:
                ok_count = sum(1 for a in grupo["alumnos"] if a["ok"])
                st.markdown(f"### {grupo['grado']} — {grupo['bimestre']}")
                st.caption(f"{ok_count}/{len(grupo['alumnos'])} libretas generadas · {grupo['archivo']}")
            with col2:
                st.download_button(
                    "⬇️ Descargar grado (ZIP)",
                    data=grupo["zip_bytes"],
                    file_name=f"{nivel} - {grupo['grado']} - {grupo['bimestre']}.zip".replace("°", ""),
                    mime="application/zip",
                    use_container_width=True,
                    key=f"zip_{nivel}_{gi}",
                )

            for a in grupo["alumnos"]:
                c1, c2 = st.columns([4, 1])
                with c1:
                    if a["ok"]:
                        st.markdown(f"<div class='sj-alumno'>N°{a['n']:02d} — {a['nombre']}</div>",
                                    unsafe_allow_html=True)
                    else:
                        st.markdown(
                            f"<div class='sj-alumno'>⚠️ N°{a['n']:02d} — {a['nombre']} "
                            f"— <span style='color:#e0a03c'>{a['error']}</span></div>",
                            unsafe_allow_html=True,
                        )
                with c2:
                    if a["ok"]:
                        st.download_button(
                            "📄 PDF", data=a["pdf_bytes"], file_name=a["filename"],
                            mime="application/pdf", use_container_width=True,
                            key=f"pdf_{nivel}_{gi}_{a['n']}",
                        )
            st.write("")

    if len(st.session_state.lote) > 1:
        st.divider()
        st.download_button(
            "⬇️ Descargar TODO (todos los niveles y grados) — ZIP",
            data=st.session_state.zip_total, file_name="Libretas_SteveJobsCollege.zip",
            mime="application/zip", use_container_width=True, type="primary", key="zip_total_btn",
        )
