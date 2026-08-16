# -*- coding: utf-8 -*-
"""
Procesador Masivo de Actas de Inventario e Inconsistencias (Zona Franca)
=========================================================================
App Streamlit para cargar PDFs sueltos o ZIPs con múltiples PDFs de
"Actas de Inventario e Inconsistencias", extraer 12 campos estandarizados
mediante RegEx y exportar el resultado a Excel (.xlsx) formateado y CSV.

Ejecutar con:
    streamlit run app.py
"""

import io
import re
import zipfile
from datetime import datetime

import fitz  # PyMuPDF
import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------------------
# Configuración general
# --------------------------------------------------------------------------

st.set_page_config(
    page_title="Procesador de Actas ZF",
    page_icon="📦",
    layout="wide",
)

COLUMNAS = [
    "Usuario",
    "Documento de transporte",
    "Transito N°",
    "FMM N°",
    "FECHA INGRESO ÚLTIMO VEHÍCULO",
    "Fecha de autorización Tránsito",
    "Fecha Maxima Finalización",
    "Acta de Inventario e Inconsistencias PICIZ",
    "Fecha acta de inventario e inconsistencias",
    "No. Planilla de Recepción (FECHA)",
    "Peso Báscula ZFC",
    "OBSERVACIONES/ INCONSISTENCIAS",
    "Archivo",              # columna auxiliar de trazabilidad
    "Campos_no_encontrados",  # columna auxiliar de QA
]

# --------------------------------------------------------------------------
# Utilidades de extracción
# --------------------------------------------------------------------------


def _buscar(patron, texto, flags=re.IGNORECASE, grupo=1):
    """Devuelve el grupo capturado o None si no hay match."""
    m = re.search(patron, texto, flags)
    if m:
        try:
            return m.group(grupo).strip()
        except IndexError:
            return None
    return None


def extraer_campos_acta(texto: str, nombre_archivo: str) -> dict:
    """
    Aplica los patrones RegEx sobre el texto plano de un acta y retorna
    un diccionario con los 12 campos estandarizados + metadatos de QA.
    """
    faltantes = []

    def campo(nombre, patron, texto_fuente=texto, grupo=1, default="", flags=re.IGNORECASE):
        valor = _buscar(patron, texto_fuente, flags=flags, grupo=grupo)
        if not valor:
            faltantes.append(nombre)
            return default
        return " ".join(valor.split())  # normaliza espacios/saltos de línea

    # 1. Usuario (empresa consignataria)
    usuario = campo(
        "Usuario",
        r"consignados al\s+(.+?)\s*\n?\s*y amparado",
    )
    if not usuario:
        # Fallback: a veces el salto de línea cae justo después de "al"
        usuario = campo(
            "Usuario",
            r"consignad[oa]s?\s+al\s*\n?\s*(.+?)\s*\n?\s*y\s+amparad",
        )

    # 2 y 4. Documento de transporte y FMM N° (fila de la tabla de mercancía)
    doc_fmm = re.search(
        r"([A-Z]{2,10}\d{5,10})\s+(\d{6,10})\s+\d+\s+BULTOS", texto, re.IGNORECASE
    )
    if doc_fmm:
        documento_transporte = doc_fmm.group(1).strip()
        fmm_n = doc_fmm.group(2).strip()
    else:
        documento_transporte, fmm_n = "", ""
        faltantes += ["Documento de transporte", "FMM N°"]

    # 3. Tránsito N° / DTA (se toma el número largo tras "Número")
    transito_n = campo(
        "Transito N°",
        r"[Nn][uú]mero\s+(\d{10,20})\s+(?:se procede|de la aduana)",
    )
    if not transito_n:
        # Fallback: cualquier número largo (10-20 dígitos) tras la palabra Número
        transito_n = campo("Transito N°", r"[Nn][uú]mero\s+(\d{10,20})")

    # 5. Fecha de ingreso del último vehículo (fila de desprecintaje)
    fecha_ingreso_vehiculo = campo(
        "FECHA INGRESO ÚLTIMO VEHÍCULO",
        r"\d{2}/\d{2}/\d{4}\s+[A-Z]{2,4}\d{5,8}\s+[A-Z0-9]{4,8}\s+(\d{2}/\d{2}/\d{4})",
    )
    if not fecha_ingreso_vehiculo:
        # Fallback: segunda fecha que aparece en el bloque "ACTA DE DESPRECINTAJE"
        bloque_desprecintaje = _buscar(
            r"ACTA DE DESPRECINTAJE\s*\n(.+?)(?=ACTA DE INVENTARIO|\Z)",
            texto, flags=re.IGNORECASE | re.DOTALL,
        )
        if bloque_desprecintaje:
            fechas = re.findall(r"\d{2}/\d{2}/\d{4}", bloque_desprecintaje)
            if len(fechas) >= 2:
                fecha_ingreso_vehiculo = fechas[1]
            elif fechas:
                fecha_ingreso_vehiculo = fechas[0]
        if not fecha_ingreso_vehiculo:
            faltantes.append("FECHA INGRESO ÚLTIMO VEHÍCULO")

    # 6. Fecha de autorización de tránsito (formato YYYY/MM/DD tal cual el PDF)
    fecha_autorizacion = campo(
        "Fecha de autorización Tránsito",
        r"autorizaci[óo]n de la operaci[óo]n\s*\(YYYY/MM/DD\)\s*(\d{4}/\d{2}/\d{2})",
    )

    # 7. Fecha máxima de finalización del régimen
    fecha_max_finalizacion = campo(
        "Fecha Maxima Finalización",
        r"l[ií]mite para finalizar el r[ée]gimen\s*\(YYYY/MM/DD\)\s*(\d{4}/\d{2}/\d{2})",
    )

    # 8. Número de Acta (PICIZ)
    acta_piciz = campo(
        "Acta de Inventario e Inconsistencias PICIZ",
        r"Acta N\.?\s*(\d+)",
    )

    # 9. Fecha del acta de inventario e inconsistencias (fecha de generación)
    fecha_acta = campo(
        "Fecha acta de inventario e inconsistencias",
        r"FECHA\s*\n?\s*GENERACI[ÓO]N\s+DEL\s+ACTA:?\s*(\d{2}/\d{2}/\d{4})",
    )
    if not fecha_acta:
        # Fallback: fecha impresa en el pie de página del PDF
        fecha_acta = campo(
            "Fecha acta de inventario e inconsistencias",
            r"(\d{1,2}/\d{1,2}/\d{2,4})\s+\d{1,2}:\d{2}\s*[AP]M",
        )

    # 10. No. Planilla de Recepción -> este formato de acta no tiene planilla,
    #     se usa la misma fecha de generación del acta (según se indicó).
    no_planilla_fecha = fecha_acta if fecha_acta else "N/A"

    # 11. Peso báscula ZFC (peso total de la sección TOTALES)
    peso_bascula = campo(
        "Peso Báscula ZFC",
        r"TOTALES:?\s*[\d.,]+\s+([\d.,]+)",
    )
    if peso_bascula:
        # Normaliza "27110.00" -> "27110"
        peso_bascula = peso_bascula.split(".")[0].replace(",", "")

    # 12. Observaciones / Inconsistencias -> se toma el texto completo del bloque
    observaciones_raw = _buscar(
        r"Observaciones\s*\n(.+?)(?=DOCUMENTO\s+FORMULARIO|USUARIO OPERADOR|USUARIO TRANSPORTADOR|\Z)",
        texto,
        flags=re.IGNORECASE | re.DOTALL,
        grupo=1,
    )
    if not observaciones_raw:
        faltantes.append("OBSERVACIONES/ INCONSISTENCIAS")
        observaciones = ""
    else:
        # Limpia líneas de ruido residual de las mini-tablas (Descripción N/A, N/A sueltos, etc.)
        ruido = re.compile(r"^(descripci[oó]n\s*)?n/a$", re.IGNORECASE)
        lineas_utiles = [
            ln.strip() for ln in observaciones_raw.split("\n")
            if ln.strip() and not ruido.match(ln.strip())
        ]
        observaciones = " ".join(" ".join(lineas_utiles).split())

    return {
        "Usuario": usuario,
        "Documento de transporte": documento_transporte,
        "Transito N°": transito_n,
        "FMM N°": fmm_n,
        "FECHA INGRESO ÚLTIMO VEHÍCULO": fecha_ingreso_vehiculo,
        "Fecha de autorización Tránsito": fecha_autorizacion,
        "Fecha Maxima Finalización": fecha_max_finalizacion,
        "Acta de Inventario e Inconsistencias PICIZ": acta_piciz,
        "Fecha acta de inventario e inconsistencias": fecha_acta,
        "No. Planilla de Recepción (FECHA)": no_planilla_fecha,
        "Peso Báscula ZFC": peso_bascula,
        "OBSERVACIONES/ INCONSISTENCIAS": observaciones,
        "Archivo": nombre_archivo,
        "Campos_no_encontrados": ", ".join(faltantes) if faltantes else "",
    }


def extraer_texto_pdf(data: bytes) -> str:
    """Extrae el texto plano de un PDF usando PyMuPDF."""
    with fitz.open(stream=data, filetype="pdf") as doc:
        return "\n".join(page.get_text() for page in doc)


def obtener_pdfs_desde_upload(uploaded_file):
    """
    Recibe un archivo subido (PDF o ZIP) y retorna una lista de tuplas
    (nombre_archivo, bytes_pdf).
    """
    nombre = uploaded_file.name
    contenido = uploaded_file.read()

    if nombre.lower().endswith(".zip"):
        pdfs = []
        with zipfile.ZipFile(io.BytesIO(contenido)) as zf:
            for info in zf.infolist():
                if info.filename.lower().endswith(".pdf") and not info.is_dir():
                    pdfs.append((info.filename.split("/")[-1], zf.read(info.filename)))
        return pdfs
    elif nombre.lower().endswith(".pdf"):
        return [(nombre, contenido)]
    else:
        return []


def procesar_archivos(uploaded_files, progress_callback=None) -> pd.DataFrame:
    """Procesa la lista de archivos subidos y retorna el DataFrame consolidado."""
    filas = []
    tareas = []

    for uf in uploaded_files:
        tareas.extend(obtener_pdfs_desde_upload(uf))

    total = max(len(tareas), 1)
    for i, (nombre_pdf, data) in enumerate(tareas, start=1):
        try:
            texto = extraer_texto_pdf(data)
            fila = extraer_campos_acta(texto, nombre_pdf)
        except Exception as exc:  # noqa: BLE001
            fila = {c: "" for c in COLUMNAS}
            fila["Archivo"] = nombre_pdf
            fila["Campos_no_encontrados"] = f"ERROR AL PROCESAR: {exc}"
        filas.append(fila)
        if progress_callback:
            progress_callback(i / total, nombre_pdf)

    if not filas:
        return pd.DataFrame(columns=COLUMNAS)

    return pd.DataFrame(filas, columns=COLUMNAS)


# --------------------------------------------------------------------------
# Exportación a Excel formateado
# --------------------------------------------------------------------------

HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True)
THIN_BORDER = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)


def generar_excel(df: pd.DataFrame) -> bytes:
    """Genera un archivo Excel (.xlsx) con formato profesional a partir del DataFrame."""
    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Actas")
        ws = writer.sheets["Actas"]

        n_filas = df.shape[0]
        n_cols = df.shape[1]

        # Encabezados: negrita, fondo azul oscuro, texto blanco
        for col_idx in range(1, n_cols + 1):
            celda = ws.cell(row=1, column=col_idx)
            celda.fill = HEADER_FILL
            celda.font = HEADER_FONT
            celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            celda.border = THIN_BORDER

        # Bordes finos en todas las celdas de datos
        for row_idx in range(2, n_filas + 2):
            for col_idx in range(1, n_cols + 1):
                ws.cell(row=row_idx, column=col_idx).border = THIN_BORDER
                ws.cell(row=row_idx, column=col_idx).alignment = Alignment(vertical="top", wrap_text=True)

        # Ajuste automático de ancho de columna según longitud del texto
        for col_idx, columna in enumerate(df.columns, start=1):
            longitudes = [len(str(columna))] + [
                len(str(v)) for v in df[columna].astype(str).tolist()
            ]
            ancho = min(max(longitudes) + 3, 60)  # tope de 60 para no desbordar
            ws.column_dimensions[get_column_letter(col_idx)].width = ancho

        # Auto-filtro sobre el rango de encabezados
        ws.auto_filter.ref = ws.dimensions

        # Inmovilizar la primera fila (encabezados)
        ws.freeze_panes = "A2"

    return buffer.getvalue()


def generar_csv(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False, sep=",", encoding="utf-8-sig").encode("utf-8-sig")


# --------------------------------------------------------------------------
# Interfaz Streamlit
# --------------------------------------------------------------------------

st.title("📦 Procesador Masivo de Actas de Inventario e Inconsistencias")
st.caption("Zona Franca · Extracción automática de 12 campos estandarizados desde PDF")

if "df_resultado" not in st.session_state:
    st.session_state.df_resultado = pd.DataFrame(columns=COLUMNAS)

with st.container(border=True):
    st.subheader("1. Cargar documentos")
    uploaded_files = st.file_uploader(
        "Arrastra aquí archivos PDF individuales o un ZIP con varios PDFs",
        type=["pdf", "zip"],
        accept_multiple_files=True,
        help="Puedes combinar PDFs sueltos y archivos ZIP en la misma carga.",
    )

    procesar = st.button("🚀 Procesar Documentos", type="primary", use_container_width=True)

if procesar:
    if not uploaded_files:
        st.warning("Por favor carga al menos un archivo PDF o ZIP antes de procesar.")
    else:
        progreso = st.progress(0.0, text="Iniciando procesamiento...")

        def _cb(pct, nombre):
            progreso.progress(pct, text=f"Procesando: {nombre}")

        with st.spinner("Extrayendo información de las actas..."):
            df = procesar_archivos(uploaded_files, progress_callback=_cb)

        progreso.empty()
        st.session_state.df_resultado = df
        st.success(f"✅ Procesamiento completado. {len(df)} acta(s) extraída(s).")

# --------------------------------------------------------------------------
# Resultados
# --------------------------------------------------------------------------

df = st.session_state.df_resultado

if not df.empty:
    st.subheader("2. Resultados")

    busqueda = st.text_input("🔍 Buscar en todos los campos", "")

    df_vista = df.copy()
    if busqueda:
        mask = df_vista.apply(
            lambda fila: fila.astype(str).str.contains(busqueda, case=False, na=False).any(),
            axis=1,
        )
        df_vista = df_vista[mask]

    st.dataframe(df_vista, use_container_width=True, height=420)

    faltantes_totales = df[df["Campos_no_encontrados"] != ""]
    if not faltantes_totales.empty:
        with st.expander(f"⚠️ {len(faltantes_totales)} acta(s) con campos no detectados (revisión manual)"):
            st.dataframe(
                faltantes_totales[["Archivo", "Campos_no_encontrados"]],
                use_container_width=True,
            )

    with st.expander("👁️ Vista previa de una acta"):
        archivo_sel = st.selectbox("Selecciona un archivo", df["Archivo"].tolist())
        fila_sel = df[df["Archivo"] == archivo_sel].iloc[0]
        st.json(fila_sel.to_dict())

    st.subheader("3. Descargar resultados")
    df_export = df.drop(columns=["Campos_no_encontrados"])  # columna interna de QA, no se exporta

    col1, col2 = st.columns(2)
    with col1:
        st.download_button(
            "⬇️ Descargar Excel (.xlsx)",
            data=generar_excel(df_export),
            file_name=f"actas_procesadas_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with col2:
        st.download_button(
            "⬇️ Descargar CSV (.csv)",
            data=generar_csv(df_export),
            file_name=f"actas_procesadas_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv",
            use_container_width=True,
        )
else:
    st.info("Carga tus archivos PDF/ZIP y presiona **Procesar Documentos** para ver los resultados aquí.")
