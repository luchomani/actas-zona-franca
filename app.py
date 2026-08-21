# -*- coding: utf-8 -*-
"""
Procesador Masivo de Actas de Inventario e Inconsistencias
================================================================================
Panel Corporativo Optimizado — Zona Franca Santander / Cúcuta
"""

import io
import re
import os
import base64
from datetime import datetime

import pandas as pd
import pdfplumber
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------------------
# Configuración general y Estilos Corporativos con Fondo Personalizado
# --------------------------------------------------------------------------

st.set_page_config(
    page_title="Extractor de Actas | Zona Franca",
    page_icon="🏢",
    layout="wide",
)

fondo_css = ""
if os.path.exists("Fondo ZFC.png"):
    with open("Fondo ZFC.png", "rb") as f:
        fondo_bytes = f.read()
    fondo_base64 = base64.b64encode(fondo_bytes).decode()
    fondo_css = f"""
    .stApp {{
        background-image: linear-gradient(rgba(244, 247, 246, 0.9), rgba(244, 247, 246, 0.9)), url("data:image/png;base64,{fondo_base64}");
        background-size: cover;
        background-position: center;
        background-repeat: no-repeat;
    }}
    """
else:
    fondo_css = """
    .stApp {
        background-color: #F4F7F6;
    }
    """

st.markdown(f"""
<style>
    :root {{
        --zf-green-dark: #1B4D3E;
        --zf-green-medium: #2C6B56;
        --zf-olive: #8A9A28;
        --zf-card-bg: #FFFFFF;
        --zf-text-main: #2C3E50;
    }}

    {fondo_css}

    h1, h2, h3 {{
        color: var(--zf-green-dark) !important;
        font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
    }}

    div[data-testid="stVerticalBlock"] > div[style*="border"] {{
        background-color: var(--zf-card-bg);
        border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.05);
        border: 1px solid #E1E8E5 !important;
        padding: 20px;
    }}

    .stButton>button {{
        background-color: var(--zf-green-dark);
        color: white;
        border-radius: 6px;
        border: none;
        font-weight: 600;
        padding: 0.5rem 1rem;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        transition: all 0.3s ease;
    }}

    .stButton>button:hover {{
        background-color: var(--zf-green-medium);
        color: white;
        border: none;
        box-shadow: 0 4px 8px rgba(0,0,0,0.15);
    }}

    div[data-testid="stMetricValue"] {{
        color: var(--zf-green-dark);
        font-weight: 700;
    }}
    div[data-testid="stMetricLabel"] {{
        color: #556B2F;
        font-weight: 600;
    }}

    .stDataFrame {{
        border-radius: 8px;
        overflow: hidden;
        border: 1px solid #E1E8E5;
    }}
</style>
""", unsafe_allow_html=True)

# --------------------------------------------------------------------------
# Función de Extracción de Datos (pdfplumber)
# --------------------------------------------------------------------------

def extraer_datos_acta(pdf_file, nombre_archivo):
    with pdfplumber.open(pdf_file) as pdf:
        texto = "\n".join(
            [page.extract_text() for page in pdf.pages if page.extract_text()]
        )

    # 1. Usuario (Consignado a)
    usuario = re.search(r"consignados?\s+al\s+(.+)", texto, re.IGNORECASE)

    # 2 y 4. Documento de Transporte y FMM N°
    doc_form = re.search(
        r"DOCUMENTO\s+FORMULARIO\s+MERCANC[ÍI]A[\s\S]*?\n\s*([A-Z0-9\.\-_]+)\s+(\d+)",
        texto,
        re.IGNORECASE,
    )

    # 3. Tránsito N°
    transito = re.search(
        r"DECLARACION\s+DE\s+TRANSITO\s+ADUANERO\s*\n?\s*N[úu]mero\s*(\d+)",
        texto,
        re.IGNORECASE,
    )

    # 5. Fecha Ingreso Último Vehículo
    fecha_ingreso = re.search(
        r"ACTA\s+DE\s+DESPRECINTAJE[\s\S]*?\d{2}/\d{2}/\d{4}\s+[\w\d]+\s+[\w\d\.]+\s+(\d{2}/\d{2}/\d{4})",
        texto,
        re.IGNORECASE,
    )

    # 6. Fecha de autorización
    fecha_auto = re.search(
        r"Fecha\s+de\s+la\s+autorizaci[oó]n\s+de\s+la\s+operaci[oó]n.*?\b(\d{4}/\d{2}/\d{2})\b",
        texto,
        re.IGNORECASE,
    )

    # 7. Tránsito Fecha Máxima Finalización
    fecha_limite = re.search(
        r"Fecha\s+l[ií]mite\s+para\s+finalizar\s+el\s+r[eé]gimen.*?\b(\d{4}/\d{2}/\d{2})\b",
        texto,
        re.IGNORECASE,
    )

    # 8. Acta de Inventario e Inconsistencias PICIZ
    acta_n = re.search(r"Acta\s+N\.\s*(\d+)", texto, re.IGNORECASE)

    # 9 y 10. Fecha acta de inventario y Fecha Planilla de Recepción (misma fecha)
    fecha_acta_match = re.search(
        r"FECHA\s+GENERACI[OÓ]N\s+DEL\s+ACTA:\s*(\d{2}/\d{2}/\d{4})",
        texto,
        re.IGNORECASE,
    )
    fecha_acta = (
        fecha_acta_match.group(1).strip() if fecha_acta_match else "N/A"
    )

    # 11. Peso Báscula ZFC
    peso_match = re.search(
        r"TOTALES\s*:\s*[\d\.,]+\s+([\d\.,]+)", texto, re.IGNORECASE
    ) or re.search(
        r"DOCUMENTO\s+FORMULARIO[\s\S]*?\n[\s\S]*?\b(\d+(?:\.\d+)?)\s*\n\s*TOTALES",
        texto,
        re.IGNORECASE,
    )
    peso_bascula = peso_match.group(1).strip() if peso_match else "N/A"

    # 12. OBSERVACIONES / INCONSISTENCIAS
    obs_match = re.search(
        r"Observaciones[\s\S]*?\n([\s\S]*?)(?=\n\s*(?:DOCUMENTO\s+FORMULARIO|TOTALES|USUARIO\s+OPERADOR|\Z))",
        texto,
        re.IGNORECASE,
    )

    if obs_match:
        lineas = obs_match.group(1).split("\n")
        lineas_limpias = []
        for line in lineas:
            line_str = line.strip()
            if re.match(
                r"^(Descripción\s*N/A|Bultos|Estado|Términos|Otra)\b",
                line_str,
                re.IGNORECASE,
            ):
                continue
            if line_str:
                lineas_limpias.append(line_str)
        observaciones = " ".join(lineas_limpias) if lineas_limpias else "N/A"
    else:
        observaciones = "N/A"

    return {
        "Usuario": usuario.group(1).strip() if usuario else "N/A",
        "Documento de transporte": (
            doc_form.group(1).strip() if doc_form else "N/A"
        ),
        "Transito N°": transito.group(1).strip() if transito else "N/A",
        "FMM N°": doc_form.group(2).strip() if doc_form else "N/A",
        "FECHA INGRESO ÚLTIMO VEHÍCULO": (
            fecha_ingreso.group(1).strip() if fecha_ingreso else "N/A"
        ),
        "Fecha de autorización": (
            fecha_auto.group(1).strip() if fecha_auto else "N/A"
        ),
        "Tránsito Fecha Maxima Finalización": (
            fecha_limite.group(1).strip() if fecha_limite else "N/A"
        ),
        "Acta de Inventario e Inconsistencias PICIZ": (
            acta_n.group(1).strip() if acta_n else "N/A"
        ),
        "Fecha acta de inventario e inconsistencias": fecha_acta,
        "No. Planilla de Recepción (FECHA)": fecha_acta,
        "Peso Báscula ZFC": peso_bascula,
        "OBSERVACIONES/ INCONSISTENCIAS": observaciones,
        "Archivo": nombre_archivo,
    }

# --------------------------------------------------------------------------
# Exportación Profesional a Excel
# --------------------------------------------------------------------------

HEADER_FILL = PatternFill(start_color="1B4D3E", end_color="1B4D3E", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True)
THIN_BORDER = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)

def generar_excel(df: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Actas")
        ws = writer.sheets["Actas"]

        n_filas = df.shape[0]
        n_cols = df.shape[1]

        for col_idx in range(1, n_cols + 1):
            celda = ws.cell(row=1, column=col_idx)
            celda.fill = HEADER_FILL
            celda.font = HEADER_FONT
            celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            celda.border = THIN_BORDER

        for row_idx in range(2, n_filas + 2):
            for col_idx in range(1, n_cols + 1):
                celda = ws.cell(row=row_idx, column=col_idx)
                celda.border = THIN_BORDER
                celda.alignment = Alignment(vertical="top", wrap_text=True)

        for col_idx, columna in enumerate(df.columns, start=1):
            longitudes = [len(str(columna))] + [len(str(v)) for v in df[columna].astype(str).tolist()]
            ancho = min(max(longitudes) + 3, 45)
            ws.column_dimensions[get_column_letter(col_idx)].width = ancho

        ws.auto_filter.ref = f"A1:{get_column_letter(n_cols)}{n_filas}"
        ws.freeze_panes = "A2"

    return buffer.getvalue()

# --------------------------------------------------------------------------
# Interfaz de Usuario Corporativa (Streamlit)
# --------------------------------------------------------------------------

col_logo, col_titulo = st.columns([1.3, 3.7])

with col_logo:
    logo_encontrado = False
    for filename in ["LOGO ZFS-ZFC.jpeg", "logo.jpeg", "logo.jpg", "logo.png"]:
        if os.path.exists(filename):
            st.image(filename, width=240)
            logo_encontrado = True
            break
    
    if not logo_encontrado:
        st.info("💡 Sube tu imagen de logo al repositorio como `LOGO ZFS-ZFC.jpeg`.")

with col_titulo:
    st.markdown("### Módulo de Gestión de Inventarios e Inconsistencias")
    st.markdown("**Extractor Automático de Actas PICIZ — Tránsito Aduanero**")
    st.caption("Zona Franca de Cúcuta | Operada por Zona Franca Santander")

st.divider()

# Inicialización de estados para limpiar y reiniciar panel
if "df_resultado_actas" not in st.session_state:
    st.session_state.df_resultado_actas = pd.DataFrame()

if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0

with st.container():
    st.subheader("1. Carga de Documentación de Actas")
    
    uploaded_files = st.file_uploader(
        "Carga los archivos PDF de las actas de inventario (individuales o múltiples)",
        type=["pdf"],
        accept_multiple_files=True,
        key=f"actas_uploader_{st.session_state.uploader_key}",
    )

    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        procesar = st.button("🚀 Procesar Actas", type="primary", use_container_width=True)
    with col_btn2:
        limpiar = st.button("🧹 Limpiar y Reiniciar Panel", use_container_width=True)

if limpiar:
    st.session_state.df_resultado_actas = pd.DataFrame()
    st.session_state.uploader_key += 1  # Incrementa la llave para vaciar el uploader físicamente
    st.rerun()

if procesar:
    if not uploaded_files:
        st.warning("Por favor carga al menos un archivo PDF antes de ejecutar el procesamiento.")
    else:
        with st.spinner("Extrayendo campos clave de las actas de tránsito...[cite: 7]"):
            datos_extraidos = [extraer_datos_acta(file, file.name) for file in uploaded_files]
            st.session_state.df_resultado_actas = pd.DataFrame(datos_extraidos)
        st.success(f"✅ Extracción completada con éxito para {len(uploaded_files)} acta(s).")

# --------------------------------------------------------------------------
# Visualización de Resultados y Analítica
# --------------------------------------------------------------------------

df = st.session_state.df_resultado_actas

if not df.empty:
    st.markdown("---")
    st.subheader("2. Resultados Consolidados")

    busqueda = st.text_input("🔍 Búsqueda rápida (Usuario, Tránsito N°, Acta PICIZ, Documento...)", "")

    df_vista = df.copy()
    if busqueda:
        mask = df_vista.apply(lambda fila: fila.astype(str).str.contains(busqueda, case=False, na=False).any(), axis=1)
        df_vista = df_vista[mask]

    st.dataframe(df_vista, use_container_width=True, height=430)

    st.markdown("---")
    st.subheader("3. Exportación de Datos")
    
    col1, col2 = st.columns(2)
    with col1:
        st.download_button(
            "⬇️ Descargar Reporte en Excel (.xlsx)",
            data=generar_excel(df),
            file_name=f"actas_zona_franca_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with col2:
        csv_data = df.to_csv(index=False, sep=",", encoding="utf-8-sig").encode("utf-8-sig")
        st.download_button(
            "⬇️ Descargar Reporte en CSV (.csv)",
            data=csv_data,
            file_name=f"actas_zona_franca_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv",
            use_container_width=True,
        )
else:
    st.info("Carga tus archivos PDF de actas para habilitar el procesamiento, la tabla interactiva y las descargas.")
