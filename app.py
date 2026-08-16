import io
import re
import openpyxl
import pandas as pd
import pdfplumber
import streamlit as st

st.set_page_config(
    page_title="Extractor de Actas - Zona Franca", layout="wide"
)


def extraer_datos_acta(pdf_file):
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

    # 3. Transito N°
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

    # 7. Tránsito Fecha Maxima Finalización
    fecha_limite = re.search(
        r"Fecha\s+l[ií]mite\s+para\s+finalizar\s+el\s+r[eé]gimen.*?\b(\d{4}/\d{2}/\d{2})\b",
        texto,
        re.IGNORECASE,
    )

    # 8. Acta de Inventario e Inconsistencias PICIZ
    acta_n = re.search(r"Acta\s+N\.\s*(\d+)", texto, re.IGNORECASE)

    # 9. Fecha acta de inventario e inconsistencias
    fecha_acta = re.search(
        r"FECHA\s+GENERACI[OÓ]N\s+DEL\s+ACTA:\s*(\d{2}/\d{2}/\d{4})",
        texto,
        re.IGNORECASE,
    )

    # 10. No. Planilla de Recepción (DUTA)
    duta = re.search(
        r"DUTA\s+CON\s+NUMERO\s*(\d+)", texto, re.IGNORECASE
    ) or re.search(r"DUTA\s*:\s*(\d+)", texto, re.IGNORECASE)

    # 11. Peso Báscula ZFC (Captura flexible de peso/kilos)
    peso_match = (
        re.search(
            r"(?:Peso\s*(?:B[áa]scula|Entrada|Bruto)?|B[áa]scula)\s*[:\.\-]?\s*([\d\.,]+)",
            texto,
            re.IGNORECASE,
        )
        or re.search(
            r"ACTA\s+DE\s+DESPRECINTAJE[\s\S]*?\b(\d{3,6})\b",
            texto,
            re.IGNORECASE,
        )
        or re.search(r"\b(\d{3,6})\s*(?:KGS?|KILOS?)\b", texto, re.IGNORECASE)
    )
    peso_bascula = peso_match.group(1).strip() if peso_match else "N/A"

    # 12. OBSERVACIONES / INCONSISTENCIAS (Extrae todo el bloque dinámico de texto)
    obs_match = re.search(
        r"Observaciones[:\s]*\n?([\s\S]*?)(?=\n\s*(?:DOCUMENTO\s+FORMULARIO|FIRMAS|FECHA\s+GENERACI[OÓ]N|\Z))",
        texto,
        re.IGNORECASE,
    )

    if obs_match:
        obs_texto = obs_match.group(1).strip()
        # Limpia encabezados internos si existen
        obs_texto = re.sub(
            r"^(?:Descripci[oó]n|Inconsistencias)?\s*",
            "",
            obs_texto,
            flags=re.IGNORECASE,
        )
        observaciones = (
            re.sub(r"\s+", " ", obs_texto).strip() if obs_texto else "N/A"
        )
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
        "Fecha acta de inventario e inconsistencias": (
            fecha_acta.group(1).strip() if fecha_acta else "N/A"
        ),
        "No. Planilla de Recepción (FECHA)": (
            duta.group(1).strip() if duta else "N/A"
        ),
        "Peso Báscula ZFC": peso_bascula,
        "OBSERVACIONES/ INCONSISTENCIAS": observaciones,
    }


# --- INTERFAZ STREAMLIT ---
st.title("Extractor de Datos de Actas de Tránsito")

uploaded_files = st.file_uploader(
    "Carga los archivos PDF de las actas",
    type=["pdf"],
    accept_multiple_files=True,
)

if uploaded_files:
    datos_extraidos = [extraer_datos_acta(file) for file in uploaded_files]
    df = pd.DataFrame(datos_extraidos)

    st.subheader("2. Resultados")
    st.dataframe(df)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Actas")

    st.download_button(
        label="📥 Descargar Excel",
        data=output.getvalue(),
        file_name="reporte_actas_transito.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
