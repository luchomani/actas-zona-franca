import re
import pdfplumber
import pandas as pd


def extraer_datos_acta(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        texto = "\n".join(
            [page.extract_text() for page in pdf.pages if page.extract_text()]
        )

    # 1. Usuario (Consignado a)
    usuario = re.search(r"consignados?\s+al\s+(.+)", texto, re.IGNORECASE)

    # 2 y 4. Documento de Transporte y FMM N° (Formulario)
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

    # 5. Fecha Ingreso Último Vehículo (segunda fecha en la tabla de desprecintaje)
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

    # 11. Peso Báscula ZFC (Peso de entrada de la tabla de desprecintaje)
    peso_bascula = re.search(
        r"ACTA\s+DE\s+DESPRECINTAJE[\s\S]*?\d{2}/\d{2}/\d{4}[\s\S]*?\b(\d{4,6})\b\s*\n\s*El\s+d[ií]a",
        texto,
        re.IGNORECASE,
    )

    # 12. OBSERVACIONES/ INCONSISTENCIAS (Texto completo de la sección)
    obs_match = re.search(
        r"Observaciones\s*\n\s*Descripci[oó]n\s*N/A\s*\n([\s\S]*?)(?=DOCUMENTO\s+FORMULARIO)",
        texto,
        re.IGNORECASE,
    )

    observaciones = (
        obs_match.group(1).replace("\n", " ").strip()
        if obs_match
        else "N/A"
    )

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
        "Peso Báscula ZFC": (
            peso_bascula.group(1).strip() if peso_bascula else "N/A"
        ),
        "OBSERVACIONES/ INCONSISTENCIAS": observaciones,
    }
