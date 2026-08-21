# -*- coding: utf-8 -*-
"""
Procesador Masivo de Actas de Inventario e Inconsistencias (PICIZ) - Completo
=============================================================================
Aplicación Streamlit para extraer todos los campos de trazabilidad aduanera 
y logística de las actas PICIZ en Zona Franca y consolidarlas en Excel.
"""

import io
import re
import pandas as pd
import pdfplumber
import streamlit as st

# Configuración inicial de la página
st.set_page_config(
    page_title="Extractor Completo Actas PICIZ",
    page_icon="📋",
    layout="wide",
)

# Definición de las columnas completas solicitadas
COLUMNAS_ACTAS_COMPLETAS = [
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
    "Peso Báscula ZF",
    "BITACORA N°",
    "RAD SOLICITUD BITACORA",
    "OBSERVACIONES/ INCONSISTENCIAS",
    "SI/NO",
    "Archivo",
]

def extraer_datos_piciz_completo(pdf_bytes):
    """
    Extrae todos los campos específicos de las Actas de Inventario e Inconsistencias de PICIZ.
    """
    datos = {}
    with pdfplumber.open(pdf_bytes) as pdf:
        texto_completo = []
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                texto_completo.append(t)
        
        texto = "\n".join(texto_completo)
        
        # 1. Usuario
        m = re.search(r'(?:Usuario|Importador|Declarante|Razón\s+Social)\s*[:\-]?\s*([^\n]+)', texto, re.IGNORECASE)
        datos["Usuario"] = m.group(1).strip() if m else ""

        # 2. Documento de transporte
        m = re.search(r'(?:Documento\s+de\s+transporte|Doc\.?\s+Transporte|Manifiesto)\s*[:\-]?\s*([A-Za-z0-9\-]+)', texto, re.IGNORECASE)
        datos["Documento de transporte"] = m.group(1).strip() if m else ""

        # 3. Transito N°
        m = re.search(r'(?:Transito\s+N°|Tránsito\s+No\.?)\s*[:\-]?\s*([A-Za-z0-9\-]+)', texto, re.IGNORECASE)
        datos["Transito N°"] = m.group(1).strip() if m else ""

        # 4. FMM N°
        m = re.search(r'(?:FMM\s+N°|FMM\s+No\.?)\s*[:\-]?\s*([A-Za-z0-9\-]+)', texto, re.IGNORECASE)
        datos["FMM N°"] = m.group(1).strip() if m else ""

        # 5. FECHA INGRESO ÚLTIMO VEHÍCULO
        m = re.search(r'(?:Fecha\s+ingreso\s+último\s+vehículo|Ingreso\s+Último\s+Vehículo)\s*[:\-]?\s*([0-9]{4}[\-/][0-9]{2}[\-/][0-9]{2}(?:\s+[0-9]{2}:[0-9]{2})?)', texto, re.IGNORECASE)
        datos["FECHA INGRESO ÚLTIMO VEHÍCULO"] = m.group(1).strip() if m else ""

        # 6. Fecha de autorización Tránsito
        m = re.search(r'(?:Fecha\s+de\s+autorización\s+tránsito|Autorización\s+Tránsito)\s*[:\-]?\s*([0-9]{4}[\-/][0-9]{2}[\-/][0-9]{2})', texto, re.IGNORECASE)
        datos["Fecha de autorización Tránsito"] = m.group(1).strip() if m else ""

        # 7. Fecha Maxima Finalización
        m = re.search(r'(?:Fecha\s+maxima\s+finalización|Fecha\s+Máxima\s+Finalización)\s*[:\-]?\s*([0-9]{4}[\-/][0-9]{2}[\-/][0-9]{2})', texto, re.IGNORECASE)
        datos["Fecha Maxima Finalización"] = m.group(1).strip() if m else ""

        # 8. Acta de Inventario e Inconsistencias PICIZ
        m = re.search(r'(?:Acta\s+de\s+Inventario\s+e\s+Inconsistencias\s+PICIZ|Acta\s+No\.?)\s*[:\-]?\s*([A-Za-z0-9\-]+)', texto, re.IGNORECASE)
        datos["Acta de Inventario e Inconsistencias PICIZ"] = m.group(1).strip() if m else ""

        # 9. Fecha acta de inventario e inconsistencias
        m = re.search(r'(?:Fecha\s+acta\s+de\s+inventario\s+e\s+inconsistencias)\s*[:\-]?\s*([0-9]{4}[\-/][0-9]{2}[\-/][0-9]{2})', texto, re.IGNORECASE)
        datos["Fecha acta de inventario e inconsistencias"] = m.group(1).strip() if m else ""

        # 10. No. Planilla de Recepción (FECHA)
        m = re.search(r'(?:No\.?\s+Planilla\s+de\s+Recepción|Planilla\s+Recepción)\s*[:\-]?\s*([A-Za-z0-9\-\/\s]+)', texto, re.IGNORECASE)
        datos["No. Planilla de Recepción (FECHA)"] = m.group(1).strip() if m else ""

        # 11. Peso Báscula ZF
        m = re.search(r'(?:Peso\s+Báscula\s+ZF|Peso\s+Báscula)\s*[:\-]?\s*([\d\.\,]+)', texto, re.IGNORECASE)
        datos["Peso Báscula ZF"] = m.group(1).strip() if m else ""

        # 12. BITACORA N°
        m = re.search(r'(?:BITACORA\s+N°|Bitácora\s+No\.?)\s*[:\-]?\s*([A-Za-z0-9\-]+)', texto, re.IGNORECASE)
        datos["BITACORA N°"] = m.group(1).strip() if m else ""

        # 13. RAD SOLICITUD BITACORA
        m = re.search(r'(?:RAD\s+SOLICITUD\s+BITACORA|Radicado\s+Bitácora)\s*[:\-]?\s*([A-Za-z0-9\-]+)', texto, re.IGNORECASE)
        datos["RAD SOLICITUD BITACORA"] = m.group(1).strip() if m else ""

        # 14. OBSERVACIONES/ INCONSISTENCIAS
        m = re.search(r'(?:Observaciones\s*/\s*Inconsistencias|Observaciones)\s*[:\-]?\s*([^\n]+)', texto, re.IGNORECASE)
        datos["OBSERVACIONES/ INCONSISTENCIAS"] = m.group(1).strip() if m else ""

        # 15. SI/NO (Indicador de Inconsistencias)
        m = re.search(r'\b(SI|NO)\b', texto, re.IGNORECASE)
        datos["SI/NO"] = m.group(1).upper() if m else ""

    return datos

def main():
    st.title("📋 Procesador Masivo de Actas de PICIZ (Inventario e Inconsistencias)")
    st.markdown(
        """
        Extracción completa de los campos logísticos, de tránsito, báscula y bitácora 
        exigidos para las actas de PICIZ en Zona Franca.
        """
    )

    st.subheader("1. Cargar documentos PDF")
    uploaded_files = st.file_uploader(
        "Seleccione uno o varios archivos PDF de actas", 
        type=["pdf"], 
        accept_multiple_files=True
    )

    if not uploaded_files:
        st.info("👆 Por favor, cargue los archivos PDF para comenzar.")
        st.stop()

    st.success(f"✅ Se han cargado **{len(uploaded_files)}** archivo(s).")

    if st.button("🚀 Extraer Datos Completos y Consolidar", type="primary", use_container_width=True):
        registros = []
        barra_progreso = st.progress(0)
        estado_texto = st.empty()

        for idx, file in enumerate(uploaded_files):
            estado_texto.text(f"Procesando ({idx + 1}/{len(uploaded_files)}): {file.name}")
            try:
                res = extraer_datos_piciz_completo(file)
                res["Archivo"] = file.name
                registros.append(res)
            except Exception as e:
                st.error(f"Error procesando {file.name}: {e}")
            
            barra_progreso.progress((idx + 1) / len(uploaded_files))

        estado_texto.empty()
        barra_progreso.empty()

        df = pd.DataFrame(registros)
        
        # Ordenar asegurando todas las columnas solicitadas
        cols = [c for c in COLUMNAS_ACTAS_COMPLETAS if c in df.columns] + [c for c in df.columns if c not in COLUMNAS_ACTAS_COMPLETAS]
        df = df[cols]

        st.session_state["df_actas_completo"] = df

    if "df_actas_completo" in st.session_state:
        df = st.session_state["df_actas_completo"]

        st.subheader("2. Resultados Extraídos")
        st.dataframe(df, use_container_width=True)

        st.subheader("3. Descargar Consolidado en Excel")
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Actas_PICIZ_Completo')
        
        st.download_button(
            label="📊 Descargar Excel Consolidado",
            data=buffer.getvalue(),
            file_name="Actas_PICIZ_Completo.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

if __name__ == "__main__":
    main()
