# -*- coding: utf-8 -*-
"""
Procesador Masivo de Actas de Inventario e Inconsistencias (PICIZ)
==================================================================
Aplicación Streamlit para cargar múltiples PDFs de actas PICIZ, 
extraer sus campos clave mediante expresiones regulares y generar 
un reporte consolidado en Excel.
"""

import io
import re
import pandas as pd
import pdfplumber
import streamlit as st

# Configuración inicial de la página
st.set_page_config(
    page_title="Extractor Actas PICIZ",
    page_icon="📋",
    layout="wide",
)

COLUMNAS_ACTAS = [
    "Número de Acta",
    "Datos del Usuario",
    "Documento de Transporte",
    "Fecha de Autorización",
    "Peso de Báscula",
    "Archivo",
]

def extraer_datos_piciz(pdf_bytes):
    """
    Extrae los campos clave de un PDF de Actas de PICIZ (Inventarios e Inconsistencias).
    """
    datos = {}
    with pdfplumber.open(pdf_bytes) as pdf:
        texto_completo = []
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                texto_completo.append(t)
        
        texto = "\n".join(texto_completo)
        
        # 1. Número de Acta
        m = re.search(r'(?:Número\s+de\s+Acta|Acta\s+No\.?|Acta\s+Número)\s*[:\-]?\s*([A-Za-z0-9\-]+)', texto, re.IGNORECASE)
        datos["Número de Acta"] = m.group(1).strip() if m else ""

        # 2. Datos del Usuario / Importador
        m = re.search(r'(?:Usuario|Importador|Declarante|Razón\s+Social)\s*[:\-]?\s*([^\n]+)', texto, re.IGNORECASE)
        datos["Datos del Usuario"] = m.group(1).strip() if m else ""

        # 3. Documento de Transporte
        m = re.search(r'(?:Documento\s+de\s+Transporte|Doc\.?\s+Transporte|Manifiesto)\s*[:\-]?\s*([A-Za-z0-9\-]+)', texto, re.IGNORECASE)
        datos["Documento de Transporte"] = m.group(1).strip() if m else ""

        # 4. Fecha de Autorización
        m = re.search(r'(?:Fecha\s+de\s+Autorización|Fecha\s+Autorización)\s*[:\-]?\s*([0-9]{4}[\-/][0-9]{2}[\-/][0-9]{2})', texto, re.IGNORECASE)
        datos["Fecha de Autorización"] = m.group(1).strip() if m else ""

        # 5. Peso de Báscula
        m = re.search(r'(?:Peso\s+de\s+Báscula|Peso\s+Báscula|Peso\s+Bruto)\s*[:\-]?\s*([\d\.\,]+)', texto, re.IGNORECASE)
        datos["Peso de Báscula"] = m.group(1).strip() if m else ""

    return datos

def main():
    st.title("📋 Procesador Masivo de Actas de PICIZ")
    st.markdown(
        """
        Esta herramienta procesa actas de inventario e inconsistencias de **PICIZ** en formato PDF, 
        extrayendo la trazabilidad aduanera de mercancías[cite: 1] para consolidarlas automáticamente en un archivo Excel.
        """
    )

    st.subheader("1. Cargar documentos PDF")
    uploaded_files = st.file_uploader(
        "Seleccione uno o varios archivos PDF de actas", 
        type=["pdf"], 
        accept_multiple_files=True
    )

    if not uploaded_files:
        st.info("👆 Por favor, cargue archivos PDF para comenzar el procesamiento.")
        st.stop()

    st.success(f"✅ Se han cargado **{len(uploaded_files)}** archivo(s).")

    if st.button("🚀 Extraer Datos y Consolidar", type="primary", use_container_width=True):
        registros = []
        barra_progreso = st.progress(0)
        estado_texto = st.empty()

        for idx, file in enumerate(uploaded_files):
            estado_texto.text(f"Procesando ({idx + 1}/{len(uploaded_files)}): {file.name}")
            try:
                res = extraer_datos_piciz(file)
                res["Archivo"] = file.name
                registros.append(res)
            except Exception as e:
                st.error(f"Error procesando el archivo {file.name}: {e}")
            
            barra_progreso.progress((idx + 1) / len(uploaded_files))

        estado_texto.empty()
        barra_progreso.empty()

        df = pd.DataFrame(registros)
        
        # Ordenar columnas
        cols = [c for c in COLUMNAS_ACTAS if c in df.columns] + [c for c in df.columns if c not in COLUMNAS_ACTAS]
        df = df[cols]

        st.session_state["df_actas"] = df

    if "df_actas" in st.session_state:
        df = st.session_state["df_actas"]

        st.subheader("2. Resultados Extraídos")
        st.dataframe(df, use_container_width=True)

        st.subheader("3. Descargar Consolidado")
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Actas_PICIZ')
        
        st.download_button(
            label="📊 Descargar Archivo Excel Consolidado",
            data=buffer.getvalue(),
            file_name="Actas_PICIZ_Consolidadas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

if __name__ == "__main__":
    main()
