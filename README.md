# Procesador Masivo de Actas de Inventario e Inconsistencias (Zona Franca)

Aplicación web interna para extraer automáticamente 12 campos estandarizados
desde "Actas de Inventario e Inconsistencias" en PDF (individuales o dentro
de un ZIP) y exportar el resultado a Excel/CSV.

## Arquitectura

Se eligió **Streamlit puro** (sin separar frontend/backend en FastAPI + HTML)
por tres razones prácticas para este caso de uso interno:

- **Velocidad de implementación**: un solo archivo `app.py` cubre carga de
  archivos, procesamiento, tabla interactiva y descargas — sin necesidad de
  levantar dos servidores ni escribir JS.
- **Despliegue gratuito inmediato**: Streamlit Community Cloud despliega
  directamente desde un repo de GitHub, sin configurar CORS ni endpoints REST.
- **Uso interno, no de alto tráfico**: no se requiere una API pública
  reutilizable por terceros, por lo que el overhead de FastAPI no se justifica.

> Si en el futuro necesitas que otros sistemas (ERP, WMS) consuman esta
> extracción vía API, la función `extraer_campos_acta()` está desacoplada de
> la interfaz y se puede envolver fácilmente en un endpoint FastAPI sin
> reescribir la lógica.

### Flujo de datos

```
Usuario carga PDFs/ZIP
        │
        ▼
obtener_pdfs_desde_upload()  → descomprime ZIP en memoria (zipfile + io.BytesIO)
        │
        ▼
extraer_texto_pdf()          → PyMuPDF (fitz) obtiene el texto plano
        │
        ▼
extraer_campos_acta()        → RegEx sobre el texto → dict de 12 campos + QA
        │
        ▼
pandas.DataFrame             → tabla consolidada, filtrable en la UI
        │
        ├──► generar_excel()  → openpyxl (estilos, bordes, autofiltro, freeze)
        └──► generar_csv()    → pandas.to_csv()
```

### Campos extraídos

| # | Campo | Fuente en el PDF |
|---|-------|-------------------|
| # | Campo | Fuente en el PDF |
|---|-------|-------------------|
| 1 | Usuario | "consignados al ..." |
| 2 | Documento de transporte | Tabla MERCANCÍA |
| 3 | Tránsito N° (DTA) | "Número ... se procede" |
| 4 | FMM N° | Tabla MERCANCÍA |
| 5 | Fecha ingreso último vehículo | Fila de desprecintaje |
| 6 | Fecha autorización tránsito | "(YYYY/MM/DD)" |
| 7 | Fecha máxima finalización | "(YYYY/MM/DD)" |
| 8 | Acta PICIZ | "Acta N." |
| 9 | Fecha acta de inventario e inconsistencias | "FECHA GENERACIÓN DEL ACTA" |
| 10 | No. Planilla de Recepción (FECHA) | No existe planilla en este formato → se usa la misma fecha del campo 9 |
| 11 | Peso báscula ZFC | "TOTALES:" |
| 12 | Observaciones/Inconsistencias | Bloque "Observaciones" (texto completo) |

Cada fila procesada incluye además una columna interna `Campos_no_encontrados`
(no se exporta) que se muestra en la app para que el usuario revise
manualmente cualquier acta cuyo formato no haya podido leerse por completo.

## 1. Instalación local

Requiere Python 3.10+.

```bash
# 1. Crear entorno virtual (recomendado)
python -m venv venv
source venv/bin/activate        # En Windows: venv\Scripts\activate

# 2. Instalar dependencias
pip install -r requirements.txt
```

## 2. Ejecutar en local

```bash
streamlit run app.py
```

Esto abre automáticamente `http://localhost:8501` en tu navegador.

## 3. Despliegue gratuito en la nube

### Opción A — Streamlit Community Cloud (recomendada, la más simple)

1. Sube este proyecto (`app.py`, `requirements.txt`) a un repositorio de
   GitHub (puede ser privado).
2. Entra a [share.streamlit.io](https://share.streamlit.io) con tu cuenta
   de GitHub.
3. Clic en **"New app"** → selecciona el repo, la rama y `app.py` como
   archivo principal.
4. Deploy. En 1-2 minutos tendrás una URL pública tipo
   `https://tu-app.streamlit.app`.

Límite del plan gratuito: la app "duerme" tras un período de inactividad y
despierta automáticamente con la siguiente visita (puede tardar ~30s).

### Opción B — Render (Web Service gratuito)

1. Sube el proyecto a GitHub.
2. En [render.com](https://render.com) → **New → Web Service** → conecta el
   repo.
3. Configuración:
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `streamlit run app.py --server.port $PORT --server.address 0.0.0.0`
4. Selecciona el plan **Free** y despliega.

> Nota: el plan gratuito de Render también "duerme" el servicio tras
> inactividad prolongada.

### Recomendación para uso interno con datos sensibles

Si las actas contienen información comercial confidencial (como en este
caso, movimientos de zona franca), usa **repositorio privado** en GitHub y,
si tu organización lo permite, considera activar autenticación básica en
Streamlit Community Cloud (disponible en el plan de equipos) o desplegar en
un servidor propio/VPN en lugar de la capa gratuita pública.

## Estructura del proyecto

```
acta_zf_app/
├── app.py              # Aplicación completa (UI + extracción + exportación)
├── requirements.txt    # Dependencias
└── README.md           # Este archivo
```
