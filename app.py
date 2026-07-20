import os
import io
import zipfile
import pandas as pd
import streamlit as st
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(
    page_title="PDF Batch Renamer Professional",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- ESTILOS PERSONALIZADOS (CSS) ---
st.markdown("""
<style>
    .main { background-color: #F8FAFC; }
    .stApp { max-width: 1100px; margin: 0 auto; }
    .main-header { font-family: 'Inter', sans-serif; font-weight: 700; color: #0F172A; }
    .sub-header { font-family: 'Inter', sans-serif; color: #475569; font-size: 1.05rem; margin-bottom: 20px; }
    .stButton>button { background-color: #2563EB; color: white; border-radius: 8px; font-weight: 600; }
    .stButton>button:hover { background-color: #1D4ED8; }
</style>
""", unsafe_allow_html=True)

# --- GENERADOR DE PDF DE NOVEDADES ---
def generar_pdf_novedades(novedades):
    pdf_buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        pdf_buffer, 
        pagesize=letter, 
        leftMargin=36, 
        rightMargin=36, 
        topMargin=36, 
        bottomMargin=36
    )
    story = []

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'TitleStyle', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=16, textColor=colors.HexColor('#0F172A'), spaceAfter=10
    )
    subtitle_style = ParagraphStyle(
        'SubtitleStyle', parent=styles['Normal'], fontName='Helvetica', fontSize=10, textColor=colors.HexColor('#475569'), spaceAfter=15
    )
    cell_header_style = ParagraphStyle('HeaderStyle', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=9, textColor=colors.white)
    cell_body_style = ParagraphStyle('BodyStyle', parent=styles['Normal'], fontName='Helvetica', fontSize=8.5, textColor=colors.HexColor('#334155'))

    story.append(Paragraph("Reporte de Novedades - Procesamiento de PDFs", title_style))
    story.append(Paragraph("A continuación se detallan los registros que presentaron inconsistencias durante el proceso:", subtitle_style))

    data = [[
        Paragraph("Fila Excel", cell_header_style),
        Paragraph("Buscado (Columna A)", cell_header_style),
        Paragraph("Nuevo Nombre (Columna B)", cell_header_style),
        Paragraph("Detalle de Novedad", cell_header_style)
    ]]

    for item in novedades:
        data.append([
            Paragraph(str(item['fila']), cell_body_style),
            Paragraph(item['origen'], cell_body_style),
            Paragraph(item['destino'], cell_body_style),
            Paragraph(item['motivo'], cell_body_style)
        ])

    table = Table(data, colWidths=[55, 175, 175, 135])
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#0F172A')),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CBD5E1')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F8FAFC')])
    ]))

    story.append(table)
    doc.build(story)
    pdf_buffer.seek(0)
    return pdf_buffer

# --- SIDEBAR / INSTRUCCIONES ---
with st.sidebar:
    st.title("📄 PDF Batch Renamer")
    st.caption("Herramienta Web | Producción")
    st.markdown("---")
    st.markdown("""
    ### 📌 Guía rápida:
    1. Prepara tu **Excel** con la lista (Columna A = Nombre actual, Columna B = Nombre nuevo).
    2. Carga el Excel y selecciona todos los **archivos PDF**.
    3. Presiona **Procesar y Renombrar**.
    4. Descarga el archivo **.ZIP** procesado.
    """)
    st.markdown("---")

# --- INTERFAZ PRINCIPAL ---
st.markdown("<h1 class='main-header'>📂 Renombrador Masivo de PDFs</h1>", unsafe_allow_html=True)
st.markdown("<p class='sub-header'>Procesamiento automatizado de paquetes documentales en formato PDF mediante reglas de Excel.</p>", unsafe_allow_html=True)

col1, col2 = st.columns(2, gap="medium")

with col1:
    st.markdown("### 1. Archivo de Mapeo")
    excel_file = st.file_uploader("Libro de Excel (.xlsx, .xls)", type=["xlsx", "xls"])

with col2:
    st.markdown("### 2. Documentos PDF")
    pdf_files = st.file_uploader("Selecciona los archivos PDF", type=["pdf"], accept_multiple_files=True)

st.markdown("---")

if excel_file and pdf_files:
    if st.button("🚀 Procesar y Renombrar PDFs", use_container_width=True):
        try:
            df = pd.read_excel(excel_file, header=None)
            
            mapeo_nombres = {}
            mapa_filas = {}
            
            for index, row in df.iterrows():
                num_fila = index + 1
                val_a = str(row[0]).strip() if pd.notna(row[0]) else ''
                val_b = str(row[1]).strip() if len(row) > 1 and pd.notna(row[1]) else ''

                if val_a and val_a.lower() != 'nan' and val_b and val_b.lower() != 'nan':
                    orig = val_a if val_a.lower().endswith('.pdf') else val_a + '.pdf'
                    nuev = val_b if val_b.lower().endswith('.pdf') else val_b + '.pdf'
                    
                    mapeo_nombres[orig.lower()] = nuev
                    mapa_filas[orig.lower()] = (num_fila, orig, nuev)

            zip_buffer = io.BytesIO()
            procesados = 0
            novedades = []
            archivos_procesados_keys = set()

            # Creación del paquete ZIP
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                for pdf in pdf_files:
                    nombre_orig = pdf.name.strip()
                    clave = nombre_orig.lower()

                    if clave in mapeo_nombres:
                        nuevo_nombre = mapeo_nombres[clave]
                        zip_file.writestr(nuevo_nombre, pdf.read())
                        procesados += 1
                        archivos_procesados_keys.add(clave)
                    else:
                        novedades.append({
                            'fila': '-',
                            'origen': nombre_orig,
                            'destino': '-',
                            'motivo': 'PDF subido no encontrado en la lista de Excel'
                        })

            # Detectar archivos listados en Excel que no se llegaron a subir
            for key, (num_fila, orig, nuev) in mapa_filas.items():
                if key not in archivos_procesados_keys:
                    novedades.append({
                        'fila': num_fila,
                        'origen': orig,
                        'destino': nuev,
                        'motivo': 'Archivo no cargado en la aplicación'
                    })

            # --- RESULTADOS ---
            st.markdown("## 📊 Resultados del Proceso")
            m_col1, m_col2 = st.columns(2)
            m_col1.metric("PDFs Procesados Exitosamente", procesados)
            m_col2.metric("Novedades Detectadas", len(novedades))

            # Descarga de ZIP
            if procesados > 0:
                zip_buffer.seek(0)
                st.download_button(
                    label="📦 Descargar Paquete de PDFs Renombrados (.ZIP)",
                    data=zip_buffer,
                    file_name="PDFs_Renombrados.zip",
                    mime="application/zip",
                    use_container_width=True
                )

            # Descarga de Reporte PDF si hay novedades
            if novedades:
                st.warning("⚠️ Se registraron novedades durante el proceso.")
                pdf_reporte = generar_pdf_novedades(novedades)
                
                st.download_button(
                    label="📄 Descargar Reporte de Novedades (.PDF)",
                    data=pdf_reporte,
                    file_name="Reporte_Novedades.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
                
                with st.expander("Ver tabla de novedades"):
                    st.dataframe(pd.DataFrame(novedades), use_container_width=True)

        except Exception as e:
            st.error(f"Error crítico al procesar los datos: {e}")