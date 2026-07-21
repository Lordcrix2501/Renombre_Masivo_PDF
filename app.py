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

# --- NORMALIZADOR DE NOMBRES ---
def normalizar_nombre(texto):
    """
    Limpia espacios al inicio/final, fuerza minúsculas y remueve .pdf 
    para hacer una comparación limpia e infalible.
    """
    if not texto:
        return ""
    txt = str(texto).strip().lower()
    if txt.endswith('.pdf'):
        txt = txt[:-4].strip()
    return txt

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

# --- SIDEBAR ---
with st.sidebar:
    st.title("📄 PDF Batch Renamer")
    st.caption("Herramienta Web | Producción")
    st.markdown("---")
    st.markdown("""
    ### 📌 Modos de Carga:
    * **Opción A (Arrastrar Carpeta):** Arrastra tu carpeta de PDFs directamente al recuadro.
    * **Opción B (Archivo ZIP):** Comprime tu carpeta en un archivo `.zip` y súbelo.
    """)
    st.markdown("---")

# --- INTERFAZ PRINCIPAL ---
st.markdown("<h1 class='main-header'>📂 Renombrador Masivo de PDFs</h1>", unsafe_allow_html=True)
st.markdown("<p class='sub-header'>Procesamiento automatizado de paquetes documentales en formato PDF mediante reglas de Excel.</p>", unsafe_allow_html=True)

col1, col2 = st.columns(2, gap="medium")

with col1:
    st.markdown("### 1. Archivo de Mapeo")
    excel_file = st.file_uploader("Sube el libro de Excel (.xlsx, .xls)", type=["xlsx", "xls"])

with col2:
    st.markdown("### 2. Carpeta o Archivos PDF")
    archivos_subidos = st.file_uploader(
        "Sube tu archivo .ZIP de la carpeta o arrastra la carpeta de PDFs aquí", 
        type=["pdf", "zip"], 
        accept_multiple_files=True
    )

st.markdown("---")

if excel_file and archivos_subidos:
    if st.button("🚀 Procesar y Renombrar PDFs", use_container_width=True):
        try:
            df = pd.read_excel(excel_file, header=None)
            
            mapeo_nombres = {}      # { clave_normalizada: nuevo_nombre_con_pdf }
            mapa_excel_info = {}   # { clave_normalizada: (num_fila, nombre_original_mostrar, nuevo_nombre_mostrar) }
            
            # 1. Cargar datos del Excel limpiando espacios
            for index, row in df.iterrows():
                num_fila = index + 1
                val_a = str(row[0]).strip() if pd.notna(row[0]) else ''
                val_b = str(row[1]).strip() if len(row) > 1 and pd.notna(row[1]) else ''

                if val_a and val_a.lower() != 'nan' and val_b and val_b.lower() != 'nan':
                    clave_norm = normalizar_nombre(val_a)
                    
                    nombre_orig_pdf = val_a if val_a.lower().endswith('.pdf') else val_a + '.pdf'
                    nombre_nuevo_pdf = val_b if val_b.lower().endswith('.pdf') else val_b + '.pdf'
                    
                    mapeo_nombres[clave_norm] = nombre_nuevo_pdf
                    mapa_excel_info[clave_norm] = (num_fila, nombre_orig_pdf, nombre_nuevo_pdf)

            # 2. Extraer archivos ignorando rutas/subcarpetas internas
            archivos_pdf_dict = {}  # { clave_normalizada: (nombre_real_archivo, contenido_bytes) }

            for item in archivos_subidos:
                if item.name.lower().endswith('.zip'):
                    with zipfile.ZipFile(item, 'r') as z:
                        for filename in z.namelist():
                            # Filtrar archivos ocultos de Mac o del sistema
                            if filename.lower().endswith('.pdf') and not filename.startswith('__MACOSX') and not os.path.basename(filename).startswith('.'):
                                nombre_solo = os.path.basename(filename)
                                clave_norm = normalizar_nombre(nombre_solo)
                                archivos_pdf_dict[clave_norm] = (nombre_solo, z.read(filename))
                elif item.name.lower().endswith('.pdf'):
                    nombre_solo = os.path.basename(item.name)
                    clave_norm = normalizar_nombre(nombre_solo)
                    archivos_pdf_dict[clave_norm] = (nombre_solo, item.read())

            zip_buffer = io.BytesIO()
            procesados = 0
            novedades = []
            claves_procesadas = set()

            # 3. Cruzar y renombrar
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                for clave_norm, (nombre_real, contenido_bytes) in archivos_pdf_dict.items():
                    if clave_norm in mapeo_nombres:
                        nuevo_nombre = mapeo_nombres[clave_norm]
                        zip_file.writestr(nuevo_nombre, contenido_bytes)
                        procesados += 1
                        claves_procesadas.add(clave_norm)
                    else:
                        novedades.append({
                            'fila': '-',
                            'origen': nombre_real,
                            'destino': '-',
                            'motivo': 'PDF subido pero no está listado en la Columna A del Excel'
                        })

            # 4. Detectar cuáles del Excel faltaron por subir
            for clave_norm, (num_fila, orig_pdf, nuev_pdf) in mapa_excel_info.items():
                if clave_norm not in claves_procesadas:
                    novedades.append({
                        'fila': num_fila,
                        'origen': orig_pdf,
                        'destino': nuev_pdf,
                        'motivo': 'Nombre en Excel no encontrado en la carpeta subida'
                    })

            # --- RESULTADOS ---
            st.markdown("## 📊 Resultados del Proceso")
            m_col1, m_col2 = st.columns(2)
            m_col1.metric("PDFs Procesados Exitosamente", procesados)
            m_col2.metric("Novedades Detectadas", len(novedades))

            if procesados > 0:
                zip_buffer.seek(0)
                st.download_button(
                    label="📦 Descargar Paquete de PDFs Renombrados (.ZIP)",
                    data=zip_buffer,
                    file_name="PDFs_Renombrados.zip",
                    mime="application/zip",
                    use_container_width=True
                )

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
                
                with st.expander("Ver tabla detallada de novedades"):
                    st.dataframe(pd.DataFrame(novedades), use_container_width=True)

        except Exception as e:
            st.error(f"Error crítico al procesar los datos: {e}")
