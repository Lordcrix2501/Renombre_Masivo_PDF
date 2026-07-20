# 📄 PDF Batch Renamer & Reporter (Web App)

Aplicación web profesional construida con **Streamlit**, **Pandas** y **ReportLab** para el renombrado masivo y automatizado de archivos PDF mediante listas de mapeo en **Excel**, con generación automática de **Reportes de Novedades en PDF**.

---

## 🌟 Características

- **Interfaz Web Moderna:** Diseño limpio, responsive y estilizado con CSS personalizado.
- **Procesamiento Masivo:** Carga múltiple de archivos PDF y libros de Excel (`.xlsx` / `.xls`).
- **Empaquetado Automático:** Descarga directa de todos los PDF renombrados en un único archivo comprimido `.ZIP`.
- **Generación de Informes PDF:** Generación automática de un reporte PDF estilizado cuando se detectan inconsistencias (archivos no encontrados, celdas vacías, etc.).
- **Totalmente en Memoria:** No almacena archivos en disco, garantizando privacidad y rapidez.

---

## 🚀 Instalación Local

1. Clonar el repositorio:
   ```bash
   git clone https://github.com/TU_USUARIO/renombrador-pdf-web.git
   cd renombrador-pdf-web
   ```

2. Crear y activar un entorno virtual:
   ```bash
   python -m venv venv
   # En Windows:
   venv\Scripts\activate
   # En macOS/Linux:
   source venv/bin/activate
   ```

3. Instalar dependencias:
   ```bash
   pip install -r requirements.txt
   ```

4. Ejecutar la aplicación:
   ```bash
   streamlit run app.py
   ```

---

## 📁 Estructura del Proyecto

```text
renombrador-pdf-web/
│
├── app.py                  # Código principal de la aplicación Streamlit
├── requirements.txt        # Librerías necesarias para el entorno
├── README.md               # Documentación del repositorio
└── .gitignore              # Archivos excluidos de Git
```

---

## 🛠️ Tecnologías Utilizadas

- [Python 3.10+](https://www.python.org/)
- [Streamlit](https://streamlit.io/)
- [Pandas](https://pandas.pydata.org/)
- [OpenPyXL](https://openpyxl.readthedocs.io/)
- [ReportLab](https://www.reportlab.com/)
