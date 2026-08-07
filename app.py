import streamlit as st
import pandas as pd
from docx import Document
import pdfplumber
import io

# Configuración de la página
st.set_page_config(page_title="Conversor de Documentos a Markdown", page_icon="📄")
st.title("📄 Conversor de Documentos a Markdown")
st.write("Sube tus archivos (Excel, CSV, Word o PDF) para convertirlos automáticamente a Markdown.")

# Función para convertir Word a Markdown
def docx_to_markdown(file):
    doc = Document(file)
    md_content = []
    for para in doc.paragraphs:
        if para.style.name.startswith('Heading'):
            level = para.style.name[-1]
            try:
                hashes = '#' * int(level)
                md_content.append(f"{hashes} {para.text}")
            except ValueError:
                md_content.append(para.text)
        else:
            if para.text.strip():
                md_content.append(para.text)
    return "\n\n".join(md_content)

# Función para convertir Excel/CSV a Markdown
def dataframe_to_markdown(file, file_type):
    if file_type == "csv":
        df = pd.read_csv(file)
    else:
        df = pd.read_excel(file)
    return df.to_markdown(index=False)

# Función para convertir PDF a Markdown
def pdf_to_markdown(file):
    md_content = []
    with pdfplumber.open(file) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text()
            if text:
                md_content.append(f"## Página {i+1}\n\n{text}")
    return "\n\n".join(md_content)

# Interfaz de carga de archivos
uploaded_file = st.file_uploader("Elige un archivo", type=["csv", "xlsx", "docx", "pdf"])

if uploaded_file is not None:
    file_extension = uploaded_file.name.split('.')[-1].lower()
    markdown_result = ""
    
    # Procesar según el tipo de archivo
    if file_extension == "docx":
        markdown_result = docx_to_markdown(uploaded_file)
    elif file_extension == "csv":
        markdown_result = dataframe_to_markdown(uploaded_file, "csv")
    elif file_extension == "xlsx":
        markdown_result = dataframe_to_markdown(uploaded_file, "xlsx")
    elif file_extension == "pdf":
        markdown_result = pdf_to_markdown(uploaded_file)
    else:
        st.error("Formato no soportado.")

    if markdown_result:
        st.success("¡Conversión exitosa!")
        
        # Mostrar vista previa
        st.subheader("Vista previa en Markdown")
        st.code(markdown_result, language="markdown")
        
        # Botón de descarga
        st.download_button(
            label="Descargar archivo .md",
            data=markdown_result,
            file_name=f"{uploaded_file.name.rsplit('.', 1)[0]}.md",
            mime="text/markdown"
        )