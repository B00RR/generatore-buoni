import streamlit as st
import os
import re
import tempfile
from zipfile import ZipFile, ZIP_DEFLATED
from docx import Document
from docx.shared import Inches
from docx.oxml.ns import qn
from PIL import Image
import fitz  # PyMuPDF
import subprocess
from PyPDF2 import PdfMerger

def sovrascrivi_qr(pdf_path, qr_path, position, output):
    pdf = fitz.open(pdf_path)
    page = pdf[0]
    img = Image.open(qr_path)
    rect = fitz.Rect(position[0], position[1], position[0] + img.width, position[1] + img.height)
    page.insert_image(rect, filename=qr_path)
    pdf.save(output)
    pdf.close()
    return output

def convert_docx_to_pdf(docx_path, output_folder):
    cmd = [
        'libreoffice',
        '--headless',
        '--convert-to', 'pdf',
        '--outdir', output_folder,
        docx_path
    ]
    subprocess.run(cmd, check=True)
    pdf_path = os.path.join(output_folder, os.path.basename(docx_path).replace('.docx', '.pdf'))
    return pdf_path

def process_buoni(template_file, zip_file, qr_position=(180, 245)):
    temp_dir = tempfile.mkdtemp()
    qr_folder = os.path.join(temp_dir, 'qr')
    os.makedirs(qr_folder)
    output_pdf_list = []
    with open(os.path.join(temp_dir, 'template.docx'), 'wb') as f:
        f.write(template_file.getbuffer())
    with open(os.path.join(temp_dir, 'qrcodes.zip'), 'wb') as f:
        f.write(zip_file.getbuffer())
    with ZipFile(os.path.join(temp_dir, 'qrcodes.zip'), 'r') as zf:
        svg_files = [f for f in zf.namelist() if f.lower().endswith('.svg')]
        for svg_path in svg_files:
            # Estrai numero buono
            folder = svg_path.split('/')[0]
            match = re.search(r'-([a-zA-Z0-9]+)-', folder)
            num = match.group(1) if match else folder
            svg_data = zf.read(svg_path)
            qr_png_path = os.path.join(qr_folder, f"{num}.png")
            import cairosvg
            cairosvg.svg2png(bytestring=svg_data, write_to=qr_png_path)
            # Sostituisci numero nel Word
            doc = Document(os.path.join(temp_dir, 'template.docx'))
            for para in doc.paragraphs:
                if 'Buono n.' in para.text and 'valido fino al' in para.text:
                    para.text = re.sub(r'Buono n\.\s*\w+', f'Buono n. {num}', para.text)
            word_path = os.path.join(temp_dir, f"Buono_{num}.docx")
            doc.save(word_path)
            # Converti Word in PDF (layout identico incluse ombre)
            pdf_path = convert_docx_to_pdf(word_path, temp_dir)
            # Sovrascrivi QR nel PDF in posizione perfetta (px)
            final_pdf = os.path.join(temp_dir, f"Buono_{num}_finale.pdf")
            sovrascrivi_qr(pdf_path, qr_png_path, qr_position, final_pdf)
            output_pdf_list.append(final_pdf)
    # Unisci tutti i PDF finali
    merger = PdfMerger()
    for pdf in output_pdf_list:
        merger.append(pdf)
    final_pdf_path = os.path.join(temp_dir, "buoni_finali.pdf")
    merger.write(final_pdf_path)
    merger.close()
    return final_pdf_path

# Streamlit UI
st.title("Generatore Buoni Carburante PDF (Ultra-Fedele)")
template_file = st.file_uploader("Carica il template Word", type='docx')
zip_file = st.file_uploader("Carica ZIP con QR (SVG)", type='zip')
if st.button("Genera PDF finale") and template_file and zip_file:
    st.write("Elaborazione in corso...")
    final_pdf = process_buoni(template_file, zip_file, qr_position=(180, 245))  # <-- PERSONALIZZA COORDINATE QR!
    with open(final_pdf, "rb") as f:
        st.download_button("Scarica PDF finale", f.read(), file_name="buoni_finali.pdf", mime="application/pdf")
    st.success("PDF generato con layout e ombre IDENTICHE!")
