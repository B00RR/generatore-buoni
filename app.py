import fitz  # PyMuPDF
import streamlit as st
import os
import re
import tempfile
from zipfile import ZipFile
from docx import Document
import subprocess
from PIL import Image
from PyPDF2 import PdfMerger
import cairosvg

def sovrascrivi_qr(pdf_path, qr_path, top_left_mm, size_mm, output):
    # Conversione mm → pt (1 mm ≈ 2.83465 pt)
    def mm2pt(mm):
        return mm * 2.83465

    x0 = mm2pt(top_left_mm[0])
    y0 = mm2pt(top_left_mm[1])
    x1 = x0 + mm2pt(size_mm[0])
    y1 = y0 + mm2pt(size_mm[1])
    rect = fitz.Rect(x0, y0, x1, y1)
    pdf = fitz.open(pdf_path)
    page = pdf[0]
    page.insert_image(rect, filename=qr_path, overlay=True)
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

def process_buoni(template_file, zip_file, qr_mm_pos=(70, 90), qr_mm_size=(32, 32)):
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
            folder = svg_path.split('/')[0]
            match = re.search(r'-([a-zA-Z0-9]+)-', folder)
            num = match.group(1) if match else folder
            svg_data = zf.read(svg_path)
            qr_png_path = os.path.join(qr_folder, f"{num}.png")
            cairosvg.svg2png(bytestring=svg_data, write_to=qr_png_path, output_width=400, output_height=400)
            doc = Document(os.path.join(temp_dir, 'template.docx'))
            for para in doc.paragraphs:
                if 'Buono n.' in para.text and 'valido fino al' in para.text:
                    para.text = re.sub(r'Buono n\.\s*\w+', f'Buono n. {num}', para.text)
            word_path = os.path.join(temp_dir, f"Buono_{num}.docx")
            doc.save(word_path)
            pdf_path = convert_docx_to_pdf(word_path, temp_dir)
            final_pdf = os.path.join(temp_dir, f"Buono_{num}_finale.pdf")
            sovrascrivi_qr(pdf_path, qr_png_path, qr_mm_pos, qr_mm_size, final_pdf)
            output_pdf_list.append(final_pdf)
    merger = PdfMerger()
    for pdf in output_pdf_list:
        merger.append(pdf)
    final_pdf_path = os.path.join(temp_dir, "buoni_finali.pdf")
    merger.write(final_pdf_path)
    merger.close()
    return final_pdf_path

# --- Streamlit UI

st.title("Generatore Buoni Carburante PDF")
template_file = st.file_uploader("Carica il template Word", type='docx')
zip_file = st.file_uploader("Carica ZIP con QR (SVG)", type='zip')

# Permetti di impostare dimensioni/posizione (modifica default se servono)
qr_x = st.number_input("QR - Margine sinistro (mm)", value=70)
qr_y = st.number_input("QR - Margine alto (mm)", value=90)
qr_w = st.number_input("QR - Larghezza (mm)", value=32)
qr_h = st.number_input("QR - Altezza (mm)", value=32)

if st.button("Genera PDF finale") and template_file and zip_file:
    st.write("Elaborazione in corso...")
    final_pdf = process_buoni(template_file, zip_file, qr_mm_pos=(qr_x, qr_y), qr_mm_size=(qr_w, qr_h))
    with open(final_pdf, "rb") as f:
        st.download_button("Scarica PDF Finale", f.read(), file_name="buoni_finali.pdf", mime="application/pdf")
    st.success("PDF generato: layout, ombre e QR come da template!")

