import streamlit as st
import os
import re
import tempfile
from zipfile import ZipFile
from docx import Document
import subprocess
from PyPDF2 import PdfMerger
import cairosvg
from PIL import Image
import fitz  

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

def get_qr_rect_auto(pdf_path):
    pdf = fitz.open(pdf_path)
    page = pdf[0]
    imglist = page.get_images(full=True)
    if not imglist:
        raise Exception("Nessuna immagine trovata nel PDF! Usa un template Word con QR finto.")
    xref = imglist[0][0]
    for b in page.get_text("dict")["blocks"]:
        if "image" in b:
            return fitz.Rect(b["bbox"])
    # In caso non funzioni, fallback a bbox della prima immagine
    pix = fitz.Pixmap(pdf, xref)
    rect = fitz.IRect(0, 0, pix.width, pix.height)
    return rect

def sovrascrivi_qr_auto(pdf_path, qr_path, output):
    rect = get_qr_rect_auto(pdf_path)
    pdf = fitz.open(pdf_path)
    page = pdf[0]
    page.insert_image(rect, filename=qr_path, overlay=True)
    pdf.save(output)
    pdf.close()
    return output

def process_buoni(template_file, zip_file):
    temp_dir = tempfile.mkdtemp()
    qr_folder = os.path.join(temp_dir, 'qr')
    os.makedirs(qr_folder)
    output_pdf_list = []
    template_path = os.path.join(temp_dir, 'template.docx')
    with open(template_path, 'wb') as f:
        f.write(template_file.getbuffer())
    zip_path = os.path.join(temp_dir, 'qrcodes.zip')
    with open(zip_path, 'wb') as f:
        f.write(zip_file.getbuffer())
    with ZipFile(zip_path, 'r') as zf:
        svg_files = [f for f in zf.namelist() if f.lower().endswith('.svg')]
        if not svg_files:
            st.error("Nessun file SVG trovato nello ZIP!")
            return None
        for svg_path in svg_files:
            folder = svg_path.split('/')[0]
            match = re.search(r'-([a-zA-Z0-9]+)-', folder)
            num = match.group(1) if match else folder
            svg_data = zf.read(svg_path)
            qr_png_path = os.path.join(qr_folder, f"{num}.png")
            cairosvg.svg2png(bytestring=svg_data, write_to=qr_png_path)
            doc = Document(template_path)
            for para in doc.paragraphs:
                if 'Buono n.' in para.text and 'valido fino al' in para.text:
                    para.text = re.sub(r'Buono n\.\s*\w+', f'Buono n. {num}', para.text)
            temp_word = os.path.join(temp_dir, f"Buono_{num}.docx")
            doc.save(temp_word)
            pdf_path = convert_docx_to_pdf(temp_word, temp_dir)
            final_pdf = os.path.join(temp_dir, f"Buono_{num}_finale.pdf")
            sovrascrivi_qr_auto(pdf_path, qr_png_path, final_pdf)
            output_pdf_list.append(final_pdf)
    merger = PdfMerger()
    for pdf in output_pdf_list:
        merger.append(pdf)
    final_pdf_path = os.path.join(temp_dir, "buoni_finali.pdf")
    merger.write(final_pdf_path)
    merger.close()
    return final_pdf_path

# --- STREAMLIT UI ---

st.title("Generatore Buoni Carburante PDF — QR automatico dal template")
st.write("Carica il template Word (con un QR generico/placeholder) e ZIP con QR code in SVG. L'app sostituirà automaticamente ogni QR sulla base del layout originale.")

template_file = st.file_uploader("Template Word (.docx con QR placeholder)", type='docx')
zip_file = st.file_uploader("ZIP con QR SVG (uno per buono)", type='zip')

if st.button("Genera PDF finale") and template_file and zip_file:
    st.write("Elaborazione in corso...")
    try:
        final_pdf = process_buoni(template_file, zip_file)
        if final_pdf:
            with open(final_pdf, "rb") as f:
                st.download_button("Scarica PDF Finale", f.read(), file_name="buoni_finali.pdf", mime="application/pdf")
            st.success("PDF creato: QR in posizione e dimensione come nel template Word!")
    except Exception as e:
        st.error(f"Errore durante la generazione: {e}")

