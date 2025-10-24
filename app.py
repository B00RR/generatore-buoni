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
import fitz  # PyMuPDF

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
    images = page.get_images(full=True)
    if len(images) == 0:
        raise Exception("Nessuna immagine trovata nella prima pagina del PDF! Usa un template Word con un QR placeholder visibile.")
    # Individua la posizione della PRIMA immagine (presume sia il QR placeholder)
    for img in images:
        xref = img[0]
        # Cerca tutte le bbox di immagini su quella pagina con get_image_bbox
        rects = page.get_image_bbox(xref)
        if rects:  # Prende il primo bbox trovato
            return rects[0]
    raise Exception("QR/Placeholder non trovato tra le immagini.")

def sovrascrivi_qr_auto(pdf_path, qr_path, output):
    rect = get_qr_rect_auto(pdf_path)
    pdf = fitz.open(pdf_path)
    page = pdf[0]
    # Sovrascrivi l'immagine nella posizione del QR placeholder
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
            # Estrai numero buono dalla cartella o dal nome come prima
            folder = svg_path.split('/')[0]
            match = re.search(r'-([a-zA-Z0-9]+)-', folder)
            num = match.group(1) if match else folder
            svg_data = zf.read(svg_path)
            qr_png_path = os.path.join(qr_folder, f"{num}.png")
            cairosvg.svg2png(bytestring=svg_data, write_to=qr_png_path)
            # Sostituisci il numero nel Word
            doc = Document(template_path)
            for para in doc.paragraphs:
                if 'Buono n.' in para.text and 'valido fino al' in para.text:
                    para.text = re.sub(r'Buono n\.\s*\w+', f'Buono n. {num}', para.text)
            temp_word = os.path.join(temp_dir, f"Buono_{num}.docx")
            doc.save(temp_word)
            # Converti in PDF (il QR placeholder del template Word resta visibile)
            pdf_path = convert_docx_to_pdf(temp_word, temp_dir)
            final_pdf = os.path.join(temp_dir, f"Buono_{num}_finale.pdf")
            sovrascrivi_qr_auto(pdf_path, qr_png_path, final_pdf)
            output_pdf_list.append(final_pdf)
    # Unisci tutti i PDF finali
    merger = PdfMerger()
    for pdf in output_pdf_list:
        merger.append(pdf)
    final_pdf_path = os.path.join(temp_dir, "buoni_finali.pdf")
    merger.write(final_pdf_path)
    merger.close()
    return final_pdf_path

# --- STREAMLIT UI ---

st.title("Generatore Buoni Carburante PDF — Rilevamento QR automatico")
st.write("Carica template Word (con QR fittizio) e ZIP con QR .svg. Il sistema rileva posizione/dimensione del QR placeholder e sostituisce ogni QR in automatico.")

template_file = st.file_uploader("Carica template Word (deve avere QR placeholder)", type='docx')
zip_file = st.file_uploader("Carica ZIP con QR (SVG)", type='zip')

if st.button("Genera PDF finale") and template_file and zip_file:
    st.write("Elaborazione in corso...")
    try:
        final_pdf = process_buoni(template_file, zip_file)
        if final_pdf:
            with open(final_pdf, "rb") as f:
                st.download_button("Scarica PDF Finale", f.read(), file_name="buoni_finali.pdf", mime="application/pdf")
                st.success("PDF generato con QR esattamente nella posizione/template originale!")
    except Exception as e:
        st.error(f"Errore durante la generazione: {e}")
