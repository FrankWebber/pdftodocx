import os
from flask import Flask, request, redirect, send_from_directory, send_file, render_template, jsonify
import pytesseract
from pdf2image import convert_from_path
from PIL import Image
from docx import Document
import pdfplumber
import io

app = Flask(__name__)

# Configurações
UPLOAD_FOLDER = os.getenv('UPLOAD_FOLDER', 'uploads')
RESULT_FOLDER = os.getenv('RESULT_FOLDER', 'results')
ALLOWED_EXTENSIONS = {'pdf'}
pytesseract.pytesseract.tesseract_cmd = os.getenv('TESSERACT_CMD', '/usr/bin/tesseract')
poppler_path = os.getenv('POPPLER_PATH', '/usr/bin/poppler')

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['RESULT_FOLDER'] = RESULT_FOLDER

# Funções auxiliares
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def aplicar_ocr_pdf(caminho_arquivo):
    paginas = convert_from_path(caminho_arquivo, 300, poppler_path=poppler_path)
    texto_completo = ""
    for pagina in paginas:
        texto = pytesseract.image_to_string(pagina, lang='por')
        texto_completo += texto + "\n\n"
    return texto_completo

def salvar_como_docx(texto, caminho_saida):
    doc = Document()
    doc.add_paragraph(texto)
    doc.save(caminho_saida)

def convert_pdf_to_docx(pdf_bytes):
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        doc = Document()
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                doc.add_paragraph(text)
        doc_io = io.BytesIO()
        doc.save(doc_io)
        doc_io.seek(0)
        return doc_io

# Rotas do Flask
@app.route('/')
def home():
    return render_template('upload.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return redirect(request.url)
    file = request.files['file']
    if file.filename == '':
        return redirect(request.url)
    if file and allowed_file(file.filename):
        filename = 'uploaded.pdf'
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)

        # Verifica se o PDF contém imagens
        try:
            texto = aplicar_ocr_pdf(file_path)
            docx_path = os.path.join(app.config['RESULT_FOLDER'], 'resultado.docx')
            salvar_como_docx(texto, docx_path)
            return send_from_directory(app.config['RESULT_FOLDER'], 'resultado.docx', as_attachment=True)
        except Exception as e:
            with open(file_path, 'rb') as f:
                pdf_bytes = f.read()
            docx_file = convert_pdf_to_docx(pdf_bytes)
            return send_file(
                docx_file,
                as_attachment=True,
                download_name=file.filename.replace('.pdf', '.docx')
            )

    return redirect(request.url)

@app.route('/convert', methods=['POST'])
def convert():
    file = request.files.get('file')
    if file and allowed_file(file.filename):
        filename = 'uploaded.pdf'
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)

        try:
            texto = aplicar_ocr_pdf(file_path)
            docx_path = os.path.join(app.config['RESULT_FOLDER'], 'resultado.docx')
            salvar_como_docx(texto, docx_path)
        except Exception as e:
            with open(file_path, 'rb') as f:
                pdf_bytes = f.read()
            docx_file = convert_pdf_to_docx(pdf_bytes)
            docx_path = os.path.join(app.config['RESULT_FOLDER'], 'resultado.docx')
            with open(docx_path, 'wb') as f:
                f.write(docx_file.getvalue())

        return send_from_directory(app.config['RESULT_FOLDER'], 'resultado.docx', as_attachment=True)

    return jsonify({"message": "Falha no upload ou arquivo inválido!"})

if __name__ == '__main__':
    if not os.path.exists(UPLOAD_FOLDER):
        os.makedirs(UPLOAD_FOLDER)
    if not os.path.exists(RESULT_FOLDER):
        os.makedirs(RESULT_FOLDER)
    app.run(host='0.0.0.0', port=5000)
