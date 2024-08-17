from flask import Flask, request, send_file, render_template
import pdfplumber
from docx import Document
import io

app = Flask(__name__)

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

@app.route('/')
def home():
    return render_template('upload.html')

@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return 'Nenhum arquivo foi enviado.', 400
    file = request.files['file']
    if file.filename == '':
        return 'Nenhum arquivo selecionado.', 400
    if file and file.filename.endswith('.pdf'):
        docx_file = convert_pdf_to_docx(file.read())
        return send_file(
            docx_file,
            as_attachment=True,
            download_name=file.filename.replace('.pdf', '.docx')
        )
    else:
        return 'Formato de arquivo não suportado.', 400

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
