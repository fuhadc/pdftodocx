from flask import Flask, request, redirect, url_for, send_file
import PyPDF2
import docx

app = Flask(__name__)

@app.route('/')
def display_form():
    return '''
    <html>
        <head>
            <link rel="stylesheet" type="text/css" href="styles.css">
        </head>
        <body>
            <h1>PDF-to-Word Converter</h1>
            <form action="/convert" method="POST" enctype="multipart/form-data">
                <input type="file" name="pdf_file">
                <input type="submit" value="Convert">
            </form>
        </body>
    </html>
    '''

@app.route('/convert', methods=['POST'])
def convert_pdf():
    pdf_file = request.files['pdf_file']
    reader = PyPDF2.PdfFileReader(pdf_file)
    text = ''
    for page in reader.pages:
        text += page.extractText()

    # Create a new Word document
    document = docx.Document()

    # Add the text to the document
    document.add_paragraph(text)

    # Save the document
    document.save('converted.docx')

    return redirect(url_for('download_file'))

@app.route('/download')
def download_file():
    return send_file('converted.docx', as_attachment=True)

if __name__ == '__main__':
    app.run(port=4444)