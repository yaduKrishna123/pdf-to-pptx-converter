from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from pdf2image import convert_from_path
from pptx import Presentation
from pptx.util import Inches
import os
import uuid

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}}, supports_credentials=True)  # Enable CORS globally

def pdf_to_ppt(pdf_path, pptx_path='output.pptx', dpi=200):
    images = convert_from_path(pdf_path, dpi=dpi)

    prs = Presentation()
    prs.slide_width = Inches(11.69)
    prs.slide_height = Inches(8.27)

    for img in images:
        temp_image = f'temp_{uuid.uuid4().hex}.png'
        img.save(temp_image, 'PNG')

        slide_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(slide_layout)

        slide.shapes.add_picture(temp_image, 0, 0, width=prs.slide_width, height=prs.slide_height)
        os.remove(temp_image)

    prs.save(pptx_path)
def add_cors_headers(response):
    response.headers['Access-Control-Allow-Origin'] = '*'
    response.headers['Access-Control-Allow-Headers'] = 'Content-Type,Authorization'
    response.headers['Access-Control-Allow-Methods'] = 'GET,POST,OPTIONS'
    return response
@app.route('/convert', methods=['POST'])
def convert_pdf():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part in the request'}), 400

    file = request.files['file']

    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400

    if file:
        input_pdf_path = f"temp_{uuid.uuid4().hex}.pdf"
        output_pptx_path = f"output_{uuid.uuid4().hex}.pptx"

        file.save(input_pdf_path)

        try:
            pdf_to_ppt(input_pdf_path, output_pptx_path)
            return send_file(output_pptx_path, as_attachment=True)
        finally:
            if os.path.exists(input_pdf_path):
                os.remove(input_pdf_path)
            if os.path.exists(output_pptx_path):
                os.remove(output_pptx_path)

    return jsonify({'error': 'File processing failed'}), 500

if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
