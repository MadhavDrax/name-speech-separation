import os
from flask import Flask, request, render_template, send_file
from openpyxl import Workbook
import re

# Set up Flask and point the template folder to the current directory
app = Flask(__name__, template_folder=os.path.dirname(os.path.abspath(__file__)))
UPLOAD_FOLDER = 'public'
OUTPUT_FOLDER = 'public/output'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER

# Ensure upload and output folders exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# Step 1: Parse the DOCX file and extract names and speeches
def parse_doc(input_file):
    from docx import Document  # Import here to keep dependencies minimal
    doc = Document(input_file)
    data = []
    current_name = ""
    current_speech = []

    name_pattern = re.compile(r"#\d+#")

    for para in doc.paragraphs:
        detected_name = False
        
        if re.match(name_pattern, para.text.strip()):
            detected_name = True

        for run in para.runs:
            if run.bold or run.font.highlight_color is not None:
                detected_name = True
                break

        if detected_name:
            if current_name:
                data.append((current_name, " ".join(current_speech)))
            current_name = para.text.strip('#').strip()
            current_speech = []
        else:
            current_speech.append(para.text)

    if current_name:
        data.append((current_name, " ".join(current_speech)))

    return data

# Step 2: Create an Excel file with the extracted data
def create_output_excel(data, output_file):
    wb = Workbook()
    ws = wb.active
    ws.title = "Name & Speech"

    # Write headers
    ws.append(["Name", "Speech"])

    # Write data rows
    for name, speech in data:
        ws.append([name, speech])

    wb.save(output_file)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return 'No file part'
    
    file = request.files['file']
    if file.filename == '':
        return 'No selected file'
    
    if file and file.filename.endswith('.docx'):
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(file_path)
        
        # Process the uploaded file
        data = parse_doc(file_path)
        output_file = os.path.join(app.config['OUTPUT_FOLDER'], 'output.xlsx')
        create_output_excel(data, output_file)
        
        # Send the Excel file back to the user
        return send_file(output_file, as_attachment=True)

    return 'Invalid file format. Please upload a .docx file.'

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=True)
