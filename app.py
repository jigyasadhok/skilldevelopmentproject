from flask import Flask, request, render_template, redirect, url_for
import os
from skilldevelopment import preprocess_image, extract_text, parse_bill_data, organize_data, export_to_excel

app = Flask(__name__)

# Configure the upload folder
app.config['UPLOAD_FOLDER'] = 'uploads/'
if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return redirect(request.url)
    file = request.files['file']
    if file.filename == '':
        return redirect(request.url)
    if file:
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(filepath)
        # Call your backend processing function here
        processed_image = preprocess_image(filepath)
        text = extract_text(processed_image)
        bill_data = parse_bill_data(text)
        df = organize_data(bill_data)
        
        excel_path = os.path.join(app.config['UPLOAD_FOLDER'], 'out.xlsx')
        export_to_excel(df, excel_path, file.filename)
        # Example: process_image(filepath)
        return redirect(url_for('result', filename=file.filename))

@app.route('/result/<filename>')
def result(filename):
    return f"File {filename} processed successfully!"

if __name__ == '__main__':
    app.run(debug=True)
