from flask import Flask, render_template, request, send_from_directory, session
import os
from google.cloud import vision
from skilldevelopment1 import extract_text_from_image, parse_bill_data, organize_data, export_to_excel

app = Flask(__name__)

# Set a secret key to use sessions
app.secret_key = 'your_secret_key'

# Directory to store Excel files
EXCEL_FOLDER = os.path.join(app.root_path, 'excel_files')
os.makedirs(EXCEL_FOLDER, exist_ok=True)  # Create folder if not exists

# Initialize Google Cloud Vision client
def initialize_vision_client():
    return vision.ImageAnnotatorClient()

# Route for the homepage
@app.route('/')
def index():
    # Get list of all existing Excel files
    excel_files = [f for f in os.listdir(EXCEL_FOLDER) if f.endswith('.xlsx')]
    return render_template('index.html', excel_files=excel_files)

# Route to handle image upload and processing
@app.route('/upload', methods=['POST'])
def upload_file():
    if request.method == 'POST':
        # Save the uploaded file
        image = request.files['file']
        image_path = os.path.join(app.root_path, image.filename)
        image.save(image_path)

        # Initialize Vision client
        client = initialize_vision_client()

        # Check if user wants to create a new Excel file
        create_new = request.form.get('create_new') == 'on'

        # Check if user selected an existing file or entered a new one
        existing_file = request.form.get('existing_file')
        new_file_name = request.form.get('new_file_name')

        # Determine the Excel file to use
        if create_new and new_file_name:
            # Create a new file with the user-entered name
            excel_file_name = f"{new_file_name}.xlsx"
        elif existing_file:
            # Use the selected existing file
            excel_file_name = existing_file
        else:
            # Default to a new file if no existing file is selected and no new name is provided
            excel_file_name = "output.xlsx"

        # Full path for the Excel file
        excel_path = os.path.join(EXCEL_FOLDER, excel_file_name)

        # Process the image and pass the client
        text = extract_text_from_image(image_path, client)
        bill_data = parse_bill_data(text)
        df = organize_data(bill_data)

        # Export the data to Excel
        export_to_excel(df, excel_path, create_new=create_new)

        # Store the Excel file name in the session
        session['excel_file'] = excel_file_name

        # Fetch the updated list of Excel files after upload
        excel_files = [f for f in os.listdir(EXCEL_FOLDER) if f.endswith('.xlsx')]

        # Render the template with the updated list of Excel files
        return render_template('index.html', excel_files=excel_files, excel_available=True)

# Route to display the Excel file
@app.route('/view_excel')
def view_excel():
    # Retrieve the Excel file name from the session
    excel_file = session.get('excel_file', 'output.xlsx')
    excel_path = os.path.join(EXCEL_FOLDER, excel_file)

    # Check if the file exists
    if os.path.exists(excel_path):
        return send_from_directory(EXCEL_FOLDER, excel_file)
    else:
        return "Excel file not found", 404

if __name__ == "__main__":
    app.run(debug=True, use_reloader=False)
