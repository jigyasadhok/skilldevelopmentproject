import cv2
import pytesseract
import pandas as pd
import re
import openpyxl
from openpyxl.utils import get_column_letter

# Set the path to Tesseract-OCR executable
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Preprocess image for better OCR accuracy
def preprocess_image(image_path):
    image = cv2.imread(image_path)
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY)
    return thresh

# Extract text from the image using Tesseract OCR
def extract_text(image):
    return pytesseract.image_to_string(image)

# Parse extracted text to get necessary bill information
def parse_bill_data(text):
    lines = text.split('\n')
    bill_data = {
        "Department Name": None,
        "Bill No": None,
        "Purchase Order No": None,
        "Date": None,
        "Total Amount": None
    }
    
    for line in lines:
        if re.search(r"Department\s*Name", line, re.IGNORECASE):
            bill_data["Department Name"] = line.split(":")[-1].strip()
        elif re.search(r"Bill\s*No", line, re.IGNORECASE):
            bill_data["Bill No"] = line.split(":")[-1].strip()
        elif re.search(r"Purchase\s*Order\s*No", line, re.IGNORECASE):
            bill_data["Purchase Order No"] = line.split(":")[-1].strip()
        elif re.search(r"Date", line, re.IGNORECASE):
            bill_data["Date"] = line.split(":")[-1].strip()
        elif re.search(r"Total\s*Amount", line, re.IGNORECASE):
            bill_data["Total Amount"] = re.sub(r'%', '', line.split(":")[-1].strip())
    
    return bill_data

# Organize parsed data into a DataFrame
def organize_data(bill_data):
    return pd.DataFrame([bill_data])

# Export DataFrame to an Excel file, appending new data without erasing old data
def export_to_excel(df, excel_path, image_name):
    try:
        existing_df = pd.read_excel(excel_path)
        final_df = pd.concat([existing_df, df], ignore_index=True)
    except FileNotFoundError:
        final_df = df

    with pd.ExcelWriter(excel_path, engine='openpyxl', mode='w') as writer:
        final_df.to_excel(writer, index=False)

    # Maintain the existing cell width of 30
    workbook = openpyxl.load_workbook(excel_path)
    sheet = workbook.active
    for column in sheet.columns:
        sheet.column_dimensions[get_column_letter(column[0].column)].width = 30  # Default width
    
    workbook.save(excel_path)
    
    # Print confirmation on the console
    print(f"Data exported to {excel_path} from image {image_name}")

# Main function to execute the workflow for multiple images
def main(image_paths, excel_path):
    all_data = []
    
    for image_path in image_paths:
        processed_image = preprocess_image(image_path)
        text = extract_text(processed_image)
        bill_data = parse_bill_data(text)
        df = organize_data(bill_data)
        all_data.append(df)
    
    # Concatenate all DataFrames into a single DataFrame
    final_df = pd.concat(all_data, ignore_index=True)
    export_to_excel(final_df, excel_path, image_paths[-1])  # Pass the last image name for the export message

# Example usage
image_paths = [
    r'C:\Users\HP\Pictures\Screenshots\Screenshot 2024-08-13 082331.png',
    r'C:\Users\HP\Pictures\Screenshots\Screenshot 2024-08-12 195804.png'
]  # List of image paths
excel_path = r'C:\Users\HP\Desktop\out.xlsx'  # Path to save the Excel file
main(image_paths, excel_path)
