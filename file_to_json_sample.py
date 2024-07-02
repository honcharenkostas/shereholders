import os
import json
import pytesseract
from PIL import Image
import fitz  # PyMuPDF
import google.generativeai as genai

def extract_text_from_tiff(file_path):
    # Open the TIFF file
    img = Image.open(file_path)
    # Use Tesseract to do OCR on the image
    text = pytesseract.image_to_string(img)
    return text

def extract_text_from_pdf(file_path):
    # Open the PDF file
    pdf_document = fitz.open(file_path)
    text = ""
    # Extract text from each page
    for page_num in range(len(pdf_document)):
        page = pdf_document.load_page(page_num)
        text += page.get_text()
    return text

# Function to handle both file types
def extract_text(file_path):
    if file_path.lower().endswith('.tiff') or file_path.lower().endswith('.tif'):
        return extract_text_from_tiff(file_path)
    elif file_path.lower().endswith('.pdf'):
        return extract_text_from_pdf(file_path)
    else:
        raise ValueError("Unsupported file type. Please provide a TIFF or PDF file.")

# Example usage
file_path = '/Users/stas/code/shereholders/data/SN-Dresden_HRB_29795+List_of_shareholders_ _entry_in_the_register_folde-20240701220923.tiff'  # or 'path_to_your_file.pdf'
text = extract_text(file_path)
genai.configure(api_key=os.environ['GEMINI_API_KEY'])
model = genai.GenerativeModel('gemini-1.5-flash')
prompt = f'''
Here is a content of file:
{text}
######## 
Please analyse a text I provide above 
and return a json of shareholders as list of objects with next fields per object:
1. name - (str) shareholder name;
2. percentage - (int) percentage of shareholdings;
3. date_of_birth - (str) date of birth of shareholder in format 'month.day.year';
4. age - (int) age of shareholder in years;

No need to explain, no need to wrap in "```json". Return plain json only.
'''.strip()


response = model.generate_content(prompt)
result = json.loads(response.text)
print(result)
# f = open("resp.json", "w")
# f.write(result)
# f.close()
