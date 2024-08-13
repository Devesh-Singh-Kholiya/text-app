import pytesseract
from PIL import Image
import docx
import io

# Set the path to the Tesseract OCR executable
pytesseract.pytesseract.tesseract_cmd = r'C:\Users\dkholiya\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'

def ocr_image_from_stage(session, stage_path):
    # Get the image file from the Snowflake stage
    image_data = session.file.get_stream(stage_path)
    image_bytes = image_data.read()
    
    # Open the image file from the bytes
    img = Image.open(io.BytesIO(image_bytes))
    
    # Use Tesseract to do OCR on the image
    text = pytesseract.image_to_string(img)
    return text

def ocr_word_from_stage(session, stage_path):
    # Get the Word document file from the Snowflake stage
    docx_data = session.file.get_stream(stage_path)
    docx_bytes = docx_data.read()
    
    # Open the Word document from the bytes
    doc = docx.Document(io.BytesIO(docx_bytes))
    
    # Extract text from each paragraph
    text = '\n'.join([para.text for para in doc.paragraphs])
    return text

# Example usage in Snowflake
session = get_active_session()  # Get the active Snowflake session

# Specify the paths to your files in the Snowflake stage
image_stage_path = '@MY_STAGE/test_image.png'
docx_stage_path = '@MY_STAGE/test_doc.docx'

# Perform OCR on the image and Word document
image_text = ocr_image_from_stage(session, image_stage_path)
word_text = ocr_word_from_stage(session, docx_stage_path)

# Print the extracted text
print("Image Text:\n", image_text)
print("\nWord Text:\n", word_text)
