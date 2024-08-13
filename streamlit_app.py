import streamlit as st
import pytesseract
from PIL import Image
import docx
import io
from snowflake.snowpark.context import get_active_session

# Define OCR functions for images and Word documents
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

# Set up the Streamlit interface
st.title("OCR Web Application")
st.caption("Upload an image or Word document to perform OCR.")

# Get the active Snowflake session
session = get_active_session()

# File upload options
uploaded_file = st.file_uploader("Choose an image or Word document", type=["png", "jpg", "jpeg", "docx"])

if uploaded_file is not None:
    # Save the uploaded file to Snowflake stage (optional, for future retrieval)
    file_stage_path = f'@MY_STAGE/{uploaded_file.name}'
    session.file.put_stream(file_stage_path, uploaded_file)
    
    # Determine file type and perform OCR
    if uploaded_file.name.endswith(('.png', '.jpg', '.jpeg')):
        st.write("Performing OCR on the uploaded image...")
        image_text = ocr_image_from_stage(session, file_stage_path)
        st.text_area("Extracted Text", value=image_text, height=200)
    
    elif uploaded_file.name.endswith('.docx'):
        st.write("Performing OCR on the uploaded Word document...")
        word_text = ocr_word_from_stage(session, file_stage_path)
        st.text_area("Extracted Text", value=word_text, height=200)
