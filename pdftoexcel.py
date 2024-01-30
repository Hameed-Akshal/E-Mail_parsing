import os
import spacy
import pdfplumber
import pandas as pd
import fitz 
import re
import phonenumbers
import validate_email_address

nlp = spacy.load("en_core_web_sm")

def extract_email_addresses(text):
    email_pattern = re.compile(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b')
    emails = re.findall(email_pattern, text)
    valid_emails = [email for email in emails if is_valid_email(email)]
    return valid_emails

def is_valid_email(email):
    try:
        v = validate_email(email)
        return v.is_valid
    except Exception:
        print(f"Error validating email {email}: {e}")
        return False

def extract_phone_numbers(text):
    phone_pattern = re.compile(r'\b(?:\d{3}[-.\s]?){1,2}\d{4}\b')
    phones = re.findall(phone_pattern, text)
    return phones
    
def parse_pdf_resume(file_path):
    try:
        with fitz.open(file_path) as pdf:
            resume_text = ""
            for page_number in range(pdf.page_count):
                page = pdf[page_number]
                resume_text += page.get_text()
    except Exception as e:
        print(f"Error opening PDF {file_path}: {e}")
        return None
    
    try:
        doc = nlp(resume_text)
    except Exception as e:
        print(f"Error processing document {file_path}: {e}")
        return None
    
    # Extract relevant information from the spaCy doc
    names = [ent.text for ent in doc.ents if ent.label_ == "PERSON"]
    name = " ".join(names)

    # Extract emails and phones using the updated functions
    emails = extract_email_addresses(resume_text)
    email = emails[0] if emails else ""

    # Use the custom function to extract phone numbers
    phones = extract_phone_numbers(resume_text)
    phone = phones[0] if phones else ""

    print(email)
    return {
        'Email': email,
        'Phone': phone,
        'Content': doc
    }

resumes_directory = r'D:\project\resumes'
parsed_resumes = []

for filename in os.listdir(resumes_directory):
    if filename.endswith('.pdf'):
        file_path = os.path.join(resumes_directory, filename)
        parsed_data = parse_pdf_resume(file_path)
        if parsed_data:
            parsed_resumes.append(parsed_data)

df = pd.DataFrame(parsed_resumes)

excel_file_path = r'D:\project\output_data_demo.xlsx'
df.to_excel(excel_file_path, index=False, sheet_name='Resumes', engine='xlsxwriter')
