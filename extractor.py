import os
import re
import PyPDF2
import xlsxwriter
from urllib.parse import urlparse, unquote

def extract_information_from_cv(cv_path):
    # Extract text from PDF
    with open(cv_path, 'rb') as f:
        reader = PyPDF2.PdfReader(f)
        text = ""
        for page_num in range(len(reader.pages)):
            text += reader.pages[page_num].extract_text()

    # Extract email IDs
    emails = re.findall(r'[\w\.-]+@[\w\.-]+', text)

    # Extract phone numbers
    phone_numbers = re.findall(r'(\+\d{1,2}\s?)?(\d{3}[-\.\s]??\d{3}[-\.\s]??\d{4}|\(\d{3}\)\s*\d{3}[-\.\s]??\d{4}|\d{3}[-\.\s]??\d{4})', text)

    return emails, phone_numbers, text

def save_to_excel(emails, phone_numbers, text, output_file):
    workbook = xlsxwriter.Workbook(output_file)
    worksheet = workbook.add_worksheet()

    # Write headers
    worksheet.write(0, 0, 'Email ID')
    worksheet.write(0, 1, 'Phone Number')
    worksheet.write(0, 2, 'Text')

    # Write data
    for i in range(len(emails)):
        worksheet.write(i+1, 0, emails[i])

    for i in range(len(phone_numbers)):
        # Convert phone number tuple to a string
        phone_number_str = ' '.join(phone_numbers[i])
        worksheet.write(i+1, 1, phone_number_str)

    worksheet.write(1, 2, text)

    workbook.close()

def main():
    # URL to the PDF file
    pdf_url = 'file:///D:/Sample2-20240406T093029Z-001/Sample2/AarushiRohatgi.pdf'

    # Parse the URL to extract the local file path and unquote it
    parsed_url = urlparse(pdf_url)
    cv_path = unquote(parsed_url.path)

    # Convert path to the correct format for Windows
    cv_path = cv_path.lstrip('/').replace('/', '\\')

    # Extract information from the CV
    emails, phone_numbers, text = extract_information_from_cv(cv_path)

    # Save information to Excel file
    output_file = 'cv_information.xls'
    save_to_excel(emails, phone_numbers, text, output_file)
    saved_file_path = os.path.abspath(output_file)

    print(f"CV information has been extracted and saved to {output_file} {saved_file_path}")

if __name__ == "__main__":
    main()
