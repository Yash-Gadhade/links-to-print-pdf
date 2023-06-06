import subprocess
from openpyxl import load_workbook
import os
from urllib.parse import urlparse
from bs4 import BeautifulSoup
import requests

# Load the Excel file
workbook = load_workbook('input.xlsx')
sheet = workbook.active

# Get the column of links
column = 'A'
start_row = 47
end_row = 87  # Set the end row according to your data

# Create a directory to store the PDF files
output_directory = 'pdf_files'
os.makedirs(output_directory, exist_ok=True)

# Iterate over the links in the column
for index, row in enumerate(range(start_row, end_row + 1), 1):
    link = sheet[f'{column}{row}'].value

    # Skip empty cells
    if not link:
        continue

    try:
        # Send a GET request to the website and get the HTML content
        response = requests.get(link, verify=False)
        html_content = response.text

        # Parse the HTML content using BeautifulSoup
        soup = BeautifulSoup(html_content, 'html.parser')

        # Extract the website title from the HTML
        title_tag = soup.title
        if title_tag is None:
            title = f'{row:02d}'
        else:
            title = title_tag.string.strip()

        # Generate a file name with the row number and title
        file_name = f'{row:02d}_{title}.pdf'
        file_path = os.path.join(output_directory, file_name)

        # Convert the web page to PDF and save it
        subprocess.run(['wkhtmltopdf', '--page-size', 'A4', '--margin-top', '0mm', '--margin-bottom', '0mm', '--margin-left', '0mm', '--margin-right', '0mm', link, file_path])

        print(f"Saved page {row} as PDF: {link}")

    except requests.exceptions.SSLError as e:
        print(f"Error occurred while accessing {link}: {e}")

print("All pages saved as PDF successfully!")
