import json
import re
import pandas as pd
from bs4 import BeautifulSoup
from tqdm import tqdm
from argparse import ArgumentParser

parser = ArgumentParser(description='HTML to Excel data extraction for bibliographic data.')
parser.add_argument(
    "-i", "--input", dest="inputFile", required=True, type=str, help="The input HTML file"
)
# parser.add_argument(
#     "-o", "--out", dest="outFile", required=True, type=str, help="Output Excel file"
# )
parser.add_argument(
    "-l", "--limit", dest="limit", required=False, type=int, help="Limit to number of records"
)
args = parser.parse_args()

def html_to_excel(html_file):
    """
    Extracts book data from an HTML file, joining data across tables.

    Args:
        html_file (str): Path to the HTML file.

    Returns:
        list: A list of dictionaries, where each dictionary represents a book.
    """
    try:
        with open(html_file, 'r', encoding='utf-8') as file:
            soup = BeautifulSoup(file, 'html.parser')
    except FileNotFoundError:
        print(f"Error: File '{html_file}' not found.")
        return []

    book_data = []
    tables = soup.find_all('table', class_='jrPage')

    all_book_entries = []
    for table in tables:
        rows = table.find_all('tr')
        bs_count = 0
        for i, row in enumerate(rows):
          if i>3:
            cols = row.find_all('td')
            if len(cols) > 0:
              bs_test = row.find_all('span', string="BS")
              if len(bs_test)>0:
                bs_count+=1
              if bs_count>0:
                all_book_entries.append(row)

    book_entries_list = []
    current_book_entries = []
    for row in all_book_entries:
      bs_test = row.find_all('span', string="BS")
      if len(bs_test)>0 and len(current_book_entries)>0:
        book_entries_list.append(current_book_entries.copy())
        current_book_entries.clear()
      current_book_entries.append(row)
    if len(current_book_entries) > 0:
      book_entries_list.append(current_book_entries.copy())

    for book_entries in book_entries_list:
      book = process_book_entries(book_entries)
      book_data.append(book)

    return book_data

def process_book_entries(book_entries):
    book = {}
    number = ""
    author = ""
    call_number = ""
    title = ""
    description = ""
    location = ""
    barcode_accn_data = []  # Initialize an empty list for barcode/accn dat

    for row in book_entries:
        cols = row.find_all('td')
        if len(cols) > 0:
            number_span = row.find_all('span')
            if len(number_span) > 0:
                if number == "":
                    for span in number_span:
                        if span.text.strip().replace('.', '').isdigit():
                            number = span.text.strip()
                if author == "":
                    for span in number_span:
                        if not span.text.strip().replace('.', '').isdigit() and call_number == "":
                            if len(span.text.strip()) > 0 and "BS" not in span.text.strip() and "Loc" not in span.text.strip():
                                author = span.text.strip()
                if call_number == "":
                    for span in number_span:
                        if len(span.text.strip().replace('.', '')) > 0 and author not in span.text.strip() and "BS" not in span.text.strip() and "Loc" not in span.text.strip():
                            if not span.text.strip().replace('.', '').isdigit():
                                call_number = span.text.strip()
                if title == "":
                    for span in number_span:
                        if len(span.text.strip()) > 0 and author not in span.text.strip() and call_number not in span.text.strip() and "BS" not in span.text.strip() and "Loc" not in span.text.strip():
                            if not span.text.strip().replace('.', '').isdigit():
                                title = span.text.strip()
                if description == "":
                    for span in number_span:
                        if len(span.text.strip()) > 0 and author not in span.text.strip() and call_number not in span.text.strip() and title not in span.text.strip() and "BS" not in span.text.strip() and "Loc" not in span.text.strip():
                            if not span.text.strip().replace('.', '').isdigit():
                                description += span.text.strip()
                        barcode_accn_list = extract_barcode_accn_date(str(span))
                        if barcode_accn_list:
                            barcode_accn_data.extend(barcode_accn_list)

                if location == "":
                    for span in number_span:
                        if "Loc" in span.text.strip():
                            location = span.text.strip()

    book['number'] = number
    book['author'] = author
    book['call_number'] = call_number
    book['title'] = title
    book['description'] = description
    book['location'] = location
    book['barcode_accn_data'] = barcode_accn_data

    return book

def extract_barcode_accn_date(html_string):
    """
    Extracts barcode and accession date from an HTML string.

    Args:
        html_string (str): The HTML string containing the barcode and accn date.

    Returns:
        list: A list of dictionaries, where each dictionary contains barcode and accn date.
    """
    soup = BeautifulSoup(html_string, 'html.parser')
    text = soup.get_text(separator=' ')
    pattern = re.compile(r'(\d+)\s*\([^\)]*\)\s*Accn Date\s*:\s*(\d{2}/\d{2}/\d{4})')
    matches = pattern.findall(text)
    result = []
    for barcode, accn_date in matches:
        result.append({'Barcode': barcode, 'Accn Date': accn_date})
    return result

# Main usage:
book_list = html_to_excel(args.inputFile)
json_string = json.dumps(book_list, indent=4)

print(json_string)
