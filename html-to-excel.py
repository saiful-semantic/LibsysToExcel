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
parser.add_argument(
    "-o", "--out", dest="outFile", required=False, type=str, help="Output Excel file"
)
parser.add_argument(
    "-l", "--limit", dest="limit", required=False, type=int, help="Limit to number of records"
)
args = parser.parse_args()

with open("config.json", "r", encoding="utf-8") as file:
    config = json.load(file)

# Load local overrides if the file exists
try:
    with open("config.local.json", "r", encoding="utf-8") as file:
        local_config = json.load(file)
        config.update(local_config)  # Merge & Override
except FileNotFoundError:
    pass  # No local config, keep defaults

def start_of_record(tr):
    # Find all <td> elements in the current <tr>
    td_elements = tr.find_all('td')
    
    # Check if the second <td> exists and has the desired attributes
    if len(td_elements) > 1:  # Ensure there is at least a second <td>
        second_td = td_elements[1]  # Get the second <td>
        
        # Check if the second <td> has colspan="9" and contains a <span>
        if second_td.get('colspan') == '9':
            span = second_td.find('span')  # Find the <span> inside the <td>
            if span:
                # Capture the text inside the <span>
                span_text = span.text.strip()
                return True, span_text
    
    return False, ''

def is_a_header_row(rows):
    if len(rows) == 1:
        td_elements = rows[0].find_all('td')

        if len(td_elements) > 3:
            second_td = td_elements[1]
            fourth_td = td_elements[3]
            
            if second_td and fourth_td:
                second_span = second_td.find('span')
                fourth_span = fourth_td.find('span')

                library_name = second_span.text.strip() if second_span else ''
                date_label = fourth_span.text.strip() if fourth_span else ''
                
                if config["libraryName"] in library_name and 'Date' in date_label:
                    return True
    
    return False

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
            html_content = file.read().replace("\n", "")
        
        # Remove newlines and replace multiple spaces/tabs with a single space
        html_content = re.sub(r"\s+", " ", html_content) 
        soup = BeautifulSoup(html_content, 'html.parser')
    except FileNotFoundError:
        print(f"Error: File '{html_file}' not found.")
        return []

    tables = soup.find_all('table', class_='jrPage')

    ## Extract all non-empty tr elements
    # all_table_rows = []
    # for table in tables:
    #     rows = table.find_all('tr')
    #     for row in rows:
    #         # Check if the row contains any non-empty text
    #         if any(cell.get_text(strip=True) for cell in row.find_all(['td', 'th'])):
    #             all_table_rows.append(row)
    
    ## ChatGPT simplification :)
    all_table_rows = [
        row for table in tables
        for row in table.find_all('tr')
        if any(cell.get_text(strip=True) for cell in row.find_all(['td', 'th']))
    ]
    # print(all_table_rows)

    ## Prepare master records object
    records = {}

    ## Identify records using a record identifier
    new_record_rows = []
    new_record_type = ''
    record_count = 0
    for tr in tqdm(all_table_rows, desc="Extracting from HTML"):
        new_record, record_type = start_of_record(tr)

        # Found start of a new record
        if new_record and record_type:

            # Ignore header row
            if len(new_record_rows) and not is_a_header_row(new_record_rows):
                record_count += 1
                records[record_count] = {
                    'record_type': new_record_type,
                    'rows': new_record_rows
                }

            new_record_rows = []
            new_record_type = record_type
        else:
            new_record_rows.append(tr)

        if args.limit and record_count == args.limit:
            break
    
    # Save the last row too
    if len(new_record_rows):
        record_count += 1
        records[record_count] = {
            'record_type': new_record_type,
            'rows': new_record_rows
        }

    ## Parse and extract book elements
    book_data = []
    for record_num, record_data in tqdm(records.items(), desc="Parsing data"):
        if not args.outFile:
            print(f"Record# {record_num}")
            print(f"Type: {record_data['record_type']}")
            print(f"Rows: {len(record_data['rows'])}")
        
        barcodes = extract_barcodes(record_data['rows'])
        book = extract_metadata(record_data['rows'])
        for item in barcodes:
            item_record = {
                'Record No.': record_num,
                'Record Type': record_data['record_type'],
                **item,
                **book
            }
            if not args.outFile:
                print(item_record)
        
        if not args.outFile:
            print("-" * 40)  # Separator for readability

    return book_data

def extract_metadata(book_entries):
    book = {}
    number = ""
    author = ""
    call_number = ""
    title = ""

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
                            if len(span.text.strip()) > 0:
                                author = span.text.strip()
                if call_number == "":
                    for span in number_span:
                        if len(span.text.strip().replace('.', '')) > 0 and author not in span.text.strip():
                            if not span.text.strip().replace('.', '').isdigit():
                                call_number = span.text.strip()
                if title == "":
                    for span in number_span:
                        if len(span.text.strip()) > 0 and author not in span.text.strip() and call_number not in span.text.strip():
                            if not span.text.strip().replace('.', '').isdigit():
                                title = span.text.strip()

    book['number'] = number
    book['author'] = author
    book['call_number'] = call_number
    book['title'] = title

    return book

def extract_barcodes(html_string):
    """
    Extracts barcode and accession date from an HTML string.

    Args:
        html_string (str): The HTML string containing the barcode and accn date.

    Returns:
        list: A list of dictionaries, where each dictionary contains barcode and accn date.
    """
    text = "\n".join(tr.get_text(separator=" ", strip=True) for tr in html_string)
    pattern = re.compile(r'(\d+)\s*\([^\)]*\)\s*Accn Date\s*:\s*(\d{2}/\d{2}/\d{4})')
    matches = pattern.findall(text)
    result = []
    for barcode, accn_date in matches:
        result.append({'Barcode': barcode, 'Accn Date': accn_date})
    return result

# Main usage:
book_list = html_to_excel(args.inputFile)

if args.outFile:
    df = pd.DataFrame(book_list)

    # Save the DataFrame to an Excel file
    df.to_excel(args.outFile, index=False)
    print(f"Data saved to {args.outFile}")
