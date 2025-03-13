import fitz
import re
import pandas as pd
from tqdm import tqdm
from argparse import ArgumentParser

parser = ArgumentParser(description='PDF to Excel data extraction for bibliographic data.')
parser.add_argument(
    "-i", "--input", dest="inputFile", required=True, type=str, help="The input PDF file"
)
parser.add_argument(
    "-o", "--out", dest="outFile", required=True, type=str, help="Output Excel file"
)
parser.add_argument(
    "-l", "--limit", dest="limit", required=False, type=int, help="Limit to number of records"
)
args = parser.parse_args()

def pdf_to_excel(pdf_path):
    book_data = []
    doc = fitz.open(pdf_path)
    total_pages = args.limit if args.limit else len(doc)

    for page_num in tqdm(range(total_pages), desc="Processing Pages"):
        page = doc[page_num]
        text = page.get_text("text")
        entries = re.split(r'[BS|TH]\n(\d+)\.\n', text)

        # Print each entry separately
        for i in range(1, len(entries), 2):
            recordNum = entries[i]
            entry = entries[i + 1].strip()
            if entry:  # Ensure the entry is not empty
                # print(f'===== Entry {recordNum} =====\n{entry}\n-----')

                data = {}
                data["Number"] = recordNum
                data["Full Record"] = entry
                # print(f'RecordNum: {data["Number"]}')

                # Author
                # data["Author"] = entry.split('\n')[0]
                # print(f'Author: {data["Author"]}')

                # Extract barcode and accn date using regex pattern
                full_entry = entry.replace('\n', ' ')
                pattern = re.compile(r'(.[^\s]\d+)\s*\([^\(]+\)?\s*Accn Date : (\d{2}\/\d{2}\/\d{4})')

                # Print the extracted barcode and accn date
                barcode_matches = pattern.findall(full_entry)
                for i, match in enumerate(barcode_matches):
                    # print(f'Iterator: {i}, Barcode: {match[0]}, Accn Date: {match[1]}')
                    data_copy = data.copy()  # Copy existing data
                    data_copy["Barcode"] = match[0]
                    data_copy["Accn Date"] = match[1]
                    if i > 0:
                        data_copy["Full Record"] = "Same as above"
                    book_data.append(data_copy)
                
                # print('............\n')

    # Create a pandas DataFrame from the extracted data
    df = pd.DataFrame(book_data)

    # Define the desired column order
    column_order = [
        "Number",
        "Barcode",
        "Accn Date",
        # "Author",
        "Full Record"
    ]

    # Add missing columns to the DataFrame and reorder
    for col in column_order:
        if col not in df.columns:
            df[col] = None

    df = df[column_order]

    # Save the DataFrame to an Excel file
    df.to_excel(args.outFile, index=False)
    print(f"Data saved to {args.outFile}")
    # print(book_data)

pdf_to_excel(args.inputFile)
