# Extract data from Libsys

Tools to extract data from Libsys catalogue report in PDF or HTML to Excel sheets

## Requirements

```bash
pip install -r requirements.txt
```

## From PDF to Excel
```bash
python pdf-to-excel.py -h
python pdf-to-excel.py -i sample-libsys.pdf -o Test1.xlsx
```

## From HTML to Excel

Update `config.json`: set `libraryName` as per the header in HTML pages to ignore the first empty record

```bash
python html-to-excel.py -h
python html-to-excel.py -i sample-libsys.html -o Test2.xlsx
```
