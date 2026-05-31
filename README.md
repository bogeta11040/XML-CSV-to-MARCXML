# XML-CSV-to-MARCXML
Book List Converter (XLS/XLSX/CSV/XML) to MARCXML Format.

Supported input formats:
  - CSV  (.csv)
  - XLS  (.xls)  — converted internally
  - XLSX (.xlsx)
  - XML  (.xml)  — first converted to CSV, then to MARCXML

Usage:


  python3 books_to_marcxml.py input.csv
  python3 books_to_marcxml.py input.xls
  python3 books_to_marcxml.py input.xlsx
  python3 books_to_marcxml.py input.xml
  python3 books_to_marcxml.py input.csv --output my_library.xml
