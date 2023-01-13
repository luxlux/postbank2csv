# postbank2csv
Convert Postbank account statements to a csv file.

Inspired by https://github.com/FlatheadV8/Postbank_PDF2CSV.

## Dependencies
- python >= 3.6
- pdftotext from poppler (brew install poppler) and then (pip install pdftotext)

## Input
Postbank account statement pdf files from July 2017 or later. Before the format was different.

## Output
On stdout, all statements, in the order they appeared in the pdf file(s).
The columns are: Buchungsdatum, Empfaenger, Verwendungszweck, Betrag, Typ, Einreicher-id, Referenz, Mandat ,Zusammen

Values are comma separated and, if needed, inside double quotes.

## Usage
postbank2csv.py [-h] pdf_file [pdf_file ...]
