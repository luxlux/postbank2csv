#!/usr/bin/env python3

import argparse
import csv
import subprocess
import sys
import tempfile

parser = argparse.ArgumentParser(description='Convert Postbank account statement pdf files to a single csv file.')
parser.add_argument('pdf_files', metavar='pdf_file', type=argparse.FileType('r'), nargs='+', help='pdf files to convert')
args = parser.parse_args()


def main():
    statements = []

    for file in args.pdf_files:
        #print(file.name)
        statements += parse_statements_from_file(str(file.name))

    write_statements_as_csv(statements)

# get text between 2 keywords
def get_between (text, first_needle, second_needle):
    
    if first_needle is None:
        first_needle_idx = 0
        first_needle = "" # needed to messure the len of 0
    else:
        first_needle_idx = None
        try:
            first_needle_idx = text.index(first_needle)
        except ValueError:
            first_needle_idx = text.index(first_needle.replace("-", ""))

    if second_needle is None:
        second_needle_idx = -1
    else:
        second_needle_idx = None
        try:
            second_needle_idx = text.index(second_needle)
        except ValueError:
            second_needle_idx = text.index(second_needle.replace("-", ""))

    return text[first_needle_idx + len(first_needle):second_needle_idx].strip()

def sub_parse_other(statement):

    if statement['Type'] == "SDD Lastschr" or statement['Type'] == "Kartenzahlung":
        statement['Empfaenger'] = get_between(statement['other'], None,"Referenz")
        einreicher_verwendungszweck = get_between(statement['other'],"Einreicher-ID",None)
        statement['Verwendungszweck'] = get_between(einreicher_verwendungszweck, " ", None) # alles nach dem ersten leerzeichen
        statement['Einreicher-id'] = get_between(einreicher_verwendungszweck, None, " ") # nur bis zum ersten leerzeichen
        statement['Referenz'] = get_between(statement['other'],"Referenz","Mandat")
        statement['Mandat'] = get_between(statement['other'],"Mandat","Einreicher-ID")

    elif statement['Type'] == "Gutschr.SEPA" or statement['Type'] == "D Gut SEPA" or statement['Type'] == "Echtzeitüberw Gutschrift" :
        statement['Empfaenger'] = get_between(statement['other'], None,"Referenz")
        statement['Verwendungszweck'] = get_between(statement['other'],"Verwendungszweck",None)
        statement['Referenz'] = get_between(statement['other'],"Referenz","Verwendungszweck")

    else:
        statement['Verwendungszweck'] = statement['other']

    return statement

def parse_statements_from_file(pdf_filename):
    txt_filename = next(tempfile._get_candidate_names()) + ".txt"

    bashCommand = f"pdftotext -layout -x 70 -y 100 -W 500 -H 700 {pdf_filename} {txt_filename}"
    process = subprocess.Popen(bashCommand.split(), stdout=subprocess.PIPE)
    output, error = process.communicate()

    with open(txt_filename, 'r') as f:
        filecontent = f.read()

    in_toprow_area = False
    in_statement_area = False
    in_statement = False
    space_needed = " " #Leerzeichen bei Zeilenumbruch aktivien (wird bei "-" am Zeilenende weggelassen)
    last_line_token = []
    statements = []
    statement = {}

    for line in filecontent.splitlines():
        line_token = [token.strip() for token in line.split()]
        #print(line_token)

        if line_token == ['Buchung/Wert', 'Vorgang/Buchungsinformation', 'Soll', 'Haben']:
            in_toprow_area = False
            in_statement_area = True
            last_line_token = line_token
            continue

        if line_token[:4] == ['Auszug', 'Jahr', 'Seite', 'von']:
            in_toprow_area = True
            in_statement_area = False
            continue

        if line_token == ['Kontonummer', 'BLZ', 'Summe', 'Zahlungseingänge']:
            in_statement_area = False
            break

        if in_toprow_area:
            file_number = int(line_token[0])
            file_year = int(line_token[1])
            in_toprow_area = False

        if in_statement_area:
            #print(line_token)
            if line_token and not last_line_token: # if non empty and last one was empty
                in_statement = True
                statement = {}
                statement_first_line = True

            if not line_token:
                in_statement = False
                if statement: # if dict not empty
                    statements.append(sub_parse_other(statement))
                    #print("new statement written", statement)
                    statement = {}

            if in_statement:
                if statement_first_line:
                    try:
                        statement['Value'] = ''.join(line_token[-2:]).replace('.', '').replace(',', '.')
                        statement['Value'] = statement['Value'].replace(".",",") # Beträge wieder in Komma Schreibweise zurückführen
                        # print (statement['value'])
                    except ValueError:
                        in_statement = False
                        continue
                    date_day, date_month = line_token[0].split('/')[0][:-1].split('.')
                    if file_number == 1 and date_month not in ['12', '01']:
                        Exception(f"There is a statement from something else than Dec or Jan in the first document of {file_year}!")
                    elif file_number == 1 and date_month == '12':
                        date_year = file_year - 1
                    else:
                        date_year = file_year

                    statement['Date'] = f"{date_day}.{date_month}.{str(date_year)}"
                    #print (line_token[1:-2], line_token)

                    statement['Type'] = ' '.join(line_token[1:-2])
                    statement['other'] = ""
                    statement_first_line = False
                else:
                    #print(line_token, line_token[-1][-1])
                    if line_token[-1][-1] == "-":  # letzes Zeichen ist "-"
                        line_token[-1] = line_token[-1][:-1] # Zeilenumbruchs "-" entfernen
                        space_needed = "" # kein leerzeichen
                    #print(line_token, line_token[-1][-1])

                    statement['other'] += ' '.join(line_token)
                    statement['other'] += space_needed
                    space_needed = " "
        #print(line_token)
        last_line_token = line_token

    bashCommand = f"rm {txt_filename}"
    process = subprocess.Popen(bashCommand.split(), stdout=subprocess.PIPE)
    output, error = process.communicate()
    
    #print(statements)

    return statements


def write_statements_as_csv(statements):

    fieldnames = ['Date', 'Type', 'Value', 'Empfaenger','Verwendungszweck', 'Einreicher-id', 'Referenz', 'Mandat' ,'other']
    writer = csv.DictWriter(sys.stdout, fieldnames=fieldnames, restval='',extrasaction='ignore')
    writer.writeheader()
    for statement in statements:
        writer.writerow(statement)


if __name__ == "__main__":
    main()
