#!/usr/bin/env python3

import argparse
import csv
import subprocess
import sys
import tempfile
import re

parser = argparse.ArgumentParser(description='Convert Postbank account statement pdf files to a single csv file.')
parser.add_argument('pdf_files', metavar='pdf_file', type=argparse.FileType('r'), nargs='+', help='pdf files to convert')
args = parser.parse_args()


def main():
    statements = []

    # class to collect data for sorting file names by date in name
    class myFile:
        def __init__(self, name, sortable):
            self.name = name
            self.sortable = sortable
        def __repr__(self): # steuert ausgabe z.b. mittels print
            return str(self.name + " " + self.sortable)  

    files = []

    # get about date in name of Auszug
    for file in args.pdf_files:
        
        #print(file.name)
        # works with filenames with a date in it like this:
        x = re.search(r"(\d{2})-(\d{2})-(\d{4})", file.name) # "PB_Kontoauszug_KtoNr_012334455_05-12-2022_104523.pdf"
        if x is None:
            x = re.search(r"((\d{4})-\d{2})-(\d{2})", file.name) # "Konto_012334455_2023-01-01_PB34553393.pdf"
            if x is None:
                sortable = 0 # casse no date found in filename
            else:    
                sortable =  x.group(1) + x.group(2) + x.group(3)
        else:
            sortable = x.group(3) + x.group(2) + x.group(1)

        #print(sortable)
        files.append(myFile(file.name,sortable))
    
    def get_sortable(myfile):
        return myfile.sortable

    files.sort(key=get_sortable, reverse=False)
    #print(files)

    for file in files:
        
        statements += parse_statements_from_file(str(file.name))
        #print(statements)

    write_statements_as_csv(statements)

# get text between 2 keywords
def get_between (text, first_needle, second_needle):
    
    if first_needle is None:
        first_needle_idx = 0
        first_needle = "" # needed to messure the len of 0
    else:
        first_needle_idx = None
        try:
            first_needle_idx = text.lower().index(first_needle.lower())
            #print("FIRST:",first_needle,first_needle_idx, "TXT:", text)
        except ValueError:
            first_needle_idx = text.lower().index(first_needle.lower().replace("-", ""))

    if second_needle is None:
        second_needle_idx = -1
    else:
        second_needle_idx = None
        try:
            second_needle_idx = text.lower().index(second_needle.lower(), first_needle_idx + len(first_needle) + 1) #suche ab nach dem 1. Zeichen (=Leerzeichen) nach fist needle
        except ValueError:
            second_needle_idx = text.lower().index(second_needle.lower().replace("-", ""),first_needle_idx + len(first_needle) + 1)
    
    mytxt = text[first_needle_idx + len(first_needle):second_needle_idx].strip()
    #print(first_needle_idx ,second_needle_idx ,">> ",mytxt )

    return mytxt

# Liste mit regelmäßig vorkommenden Namen, z.B. in Daueraufträgen, um diese Namen vom Begin des Verwendungszweckes zu entfernen und in "Empfänger" zu übernehmen
names_list = ["MIETERKONTO", "Advanzia Bank S.A", "Amazon.de KartenService", "Rundfunk ARD, ZDF, DRadio"] # liste der empänger von Daueraufträge, für die Namenstrennung aus Verwendungszweck
key_words0 = ["Referenz", "Mandat", "Einreicher-ID","Verwendungszweck"]
keys = []
# keys.append({"word":"Referenz", "extract":"next"})
# keys.append({"word":"Mandat", "extract":"next"})
# keys.append({"word":"Einreicher-ID", "extract":"next"})
# keys.append({"word":"Verwendungszweck", "extract":"next"})

class key:
    def __init__(self, word, field_name, extract):
        self.word = word
        self.field_name = field_name
        self.extract = extract
    def __repr__(self): # steuert ausgabe z.b. mittels print
        return str(self.word + " " + self.extract)  

keys.append(key("Referenz","Empfaenger","beforeKey"))
keys.append(key("Referenz","Referenz","untilNextKey"))
keys.append(key("Mandat","Mandat","untilNextKey"))
keys.append(key("Einreicher-ID","Einreicher-ID","naechsteWort"))
keys.append(key("Einreicher-ID","Verwendungszweck","restOhneErstesWort")) # Rest der Buchung, ohne erstes Wort
#keys.append(key("EinreicherID","Einreicher-ID","naechsteWort")) # manchmal fehlt der Bindestrich
keys.append(key("Verwendungszweck","Verwendungszweck","untilNextKey"))

# teilt die in der Buchung enthaltenen Infos in 'Zusammen' in weitere Spalten auf 
def sub_parse_zusammen2(statement):
    
    # in buchung vorkommende keys ermitteln 
    contained_keys = []
    for key in keys:
        if statement['Zusammen'].replace("-","").find(key.word.replace("-", "")) > -1: # search but irnore "-"
            contained_keys.append(key)
    #print("CON: ",len(contained_keys),contained_keys)

    # wenn keine keys gefunden, zumindest namen aus names_list separieren
    if len(contained_keys) < 1:
        # alles in Verwendungszweck übernehmen
        statement['Verwendungszweck'] = statement['Zusammen']
        #print("VER:",statement['Verwendungszweck'])

        for name in names_list:
            if statement['Zusammen'].lower().find(name.lower()) == 0:
                statement['Empfaenger'] = name
                statement['Verwendungszweck'] = get_between(statement['Zusammen'],name, None)
                #print("VERW:", statement['Verwendungszweck'])
            
    else:
        # wenn nur 1 key gefunden, vor Schlüßel immer als Name und danach auch als Verwendungszweck verwenden
        if len(contained_keys) == 1 or len(contained_keys) == 2 :
            statement['Empfaenger'] = get_between(statement['Zusammen'],None ,contained_keys[0].word)
            statement['Verwendungszweck'] = get_between(statement['Zusammen'], contained_keys[0].word, None)


        #vorkommende keys zum separieren verwenden
        for idx, key in enumerate(contained_keys):
            if key.extract == "beforeKey":
                statement[key.field_name] = get_between(statement['Zusammen'],None ,key.word)
            if key.extract == "naechsteWort":
                statement[key.field_name] = get_between(statement['Zusammen'],key.word ," ")    
            if key.extract == "restOhneErstesWort": # Rest der Buchung, ohne erstes Wort
                rest = get_between(statement['Zusammen'],key.word, None)
                statement[key.field_name] = get_between(rest," ", None)    
                            
            if key.extract == "untilNextKey":
                try:
                    statement[key.field_name] = get_between(statement['Zusammen'],key.word ,contained_keys[idx+1].word)
                    #print(statement[key.field_name])

                except IndexError:
                    #print(key.field_name, "IndexError: ",key.word, len(contained_keys), idx, contained_keys  )
                    statement[key.field_name] = get_between(statement['Zusammen'],key.word, None)
                    #print(statement[key.field_name])

    if statement['Typ'] == "Zinsen/Entg.":
        #print("ZIN:",statement)
        statement['Verwendungszweck'] = statement['Typ']
        statement['Empfaenger'] = "Postbank"
    
    if statement['Typ'] == "Rechnungsabschluss -" and statement['Betrag'] == "sieheHinweis":
        statement['Betrag'] = "" 
        

    return statement      

# OLD Version - not in use
def sub_parse_zusammen(statement):

    if statement['Typ'] == "SDD Lastschr" or statement['Typ'] == "Kartenzahlung":
        statement['Empfaenger'] = get_between(statement['Zusammen'], None,"Referenz")
        einreicher_verwendungszweck = get_between(statement['Zusammen'],"Einreicher-ID",None)
        statement['Verwendungszweck'] = get_between(einreicher_verwendungszweck, " ", None) # alles nach dem ersten leerzeichen
        statement['Einreicher-id'] = get_between(einreicher_verwendungszweck, None, " ") # nur bis zum ersten leerzeichen
        statement['Referenz'] = get_between(statement['Zusammen'],"Referenz","Mandat")
        statement['Mandat'] = get_between(statement['Zusammen'],"Mandat","Einreicher-ID")

    elif statement['Typ'] == "Gutschr.SEPA" or statement['Typ'] == "D Gut SEPA" or statement['Typ'] == "Echtzeitüberw Gutschrift" :
        statement['Empfaenger'] = get_between(statement['Zusammen'], None,"Referenz")
        statement['Verwendungszweck'] = get_between(statement['Zusammen'],"Verwendungszweck",None)
        statement['Referenz'] = get_between(statement['Zusammen'],"Referenz","Verwendungszweck")
    
    # ausgewählte Empfänger aus dauerauftrag_list erkennen und abtrennen
    elif statement['Typ'].find("Dauerauftrag") == 0 or statement['Typ'] == "SEPA Überw. Einzel":
        for name in names_list:
            if statement['Zusammen'].find(name) == 0:
                statement['Empfaenger'] = name
                statement['Verwendungszweck'] = get_between(statement['Zusammen'], name, None)
    else:
        statement['Verwendungszweck'] = statement['Zusammen']

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
                    statements.append(sub_parse_zusammen2(statement))
                    #print("new statement written", statement)
                    statement = {}

            if in_statement:
                if statement_first_line:
                    try:
                        statement['Betrag'] = ''.join(line_token[-2:]).replace('.', '').replace(',', '.')
                        statement['Betrag'] = statement['Betrag'].replace(".",",") # Beträge wieder in Komma Schreibweise zurückführen
                        # print (statement['Betrag'])
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

                    statement['Buchungsdatum'] = f"{date_day}.{date_month}.{str(date_year)}"
                    #print (line_token[1:-2], line_token)

                    statement['Typ'] = ' '.join(line_token[1:-2])
                    statement['Zusammen'] = ""
                    statement_first_line = False
                else:
                    #print(line_token, line_token[-1][-1])
                    if line_token[-1][-1] == "-":  # letzes Zeichen ist "-"
                        line_token[-1] = line_token[-1][:-1] # Zeilenumbruchs "-" entfernen
                        space_needed = "" # kein leerzeichen
                    #print(line_token, line_token[-1][-1])

                    statement['Zusammen'] += ' '.join(line_token)
                    statement['Zusammen'] += space_needed
                    space_needed = " "
        #print(line_token)
        last_line_token = line_token

    bashCommand = f"rm {txt_filename}"
    process = subprocess.Popen(bashCommand.split(), stdout=subprocess.PIPE)
    output, error = process.communicate()
    
    #print(statements)

    return statements


def write_statements_as_csv(statements):

    fieldnames = ['Buchungsdatum','Empfaenger' ,'Verwendungszweck', 'Notiz', 'Betrag', 'Typ', 'Einreicher-ID', 'Referenz', 'Mandat' ,'Zusammen']
    writer = csv.DictWriter(sys.stdout, fieldnames=fieldnames, restval='',extrasaction='ignore')
    writer.writeheader()

    statementsReversed = reversed(statements) # Reihenfolge der Buchungen umdehen

    for statement in statementsReversed:
        writer.writerow(statement)


if __name__ == "__main__":
    main()
