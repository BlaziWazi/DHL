"""
This module contains functions required to gain the declaration and air waybill values form a certain Excel file.
"""


# Built-in modules:
# https://docs.python.org/3/py-modindex.html
from io import StringIO
from csv import reader
from re import sub

# Third-party modules:
from xlsx2csv import Xlsx2csv


# Module variables:

# Cat art by Marcin Glinski:
ascii_cat = """
⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⠟⠉⡉⠙⢿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿
⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⠟⠀⣼⠙⡆⠈⢻⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿
⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⠟⠀⣼⠃⡆⢻⡆⠀⢻⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⡿⣿⣿⣿⣿⣿⣿⣿⣿⣿
⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⠀⣰⡟⢰⣷⡘⣿⡄⠈⠿⠿⠟⠛⠛⠛⠛⠛⠿⢿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⠛⠋⢀⣤⠀⢸⣿⣿⣿⣿⣿⣿⣿⣿
⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⠃⠀⣿⠁⣛⣉⣅⣿⣧⣤⣤⣴⣾⣿⣿⣿⣿⣿⣷⣦⣤⣀⠉⠙⠻⢿⣿⣿⠟⠉⣤⡶⠛⢛⣿⠀⢸⣿⣿⣿⣿⣿⣿⣿⣿
⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⡟⠀⣾⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣶⣤⡀⠉⠡⣴⠟⣡⣴⣿⢸⡏⠀⣼⣿⣿⣿⣿⣿⣿⣿⣿
⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⡟⠀⣰⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣶⣄⠉⣄⠙⠿⡏⢸⡇⢠⣿⣿⣿⣿⣿⣿⣿⣿⣿
⣿⣿⣿⣿⣿⣿⣿⣿⣿⡿⠁⢰⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣷⣶⣤⡾⠀⢸⣿⣿⣿⣿⣿⣿⣿⣿⣿
⣿⣿⣿⣿⣿⣿⣿⣿⣿⠇⢰⣿⣿⣿⣿⣿⣿⡟⠉⠉⠉⠉⠉⠛⠻⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⡇⠀⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿
⣿⣿⣿⣿⣿⣿⣿⣿⡇⢠⣿⣿⣿⣿⣿⣿⣿⣿⣄⠈⠻⠀⢰⣦⠀⠘⣿⣿⣿⣿⣿⣿⡿⠛⠋⠉⠉⠉⠙⠛⢿⣿⣿⣿⣿⡇⠰⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿
⣿⣿⣿⣿⣿⣿⣿⡟⢀⣾⣿⣿⣿⣿⣿⣿⣿⣿⣿⣷⣶⣤⣤⣤⣤⠀⢿⣿⣿⣿⡿⠋⠀⠀⠺⠇⠀⠟⠂⠀⢀⣿⣿⣿⣿⡇⢸⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿
⣿⣿⣿⣿⣿⣿⣿⠁⣸⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣤⣤⣽⣿⣿⠇⢰⣾⣷⣶⣤⣤⣤⣴⣿⣿⣿⣿⣿⣿⡇⠈⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿
⣿⣿⣿⣿⣿⣿⡇⢠⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⡇⠀⣀⡈⠉⠓⢺⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⡇⢸⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿
⣿⣿⣿⣿⣿⡟⠀⣸⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣟⠉⢻⡉⠻⣿⣿⡄⠙⠿⠋⣀⣼⣿⣿⢿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⠁⢸⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿
⣿⣿⣿⣿⣿⠇⢠⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣄⠈⠻⠷⠬⣿⡿⠀⢀⣾⡿⢿⣟⠙⣦⠈⢻⣿⣿⣿⣿⣿⣿⣿⡿⠀⢸⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿
⣿⣿⣿⣿⡟⠀⣼⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣷⣶⣤⣄⣀⣠⣤⡈⠉⠓⠚⠛⠛⠉⢀⣼⣿⣿⣿⣿⣿⣿⣿⡇⢀⣾⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿
⣿⣿⣿⣿⠇⢠⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣷⡌⠙⠿⣿⣿⣿⠏⢶⣶⣶⣶⣶⣿⣿⣿⣿⣿⣿⣿⣿⣿⠃⢸⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿
⣿⣿⣿⡟⠀⣾⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣷⣤⣄⣀⣉⣁⣤⣾⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⠏⠀⢸⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿
⣿⣿⣿⠁⢠⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣧⡀⠘⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿
⣿⣿⡿⠀⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣧⡀⠘⢿⣿⣿⣿⣿⣿⣿⣿⣿⣿
⣿⣿⠁⣰⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣷⡄⠈⢻⣿⣿⣿⣿⣿⣿⣿⣿
⣿⡟⠀⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⡄⠀⢻⣿⣿⣿⣿⣿⣿⣿
⣿⠃⢸⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⡄⠈⣿⣿⣿⣿⣿⣿⣿
⣿⠀⣾⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⠟⢉⡉⠉⣹⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣧⠀⢸⣿⣿⣿⣿⣿⣿
⣿⠀⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⡟⠛⠛⠟⢁⣴⡟⢃⣴⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⡇⠀⣿⣿⣿⣿⣿⣿
⡟⠀⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⠟⢁⡍⢠⣾⣦⣤⡟⠋⢰⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣷⠀⢿⣿⣿⣿⣿⣿
⠇⢸⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⠟⢿⣿⣿⣿⣿⣿⣿⣿⣿⣿⠏⠀⣾⣿⣿⣿⣿⣿⣿⣧⡄⢨⣿⣿⣿⣿⣿⢁⣄⠹⣿⣿⣿⣿⣿⣿⣿⠀⠀⠁⠀⠀⠈⢻
⠀⢸⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⠀⢸⣿⣿⣿⣿⣿⣿⣿⡿⠃⢀⣼⣿⣿⣿⣿⣿⣿⣿⡟⢀⣼⣿⣿⣿⣿⠇⣼⣿⠀⣿⣿⣿⣿⣿⣿⣿⠀⠀⣾⣿⣿⠀⢸
⠀⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⠀⢸⣿⣿⣿⣿⣿⡿⠋⢀⣴⣿⣿⣿⣿⣿⣿⣿⡿⠋⢀⣼⣿⣿⣿⣿⣿⠀⣿⡏⢸⣿⣿⣿⣿⣿⣿⣿⠀⠀⣿⣿⡿⠀⣸
⠀⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⠀⢸⣿⣿⡿⠟⠉⣠⣴⣿⣿⣿⣿⣿⣿⣿⠟⠁⣀⣴⣿⣿⣿⣿⡟⠛⢁⣀⣿⣃⠀⠹⢿⣿⣿⣿⣿⣿⠀⠀⣿⣿⠁⢀⣿
⠀⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⡿⠀⠀⠋⢁⣠⣴⣿⣿⣿⣿⣿⣿⣿⣿⡿⠁⣴⣾⣿⣿⣿⣿⣿⣿⠁⢰⣿⣿⣿⣿⣷⣄⠀⢿⣿⣿⣿⣿⠀⠀⣿⡏⠀⣼⣿
⠀⢿⣿⣿⣿⣿⣿⣿⣿⡿⠛⠁⠀⣤⣶⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⠋⢀⣾⣿⣿⣿⣿⣿⣿⣿⡇⠀⢸⣿⣿⣿⣿⣿⡿⠀⣼⣿⣿⣿⣿⠀⠀⣿⠀⢰⣿⣿
⠀⠸⣿⣿⣿⣿⣿⣿⣇⠀⣠⣶⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⠟⢁⣴⣿⣿⣿⣿⣿⣿⣿⣿⣿⣷⠀⢸⣿⣿⣿⣿⣿⠇⢀⣿⣿⣿⣿⣿⠀⠀⡇⠀⣿⣿⣿
⡇⠀⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⠟⢁⣰⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⡄⢸⣿⣿⣿⣿⡟⠀⢸⣿⣿⣿⣿⣿⠀⠀⠀⣰⣿⣿⣿
⡇⠀⠘⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⡿⠟⢁⣴⣾⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⡇⠸⣿⣿⣿⣿⣧⠀⠸⣿⣿⣿⣿⣿⠀⠀⠀⣿⣿⣿⣿
⡇⠀⠀⠘⢿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⠟⠉⣀⣴⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⠁⠀⣿⣿⣿⣿⣿⡄⠀⢻⣿⣿⣿⡟⠀⠀⢰⣿⣿⣿⣿
⠇⢰⣧⡀⠀⠙⠻⠿⠿⠿⠿⠿⠿⠛⠋⠀⣠⣾⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⡇⠀⠛⣿⣿⣿⣿⡇⠀⢸⣿⣿⣿⠃⠀⠀⣸⣿⣿⣿⣿
⣷⣶⣿⣿⣿⣿⣿⣶⣶⣶⣶⣶⣾⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣷⣶⣶⣿⣿⣿⣿⣿⣶⣶⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿
"""

# Module functions:

# Gets the cat:
def get_cat():
    return ascii_cat.strip().encode("utf-8")


# Checks if the file is a valid Excel file:
def check_excel_file(filename):
    if filename.endswith(".xlsx") and filename[0:2] != "~$":
        return True
    else:
        return False


# Creates a new text file:
def create_text_file(excel_file, content):
    if content:
        f = open(excel_file[0:-4] + "txt", "w")
        f.write(content.strip())
    else:
        f = open(excel_file[0:-4] + "txt", "wb")
        f.write(get_cat())
    f.close()


# Gains the declaration and air waybill values form a certain Excel file:
def get_declawb(excel_file):
    content = ""
    # Converts the Excel file to a CSV file:
    csv = StringIO()
    Xlsx2csv(excel_file, skip_empty_lines=True).convert(csv)
    csv.seek(0)
    # Finds the indexes of the necesarry CSV columns (MRN and PAVADDOK):
    csv_reader = reader(csv)
    i_mrn = None
    i_pavaddok = None
    for row in csv_reader:
        # Finds the MRN column:
        if "MRN" in row:
            i_mrn = row.index("MRN")
            # Finds the PAVADDOK column:
            i = 0
            for column in row:
                if len(column) >= 8:
                    if column[0:8].upper() == "PAVADDOK":
                        i_pavaddok = i
                i += 1
            break
    # Filters the necesarry CSV columns (MRN and PAVADDOK), filters the necessary 10-digit numbers from PAVADDOK:
    csv.seek(0)
    if i_mrn and i_pavaddok:
        for row in csv_reader:
            mrn = row[i_mrn]
            pavaddok = row[i_pavaddok]
            if mrn != "" and pavaddok != "":
                if mrn != "MRN" and pavaddok[0:8].upper() != "PAVADDOK":
                    # Filters the necessary 10-digit number from PAVADDOK:
                    pavaddok = sub("[-,]", "", pavaddok)
                    pavaddok = pavaddok.split(" ")
                    for code in pavaddok:
                        if len(code) == 10:
                            pavaddok = code
                            break
                    # Adds the results to the content variable:
                    content += pavaddok + "\n" + mrn + "\n"
    return content
