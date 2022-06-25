# Built-in modules:
# https://docs.python.org/3/py-modindex.html
from io import StringIO
from csv import reader
from re import sub
import os

# Third-party modules:
from xlsx2csv import Xlsx2csv


# Module variables:
desktop_path = os.path.join(os.getenv("SystemDrive"), os.environ["HOMEPATH"], "Desktop")
eds_folder = "EDS"


# Module functions:

# Creates the EDS folder:
def create_eds_folder():
    if not os.path.exists(os.path.join(desktop_path, eds_folder)):
        os.makedirs(os.path.join(desktop_path, eds_folder))


# Gets all Excel files in a directory (also the temperary ones):
def get_excel_files():
    excel_files = []
    for file in os.listdir(os.path.join(desktop_path, eds_folder)):
        # Ignores the temporary Excel files (the ones that are getting edited):
        if file.endswith(".xlsx") and file[0:2] != "~$":
            excel_files.append(os.path.join(desktop_path, eds_folder, file))
    return excel_files


# Gets all 10-digit numbers from the provided Excel files and writes the results to the text file:
def get_numbers():
    # Opens the text file:
    f = open(os.path.join(desktop_path, eds_folder, "result.txt"), "w")
    # Checks if any Excel files exist:
    if not get_excel_files():
        # Writes a line in the text file:
        f.write("Please add Excel files to \"" + os.path.join(desktop_path, eds_folder) + "\" and run the code again!")
    # Checks each Excel file:
    for excel_file in get_excel_files():
        # Writes the Excel file name in the text file:
        f.write("Results from \"" + excel_file + "\":\n")
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
            i = 0
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
                        # Writes the results to the text file:
                        f.write(pavaddok + "\n" + mrn + "\n")
                i += 1
        # Writes a new line in the text file:
        f.write("\n")
    # Closes the text file:
    f.close()


# Runs the code:
create_eds_folder()
get_numbers()
