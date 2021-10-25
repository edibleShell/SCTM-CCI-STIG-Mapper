#  Import Standard Libraries 
from colorama import init, Fore, Back, Style
import fnmatch
import os
from os import system, name
import pathlib


#  Import Third Party Libraries
from datetime import datetime, date
import openpyxl
import pandas as pd
import xlrd3 as xlrd


#  Variables
csv_pattern = "*.csv"  #  CSV File type variable
xlsx_pattern = "*.xlsx"  #  Excel File type variable
quotation_input = '"*"'  #  Quotation Mark variable
stig_export = []  #  A list to store all STIGs for comparison
today = date.today()  #  Obtain todays date for date variable
date = today.strftime("%Y-%m-%d")  #  Date Function Format used to title the finished report
results_dir = "Results"  #  Results will be used when creating the Results directory
stig_dir = "STIGs"  #  Used to locate the STIGs directory
stig_dir_casefold = stig_dir.casefold()  #  Used to remove case sensitivity for the STIG folder


#  Script Main
init(autoreset=True)  #  autoresets the font color when run on Windows OS


#  Retrieve the path to the current working directory for the STIG exports
path = pathlib.Path().absolute()  #  Store the absolute path to the working directory of the script
results = os.path.join(path, results_dir)  #  Append results_dir to the absolute path of the working DIR
stig_path = os.path.join(path, stig_dir)  #  Store the path to the STIGs directory


try:
    os.mkdir(results)  #  Create the Results directory
except OSError as error:
    print("\n" + Fore.RED + f"{error}" + "\n\n")  #  If the results directory already exists the script will print the error to the scrren and continue


#  Enumerate STIG exports and retrieve column names
for file in os.listdir(stig_path):
    if file.endswith(".csv"):
        stig_export.append(file)
    elif file.endswith(".xlsx"):
        stig_export.append(file)


print(Fore.GREEN + "Your Company Name Here\n" + Style.BRIGHT )
print(Fore.RED + "Author: " + Fore.WHITE + "Richard Fontaine\n" 
    + Fore.RED + "Version: " + Fore.WHITE + "1.1\n"
    + Fore.RED + "Published: "+ Fore.WHITE +"2021-04-27")


data_file = input("\n\nCNS 1253 Security Controls Assessment Procedures:> ")


if fnmatch.fnmatch(data_file, quotation_input):  #  Check file for Quotation marks
    data_file = data_file.strip('""')  #  Remove Quotation marks


if fnmatch.fnmatch(data_file, xlsx_pattern):  #  Determine if the file is a csv or xlsx file type
    data_file = pd.read_excel(data_file, header=2, usecols=[2,3,4,5,6,7])  #  Read xlsx file type
else:
    data_file = pd.read_csv(data_file, header=2, usecols=[2,3,4,5,6,7])  #  Read csv file type


#  Begin stig list import to retrieve vuln IDs
for stig in stig_export:
    data_file_stig = stig_dir_casefold + "\\" + stig
    if fnmatch.fnmatch(data_file_stig, xlsx_pattern):  #  Determine if the file is a csv or xlsx file type
        data_file_stig = pd.read_excel(data_file_stig, header=1)  #  Read xlsx file type
        column_name = stig.replace(".xlsx", "")
    else:
        data_file_stig = pd.read_csv(data_file_stig, header=1)   #  Read csv file type
        column_name = stig.replace(".csv", "")


    data_file[column_name] = ""  #  Create column for the specific STIG in the CNSS 1253 data file
    data_file[column_name] = data_file[column_name].fillna("N/A")  #  Fill blank cells with the N/A value
    

    #  Split comma sperated values into new rows.
    data_file_stig = (data_file_stig.set_index(["Vuln ID"])
        .stack()
        .str.split(",", expand=True)
        .stack()
        .unstack(-2)
        .reset_index(-1, drop=True)
        .reset_index()
    )


    #  Compare each CCI to a CCI in the CNSS 1253 if match add Vuln ID.
    for index, row in data_file_stig.iterrows():
        vuln_id = row["Vuln ID"]  #  Store the current rows vuln ID
        cci_id = row["CCI"]  #  Store the current rows cci ID
        for index, row2 in data_file.iterrows():
            if row2["CCI"] == cci_id:  #  Check to see if current STIG cci ID is equal to CNSS 1253 cci ID
                if row2[column_name] != "N/A":  #  If N/A is not present then a Vuln ID must exist in the current cell
                    existing_vuln_id = row2[column_name]  #  Store current Vuln ID in existing_vuln_id variable
                    data_file.loc[data_file["CCI"] == cci_id, column_name] = existing_vuln_id + "," + vuln_id  #  comma seperate existing_vuln_id and vuln ID from STIG file current row
                else:
                    data_file.loc[data_file["CCI"] == cci_id, column_name] = vuln_id  #  N/A exists replace with vuln ID since the cci ID matches this row


    data_file[column_name] = data_file[column_name].str.lstrip(",")


    print(Fore.YELLOW + f" -  {column_name}:>" + Fore.GREEN + " Processing Complete")


excel_writer = pd.ExcelWriter(f"{results}\\{date}_CNSS_1253_STIGs.xlsx")  # Title Report and store results in writer variable
data_file.to_excel(excel_writer, sheet_name = "SecurityControls", index = False)  # Write data_file dataframe to sheet Compliance Overview
excel_writer.save()  # Write excel file to working directory


print(f"\n" + Fore.YELLOW + "Results saved to:> " + Fore.CYAN + f"{results}\\{date}_CNS_1253_STIGs.xlsx")


input("\nPress [Enter] key to exit:> ")
