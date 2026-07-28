import pandas as pd
import pdfplumber
import re
from pathlib import Path
import os
import csv

#read in excel data
path_excel_raw = Path(r"path_to_cdcp_file") #FILL IN

path_excel = path_excel_raw.as_posix()

df = pd.read_excel(path_excel,sheet_name="DH")

df["2026 PT Fee"] = pd.to_numeric(df["2026 PT Fee"],errors="coerce").round(2)

df_select = df[["PT", "Specialty", "Procedure Code", "2026 PT Fee"]]
df_filter = df_select[~df_select["Procedure Code"].isin(["SK", "NL", "NS", "QC"])] #this just filters out the places that aren't pdfs for the sake of comparison

#iterate through folder of pdfs and process all of them
p = r"path_to_folder_of_fee_guides" #FILL IN

#df that will be filled in at the end of each loop
extracted_info = {'PT': [],
                  'Specialty': [], 
                  'Procedure Code':[],
                  'Fee': []}

extracted_data = pd.DataFrame(extracted_info)

#loops each pdf within the folder to extract info 
for e in os.scandir(p):
    #open pdf file
    path_raw_pdf = Path(e)

    filename = path_raw_pdf.stem

    parts = filename.split()

    location = parts[0] #collecting location for data organization. Example: file is called "ON DH Fee Guide 2026", this collects "ON"

    specialty = parts[1] #collecting specialty for data organization Example: file is called "ON DH Fee Guide 2026", this collects "DH"

    path_pdf = path_raw_pdf.as_posix()

    text = ""

    #removes all text from the pdf
    with pdfplumber.open(path_pdf) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:  # handles pages with no text
                text += page_text + "\n"

    #extract all numbers from text
    numbers = re.findall(r'\d+\.?\d*', text)

    #handle case with range of costs. ex: "$44.66 to $89.32". we always take the higher number as worst case scenario
    #the for loop ignores each number that contains a decimal and the next number in line contains a decimal, all the rest are accepted.
    numbers_filtered1 = []

    for i in range(len(numbers) - 1):
        if not (("." in numbers[i]) and ("." in numbers[i+1])):
            numbers_filtered1.append(numbers[i])

    #remove 00100 and single character numbers
    numbers_filtered = []

    for number in numbers_filtered1:
        if ("00100" not in number) and (len(number) > 2):
            numbers_filtered.append(number)

    #extract pairs of code and cost
    #the for loop finds checks if i is a five digit code and if i+1 contains a decimal, if both are true it pairs the numbers up in the df
    pairs_data = {'PT': [], 
                  'Fee': []}

    pairs = pd.DataFrame(pairs_data)

    for i in range(len(numbers_filtered) - 1):
        if (len(numbers_filtered[i]) == 5) and ("." in numbers_filtered[i + 1]):
            pairs.loc[len(pairs)] = [numbers_filtered[i], numbers_filtered[i + 1]]

    #sometimes the pdf will reference previous codes in the text between the number and the fee, so we will remove the duplicate codes and 
    #these outlier cases can be manually filled in while the rest should work properly 
    pairs_filtered = pd.DataFrame(columns=["PT", "Fee"])

    for _, row in pairs.iterrows():
        if (row["PT"] not in pairs_filtered["PT"].values) and ("." not in row["PT"]):
            pairs_filtered.loc[len(pairs_filtered)] = row

    pairs_filtered["PT"] = pairs_filtered["PT"].astype("int64")
    pairs_filtered["Fee"] = pairs_filtered["Fee"].astype("float64")

    for i, pt in enumerate(pairs_filtered["PT"]):
        if str(pt).startswith("1"):
            pairs_filtered = pairs_filtered.iloc[i:].reset_index(drop=True)
            break

    pairs_filtered["Specialty"] = specialty
    pairs_filtered["Procedure Code"] = location
    extracted_fees = pairs_filtered[['PT', 'Specialty', 'Procedure Code', "Fee"]]

    extracted_data = pd.concat([extracted_data, extracted_fees],ignore_index=True)

#see the results
merge_df = df_filter.merge(extracted_data, on= ['PT', 'Specialty', 'Procedure Code'], how='left')

inconsistencies = merge_df.loc[merge_df['2026 PT Fee'] != merge_df['Fee']]

extracted_data.to_csv(r"path_to_folder_of_choice\extracted_data.csv", index=False) #FILL IN
inconsistencies.to_csv(r"path_to_folder_of_choice\inconsistencies.csv",index=False) #FILL IN
