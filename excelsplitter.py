import pandas as pd

path = input("Path of the file: ")
path_export = input('Export path and file name, for example: C:\Downloads\\file_name: ')
sheet = input("Sheet number: ")
headerNr = int(input("Header row number: "))
includeHeader = input("Include header (Y or N): ")
nr = input("Split by: ")

headerList = []

a = 0

while(a < headerNr):
    headerList.append(a)
    a = a + 1

if (includeHeader == 'Y' or includeHeader == "y" or includeHeader == 'Yes' or includeHeader == 'yes'):
    includeHeader = True
elif (includeHeader == 'N' or includeHeader == "n" or includeHeader == 'No' or includeHeader == 'no'):
    includeHeader = False

df = pd.read_excel(path, sheet_name = int(sheet) - 1, header = headerList)

i = 0
j = int(nr)
k = 1

for row in df.itertuples():
    export_path = f"{path_export}_{k}.txt"
    data = df.iloc[i:j]
    data.to_csv(export_path, sep = '\t', encoding ='utf-8', header = includeHeader, index = False)
    k = k + 1
    if (j <= len(df)):
        i = i + int(nr)
        j = j + int(nr)
    else:
        data = df.iloc[i:]
        data.to_csv(export_path, sep = '\t', encoding ='utf-8', header = includeHeader, index = False)
        break




