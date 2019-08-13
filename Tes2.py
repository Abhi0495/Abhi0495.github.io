import openpyxl
import pandas as pd
Locate = ['Telangana']
filepath_Invoice_number = r"C:\Users\pv.abhilash\Desktop\vinutha\Import of services invoice series.xlsx"
names = list()
ws = openpyxl.load_workbook(filepath_Invoice_number)
#file_read = pd.read_excel(filepath_Invoice_number)
file_read = pd.ExcelFile(filepath_Invoice_number)
names = file_read.sheet_names
no = len(names)
Latest_sheetname = names[no-1]
file_read = pd.read_excel(filepath_Invoice_number,Latest_sheetname)
df = pd.DataFrame(file_read)
print(df)
dr = df[df.Location.isin(Locate)]
#dr = df.loc[(df.Location == Locate)]
print (dr)