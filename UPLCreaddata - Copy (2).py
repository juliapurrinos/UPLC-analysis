import pandas as pd
import openpyxl
from openpyxl import load_workbook
import xlwings as xw

directory = "C:/Users/ba03/OneDrive/Desktop/biome/" #<- change file directory/folder here "C:\Users\ba03\OneDrive\Desktop\biome\Results_20240829_1.xlsx"
csv_file = pd.ExcelFile(directory + "20240905_test2wb.xlsx")#put specific file name to be read
csv_file = pd.read_excel(csv_file, names= ['samplename_RT', 'Type', 'Width', 'Area', 'Height', 'Areapercent', 'Name'])

dfdict = {}
samplenames = []
data =[]

for i, row in enumerate(csv_file['samplename_RT']):
    if row == 'Sample name:':
        samplename = csv_file.loc[i, 'Type']
        if samplename in samplenames:
            samplename = str(samplename) + 'i'
        samplenames.append(samplename)

for i, row in enumerate(csv_file['samplename_RT']):
    
    datalist = [None, None, None, None]
    if isinstance(row, float):#should be in order of samples, can always fill in missing with none 
        if row > 9.2 and row <= 9.6:#HBA window
            datalist[0] = float(csv_file.loc[i, 'Area'])

        if row > 7.45 and row <= 8.0:#PCA window
            datalist[1] = float(csv_file.loc[i, 'Area'])

        if row > 1.5 and row < 1.6: #2,4 PDCA
            datalist[2] = float(csv_file.loc[i, 'Area'])

        if row > 1.9 and row <= 2.1:#2,5 PDCA
            datalist[3] = float(csv_file.loc[i, 'Area'])

    data.append(datalist)

print(len(data))
# alldata = pd.DataFrame(dfdict, index=['4-HBA', 'PCA', '2,4 PDCA', '2,5 PDCA'])
# formattedtable = alldata.transpose()
#formattedtable.to_csv(directory + 'secondstandards.csv', index=True)#

# for sheetname in csv_file.sheet_names:
#     pdfile = pd.read_excel(csv_file, sheet_name=sheetname)
#     samplename = pdfile.loc[2, 'Unnamed: 6']
#     if samplename in samplenames:
#         samplename = samplename + 'i'
    
#     samplenames.append(samplename)
    
#     try:
#         if 'Blank' in samplename or 'blank' in samplename or 'water' in samplename: #sheets in file that the code will avoid in data output
#             continue

#     except TypeError:
#         continue

#     signal260 = 0
#     signal280 = 0
    
#     for i, row in enumerate(pdfile['Unnamed: 2']):
    
#         if row == 'DAD1 C, Sig=260,4 Ref=off':
#             signal260 = i
        
#         elif row == 'DAD1 D, Sig=280,4 Ref=off':
#             signal280 = i
        
#         else:
#             continue
    
#     datalist = [None, None, None]
 
#     for i, rt260 in enumerate(pdfile['Area Percent Report'].iloc[signal260+2:signal280-1]):
#         if float(rt260) >= 9.2 and float(rt260) <= 9.3: #4-HBA window
#             datalist[0] = float(pdfile.loc[i+signal260+2, 'Unnamed: 7'])
            
#         if float(rt260) >= 7.45 and float(rt260) <= 8.0: #PCA window
#             datalist[1] = float(pdfile.loc[i+signal260+2, 'Unnamed: 7'])
            
#     for i, rt280 in enumerate(pdfile['Area Percent Report'].iloc[signal280+2:]):
#         if float(rt280) >= 1.9 and float(rt280) <= 2.1: #PDCA window
#             datalist[2] = float(pdfile.loc[i+signal280+2, 'Unnamed: 7'])

      
#     dfdict[samplename] = datalist
 

# alldata = pd.DataFrame(dfdict, index=['4-HBA', 'PCA', '2,5 PDCA'])
# formatedtable = alldata.transpose()
   
# formatedtable.to_csv(directory + 'secondstandards.csv', index=True)#<- change new file name here
