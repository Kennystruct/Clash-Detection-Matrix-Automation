#DEVELOPED BY: KEHINDE AYOBADE

#Importing the required packages
import os
from textwrap import fill
import pandas as pd
import numpy as np
import xlsxwriter
import itertools
import openpyxl

#Import files from file path
file_path = r'C:\Users\m\Documents\PROGRAMMING SCRIPTS\CLASH MATRIX AUTOMATION\Excel Runtime Files'
os.chdir(file_path)
ExcelRevitData_File_Name = "All Model Elements Data"
ExcelRevitData_File_Path = r'C:\Users\m\Documents\PROGRAMMING SCRIPTS\CLASH MATRIX AUTOMATION\Clash Matrix Data\All Model Elements Data.xlsx'
ExcelMatrixFile = r'C:\Users\m\Documents\PROGRAMMING SCRIPTS\CLASH MATRIX AUTOMATION\Clash Matrix Data\Navisworks_Clash_Matrix.xlsx'

#Read 'All Model Elements Data' file
excel_df = pd.read_excel(ExcelRevitData_File_Path, sheet_name=None)

#Create Dataframe to concatenate multiple sheets
df = pd.concat(excel_df,ignore_index=False,sort=False,axis= 1)

df.to_excel('01_Merged_Sheets.xlsx', sheet_name='All Model ELements')

#Obtain all values in dataframe and convert to list
df_val = df.values.tolist()

#Transpose 2D lists and create a new list of lists
np_array = np.array(df_val)
transpose = np_array.T
df_transpose = transpose.tolist()

#Create Workbook
book = xlsxwriter.Workbook('02_Data_Write_Matrix.xlsx')
sheet = book._add_sheet('All Model Elements')

#Start rows and columns indices 
row1 = 2
row2 = 3
column1 = 2
column2 = 3

full_list = []

#Elements Data on the rows
for element in df_transpose:
    for item in element:
        #write operation perform
        sheet.write(row2, column1, item)
        #incrementing the value of the rows by 1 with each iteration
        row2 +=1
        full_list.append(item)

#Elements Data on the Columns
for element in df_transpose:
    for item in element:
        #write operation perform
        sheet.write(row1, column2, item)
        #incrementing the value of the rows by 1 with each iteration
        column2 +=1

book.close()

#Read excel file into dataframe to remove null values 
Edf = pd.read_excel(
    '02_Data_Write_Matrix.xlsx'
    )

#Remove rows with all null values
EdR = Edf.dropna(axis=0, how="all")
#Remove columns with all null values
EdC = EdR.dropna(axis=1, how="all")

EdC.to_excel('03_Matrix_Data.xlsx', sheet_name='All Model ELements')

keys_list = df.columns.tolist()

#Get all Model Categories
keys_category = []
for item in keys_list:
    keys_category.append(item[1])

#Get all project disciplines
keys_discipline = []
for item in keys_list:
    keys_discipline.append(item[0])

keys_count = df.count().tolist()
keys_cycle1 = []
keys_cycle2 = []

#Combine list in Categories and Disciplines
catlist = ['Element Category']
for i in itertools.zip_longest(keys_category, keys_count):
    keys_cycle1.append(i)

repeat_elements_category1 = list(itertools.chain.from_iterable(itertools.repeat(x, y) for x, y in keys_cycle1))
repeat_elements_category2 = catlist + repeat_elements_category1


displist = ['Discipline']
for i in itertools.zip_longest(keys_discipline, keys_count):
    keys_cycle2.append(i)


repeat_elements_discipline1 = list(itertools.chain.from_iterable(itertools.repeat(x, y) for x, y in keys_cycle2))
repeat_elements_discipline2 = displist + repeat_elements_discipline1

#Set Indices for cateories and disciplines
disp_cat1 = [repeat_elements_discipline2, repeat_elements_category2]
disp_cat2 = [repeat_elements_discipline1, repeat_elements_category1]

index_column1 = pd.MultiIndex.from_arrays(disp_cat1)
index_column2 = pd.MultiIndex.from_arrays(disp_cat2)

EdCa = EdC.T.reset_index(drop=True).T

EdCa.columns = index_column1
EdCa.set_index([index_column1], inplace=True, drop=True)

EdCa.dropna(how='all')

ET = pd.DataFrame({'TYPES': ['Element Type']})
EC = pd.DataFrame({'TEST': ['NA']})


with pd.ExcelWriter('05_Clash_Matrix.xlsx') as writer:
    EdCa.to_excel(
        writer, 
        sheet_name='All Model ELements',
        startrow= 0,
        startcol= 0,
        header= True,
        index = True,
    )
    ET.to_excel(
        writer, 
        sheet_name='All Model ELements',
        startrow= 3,
        startcol= 2,
        header= False,
        index = False,
    )
    EC.to_excel(
        writer, 
        sheet_name='All Model ELements',
        startrow= 4,
        startcol= 3,
        header= False,
        index = False,
    )

EdCa.head(5)

path1 = r'E:\NAVISWORKS CLASH TEST FILES\Excel Files\05_Clash_Matrix.xlsx'
path = '05_Clash_Matrix.xlsx'

# load excel file using openpyxl
book = openpyxl.load_workbook(path)
sheet = book['All Model ELements']
sheet.delete_rows(3, 1)
book.save('05_Clash_Matrix.xlsx')

NCM = pd.read_excel(
    '05_Clash_Matrix.xlsx',
    header=[0,1,2],
    index_col=[0,1,2],
    keep_default_na=True
)

with pd.ExcelWriter(ExcelMatrixFile) as writer:
    NCM.to_excel(
        writer, 
        sheet_name='All Model ELements',
        startrow= 0,
        startcol= 0,
        header= True,
        index = True,
    )

NCM.head(10)