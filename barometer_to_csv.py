#Here we are taking the following steps:
#   1. Parse the excel file and assign each sheet name to a list.
#   2. If 'Base Redressee' is not in cell [8,1] but is in cell [9,1], delete row 8. This ensures each sheet is of the same structure.
#   3. We then extract the list of questions from the Summary page to a list.
#   4. Finally, we add each answer for each year from every question to a dictionary.
#   5. This dictionary is then output to a CSV file.

# ## Extract each sheet and remove inconsistencies

import os
import numpy as np
import pandas as pd
from collections import defaultdict

xl = pd.ExcelFile('Datasets/download.xlsx') #parse Excel file
sheet_names = xl.sheet_names #assign sheet names to a list
k = len(sheet_names)
sheets_consistent=[]

for x in range(k):
    sheet_x = xl.parse(sheet_names[x]) #parse sheet with name x, where i is defined in sheet_names
    if sheet_x.iat[8,1] != "Base redressée" and sheet_x.iat[9,1] == "Base redressée":
        sheet_x = sheet_x.drop([8]) #drop row 8
        i = sheet_x.shape[0] #number of rows in sheet_x
        sheet_x.index = range(i) #reassign index
        sheets_consistent.append(sheet_x) #append sheet_x to sheet_names_2
    else:
        sheets_consistent.append(sheet_x) #if 'base redressee' in [8,1], leave as is and append

for x in range(k):
    if sheets_consistent[x].iat[8,1] != "Base redressée":
        print(x) #check to see if all sheets have 'base redressee' in cell [8,1]. (sheets 0 and 1 are exceptions)


# # Extract questions from sheet[1] to a list
summary_page = sheets_consistent[1].set_index('SOMMAIRE').dropna()#set index column to be SOMMAIRE column and drop nan values
y = len(summary_page) #define number of non-empty values in summary_page

questions=[] #create empty list
for x in range(y):
    questions.append(summary_page.iat[x,0]) #append questions to a list

    
# # Data into structured dictionary

my_dict = defaultdict(list) #create an empty dictionary of lists
for t in range(2,y+2):
    my_dict[t].append({'question' : questions[t-2]}) #append questions from questions list to dictionary
    p = sheets_consistent[t]['Baromètre d’opinion de la Drees'].count() #count the number of values in answers column
    g = sheets_consistent[t].iloc[6].count() #count the number of values in years row
    for i in range(g):
        for j in range(p-1):
            #append year, answer and volume to dictionary for every question
            my_dict[t].append({("{0}".format(sheets_consistent[t].iat[6,2+(2*i)]),"{0}".format(sheets_consistent[t].iat[8+j,1].replace('  ',''))):"{0}".format(sheets_consistent[t].iat[8+j,2+(2*i)])})
            
#export to CSV
with open('barometer_dataset.csv', 'w') as f:
    writer = csv.writer(f)
    for k,v in my_dict.items():
        writer.writerow(v)

