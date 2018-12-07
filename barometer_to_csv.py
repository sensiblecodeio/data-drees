
# coding: utf-8

# ## Extract each sheet and remove inconsistencies

# In[137]:


import os
import numpy as np
import pandas as pd
from collections import defaultdict
import json


# In[2]:


xl = pd.ExcelFile('Datasets/download.xlsx') #parse Excel file
sheet_names = xl.sheet_names #assign sheet names to a list
k = len(sheet_names)
sheets_consistent=[]


# In[3]:


for x in range(k):
    sheet_x = xl.parse(sheet_names[x]) #parse sheet with name x, where i is defined in sheet_names
    if sheet_x.iat[8,1] != "Base redressée" and sheet_x.iat[9,1] == "Base redressée":
        sheet_x = sheet_x.drop([8]) #drop row 8
        i = sheet_x.shape[0] #number of rows in sheet_x
        sheet_x.index = range(i) #reassign index
        sheets_consistent.append(sheet_x) #append sheet_x to sheet_names_2
    else:
        sheets_consistent.append(sheet_x) #if 'base redressee' in [8,1], leave as is and append


# In[4]:


for x in range(k):
    if sheets_consistent[x].iat[8,1] != "Base redressée":
        print(x) #check to see if all sheets have 'base redressee' in cell [8,1]. (sheets 0 and 1 are exceptions)


# # Extract questions from sheet[1] to a list

# In[5]:


summary_page = sheets_consistent[1].set_index('SOMMAIRE').dropna()#set index column to be SOMMAIRE column and drop nan values
y = len(summary_page)


# In[6]:


questions=[]
for x in range(y):
    questions.append(summary_page.iat[x,0]) #append questions to a list


# # Data into structured dictionary

# In[120]:


my_dict = defaultdict(list)
for t in range(2,y+2):
    my_dict[t].append({'question' : questions[t-2]})
    p = sheets_consistent[t]['Baromètre d’opinion de la Drees'].count()
    g = sheets_consistent[t].iloc[6].count()
    for i in range(g):
        for j in range(p-1):
            my_dict[t].append({("{0}".format(sheets_consistent[t].iat[6,2+(2*i)]),"{0}".format(sheets_consistent[t].iat[8+j,1].replace('  ',''))):"{0}".format(sheets_consistent[t].iat[8+j,2+(2*i)])})


# In[148]:


with open('barometer_dataset.csv', 'w') as f:
    writer = csv.writer(f)
    for k,v in my_dict.items():
        writer.writerow(v)

