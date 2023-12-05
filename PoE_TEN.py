#!/usr/bin/env python
# coding: utf-8

# 
# 
# ---
# ## Importing Libraries
# ---
# 
# 
# 

# In[19]:


#!pip install -q streamlit
#!pip install streamlit --q

import streamlit as st
import pandas as pd
import warnings
import re
import math
from collections import defaultdict
warnings.filterwarnings('ignore')


# In[20]:


from openpyxl import load_workbook
#from google.colab import drive
#drive.mount('/content/drive', force_remount=True)

from openpyxl import load_workbook

#get the entries which have been struck out
def get_strikethrough_rows_in_column(file_name, sheet_name, column_name):
    workbook = load_workbook(filename=file_name, data_only=True)
    sheet = workbook[sheet_name]

    # Find the index of the column with the specified name
    column_index = None
    for i, col in enumerate(sheet.iter_rows(min_row=1, max_row=1, values_only=True)):
        if column_name in col:
            column_index = i
            break

    if column_index is None:
        raise ValueError(f"Column '{column_name}' not found.")

    struck_rows = set()
    for row in sheet.iter_rows(min_row=2):  # Assuming the first row is the header
        cell = row[column_index]
        if cell.font and cell.font.strikethrough:
            struck_rows.add(cell.row - 1)  # Subtract 1 to match pandas 0-indexing

    return struck_rows

# Replace 'your_column_name' with the name of the column you want to check
strikethrough_rows = get_strikethrough_rows_in_column('ECA.xlsm', 'EQUIPMENT LIST','TAG')


# In[21]:


#df_EqList=pd.read_csv('drive/My Drive/Colab Notebooks/ECA.csv')
df_EqList = pd.DataFrame()
EqList=pd.ExcelFile('ECA.xlsm')
read_EqList = pd.read_excel(EqList, 'EQUIPMENT LIST')
df_EqList = df_EqList.append(read_EqList)
# Drop the rows that are struck through
df_EqList.drop(strikethrough_rows, axis=0, inplace=True, errors='ignore')
#remove spaces from column titles
df_EqList.columns=[col.replace(" ","")for col in df_EqList.columns]
#remove NaN from PoE
df_EqList.dropna(subset=['PoE'], inplace=True)
#convert PoE to integers
df_EqList['PoE'] = pd.to_numeric(df_EqList['PoE'], errors='coerce')
#if there are still some NaN or NA entries in PoE replace them with 0
df_EqList['PoE'].fillna(0, inplace=True)
#aftter converting the number in MaterialCode to string a decimal/period is added to the string. Remove that decimal/period.
df_EqList['MATERIALCODE'] = df_EqList['MATERIALCODE'].astype(str).apply(lambda x: x.split('.')[0])


# In[22]:


#function to count total line items in the Equipment list
def count_items(dataframe):
    item_counter = defaultdict(int)
    item_indices = defaultdict(list)

    pattern1 = r'([A-Za-z-]+-\d+(-\d+)?)'
    pattern2 = r'([A-Za-z]+-[A-Za-z]+\d+)'
    pattern3 = r'([A-Za-z-]+\d+(-\d+)?)'

    for index, row in enumerate(dataframe['TAG']):
        match1 = re.search(pattern1, row)
        match2 = re.search(pattern2, row)
        match3 = re.search(pattern3, row)

        if match1:
            base_item1 = match1.group(0)
            item_counter[base_item1] += 1
            item_indices[base_item1].append(index)
        elif match2:
            base_item2 = match2.group(0)  # Use group(0) to get the entire matched string
            item_counter[base_item2] += 1
            item_indices[base_item2].append(index)
        elif match3:
            base_item3 = match3.group(0)  # Use group(0) to get the entire matched string
            item_counter[base_item3] += 1
            item_indices[base_item3].append(index)
    return item_counter, item_indices

# Assuming df_POE is your DataFrame
item_counts, item_indices = count_items(df_EqList)

EqList_Total = 0
# Printing the counts and indices
for item, count in item_counts.items():
    #print(f"{item}: {count}, Indices: {item_indices[item]}")
    EqList_Total += count

#print(f"Total Equipment/ Line Items (not POE): {EqList_Total}")


# In[23]:


def PoE(dataframe):
    item_counter = defaultdict(int)
    item_indices = defaultdict(list)

    # Create a subset for air cooler related items
    air_cooler_conditions = (
        dataframe['MATERIALCODE'].isin(['0710', '710']) &
        dataframe['REQUISITIONDESIGNATION'].str.lower().str.contains('air cooler|aircooler|air-cooled|air cooled')
    )
    air_cooler_df = dataframe[air_cooler_conditions]

    # Apply count_items to the subset
    air_cooler_counter, air_cooler_indices = count_items(air_cooler_df)

    for tag, indices in air_cooler_indices.items():
        if tag not in item_counter:
            # If the tag is encountered for the first time, count it as 1
            item_counter[tag] = 1
            # Add the first index to item_indices
            item_indices[tag].append(indices[0])

            # If there are additional items, count each as 0.5
            for additional_index in indices[1:]:
                item_counter[tag] += 0.5
                item_indices[tag].append(additional_index)
        else:
            # If the tag has already been encountered, count each item as 0.5
            for index in indices:
                item_counter[tag] += 0.5
                item_indices[tag].append(index)


    # Continue with the rest of the dataframe, skipping air cooler items
    for index, row in dataframe.iterrows():
        if index not in air_cooler_df.index:
            if row['TAG'] == row['PARENTTAGNUMBER']:
                increment = 0
                if row['MATERIALCODE'] == '1011' and 'compressor' in row['REQUISITIONDESIGNATION'].lower():
                    increment = 2
                elif row['MATERIALCODE'] == '1011' and 'turbine' in row['REQUISITIONDESIGNATION'].lower():
                    increment = 3
                elif row['MATERIALCODE'] == '0140' or '140' and 'thermal oxidizer' in row['REQUISITIONDESIGNATION'].lower():
                    increment = 3
                elif row['MATERIALCODE'] in ['4046', '4119', '4171', '4133', '210', '0210', '0168', '168', '0180','180', '0275', '275']:
                    increment = 1.2
                elif row['MATERIALCODE'] == '4064' and any(term in row['REQUISITIONDESIGNATION'].lower() for term in ['hoist', 'crane']):
                    increment = 0
                else:
                    increment = 1
                item_counter[row['TAG']] += increment
                item_indices[row['TAG']].append(index)

    return item_counter, item_indices

# Assuming item_counts and item_indices are obtained from the PoE function
item_counts, item_indices = PoE(df_EqList)

# Create a list of dictionaries, each representing a row in the DataFrame
data = []
for tag, count in item_counts.items():
    data.append({
        "Tag": tag,
        "Count": count,
        "Indices": item_indices[tag]
    })

# Convert the list of dictionaries to a DataFrame
df_export = pd.DataFrame(data)

# Write the DataFrame to a CSV file
#df_export.to_csv(r'drive/My Drive/Colab Notebooks/POE_final.csv')

PoE_Total = 0
# Printing the counts and indices
for item, count in item_counts.items():
    #print(f"{item}: {count}, Indices: {item_indices[item]}")
    PoE_Total += count
PoE_Total=math.ceil(PoE_Total)
#print(f"POE: {PoE_Total}")


# In[17]:


st.title("PoE Estimator")
st.sidebar.header("Instructions")
st.sidebar.info(
    '''Upload a **Equipment List** to find the Pieces of Equipment. Make sure you delete the any empty rows above the column titles.'''
    )
uploaded_files = st.file_uploader('Upload your files',accept_multiple_files=False, type=['xslx', 'xlsm', 'xls','csv'])

if st.button("Get PoE"):
  st.write(f"Total Equipment/ Line Items (not POE):\n {EqList_Total}, **Total PoE**: {PoE_Total}")


# In[ ]:




