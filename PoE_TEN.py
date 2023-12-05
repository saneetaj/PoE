#!/usr/bin/env python
# coding: utf-8

#pip install --upgrade pandas
# ## Importing Libraries
import streamlit as st
import pandas as pd
import warnings
import re
import math
import openpyxl
from collections import defaultdict
warnings.filterwarnings('ignore')


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


# In[27]:


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

    packaged_eq=['4046', '4119', '4171', '4133', '210', '0210', '0168', '168', '0180','180', '0275', '275']
    # Continue with the rest of the dataframe, skipping air cooler items
    for index, row in dataframe.iterrows():
        if index not in air_cooler_df.index:
            if row['TAG'] == row['PARENTTAGNUMBER']:
                increment = 0
                if row['MATERIALCODE'] == '1011' and 'compressor' in row['REQUISITIONDESIGNATION'].lower():
                    increment = 2
                elif row['MATERIALCODE'] == '1011' and 'turbine' in row['REQUISITIONDESIGNATION'].lower():
                    increment = 3
                elif (row['MATERIALCODE'] == '0140' or row['MATERIALCODE'] == '140') and any(term in row['REQUISITIONDESIGNATION'].lower() for term in ['thermal oxidizer', 'oxidizer']):
                    increment = 3
                elif any (term in row['MATERIALCODE'] for term in packaged_eq):
                    increment = 1.2
                elif row['MATERIALCODE'] == '4064' and any(term in row['REQUISITIONDESIGNATION'].lower() for term in ['hoist', 'crane']):
                    increment = 0
                else:
                    increment = 1
                item_counter[row['TAG']] += increment
                item_indices[row['TAG']].append(index)

    return item_counter, item_indices


# In[28]:


st.title("PoE Estimator")
st.sidebar.header("Instructions")
st.sidebar.info(
    '''Upload a **EQUIPMENT LIST** to find the Pieces of Equipment. ***Make sure you delete any empty rows above the column titles, and any tags that are struck out.***'''
    )
uploaded_files = st.file_uploader('Upload your files',accept_multiple_files=False, type=['xslx', 'xlsm', 'xls','csv'])

if st.button("Get PoE"):

    # Ensure df_EqList is a DataFrame
    if 'df_EqList' not in locals() or not isinstance(df_EqList, pd.DataFrame):
        df_EqList = pd.DataFrame()
 
    # Load data from the uploaded file
    if uploaded_files is not None:
        EqList = pd.ExcelFile(uploaded_files)
        df_EqList = pd.read_excel(EqList, 'EQUIPMENT LIST')

        #Drop NaN 
        df_EqList.dropna(subset=['TAG'], inplace=True)

        #remove spaces from column titles
        df_EqList.columns=[col.replace(" ","")for col in df_EqList.columns]
            
        #after converting the number in MaterialCode to string a decimal/period is added to the string. Remove that decimal/period.
        df_EqList['MATERIALCODE'] = df_EqList['MATERIALCODE'].astype(str).apply(lambda x: x.split('.')[0])

        df_EqList['REQUISITIONDESIGNATION'] = str(df_EqList['REQUISITIONDESIGNATION'])
        
        line_counts, line_indices = count_items(df_EqList)
        EqList_Total = 0
        # Printing the counts and indices
        for item, count in line_counts.items():
            EqList_Total += count
    
        POE_counts, POE_indices = PoE(df_EqList)
        PoE_Total = 0
        # Printing the counts and indices
        for item, count in POE_counts.items():
            st.write(f"{item}: {count}, Indices: {item_indices[item]}")
            PoE_Total += count
        PoE_Total=math.ceil(PoE_Total)
                
        st.write(f"Total Equipment/ Line Items (not POE):\n {EqList_Total}, **Total PoE**: {PoE_Total}")

# In[ ]:




