# -*- coding: utf-8 -*-
"""
Created on Wed Aug 28 12:21:12 2024

@author: tpham + ChatGPT
"""

import os
import pandas as pd
import re

#%% ===================== CONFIGURATION ===================== #
MASTER_FOLDER = r"Y:/5900/HydrogenTechFuelCellsGroup/CO2R/Nhan P/Experiments/CO2 Cell Testing/TS7/1NP91/GC Data"

# Keywords to extract from the table
KEYWORDS = ['Component', 'Chan#', 'Retention', 'ESTD', 'Norm.ESTD%']

# Define analysis gas ranges
ANALYSIS_GASES = [
    {"name": "Hydrogen", "Chan#": 1, "range": (21, 26)},
    {"name": "Carbon Monoxide", "Chan#": 1, "range": (60, 90)},
    {"name": "Oxygen", "Chan#": 1, "range": (30, 33)},
]

IDENTIFIER_GASES = {
    "Identifier Gas 1": {"name": "Nitrogen", "Chan#": 2, "range": (31, 33)},
    "Identifier Gas 2": {"name": "Carbon Dioxide", "Chan#": 3, "range": (20, 26)},
    "Identifier Gas 3": {"name": "Carbon Monoxide", "Chan#": 1, "range": (60, 90)},
}

#%% ======================================================== #
#Get Table from txt file in multiple folders. This definition works well.
def get_txt_table_data(file_path):
    try:
        with open(file_path, 'r') as f:
            lines = f.readlines()
        
        start_index = None
        end_index = None
        
        for i, line in enumerate(lines):
            if '>==CT==' in line:
                start_index = i
            if 'Totals' in line:
                end_index = i
                break
        if start_index is not None and end_index is not None:
            table_lines = lines[start_index + 1:end_index]
        else:
            print("Table Section not found. Check start_index and end_index logic")
            table_lines = []
        
        data = []
        headers = []
        for line in table_lines:
            stripped = line.strip()
            if stripped == '':
                continue        #Skip empttu lines
            parts = [p.strip() for p in line.split('\t')]   #Split by tab
            if not headers:
                headers = parts
            else: data.append (parts)
            
        if not data:
            print(f"No valid data found in file: {file_path}")
            return pd.DataFrame()       #return empty dataframe if no data extracted
        return pd.DataFrame(data, columns=headers)
    except Exception as e:
        print(f"Error processing file {file_path}: {e}")
        return pd.Dataframe() #return an empty Dataframe

#Find the gas we need
def find_gas(df, gas):
    try:
        df['Retention'] = pd.to_numeric(df['Retention'], errors='coerce')
        df['Chan#']    = pd.to_numeric(df['Chan#'],    errors='coerce', downcast='integer')
    except Exception as e:
        print(f"[{gas['name']}] Conversion error: {e}")
        return None

    # filter
    matches = df[
        (df['Chan#'] == gas['Chan#']) &
        (df['Retention'].between(gas['range'][0], gas['range'][1]))
    ]

    if matches.empty:
        print(f"[{gas['name']}] No match (Chan#={gas['Chan#']}, range={gas['range']})")
        return None

    if len(matches) > 1:
        print(f"[{gas['name']}] ERROR: {len(matches)} matches in Chan#={gas['Chan#']}, range={gas['range']}")
        print(matches.to_dict(orient='records'))
        return 'error'

    # exactly one
    row = matches.iloc[0]
    return row.to_dict()

#Sort folder by Group ID based on 3 Identifier gas value
def get_group_id(df, folder_name):
    g1 = find_gas(df, IDENTIFIER_GASES['Identifier Gas 1'])
    g2 = find_gas(df, IDENTIFIER_GASES['Identifier Gas 2'])
    g3 = find_gas(df, IDENTIFIER_GASES['Identifier Gas 3'])

#Debug Print
    print(f"\n[{folder_name}] Matching Identifier Gases:")
    print(f"  G1 ({IDENTIFIER_GASES['Identifier Gas 1']['name']}): {g1.to_dict() if hasattr(g1, 'to_dict') else g1}")
    print(f"  G2 ({IDENTIFIER_GASES['Identifier Gas 2']['name']}): {g2.to_dict() if hasattr(g2, 'to_dict') else g2}")
    print(f"  G3 ({IDENTIFIER_GASES['Identifier Gas 3']['name']}): {g3.to_dict() if hasattr(g3, 'to_dict') else g3}")
#
    if any(g == 'error' for g in [g1, g2, g3]):
        print(f"[{folder_name}] One or more identifier gases returned error.")
        return 'Error'
    
    if g1 is None and g2 is not None:
        print(f"[{folder_name}] G1 not found, G2 found → Group ID = 'Cathode'")
        return 'Cathode'
         
    try:
        estd2 = float(g2['ESTD']) if g2 is not None else None
        estd3 = float(g3['ESTD']) if g3 is not None else None

        if estd2 is not None and estd3 is not None:
            if 0.4 <= estd2 <= 0.6 and 0.4 <= estd3 <= 0.6:
                print(f"[{folder_name}] Group ID set to 'QC'")
                return 'QC'
        
        # Not QC — check Anode/Cathode by comparing Area of g1 and g2
        if g1 is not None and g2 is not None:
            area1 = float(g1['Area'])
            area2 = float(g2['Area'])

            if area1 > area2:
                print(f"[{folder_name}] Group ID set to 'Anode' (Area1 > Area2)")
                return 'Anode'
            elif area1 < area2:
                print(f"[{folder_name}] Group ID set to 'Cathode' (Area1 < Area2)")
                return 'Cathode'
            else:
                print(f"[{folder_name}] Check: Area1 == Area2 — conflict.")
                return 'Error'
        else:
            print(f"[{folder_name}] Insufficient data to compare Areas for Anode/Cathode.")
            return 'Error'
    except Exception as e:
        print(f"[{folder_name}] Error in ESTD/Area logic: {e}")
        return 'Error'


def extract_analysis_results(df, group_id):
    results = {}
    for gas in ANALYSIS_GASES:
        match = find_gas(df, gas)
        label = f"{group_id}_{gas['name']}_Chan{gas['Chan#']}"
        if match == 'error' or match is None:
            results[label] = None
        else:
            results[label] = match['ESTD']
    return results

def main():
    data_rows = []
    for folder in os.listdir(MASTER_FOLDER):
        folder_path = os.path.join(MASTER_FOLDER, folder)
        if not os.path.isdir(folder_path):
            continue

        file_path = os.path.join(folder_path, 'SAMPRSLT.TXT')
        if not os.path.exists(file_path):
            continue

        df = get_txt_table_data(file_path)
        if df.empty:
            continue

        group_id = get_group_id(df, folder)
        result_data = extract_analysis_results(df, group_id) if group_id != 'Error' else {}
        row = {'FolderName': folder, 'Group ID': group_id, **result_data}
        data_rows.append(row)

    final_df = pd.DataFrame(data_rows)
    final_df = final_df.sort_values(by='FolderName')
    final_df.to_excel(MASTER_FOLDER + '/Grouped_Analysis.xlsx', index=False)
    print("Excel file created successfully.")

if __name__ == '__main__':
    main()