import openpyxl
import json
from datetime import datetime

wb = openpyxl.load_workbook("mess.xlsx")
sheet = wb['Sheet1']
data = {}

col_check = []
for col in sheet.iter_cols(values_only=True):
    col_check.append(list(col)) 

for col in col_check:
    if col and any(value is not None for value in col):
        pass
    
    dict_day = {}
    meal = None
    
    date = None
    for cell in col:
        if isinstance(cell, datetime):
            date = cell.strftime("%Y-%m-%d")
            break
        elif isinstance(cell, str):
            try:
                date = datetime.strptime(cell, "%Y-%m-%d").strftime("%Y-%m-%d")
                break
            except ValueError:
                continue
    
    if not date:
        continue 
    for item in col[1:]: 
        if item is None:
            continue
        
        item = str(item).strip()
        if item[0] == "*":   
            continue
        
        if item.upper() in ["BREAKFAST", "LUNCH", "DINNER"]:
            meal = item.upper()
            dict_day[meal] = []
        elif current_meal:  
            dict_day[meal].append(item)
    
    if day_menu:
        data[date] = dict_day

jsonf = "mess.json"
with open(jsonf, "w") as jsonf:
    json.dump(data, jsonf, indent=2)
