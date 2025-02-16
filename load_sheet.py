import openpyxl
import json
from datetime import datetime

wb = openpyxl.load_workbook("mess.xlsx")
sheet = wb['Sheet1']
menu_data = {}

columns_data = []
for col in sheet.iter_cols(values_only=True):
    columns_data.append(list(col)) 

for col in columns_data:
    if col and any(value is not None for value in col):
        pass
    
    day_menu = {}
    current_meal = None
    
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
            current_meal = item.upper()
            day_menu[current_meal] = []
        elif current_meal:  
            day_menu[current_meal].append(item)
    
    if day_menu:
        menu_data[date] = day_menu

jsonf = "mess.json"
with open(jsonf, "w") as jsonf:
    json.dump(menu_data, jsonf, indent=2)
