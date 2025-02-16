"""""""""
My thought process for this project was: 
1) Iteration on columns 
2) Converting columns into lists (The parent node will be the date and will have children nodes as breakfast lunch and dinner of that date)
3) Split the list on the day of the week (So for Example I have a list [1,2,3,4], I'll split it into [[1,2], [3,4]])->
(What it does is first a single list is created [Breakfast:[], ..... etc.] and then I split it into [[Breafast: []], [Lunch: []], [Dinner: []]])
4) Conversion to dictionary outside the col_check list loop   
"""""""""
import openpyxl
import json
from datetime import datetime

wb = openpyxl.load_workbook("mess.xlsx")
sheet = wb['Sheet1']
data = {} # will store the extracted menu 

col_check = []
for col in sheet.iter_cols(values_only=True): #iterating the columns 
    col_check.append(list(col)) #storing the col data in col_check 

for col in col_check:
    if col and any(value is not None for value in col):
        pass #Making sure that no column is empty 
    
    dict_type_of_meal = {} #dictionary to store meal type(Breakfast, Lunch and Dinner) that will correspond to menu items
    meal = None #stores the current meal type which we are tracking 
    
    date = None #the lines from 28 to 41 iterates columns to find a date and converts it into string iterable form(yyyy, mm, dd) format
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
    
    if not date: #if date is not found skip the column 
        continue 
    for item in col[1:]:  #the for loop is to tackle any empty value in the cell 
        if item is None:
            continue
        
        item = str(item).strip() #Now that date-time problem is fixed, converting all the items in string format and using .strip() to remove spaces (if any)
        if item[0] == "*":   #if the first letter of the string is * skip it. 
            continue
        
        if item.upper() in ["BREAKFAST", "LUNCH", "DINNER"]: 
            meal = item.upper() #checks if the item is a meal type(B, L, D) and if yes, updates the var 
            dict_type_of_meal[meal] = [] #after updating, creates an empty list under the dictionary to store food items on that particular event
        elif dict_type_of_meal:
            # Since there was no day type json object in the example given by sutt, I removed the days of the week from the final output
            words = item.split()
            filtered_words = [word for word in words if word.upper() not in ["MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY", "SUNDAY"]]
            cleaned_item = " ".join(filtered_words)
            if cleaned_item: 
                dict_type_of_meal[meal].append(cleaned_item)
    
    if dict_type_of_meal:
        data[date] = dict_type_of_meal #if the dictionary contains meal info, stores under the main data dictionary under that particular date

jsonf = "mess.json"  #saving data in a new json file
with open(jsonf, "w") as jsonf: #opening file in write mode "w"
    json.dump(data, jsonf, indent=2)
#data -> required python objects to be converted to .json format
#jsonf -> File where the json format will be displayed 
