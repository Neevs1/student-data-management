import pandas as pd
import numpy as np
from tkinter import *
from tkinter import filedialog

fp=""
def selectFile():
    global fp
    filepath = filedialog.askopenfilename(filetypes=(("Excel files","*.xlsx"),))
    fp=filepath

    window.destroy()
    

window = Tk() #creates a window using Kinter
button = Button(text="Select file",command=selectFile) #button that executes a command on press
text=Label(window,text="Please select the required Excel file")
text.pack()
button.pack() #added button to window
window.geometry("400x400") #set size of window to 400 by 400 px
window.mainloop() #added window to loop




file_path = fp #file_path is assigned value selected through window
df = pd.read_excel(file_path, sheet_name='Sheet1')

score_columns = ['UT-1 (15)', 'UT-2 (15)', 'UT-3 (15)', 'UT-4 (15)', 'UT-5 (15)&6(15)']

df[score_columns] = df[score_columns].apply(pd.to_numeric, errors='coerce')

branches = df.groupby('FE \nBranch')

def check_improvement(row, test1, test2):
    if pd.notna(row[test1]) and pd.notna(row[test2]):
        if row[test2] > row[test1]:
            return 'Improved'
        elif row[test2] < row[test1]:
            return 'Declined'
        else:
            return 'No Change'
    return 'No Data'

df['Performance UT-1 to UT-2'] = df.apply(lambda row: check_improvement(row, 'UT-1 (15)', 'UT-2 (15)'), axis=1)

def categorize_marks(total_score):
    if 80 <= total_score <= 100:
        return 'Outstanding'
    elif 70 <= total_score <= 79:
        return 'Excellent'
    elif 60 <= total_score <= 69:
        return 'Very Good'
    elif 55 <= total_score <= 59:
        return 'Good'
    elif 50 <= total_score <= 54:
        return 'Above Average'
    elif 45 <= total_score <= 49:
        return 'Average'
    elif 40 <= total_score <= 44:
        return 'Pass'
    elif 0 <= total_score <= 39:
        return 'Fail'
    return 'Absent'

df['Total_Score'] = df[score_columns].sum(axis=1, min_count=1)  
df['Category'] = df['Total_Score'].apply(categorize_marks)
num = fp.rindex('/')
fp=fp[:num+1]

for branch_name, group in branches:
    output_file = f'{fp}students_{branch_name.replace("/", "_")}.xlsx'  
    group.to_excel(output_file, index=False)
    print(f'Saved {output_file}')
