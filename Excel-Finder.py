
import os
import explorer as exp
import pandas as pd

path = r'C:\Users\Flavio\Documents\Orders - Printing and Converting-20210726T191931Z-001\Orders - Printing and Converting'
# os.chdir(path)

"""
folder_list = []
a = ''

for folder in os.listdir('./'):

    a = folder
    a = a.replace(' ','\\')
    print(os.listdir(f'{a}'))


    # if file.endswith('.xls'):
    #     print(file)

"""

items_dictionary = {'SMK Order': 'B4',
                'Item Name': 'B6',
                'SMK ID or Item ID': 'B8',
                'Quantity': 'H26',
                'Overrun%': 'H27',
                'Qty with OR%': '',
                'Cylinder': 'H18',
                'Around': 'H19',
                'Repeat': '',
                'Across': 'H21',
                'Width': 'B15',
                'Length': 'B17',
                'Bottom Gusset (BM)': 'B20',
                'Lip': 'B19',
                'Zip Header': 'B18',
                'Web Width (W.W.)': '',
                'lbs/m': '',
                'Net (ft)': '',
                'Net (lbs)': '',
                'Net (kg)': '',
                'Incl OR (ft)': '',
                'Incl OR (lbs)': '',
                'Incl OR (kg)': '',
                'Film grade': 'H11',
                'Mil': 'H12'}


df = pd.DataFrame([])
name_ = ''
for root, dirs, files in os.walk(path):
    for name in files:
        if name.endswith((".xlsx")):
            name_ = name
            # name_ = name_.replace('.xls.xlsx','.xlsx')
            if exp.excel_explorer(root+'\\'+name_) != None:
                df = df.append(pd.DataFrame([exp.excel_explorer(root+'\\'+name_)]))
                print(name_)
                print(df)   
            else:
                pass

df.columns = list(items_dictionary.keys())
df.to_csv('file_name.csv')