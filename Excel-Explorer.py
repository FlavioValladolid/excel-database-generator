# from openpyxl.styles.styleable import StyleableObject
import pandas as pd
import openpyxl

file = 'Test.xlsx'
data = pd.ExcelFile(file)
#https://medium.com/analytics-vidhya/how-to-extract-information-from-your-excel-sheet-using-python-5f4f518aec49
# We need to install openpyxl before
# print(data.sheet_names) # Read excel sheet names

# Parse the sheet into a dataframe
df = data.parse('Eng-Start Here')
# df.info

### Read in the spreadsheet data

ps = openpyxl.load_workbook(file)

sheet = ps['Eng-Start Here']


# Returns the total number of rows in the sheet

## Iterate over all the rows in the sheet

info_list = []

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

items_names = list(items_dictionary.keys())

qty_w_or_operation = 0
repeat_operation = 0
net_ft_operation = 0
net_lbs_operation = 0
net_kg_operation = 0
incl_OR_ft_operation = 0
incl_OR_lbs_operation = 0
incl_OR_kg_operation = 0
ww_operation = 0
lbs_m_operation = 0

for row in items_names:
    items_dictionary[row]
    if row != 'Repeat' and row != 'Web Width (W.W.)' and row != 'lbs/m' and row != 'Qty with OR%' and row != 'Net (ft)' and row != 'Net (lbs)' and row != 'Net (kg)' and row != 'Incl OR (ft)' and row != 'Incl OR (lbs)' and row != 'Incl OR (kg)':
        info_list.append(sheet[items_dictionary[row]].value)
    elif row == 'Qty with OR%':
        qty_w_or_operation = sheet[items_dictionary['Quantity']].value * (1 + sheet[items_dictionary['Overrun%']].value)
        info_list.append(qty_w_or_operation)
    elif row == 'Repeat':
        repeat_operation = sheet[items_dictionary['Cylinder']].value / sheet[items_dictionary['Around']].value
        info_list.append(repeat_operation)
    elif row == 'Web Width (W.W.)':
        ww_operation = 2*(sheet[items_dictionary['Length']].value + sheet[items_dictionary['Zip Header']].value) + sheet[items_dictionary['Lip']].value + sheet[items_dictionary['Bottom Gusset (BM)']].value
        info_list.append(ww_operation)
    elif row == 'lbs/m':
         lbs_m_operation = (repeat_operation * ww_operation) / 30.033 * sheet[items_dictionary['Mil']].value
         info_list.append(lbs_m_operation)
    elif row == 'Net (ft)':
        net_ft_operation = repeat_operation*sheet[items_dictionary['Quantity']].value/12/sheet[items_dictionary['Across']].value
        info_list.append(net_ft_operation)
    elif row == 'Net (lbs)':
        net_lbs_operation = (sheet[items_dictionary['Quantity']].value/1000)*lbs_m_operation
        info_list.append(net_lbs_operation)
    elif row == 'Net (kg)':
        net_kg_operation = net_lbs_operation / 2.2046
        info_list.append(net_kg_operation)
    elif row == 'Incl OR (ft)':
        incl_OR_ft_operation = net_ft_operation * (1 + sheet[items_dictionary['Overrun%']].value)
        info_list.append(incl_OR_ft_operation)
    elif row == 'Incl OR (lbs)':
        incl_OR_lbs_operation = net_lbs_operation * (1 + sheet[items_dictionary['Overrun%']].value)
        info_list.append(incl_OR_lbs_operation)
    elif row == 'Incl OR (kg)':
        incl_OR_kg_operation = net_kg_operation*(1 + sheet[items_dictionary['Overrun%']].value)
        info_list.append(incl_OR_kg_operation)
    else:
        print('Error')

print(info_list)
