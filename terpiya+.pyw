import os
import sys
import pandas as pd
import csv
from tkinter import StringVar, ttk
import tkinter as tk
from tkinter import *

os.chdir(sys.path[0])
print('Starting the prorgam...')

print('Try to open terapiya+.xlsx')
read_file = pd.read_excel('terapiya+.xlsx', sheet_name='Sheet2')
print('terapiya+.xlsx: Sheet2 in opened')
print('Try to convert terapiya+.xlsx to terapiya+Sheet2.csv')
read_file.to_csv('terapiya+Sheet2.csv', index=None, header=True)
print('terapiya+.csv is created')

listDataTerapiya = []  # Data form terapiya+Sheet2.cs

listOfDrags = []  # List of drugs for combo
listOfMR = []  # List of medical representative
listOfYears = []  # List of years for combo of years
# List of month for combo of month
lisfOfMonth = ['01', '02', '03', '04', '05',
               '06', '07', '08', '09', '10', '11', '12']
listOfDays = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12', '13', '14', '15', '16', '17',
              '18', '19', '20', '21', '22', '23', '24', '25', '26', '27', '28', '29', '30', '31']  # List of days for combo of days

print('Sending data from "terapiya+Sheet2.csv" to listDataTerapiya')
with open('terapiya+Sheet2.csv', encoding='utf-8') as input:
    csv_reader = csv.DictReader(input, delimiter=',')
    for row in csv_reader:
        listDataTerapiya.append(row)
        listOfYears.append(row['Min of * Дата'][:4])

#print(listDataTerapiya[0]['Имя МП1'])
# print(listDataTerapiya[0]['Препарат'])
#print(listDataTerapiya[0]['Min of * Дата'])

for i in range(len(listDataTerapiya)):
    if listDataTerapiya[i]['Препарат'] == '':
        pass
    else:
        listOfDrags.append(listDataTerapiya[i]['Препарат'])
listOfDrags = list(set(listOfDrags))
listOfDrags.sort()

for i in range(len(listDataTerapiya)):
    if listDataTerapiya[i]['Имя МП1'] == '':
        pass
    else:
        listOfMR.append(listDataTerapiya[i]['Имя МП1'])
listOfMR = list(set(listOfMR))
listOfMR.sort()

listOfYears = list(set(listOfYears))
listOfYears.sort()
listOfYears.pop(0)


##### UI ####


def getDataFromUI():
    print('>getDataFromUI:')
    getUIName = variableOfName.get()
    getUIDrag = comboOfDrag.get()
    getUIStartYear = variableOfStartYear.get()
    getUIStartMonth = comboOfStartMonth.get()
    getUIStartDay = comboOfStartDay.get()
    getUIEndYear = comboEndOftYear.get()
    getUIEndMontn = comboEndMonth.get()
    getUIEndDay = comboEndDay.get()

    #print(getUIName, getUIDrag)
    #print(getUIStartYear, getUIStartMonth, getUIStartDay)
    #print(getUIEndYear, getUIEndMontn, getUIEndDay)

    num = countNewCards(getUIName, getUIDrag, getUIStartYear,
                        getUIStartMonth, getUIStartDay, getUIEndYear, getUIEndMontn, getUIEndDay)
    numCards.config(text=num)


def countNewCards(MRName, drug, startYear, startMonth, startDay, endYear, endMonth, endDay):
    listOfCards = []
    startDate = f'{startYear}-{startMonth}-{startDay}'
    endDate = f'{endYear}-{endMonth}-{endDay}'
    print('>countNewCards:')
    print(MRName, drug, startYear, startMonth,
          startDay, endYear, endMonth, endDay)
    for i in range(0, len(listDataTerapiya)):
        #print(listDataTerapiya[i]['Имя МП1'])
        if listDataTerapiya[i]['Имя МП1'] == MRName and listDataTerapiya[i]['Препарат'] == drug and listDataTerapiya[i]['Min of * Дата'] >= startDate and listDataTerapiya[i]['Min of * Дата'] <= endDate:
            listOfCards.append(listDataTerapiya[i])
    newCards = len(listOfCards)
    print(newCards)
    for date in listOfCards:
        print(date['Min of * Дата'])
    return newCards


root_tk = tk.Tk()
# Main settings
root_tk.config(bg='#EBEBEB')
root_tk.title('Терапія+')
root_tk.geometry('875x50+400+100')
root_tk.resizable(False, False)


# Head of table
Label(root_tk, text='№', width=3).grid(column=0, row=0)
Label(root_tk, text='Медичний представник', width=22).grid(column=1, row=0)
Label(root_tk, text='Препарат').grid(column=2, row=0)
Label(root_tk, text='', width=1).grid(column=3, row=0)
Label(root_tk, text='Дата початку').grid(
    column=4, row=0, columnspan=5)
Label(root_tk, text='', width=1).grid(column=9, row=0)
Label(root_tk, text='Дата завеншення', width=15).grid(
    column=10, row=0, columnspan=5)
Label(root_tk, text='', width=1).grid(column=15, row=0)
Label(root_tk, text='Нові пацієнти', width=15).grid(column=16, row=0)

# START TABLE
Label(root_tk, text='1').grid(column=0, row=1)
# Combo Name
variableOfName = StringVar()
comboOfSurnameOfMR = ttk.Combobox(
    root_tk, textvariable=variableOfName, values=listOfMR, width=22)
comboOfSurnameOfMR['state'] = 'readonly'
comboOfSurnameOfMR.set('Оберіть представника')
comboOfSurnameOfMR.grid(column=1, row=1)

# Combo Drug
comboOfDrag = StringVar()
comboOfDrag = ttk.Combobox(
    root_tk, textvariable=comboOfDrag, values=listOfDrags)
comboOfDrag['state'] = 'readonly'
comboOfDrag.set('Оберіть препарат')
comboOfDrag.grid(column=2, row=1)

# Combo Year start
variableOfStartYear = StringVar()
comboOfStartYear = ttk.Combobox(
    root_tk, textvariable=variableOfStartYear, values=listOfYears, width=5)
comboOfStartYear['state'] = 'readonly'
comboOfStartYear.set('2022')
comboOfStartYear.grid(column=4, row=1)

# -
Label(root_tk, justify=CENTER, text='-').grid(column=5, row=1)

# Combo Month start
variableOfStartMonth = StringVar()
comboOfStartMonth = ttk.Combobox(
    root_tk, textvariable=variableOfStartMonth, values=lisfOfMonth, width=3)
comboOfStartMonth['state'] = 'readonly'
comboOfStartMonth.set('01')
comboOfStartMonth.grid(column=6, row=1)

# -
Label(root_tk, justify=CENTER, text='-').grid(column=7, row=1)

# Combo Day start
variableOfStartDay = StringVar()
comboOfStartDay = ttk.Combobox(
    root_tk, textvariable=variableOfStartDay, values=listOfDays, width=3)
comboOfStartDay['state'] = 'readonly'
comboOfStartDay.set('31')
comboOfStartDay.grid(column=8, row=1)


# Combo end Year
variablEndeOfYear = StringVar()
comboEndOftYear = ttk.Combobox(
    root_tk, textvariable=variablEndeOfYear, values=listOfYears, width=5)
comboEndOftYear['state'] = 'readonly'
comboEndOftYear.set('2022')
comboEndOftYear.grid(column=10, row=1)
# -
Label(root_tk, justify=CENTER, text='-').grid(column=11, row=1)

# Combo Month end
variableEndMonth = StringVar()
comboEndMonth = ttk.Combobox(
    root_tk, textvariable=variableEndMonth, values=lisfOfMonth, width=3)
comboEndMonth['state'] = 'readonly'
comboEndMonth.set('12')
comboEndMonth.grid(column=12, row=1)

# -
Label(root_tk, justify=CENTER, text='-').grid(column=13, row=1)


# Combo Day start
variableOfEndDay = StringVar()
comboEndDay = ttk.Combobox(
    root_tk, textvariable=variableOfEndDay, values=listOfDays, width=3)
comboEndDay['state'] = 'readonly'
comboEndDay.set('31')
comboEndDay.grid(column=14, row=1)

numCards = Label(root_tk, text='0')
numCards.grid(row=1, column=16)

# Button calc
Button(root_tk, text='Порахувати', command=getDataFromUI).grid(row=1, column=17)

print("Start UI")
root_tk.mainloop()
