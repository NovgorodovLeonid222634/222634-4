import csv
import pandas as pd
import numpy as np
import xlwings as xw
recipes=pd.read_csv('recipes_sample.csv',parse_dates=['submitted'])
reviews=pd.read_csv('reviews_sample.csv')
recipes=pd.concat([recipes['id'],recipes['name'],recipes['minutes'],recipes['submitted'],recipes['description'],recipes['n_ingredients']],sort=False,axis=1)
print('1. Задание выполнено!')
A=recipes.sample(frac=0.05)
B=reviews.sample(frac=0.05)
obj=pd.ExcelWriter('recipes.xlsx',engine='xlsxwriter')
A.to_excel(obj, sheet_name='Рецепты',index=False)
B.to_excel(obj, sheet_name='Отзывы',index=False)
obj.close()
print('2. Задание выполнено!')
wb=xw.Book('recipes.xlsx')
sheet=wb.sheets['Рецепты']
sheet.range('G1').value='seconds_assign'
sheet.range('G2').options(pd.Series,index=False,header=False).value = A['minutes']*60
print('3. Задание выполнено!')
sheet.range('H1').value='seconds_formula'
sheet.range('H2:H1501').formula_array='=C2:C1501*60'
print('4. Задание выполнено!')
sheet.range('H1:G1501').api.Font.Bold = True
sheet['H1:G1501'].api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
print('5. Задание выполнено!')
sheet.range('C2:C1501').color=(255,255,0)
for i in sheet.range('C2:C1501'):
    print(sheet.range(i))
    if sheet.range(i).value <=5:
        sheet.range(i).color=(0,255,0)
    elif sheet.range(i).value >10:
        sheet.range(i).color=(255,0,0)
print('6. Задание выполнено!')


