import os
import pandas as pd
import xlwings as xw

month = 'Июль'

os.chdir('Excel_files')

main_file = pd.read_excel(r'Расчет Администрация2023_2.xlsx', sheet_name=month)

table_file = pd.read_excel(r'1. ТАБЕЛЬ ТЕХНИЧЕСКАЯ С СЕНТЯБРЯ 2022.xlsx', sheet_name=month)
advances_nal = pd.read_excel(r'Авнсы и ЗП.xlsx', sheet_name='Аванс Нал')
advances_karta = pd.read_excel(r'Авнсы и ЗП.xlsx', sheet_name='Белый Аванс')
advances_karta = pd.read_excel(r'Авнсы и ЗП.xlsx', sheet_name='Белая ЗП')

#

df_combaine = pd.merge(main_file, table_file, left_on='ФИО', right_on='Офис', how='left')
# df_combaine.drop(df_combaine.columns[
# 	                 [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27,
# 	                  28, 29, 30]],
#                  axis=1, inplace=True)


df_combaine.to_excel('Newfile.xlsx', index=False)
# with pd.ExcelWriter('Newfile.xlsx', mode='w') as writer:
#     df_combaine.to_excel(writer, sheet_name='sheet_1')
#     writer._save()

#df_combaine.to_excel('Newfile.xlsx')

df_combaine = pd.merge(main_file, advances_nal, left_on='ФИО', right_on='Фамилия', how='left')
df_combaine.to_excel('Newfile.xlsx', index=False)

#print(df_combaine1)

wb = xw.Book('Newfile.xlsx')
#
#
#
