import controller
import globalVar
import openpyxl
import view
import os

path = os.getcwd()
try:
    os.remove(path + '/Descryption.txt')
except OSError:
    pass
try:
    os.remove(path + '/Tool_consumption.xlsx')
except OSError:
    pass

df = controller.start_program()
view.end_message(globalVar.COUNT_KN)
# 
# answer = input('Нажмите Enter для выхода из программы')
# print(df)
# model.create_xlsx(df)