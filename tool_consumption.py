import controller
import globalVar
import openpyxl
import view


df = controller.start_program()
view.end_message(globalVar.COUNT_KN)
# 
# answer = input('Нажмите Enter для выхода из программы')
# print(df)
# model.create_xlsx(df)