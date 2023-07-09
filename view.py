import tkinter as tk
import tkinter.messagebox as mb
import sys
import globalVar



def window_keyError(directory):
    answer = mb.askyesno(
    title="Tool_consumption_v1.1", 
    message=f"Ошибка чтения файла {directory}. Неверно заполнена карта наладки. \n. Продолжить выполнение программы?\n")

    if answer:
        print('ok')
    else:
        print('Process is over by user....')
        sys.exit()

def window_ColumnValuesNanError(directory):
    answer = mb.askyesno(
    title="Tool_consumption_v1.1", 
    message=f'Ошибка чтения файла {directory}. значение в столбце "Расход инстр. На 1-ну дет." не заполнено \n. Продолжить выполнение программы?\n')

    if answer:
        print('\n')
    else:
        print('Process is over by user')
        sys.exit()

def window_dict_tool_sum_error(directory):
    answer = mb.askyesno(
    title="Tool_consumption_v1.1", 
    message=f'{directory}  - Tool {globalVar.CURRENT_TOOL}. Ошибка заполнения справочника при суммировании расхода\n')    
    
    if answer:
        print('\n')
    else:
        print('Process is over by user')
        sys.exit()

def window_dict_tool_new_item (directory):
    answer = mb.askyesno(
    title="Tool_consumption_v1.1", 
    message=f'{directory}  - Tool {globalVar.CURRENT_TOOL}. Ошибка заполнения справочника при добавлении нового инструмента\n')    
    
    if answer:
        print('\n')
    else:
        print('Process is over by user')
        sys.exit()

def end_message(kn_count):
    mb.showinfo(title='Tool_consumption_v1.1', message=f'Обработано {kn_count} файлов xlsx.')

def start_error_message():
    mb.showinfo(title='Tool_consumption_v1.1', message=f'Удалите скрытые строки в файле xlsx и преобразуйте диапазон данных в таблицу.')

def net_kn_error_message():
    mb.showinfo(title='Tool_consumption_v1.1', message=f'Карты наладки для данной заявки не найдены.')