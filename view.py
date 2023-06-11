import tkinter as tk
import tkinter.messagebox as mb
import sys


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
        print('ok')
    else:
        print('Process is over by user....')
        sys.exit()

def end_message(kn_count):
    mb.showinfo(title='Tool_consumption_v1.1', message=f'Обработано {kn_count} файлов xlsx.')