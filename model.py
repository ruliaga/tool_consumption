import pandas as pd
import glob 
import os
import re


def get_xlsx_directory (): 
    path = os.path.dirname(os.path.realpath(__file__)) #функция читает текущее расположение файла py
    xlsx_directory = glob.glob(path + "/*.xlsx") #находит файлы xlsx и создает список из названий
    print(xlsx_directory)
    return xlsx_directory #возвращает этот список



def xlsx_reading(xlsx_directory): #функция создает датафрейм из файла xlsx
    df = pd.read_excel(str(xlsx_directory[0]),sheet_name='TDSheet')
    return df

def create_xlsx(df):
    df.to_excel('1.xlsx')

def reindex_dataframe(df):
    df = df.reset_index(drop=True)
    return df

def sort_dataframe(df): # сортировка по трем столбцам
    df = df.sort_values(['Ссылка.Номер', 'Ссылка.Дата','Номер операции'])
    return df

def del_NAN(df):
    df.dropna(axis=0,how='any')
    return df

def converting_table(df):
    df = df[['Номенклатура','Количество']]
    dict = {}
    for i in range(0, df.shape[0]):
        if str(df['Номенклатура'][i]) not in dict:
            dict[df['Номенклатура'][i]] = df['Количество'][i]
        else:
            dict[df['Номенклатура'][i]] += df['Количество'][i]
    df = pd.DataFrame.from_dict(dict, orient='index').reset_index()
    df.columns = ['Продукция', 'Количество']

    return df

def add_folder_shifr_columns(df):
    df.insert(2,'Папка','')
    df.insert(3,'Шифр','')
    return df

def split_str(df):
    for i in range(0,df.shape[0]):
       # df['Папка'][i] = re.split('\s', str(df['Продукция'][i]))[0]
        df['Папка'][i] = str(df['Продукция'][i]).split(' ',1)[0]
        if len(str(df['Папка'][i])) > 8:
            df['Папка'][i] = str(df['Продукция'][i]).split('.',1)[0]
    for i in range(0,df.shape[0]):
        df['Шифр'][i] = str(df['Продукция'][i]).split(' ',1)[0]
        if len(str(df['Шифр'][i])) < 9:
            try: # Вылетает исключение IndexError
                df['Шифр'][i] = str(df['Продукция'][i]).split(' ',2)[0] + ' ' + str(df['Продукция'][i]).split(' ',2)[1]
            except IndexError as ie:
                print(f'Index error')
                
    return df    

       



