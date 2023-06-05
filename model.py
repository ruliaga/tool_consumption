import pandas as pd
import glob 
import os


def get_xlsx_directory (): 
    path = os.path.dirname(os.path.realpath(__file__)) #функция читает текущее расположение файла py
    xlsx_directory = glob.glob(path + "/*.xlsx") #находит файлы xlsx и создает список из названий
    print(xlsx_directory)
    return xlsx_directory #возвращает этот список



def xlsx_reading(xlsx_directory): #функция создает датафрейм из файла xlsx
    df = pd.read_excel(str(xlsx_directory[0]),sheet_name='TDSheet')
    return df

def create_xlsx(df):
    df.to_excel('Operations.xlsx')

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
    df = pd.DataFrame.from_dict(dict, orient='index')
  
    

    return df