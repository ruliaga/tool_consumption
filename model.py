import pandas as pd
import glob 
import os
import view
import globalVar
import openpyxl

def get_xlsx_directory (): 
    #path = os.path.dirname(os.path.realpath(__file__)) #функция читает текущее расположение файла py
    path = os.getcwd()
    xlsx_directory = glob.glob(path + "/*.xlsx") #находит файлы xlsx и создает список из названий
    print(xlsx_directory)
    return xlsx_directory #возвращает этот список



def xlsx_reading(xlsx_directory): #функция создает датафрейм из файла xlsx
    df = pd.read_excel(str(xlsx_directory[0]),sheet_name='TDSheet')
    return df
   

def create_xlsx(df, message):
    df.to_excel(message)

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
        if str(df['Номенклатура'].values[i]) not in dict:
            dict[df['Номенклатура'].values[i]] = df['Количество'].values[i]
        else:
            dict[df['Номенклатура'].values[i]] += df['Количество'].values[i]
    df = pd.DataFrame.from_dict(dict, orient='index').reset_index()
    df.columns = ['Номенклатура', 'Количество']
    for i in range(0,df.shape[0]):
        if pd.isna(df['Количество'].values[i]):
            df['Количество'].values[i] = 0
    return df

def add_folder_shifr_columns(df):
    df.insert(2,'Папка','')
    df.insert(3,'Шифр','')
    return df

def split_str(df):
    for i in range(0,df.shape[0]):
       # df['Папка'][i] = re.split('\s', str(df['Продукция'][i]))[0]
        df['Папка'].values[i] = str(df['Номенклатура'].values[i]).split(' ',1)[0]
        if len(str(df['Папка'].values[i])) > 8:
            df['Папка'].values[i] = str(df['Номенклатура'].values[i]).split('.',1)[0]
    for i in range(0,df.shape[0]):
        df['Шифр'].values[i] = str(df['Номенклатура'].values[i]).split(' ',1)[0]
        if len(str(df['Шифр'].values[i])) < 9:
            try: # Вылетает исключение IndexError
                df['Шифр'][i] = str(df['Номенклатура'].values[i]).split(' ',2)[0] + ' ' + str(df['Номенклатура'].values[i]).split(' ',2)[1]
            except IndexError as ie:
                print(f'Index error')            
    return df    

def tool_consumption(df):
    tool_dict = {}
    path = 'R:\\dmg\\MSCDATA\\NC program'
    #path = 'C:\\Users\\rulia\\Desktop\\ms_data'
    for i in range(0,df.shape[0]):
        folder = df['Папка'].values[i]
        shifr = df['Шифр'].values[i]
        kol_vo = df['Количество'].values[i]
        with open("sample.txt", "a") as file_object:
                    file_object.write(df['Шифр'].values[i] + '\n')
        xlsx_directory = glob.glob(path + f"\\{folder}\\{shifr}*\\*.xlsx")
        # if len(xlsx_directory)!=0:
        for i in range(0, len(xlsx_directory)):
              if '~$' in xlsx_directory[i]:
                continue
              else:
                print(xlsx_directory[i])
                with open("sample.txt", "a") as file_object:
                    file_object.write('*******' + xlsx_directory[i] + '\n')
                globalVar.count_kn +=1
                print('Количество деталей = ' + str(kol_vo))
                df_kn = pd.read_excel(f'{str(xlsx_directory[i])}',engine='openpyxl')
                kadrNo_str_index = df_kn.index[df_kn.isin(['Кадр №']).any(axis=1)].values[0]
                df_kn = df_kn.tail(-kadrNo_str_index)
                df_kn = df_kn.drop(df_kn.tail(2).index)
                df_kn = df_kn.reset_index(drop=True)
                df_kn.columns = df_kn.iloc[0]
                df_kn = df_kn[1:]
                try:
                    df_kn = df_kn[['Имя инструмента','Расход инстр. На 1-ну дет.']]
                    df_kn['Расход инстр. На 1-ну дет.'] = df_kn['Расход инстр. На 1-ну дет.'].astype(float)
                except KeyError as ke:
                    view.window_keyError(xlsx_directory[i])
                    continue
                for i in range(0,df_kn.shape[0]):
                    if pd.isna(df_kn['Расход инстр. На 1-ну дет.'].values[i]):
                        print(df_kn['Имя инструмента'].values[i], " ", df_kn['Расход инстр. На 1-ну дет.'].values[i])
                        try:
                            view.window_ColumnValuesNanError(xlsx_directory[i])
                        except IndexError as ie:
                            continue
                df_kn.insert(2,'Шифр', shifr)
                df_kn.insert(3,'Количество деталей', kol_vo)
                df_kn.insert(4,'Суммарный расход', float(kol_vo)*df_kn['Расход инстр. На 1-ну дет.'])
                for i in range(0, df_kn.shape[0]):
                    if str(df_kn['Имя инструмента'].values[i]) not in tool_dict:
                        tool_dict[df_kn['Имя инструмента'].values[i]] = df_kn['Суммарный расход'].values[i]
                    else:
                        tool_dict[df_kn['Имя инструмента'].values[i]] += df_kn['Суммарный расход'].values[i]
                df_tool = pd.DataFrame.from_dict(tool_dict, orient='index').reset_index()
                df_tool.columns = ['Имя инструмента', 'Суммарный расход']
                
                
                print(df_tool)
             
        
    return df_tool
