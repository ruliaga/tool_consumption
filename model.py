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
    try:
        df = pd.read_excel(str(xlsx_directory[0]),sheet_name='TDSheet')
    except KeyError:
        view.start_error_message()
    except IndexError:
        view.net_zayavok_error_message()
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
            globalVar.COUNT_NOMENKLATURA +=1
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
        if str(df['Папка'].values[i])=='ПК':
                df['Шифр'].values[i] = str(df['Номенклатура'].values[i]).split(' ',3)[1] + ' ' + str(df['Номенклатура'].values[i]).split(' ',3)[2]       
        if len(str(df['Шифр'].values[i])) < 9 and not pd.isna(str(df['Номенклатура'].values[i])):
            try: # Вылетает исключение IndexError - пустое значение столбца номенклатура
                df['Шифр'].values[i]= str(df['Номенклатура'].values[i]).split(' ',2)[0] + ' ' + str(df['Номенклатура'].values[i]).split(' ',2)[1]
            except IndexError as ie:
                continue  
                
    print(df)
    return df    

def tool_consumption(df):
    tool_dict = {}
    path = 'R:\\dmg\\MSCDATA\\NC program'
    # path = 'C:\\Users\\rulia\\Desktop\\ms_data'
    for i in range(0,df.shape[0]):
        folder = df['Папка'].values[i]
        shifr = df['Шифр'].values[i]
        kol_vo = df['Количество'].values[i]
        nomenklatura = df['Номенклатура'].values[i]
        with open("Descryption.txt", "a", encoding='utf-16') as file_object:
                    file_object.write('\n' + '_'*117 + '\n\n')
                    file_object.write(str(nomenklatura) + "\n" + "Количество: " + str(kol_vo) +' шт\n' + f'Папка для поиска: {str(folder)}\n')
        xlsx_directory = glob.glob(path + f"\\{folder}\\{shifr}*\\*.xlsx")
        if len(xlsx_directory)==0:
            xlsx_directory = glob.glob(path + f"\\{folder}\\*\\{shifr}*\\*.xlsx")   
        if len(xlsx_directory)==0 and folder=='РЦО':
            xlsx_directory = glob.glob(path + f"\\{folder}\\*\\*\\{shifr}*\\*.xlsx")
        if len(xlsx_directory)==0:
            with open("Descryption.txt", "a", encoding='utf-16') as file_object:
                    file_object.write('\nКарты наладки не обнаружены')
        for i in range(0, len(xlsx_directory)):
              if '~$' in xlsx_directory[i]:
                continue
              else:
                print(xlsx_directory[i])
                with open("Descryption.txt", "a", encoding='utf-16') as file_object:
                    file_object.write('\n\n------------Чтение-карты-наладки-------------------------------------------------------------------------------------\n')
                    file_object.write('\n' + xlsx_directory[i] + '\n')
                    file_object.write('\n------------Обнаружены_cледущие_элементы_для_добавления_в_справочник-------------------------------------------------\n\n')
                globalVar.COUNT_KN +=1
                print('Количество деталей = ' + str(kol_vo))
                df_kn = pd.read_excel(f'{str(xlsx_directory[i])}',engine='openpyxl')
                kadrNo_str_index = df_kn.index[df_kn.isin(['Кадр №']).any(axis=1)].values[0]
                df_kn = df_kn.tail(-kadrNo_str_index)
                df_kn = df_kn.drop(df_kn.tail(1).index)
                df_kn = df_kn.reset_index(drop=True)
                df_kn.columns = df_kn.iloc[0]
                df_kn = df_kn[1:]
                try:
                    df_kn = df_kn[['Имя инструмента','Расход инстр. На 1-ну дет.']]
                    df_kn['Расход инстр. На 1-ну дет.'] = df_kn['Расход инстр. На 1-ну дет.'].astype(float)
                    df_kn = df_kn.dropna(subset=['Имя инструмента'])
                except KeyError as ke:
                    view.window_keyError(xlsx_directory[i])
                    with open("Descryption.txt", "a", encoding='utf-16') as file_object:
                        file_object.write(f'TypeError: {xlsx_directory} - неправильно оформлена КН')
                    continue
                except ValueError as ve:
                    view.window_ColumnValuesNanError(xlsx_directory[i])
                    continue
                # for i in range(0,df_kn.shape[0]):
                #     if pd.isna(df_kn['Расход инстр. На 1-ну дет.'].values[i]):
                #         print(df_kn['Имя инструмента'].values[i], " ", df_kn['Расход инстр. На 1-ну дет.'].values[i])
                #         try:
                #             view.window_ColumnValuesNanError(xlsx_directory[i])
                #         except IndexError as ie:
                #             continue
                df_kn.insert(2,'Шифр', shifr)
                df_kn.insert(3,'Количество деталей', kol_vo)
                df_kn.insert(4,'Суммарный расход', float(kol_vo)*df_kn['Расход инстр. На 1-ну дет.'])
                
                for i in range(0,df_kn.shape[0]):
                    name_tool = str(df_kn['Имя инструмента'].values[i])
                    globalVar.CURRENT_TOOL = name_tool
                    sum_consump = float(df_kn['Суммарный расход'].values[i])
                    with open("Descryption.txt", "a", encoding='utf-16') as file_object:
                        file_object.write(f'{name_tool}: +' + f'{sum_consump}'+' шт\n')
                
                
                for i in range(0, df_kn.shape[0]):
                    if str(str(df_kn['Имя инструмента'].values[i])) not in tool_dict:
                        try:
                            tool_dict[str(df_kn['Имя инструмента'].values[i])] = round(float(df_kn['Суммарный расход'].values[i]),3)
                        except TypeError as te:
                            with open("Descryption.txt", "a", encoding='utf-16') as file_object:
                                file_object.write(f'TypeError: {xlsx_directory}  -  {name_tool} - ошибка при добавлении нового инструмента')
                            try:
                                view.window_dict_tool_new_item(xlsx_directory[i])
                            except IndexError as ie:
                                continue
                            continue
                    else:
                        try:
                            tool_dict[str(df_kn['Имя инструмента'].values[i])] += round(float(df_kn['Суммарный расход'].values[i]),3)
                        except TypeError as te:
                            with open("Descryption.txt", "a", encoding='utf-16') as file_object:
                                file_object.write(f'TypeError: {xlsx_directory}  -  {name_tool} - ошибка при суммировании расхода')
                            try:
                                view.window_dict_tool_sum_error(xlsx_directory[i])
                            except IndexError as ie:
                                continue
                            continue
                df_tool = pd.DataFrame.from_dict(tool_dict, orient='index').reset_index()
                df_tool.columns = ['Имя инструмента', 'Суммарный расход']
                with open("Descryption.txt", "a", encoding='utf-16') as file_object:
                        file_object.write('\n------------Обновление-справочника-расхода-инструмента--------------------------------------------------------------\n\n\n')
                        file_object.write(str(df_tool))
               
                
                print(df_tool)
             
    try:    
        return df_tool
    except UnboundLocalError:
        view.net_kn_error_message()


# С изменением начальных данных (через универсальный отчет)

# def get_df1(df):
#     df1 = df[['Заказ покупателя.Номер','Номенклатура.Наименование','Количество']]
#     df1 = df1.drop_duplicates(subset = ['Заказ покупателя.Номер'], keep = 'first')
#     dict = {}
#     for i in range(0, df1.shape[0]):
#         if str(df1['Номенклатура.Наименование'].values[i]) not in dict:
#             dict[df1['Номенклатура.Наименование'].values[i]] = df1['Количество'].values[i]
#         else:
#             dict[df1['Номенклатура.Наименование'].values[i]] += df1['Количество'].values[i]
#     df1 = pd.DataFrame.from_dict(dict, orient='index').reset_index()
#     df1.columns = ['Номенклатура', 'Количество']
#     return df1

# def get_df2(df):
#     df2 = df[['Номенклатура.Наименование','Спецификация.Исходные комплектующие.Номенклатура.Наименование','Спецификация.Исходные комплектующие.Количество','Спецификация.Исходные комплектующие.Вид воспроизводства']]
#     df2 = df2.dropna(axis=0,how='any')
#     df2 = df2.drop_duplicates(subset = ['Спецификация.Исходные комплектующие.Номенклатура.Наименование'], keep = 'first')
#     df2 = df2[(df2['Спецификация.Исходные комплектующие.Вид воспроизводства'] == 'Производство') | pd.isna(df2['Спецификация.Исходные комплектующие.Вид воспроизводства'])]
#     df2 = df2.reset_index()
#     df2 = df2[['Номенклатура.Наименование','Спецификация.Исходные комплектующие.Номенклатура.Наименование','Спецификация.Исходные комплектующие.Количество']]
#     return df2

# def get_df3(df1, df2):
#     df2.insert(3, 'Количество.Сборка', 0)
#     for i in range(0,df1.shape[0]):
#         for j in range(0,df2.shape[0]):
#             if str(df2['Номенклатура.Наименование'].values[j]) == str(df1['Номенклатура'].values[i]):
#                if not pd.isna(df1['Количество'].values[i]):
#                     df2['Количество.Сборка'].values[j] = df1['Количество'].values[i]
#                else: 
#                     df2['Количество.Сборка'].values[j] = 0
#     df2.insert(4,'Общее количество', df2['Спецификация.Исходные комплектующие.Количество']*df2['Количество.Сборка'])
#     df3 = df2[['Спецификация.Исходные комплектующие.Номенклатура.Наименование','Количество.Сборка','Общее количество']]
#     df3 = df3[df3['Общее количество'] !=0]
#     dict = {}
#     for i in range(0, df3.shape[0]):
#         if str(df3['Спецификация.Исходные комплектующие.Номенклатура.Наименование'].values[i]) not in dict:
#             dict[df3['Спецификация.Исходные комплектующие.Номенклатура.Наименование'].values[i]] = df3['Общее количество'].values[i]
#         else:
#             dict[df3['Спецификация.Исходные комплектующие.Номенклатура.Наименование'].values[i]] += df3['Общее количество'].values[i]
#     for i in range(0, df1.shape[0]):
#         if str(df1['Номенклатура'].values[i]) not in dict:
#             dict[df1['Номенклатура'].values[i]] = df1['Количество'].values[i]
#         else:
#             dict[df1['Номенклатура'].values[i]] += df1['Количество'].values[i]
#     df3 = pd.DataFrame.from_dict(dict, orient='index').reset_index()
#     df3.columns = ['Номенклатура', 'Количество']
#     with open("Descryption.txt", "a", encoding='utf-16') as file_object:
#             file_object.write('------------Список-номенклатуры-для-поиска:--------------------------------------------------------\n\n')
#     for i in range(0,df3.shape[0]):
#         name_nomenklatura = str(df3['Номенклатура'].values[i])
#         kolichestvo = float(df3['Количество'].values[i])
#         with open("Descryption.txt", "a", encoding='utf-16') as file_object:
#             file_object.write(f'{name_nomenklatura}     Количество: ' + f'{kolichestvo}'+' шт\n\n')

#     return df3


# датафрэйм формируется из отчета "Планирование производства и закупок"

def get_df_planirovanie(df):
    df = df[['Номенклатура','Количество']]
    dict = {}
    for i in range(0, df.shape[0]):
        if str(df['Номенклатура'].values[i]) not in dict:
            dict[df['Номенклатура'].values[i]] = df['Количество'].values[i]
        else:
            dict[df['Номенклатура'].values[i]] += df['Количество'].values[i]
    df = pd.DataFrame.from_dict(dict, orient='index').reset_index()
    df.columns = ['Номенклатура', 'Количество']
    return df
