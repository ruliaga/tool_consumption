import model
import view


def start_program():
  
  path = model.get_xlsx_directory()
  df = model.xlsx_reading(path)
  df = model.del_NAN(df)
  df = model.converting_table(df)
  df = model.add_folder_shifr_columns(df)
  df = model.split_str(df)
  df = model.tool_consumption(df)
  model.create_xlsx(df,'Tool_consumption_v1.1.xlsx')
  
  
  return df
