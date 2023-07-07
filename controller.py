import model
import view


def start_program():
  
  path = model.get_xlsx_directory()
  df = model.xlsx_reading(path)
  df1 = model.get_df1(df)
  df2 = model.get_df2(df)
  df3 = model.get_df3(df1, df2)

  df = df3
  # df = model.del_NAN(df)
  # df = model.converting_table(df)
  df = model.add_folder_shifr_columns(df)
  df = model.split_str(df)
  df = model.tool_consumption(df)
  model.create_xlsx(df,'Tool_consumption.xlsx')
  
  
  return df
