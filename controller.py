import model


def start_program():
  path = model.get_xlsx_directory()
  df = model.xlsx_reading(path)
  df = model.del_NAN(df)
  df = model.converting_table(df)
  # df = model.reindex_dataframe(df)
  return df
