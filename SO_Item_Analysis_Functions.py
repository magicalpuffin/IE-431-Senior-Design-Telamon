import pandas as pd

def df_to_excel(df_list, name_list, file_name = "py_excel_output.xlsx"):
    datatoexcel = pd.ExcelWriter(file_name)
    for df, name in zip(df_list, name_list):
        df.to_excel(excel_writer = datatoexcel, sheet_name = name)
    datatoexcel.save()
    return