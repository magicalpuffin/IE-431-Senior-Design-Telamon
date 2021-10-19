import pandas as pd

def df_to_excel(df_list, name_list, file_name = "py_excel_output.xlsx"):
    """
    Inputs list of dataframes and names and outputs to excel \n
    Parameters: \n
        df_list: list of dataframes to be export to excel file \n
        name_list: list of names to assigned to each dataframe exported \n
        file_name: file name of excel sheet, defaults to py_excel_output.xlsx \n
    """
    # Uses openpyxl and pandas, loops through sheets and a excel writer
    datatoexcel = pd.ExcelWriter(file_name)
    for df, name in zip(df_list, name_list):
        df.to_excel(excel_writer = datatoexcel, sheet_name = name)
    datatoexcel.save()
    return

def item_df_clean(item_df):
    """
    Cleans item dataframe, removes missing dims, weights, wrong dim, wrong weights. Returns cleaned item df \n
    Should be modified to return data with issues \n
    Returns: \n
        item_df: item dataframe, requires a standard format
    """
    # Removes unnecessary items by creating dataframes of items with issues and using df.isin to filter
    # Creates dataframe of items missing dims and weights
    # Most data with one unit dim have the rest but this filter is safer
    item_df_missing_dim = item_df[(item_df["Unit Length"].isna()) | 
                                    (item_df["Unit Width"].isna()) | 
                                    (item_df["Unit Height"].isna())]
    item_df_missing_weight = item_df[item_df["Unit Weight"].isna()]

    # Item df of all items that aren't missing dims and weights
    item_df_no_na = item_df[(item_df["Item Number"].isin(item_df_missing_weight["Item Number"]) == False) & 
                        (item_df["Item Number"].isin(item_df_missing_dim["Item Number"]) == False)]

    # Of the items that have dims and weights, filter our items in wrong units and weights
    # Very few items have mislabled units, safer to filter
    item_df_wrong_dim_UOM =  item_df_no_na[(item_df_no_na["Dimension UOM"] != "In") & (item_df_no_na["Dimension UOM"] != "Ft")]
    item_df_wrong_weight_UOM = item_df_no_na[(item_df_no_na["Weight UOM"] != "Lbs")]

    # Cleaned data item data, currently missing unit conversions
    item_df_clean = item_df_no_na[(item_df_no_na["Item Number"].isin(item_df_wrong_dim_UOM["Item Number"]) == False) & 
                        (item_df_no_na["Item Number"].isin(item_df_wrong_weight_UOM["Item Number"]) == False)]

    # Keeps only useful columns
    item_df_clean = item_df_clean[["Item Number", "Description", "Item Type", "Product Category", "Unit Length", "Unit Width", "Unit Height", "Unit Weight"]]

    return item_df_clean