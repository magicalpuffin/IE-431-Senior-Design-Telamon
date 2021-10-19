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
    Cleans item dataframe, removes missing dims, weights, wrong dim, wrong weights. Returns cleaned item df and excluded items\n
    Filters in order of missing dims, missing weights, wrong dim UOM, wrong weight UOM\n
    Parameters: \n
        item_df: item dataframe, requires a standard format \n
    Returns: \n
        item_df_clean: cleaned item data with uniform units (in, lbs) and no missing values \n
        item_df_excluded: all removed items and reason of removal \n
    """
    # Removes items with issues by making a dataframe of bad items and then using df.isin to filter
    # Filters in layers in order to minimize repeated items
    # Current use of item_df and item_df_no_na is a bit messy
    # Filter missing dims
    item_df_missing_dim = item_df[(item_df["Unit Length"].isna()) | 
                                    (item_df["Unit Width"].isna()) | 
                                    (item_df["Unit Height"].isna())]
    item_df_no_na = item_df[(item_df["Item Number"].isin(item_df_missing_dim["Item Number"]) == False)]

    # Filter missing weights
    item_df_missing_weight = item_df_no_na[item_df_no_na["Unit Weight"].isna()]
    item_df_no_na = item_df_no_na[(item_df_no_na["Item Number"].isin(item_df_missing_weight["Item Number"]) == False)]

    # Filter units that are not imperial
    item_df_wrong_dim_UOM =  item_df_no_na[(item_df_no_na["Dimension UOM"] != "In") & (item_df_no_na["Dimension UOM"] != "Ft")]
    item_df_no_na = item_df_no_na[(item_df_no_na["Item Number"].isin(item_df_wrong_dim_UOM["Item Number"]) == False)]

    # Filter weights not lbs, most data is usually fine
    item_df_wrong_weight_UOM = item_df_no_na[(item_df_no_na["Weight UOM"] != "Lbs")]
    item_df_clean = item_df_no_na[(item_df_no_na["Item Number"].isin(item_df_wrong_dim_UOM["Item Number"]) == False)]

    # Cleaned data item data, do unit conversion to be in inches
    item_df_clean.loc[item_df_clean["Dimension UOM"] == "Ft", ["Unit Length", "Unit Width", "Unit Height"]] = item_df_clean[item_df_clean["Dimension UOM"] == "Ft"][["Unit Length", "Unit Width", "Unit Height"]] * 12
    item_df_clean.loc[item_df_clean["Dimension UOM"] == "Ft", ["Dimension UOM"]] = "In"

    # Keeps only useful columns
    item_df_clean = item_df_clean[["Item Number", "Description", "Item Type", "Product Category", "Unit Length", "Unit Width", "Unit Height", "Unit Weight"]]

    # Adds new column to all dataframes to describe exlusion reason and concats together
    item_df_missing_dim = item_df_missing_dim.assign(exclude_reason = "Missing Dimension")
    item_df_missing_weight = item_df_missing_weight.assign(exclude_reason = "Missing Weight")
    item_df_wrong_dim_UOM = item_df_wrong_dim_UOM.assign(exclude_reason = "Wrong Dimension Units")
    item_df_wrong_weight_UOM = item_df_wrong_weight_UOM.assign(exclude_reason = "Wrong Weight Units")
    item_df_excluded = pd.concat([item_df_missing_dim, item_df_missing_weight, item_df_wrong_dim_UOM, item_df_wrong_weight_UOM])

    return item_df_clean, item_df_excluded