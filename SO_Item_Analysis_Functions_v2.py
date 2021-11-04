import pandas as pd
from py3dbp import Packer, Bin, Item

def df_to_excel(df_list, name_list, file_name = "py_excel_output.xlsx"):
    """
    Inputs list of dataframes and names and outputs to excel \n
    Arguments: \n
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

def remove_na_in_cols(df, col_names = []):
    """
    Removes rows with a na in the specified columns. Outputs cleaned dataframe and a dataframe what was removed. \n
    Due to the loop, the exclude reason will prioritize the first item in list. \n
    Arguments: \n
        df: Dataframe that you want remove the na from \n
        col_names: List of strings. Names of columns that should check for na values \n
    Returns: \n
        clean_df: Dataframe without na in any of the specified columns \n
        na_df: Dataframe with na in the specified columns and exclusion reason \n
    """
    clean_df = df
    na_df = pd.DataFrame()
    for col_name in col_names:
        # Creates temporary dataframe of na in column
        # Temporary dataframe assigned exclusion reason
        # Data removed from clean_df by droping the index
        # The temp_na_df is concated to the na_df
        temp_na_df = clean_df[clean_df[col_name].isna()]
        temp_na_df = temp_na_df.assign(exclude_reason = "NA in " + col_name)
        clean_df = clean_df.drop(temp_na_df.index.values)
        na_df = pd.concat([na_df, temp_na_df])
    return clean_df, na_df

def remove_val_in_cols(df, col_val_dict = {}):
    """
    Removes rows based on specified values in the specified columns. Outputs cleaned dataframe and a dataframe what was removed. \n
    Due to the loop, the exclude reason will prioritize the first item in dictionary. \n
    Arguments: \n
        df: Dataframe that you want remove rows based on values in columns \n
        col_val_dict: Dictionary with key of column names and value of list of values to remove. Value must be a list. \n
            Ex. {"Dimension UOM": ["Ft", "In"], "Weight UOM": ["Lbs"]} \n
    Returns: \n
        clean_df: Dataframe without the specified values in any of the specified columns \n
        not_val_df: Dataframe with na in the specified columns and exclusion reason \n
    """
    clean_df = df
    not_val_df = pd.DataFrame()
    for key, value in col_val_dict.items():
        # Creates temporary dataframe of incorrect val in column
        # Exclusion reason added to temporary dataframe
        # Data removed from clean_df by droping the index
        # The temp_not_val_df is concated to the not_val_df
        temp_not_val_df = clean_df[clean_df[key].isin(value) == False]
        temp_not_val_df = temp_not_val_df.assign(exclude_reason = "Unacceptable value in " + key)
        clean_df = clean_df.drop(temp_not_val_df.index.values)
        not_val_df = pd.concat([not_val_df, temp_not_val_df])
    return clean_df, not_val_df

def clean_item_df(item_df):
    """
    Cleans item dataframe, removes missing dims, weights, wrong dim, wrong weights. Returns cleaned item df and excluded items\n
    Filters in order of missing dims, missing weights, wrong dim UOM, wrong weight UOM\n
    Arguments: \n
        item_df: item dataframe, requires a standard format \n
    Returns: \n
        item_df_clean: cleaned item data with uniform units (in, lbs) and no missing values \n
        item_df_excluded: all removed items and reason of removal \n
    """
    
    item_df_clean, item_df_na = remove_na_in_cols(item_df, ["Unit Length", "Unit Width", "Unit Height", "Unit Weight", 
                                                     "Dimension UOM", "Weight UOM"])

    item_df_clean, item_df_val = remove_val_in_cols(item_df_clean, {"Dimension UOM": ["Ft", "In"], "Weight UOM": ["Lbs"]})

    # Cleaned data item data, do unit conversion to be in inches
    item_df_clean.loc[item_df_clean["Dimension UOM"] == "Ft", ["Unit Length", "Unit Width", "Unit Height"]] = item_df_clean[item_df_clean["Dimension UOM"] == "Ft"][["Unit Length", "Unit Width", "Unit Height"]] * 12
    item_df_clean.loc[item_df_clean["Dimension UOM"] == "Ft", ["Dimension UOM"]] = "In"
    item_df_clean.loc[:, ["Unit Volume"]] = item_df_clean["Unit Length"] * item_df_clean["Unit Width"] * item_df_clean["Unit Height"]

    # Keeps only useful columns
    item_df_clean = item_df_clean[["Item Number", "Description", "Item Type", "Product Category", "Unit Length", "Unit Width", "Unit Height", "Unit Weight", "Unit Volume"]]

    item_df_excluded = pd.concat([item_df_na, item_df_val])

    return item_df_clean, item_df_excluded

def clean_merged_df(shipping_df, item_df_clean):
    """
    Merges shipping dataframe with clean item dataframe to get a dataframe with shipping orders and item dims\n
    Calculates total total\n
    Arguments:\n
        shipping_df: Dataframe of all shipping orders. Needs to have Item Number column for merge \n
        item_df_clean: Item dataframe that is fully defined. All items should have all dimensions \n
    Returns:\n
        merged_df_clean: Cleaned merged dataframe with shipping orders and items \n
        shipping_df_item_na: Dataframe of shipping orders that have an item missing dimensions \n
    """
    # Merges the shipping dataframe with item dataframe on items
    merged_df = pd.merge(shipping_df, item_df_clean, on ="Item Number", how="left")
    
    # Filters out the shipping orders missing items
    shipping_df_item_na = merged_df[merged_df["Unit Length"].isna()]
    merged_df_clean = merged_df[(merged_df["Sales Order Number"].isin(shipping_df_item_na["Sales Order Number"]) == False)]

    # Finds the total volume and total weight for shipping orders
    merged_df_clean.loc[:, ["Total Volume"]] = merged_df_clean["Quantity"] * merged_df_clean["Unit Volume"]
    merged_df_clean.loc[:, ["Total Weight"]] = merged_df_clean["Quantity"] * merged_df_clean["Unit Weight"]
    
    return merged_df_clean, shipping_df_item_na

def pack_SO(df, bin_df):
    """
    Determines the best bin to pack a shipping order. \n
    Designed to be used with an apply function on a dataframe grouped by SO.\n
    Parameters: \n
        df: Dataframe of a single shipping order. Item name assumed as index. Items have quantity and dimensions.\n
        bin_df: Dataframe of bins to be used during packing. Change the default so it works with apply. \n
                I don't know if there is a better way.\n
    Returns: \n
        output_sr: Series with different output parameters
    """
    # Prepares and runs packer object
    # ----
    # Initialized packer object. From the py3dbp package.
    packer = Packer()

    # Loops through each item in SO. Will add items to packer multiple times based on quantity.
    # Index should be the item name.
    for index, row in df.iterrows():
        for i in range(int(row["Quantity"])):
            packer.add_item(Item(index, row["Unit Length"], row["Unit Width"], row["Unit Height"], row["Unit Weight"]))

    # Loops though each bins and adds them into the packer object
    for index, row in bin_df.iterrows():
        packer.add_bin(Bin(row["Bin Name"], row["Length"], row["Width"], row["Height"], row["Weight"]))
    
    # Packer evaluates how well items are packed in each bin
    packer.pack()
    # ----

    # Interpret the packed results
    # Note that values from packer use the Decimal module
    # ----
    # Finds the smallest volume bin that fits all items by looping through in order until one fits all items
    # Will default to largest bin if none can be found
    # Determines if the items were packed or not in pack_status
    bestbin = packer.bins[-1]
    pack_status = False
    for b in packer.bins:
        if len(b.unfitted_items) == 0:
            bestbin = b
            pack_status = True
            break
    
    # Creates variables for important parameters for the bin
    bin_name = bestbin.name
    bin_vol = bestbin.get_volume()
    bin_weight = bestbin.max_weight

    # Determines the total item weight and volume in packed SO for validation and comparision with bin maximums
    # Uses Decimal module. I don't know how to use it so I convert it to float
    total_item_vol = 0
    total_item_weight = 0
    for i in bestbin.items:
        total_item_vol = total_item_vol + i.get_volume()
        total_item_weight = total_item_weight + i.weight
    vol_diff = float(bin_vol - total_item_vol)
    weight_diff = float(bin_weight - total_item_weight)
    # ----

    # Output series in order to automatically create a dataframe from the apply function
    # Includes more information than needed for debugging purposes as of now
    output_sr = pd.Series({'Packer': packer, 
                           'Pack Status': pack_status, 
                           'Best Bin': bestbin, 
                           'Bin Name': bin_name, 
                           'Bin Volume': bin_vol, 
                           'Bin Weight': bin_weight, 
                           'Total Item Vol': total_item_vol, 
                           'Total Item Weight': total_item_weight, 
                           'Volume Difference': vol_diff, 
                           'Weight Difference': weight_diff})
    return output_sr