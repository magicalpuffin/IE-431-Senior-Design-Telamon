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

def filter_values_in_cols(df, col_val_dict = {}):
    """
    Filters dataframe to only keep rows with the values specified for each column. \n
    Outputs a cleaned dataframe and a dataframe of what was removed. \n
    Due to the loop, the exclude reason will prioritize the first item in dictionary. \n
    Arguments: \n
        df: Dataframe that you want remove rows based on values in columns \n
        col_val_dict: Dictionary with key of column names and value of list of values to keep. Value must be a list. \n
            Ex. {"Dimension UOM": ["Ft", "In"], "Weight UOM": ["Lbs"]} \n
    Returns: \n
        clean_df: Dataframe without the specified values in any of the specified columns \n
        removed_df: Dataframe with na in the specified columns and exclusion reason \n
    """
    clean_df = df
    removed_df = pd.DataFrame()
    for key, value in col_val_dict.items():
        # Key is the name of column
        # Value is the list of values to keep within column
        # Creates temporary dataframe of incorrect val in column
        # Exclusion reason added to temporary dataframe
        # Data removed from clean_df by droping the index
        # The temp_removed_df is concated to the removed_df
        temp_removed_df = clean_df[clean_df[key].isin(value) == False]
        temp_removed_df = temp_removed_df.assign(exclude_reason = "Unacceptable value in " + key)
        clean_df = clean_df.drop(temp_removed_df.index.values)
        removed_df = pd.concat([removed_df, temp_removed_df])
    return clean_df, removed_df

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

    item_df_clean, item_df_val = filter_values_in_cols(item_df_clean, {"Dimension UOM": ["Ft", "In"], "Weight UOM": ["Lbs"]})

    # Cleaned data item data, do unit conversion to be in inches
    item_df_clean.loc[item_df_clean["Dimension UOM"] == "Ft", ["Unit Length", "Unit Width", "Unit Height"]] = item_df_clean[item_df_clean["Dimension UOM"] == "Ft"][["Unit Length", "Unit Width", "Unit Height"]] * 12
    item_df_clean.loc[item_df_clean["Dimension UOM"] == "Ft", ["Dimension UOM"]] = "In"
    item_df_clean.loc[:, ["Unit Volume"]] = item_df_clean["Unit Length"] * item_df_clean["Unit Width"] * item_df_clean["Unit Height"]

    # Keeps only useful columns
    item_df_clean = item_df_clean[["Item Number", "Description", "Item Type", "Product Category", "Unit Length", "Unit Width", "Unit Height", "Unit Weight", "Unit Volume"]]

    item_df_excluded = pd.concat([item_df_na, item_df_val])

    return item_df_clean, item_df_excluded

def clean_merged_df(merged_df):
    """
    Merges shipping dataframe with clean item dataframe to get a dataframe with shipping orders and item dims\n
    Calculates total total\n
    Arguments:\n
        merged_df: merged data of shipping and items. Dataframe of all shipping orders. Needs to have Item Number column for merge \n
    Returns:\n
        merged_df_clean: Cleaned merged dataframe with shipping orders and items \n
        shipping_df_item_na: Dataframe of shipping orders that have an item missing dimensions \n
    """
    
    # Filters out the shipping orders missing items
    shipping_df_item_na = merged_df[merged_df["Unit Length"].isna()]
    merged_df_clean = merged_df[(merged_df["Sales Order Number"].isin(shipping_df_item_na["Sales Order Number"]) == False)]

    # Finds the total volume and total weight for shipping orders
    # I shouldn't be calculating these in this function
    merged_df_clean.loc[:, ["Total Volume"]] = merged_df_clean["Quantity"] * merged_df_clean["Unit Volume"]
    merged_df_clean.loc[:, ["Total Weight"]] = merged_df_clean["Quantity"] * merged_df_clean["Unit Weight"]
    
    return merged_df_clean, shipping_df_item_na

def get_SO_summary(SO_indexed_df):
    SO_quantities = SO_indexed_df.groupby(level = 0)["Quantity"].sum()
    SO_volume = SO_indexed_df.groupby(level = 0)["Total Volume"].sum()
    SO_weight = SO_indexed_df.groupby(level = 0)["Total Weight"].sum()
    # Quantity of items in SO, total volume of items in SO, total weight of items in SO
    SO_summary_df = pd.concat([SO_quantities, SO_volume, SO_weight], axis = 1)
    return SO_summary_df

def get_solution_bin(packer):
    """
    Finds the best bin that packed the items based on packing object. Will pick bin with smallest volume. \n
    Defaults to the largest bin if packing failed. Outputs bin and true false status. \n
    """
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
    return bestbin, pack_status

def get_excess_vol_weight(bin):
    """
    Takes a bin and finds the extra volume and weight left by the fitted items
    """
    total_item_vol = 0
    total_item_weight = 0
    for i in bin.items:
        total_item_vol = total_item_vol + i.get_volume()
        total_item_weight = total_item_weight + i.weight

    # Determines the total item weight and volume in packed SO for validation and comparision with bin maximums
    # Uses Decimal module. I don't know how to use it so I convert it to float

    vol_diff = float(bin.get_volume() - total_item_vol)
    weight_diff = float(bin.max_weight - total_item_weight)
    vol_util = float(total_item_vol / bin.get_volume())
    weight_util = float(total_item_weight / bin.max_weight)

    return vol_diff, weight_diff, vol_util, weight_util

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
    # I want to switch this to use indexes but it breaks the package for some reason
    for index, row in bin_df.iterrows():
        packer.add_bin(Bin(row["Bin Name"], row["Length"], row["Width"], row["Height"], row["Weight"]))
    
    # Packer evaluates how well items are packed in each bin
    packer.pack()
    # ----
    # Interpret the packed results
    # Note that values from packer use the Decimal module
    bestbin, pack_status = get_solution_bin(packer)
    vol_diff, weight_diff, vol_util, weight_util = get_excess_vol_weight(bestbin)
    bin_name = bestbin.name
    # ----
    # Output series in order to automatically create a dataframe from the apply function
    # packer and bestbin are included for debugging purposes
    output_sr = pd.Series({'Pack Status': pack_status, 
                           'Bin Name': bin_name, 
                           'Volume Difference': vol_diff, 
                           'Weight Difference': weight_diff,
                           'Volume Utilization': vol_util,
                           'Weight Utilization': weight_util})
    return output_sr