import pandas as pd
from py3dbp import Packer, Bin, Item

def df_to_excel(df_list, name_list, file_name = "py_excel_output.xlsx"):
    """
    Exports an excel file of all dataframes. Uses file_name for file name and name_list for sheet names. \n
    Arguments: \n
        df_list: List of dataframes to exported to excel. List length for df and name should be the same. \n
        name_list: List of names corresponding to each dataframe. Sheet name of each dataframe. \n
        file_name: Name of the excel file to be exported. Will default to py_excel_output.xlsx \n
    """
    # Uses openpyxl and pandas
    # Loops through, assigning each dataframe to corresponding sheet name
    # This function may be better done using an dictionary
    datatoexcel = pd.ExcelWriter(file_name)
    for df, name in zip(df_list, name_list):
        df.to_excel(excel_writer = datatoexcel, sheet_name = name)
    datatoexcel.save()
    return

def remove_na_in_cols(df, col_names = []):
    """
    Removes rows with na values in the specified columns from the dataframe. Outputs a cleaned dataframe and a dataframe of what was removed. \n
    Column resulting in the removal is added to the removed dataframe. Due to the loop, the exclude reason will prioritize in the order the columns were listed. \n
    Arguments: \n
        df: Dataframe that you want remove na values from \n
        col_names: Names of columns in dataframe to remove na values from. List of strings. \n
    Returns: \n
        clean_df: Dataframe after na values removed from specified columns \n
        na_df: Dataframe of the removed values. New column added for exclusion reason \n
    """
    clean_df = df
    na_df = pd.DataFrame()
    for col_name in col_names:
        # Creates temporary dataframe with na values for the column
        # dataframe with na values dropped from clean_df and and concated with full dataframe of removed values
        temp_na_df = clean_df[clean_df[col_name].isna()]
        temp_na_df = temp_na_df.assign(exclude_reason = "NA in " + col_name)
        clean_df = clean_df.drop(temp_na_df.index.values)
        na_df = pd.concat([na_df, temp_na_df])
    return clean_df, na_df

def filter_values_in_cols(df, col_val_dict = {}):
    """
    Filters dataframe, keeping only the values specified for each column. \n
    Outputs a cleaned dataframe and a dataframe of what was removed. \n
    Due to the loop, the exclude reason will prioritize the first item in dictionary. \n
    Arguments: \n
        df: Dataframe that you want filter values from \n
        col_val_dict: Dictionary with key of column names and value of list of values to keep. Value must be a list. \n
            Ex. {"Dimension UOM": ["Ft", "In"], "Weight UOM": ["Lbs"]} \n
    Returns: \n
        clean_df: Dataframe filtered to only contain the values in the list specified for the columns \n
        removed_df: Dataframe of what was removed \n
    """
    clean_df = df
    removed_df = pd.DataFrame()
    for key, value in col_val_dict.items():
        # Key is the name of column; Value is the list of values to be kept for the column
        # Creates temporary dataframe of incorrect val in column
        # Data removed from clean_df by droping the index
        # The temp_removed_df is concated to the removed_df
        temp_removed_df = clean_df[clean_df[key].isin(value) == False]
        temp_removed_df = temp_removed_df.assign(exclude_reason = "Unacceptable value in " + key)
        clean_df = clean_df.drop(temp_removed_df.index.values)
        removed_df = pd.concat([removed_df, temp_removed_df])
    return clean_df, removed_df

class Items:
    """
    Stores and manipulates item data \n
    """
    def __init__(self, df, item_num_col, width_col, height_col, depth_col, weight_col, dim_UOM, weight_UOM):
        """
        Initializes the dataframe and renames its column names. Initializes units of measure. \n
        Assumes a generic item dataframe containing: Item Number, Width, Height, Depth, Weight. \n
        Arguments: \n
            df: Dataframe with all items \n
            item_num_col: String. Name of item number column in the dataframe \n
            width_col: String. Name of the width column in the dataframe \n
            height_col: String. Name of the height column in the dataframe \n
            depth_col: String. Name of the depth column in the dataframe \n
            weight_col: String. Name of the weight column in the dataframe \n
            dim_UOM: Unit of measure for the dimensions \n
            weight_UOM: Unit of measure for the weights \n
        """
        self.item_df = df.rename(columns = {item_num_col : "Item Number", width_col : "Width", height_col : "Height", depth_col : "Depth", weight_col : "Weight"})
        self.item_df = self.item_df[["Item Number", "Width", "Height", "Depth", "Weight"]]
        self.dim_UOM = dim_UOM
        self.weight_UOM = weight_UOM
    
    def add_volume_col(self):
        """
        Adds a new volume column to the item dataframe. Calculates for volume using width, height and depth \n
        """
        self.item_df.loc[:, ["Volume"]] = self.item_df["Width"] * self.item_df["Height"] * self.item_df["Depth"]
    
    def remove_na(self):
        """
        Removes na values from item dataframe. Outputs a dataframe of what was removed \n
        Returns: \n
            removed_na_df: Dataframe of what was removed \n
        """
        self.item_df, removed_na_df = remove_na_in_cols(self.item_df, ["Width", "Height", "Depth", "Weight"])
        return removed_na_df

    def get_df(self):
        """
        Returns the current item dataframe \n
        Returns: \n
            self.item_df: Current item dataframe \n
        """
        return self.item_df
        
class Shipments:
    """
    Stores and manipulates shipment data \n
    """
    def __init__(self, df, shipment_num_col, item_num_col, quantity_col):
        """
        Initializes the dataframe and renames its column names. \n
        Assumes a generic shipment dataframe containing: Shipment Number, Item Number, Quantity. \n
        Arguments: \n
            df: Dataframe with all shipments \n
            shipment_num_col: String. Name of shipment number column in the dataframe \n
            item_num_col: String. Name of the item number column in the dataframe \n
            quantity_col: String. Name of the quantity column in the dataframe \n
        """
        self.shipment_df = df.rename(columns = {shipment_num_col : "Shipment Number", item_num_col : "Item Number", quantity_col : "Quantity"})
        self.shipment_df = self.shipment_df[["Shipment Number", "Item Number", "Quantity"]]

    def get_df(self):
        """
        Returns the current shipment dataframe \n
        Returns: \n
            self.shipment_df: Current shipment dataframe \n
        """
        return self.shipment_df

class Shipments_Items_Merge:
    """
    Stores and manipulates the merged shipments and items data \n
    I dislike the current name of this class but I have no better ideas \n
    """
    def __init__(self, df):
        """
        Initializes the dataframe and sets the Sipment Number and Item Number as indexes. \n
        Assumes a generic merged dataframe containing: Shipment Number, Item Number, Quantity, Width, Height, Depth, Weight, Volume. \n
        Arguments: \n
            df: Dataframe of merged shipments and items. Needs to have correct column names. \n
        """
        self.merge_df = df.set_index(["Shipment Number", "Item Number"])

    def add_total_vol_weight_col(self):
        """
        Adds Total Volume and Total Weight columns to the merged dataframe. Total volume and weight are of the item type in the shipment. \n
        Item volume and item weight multiplied by quantity. \n
        """
        self.merge_df.loc[:, ["Total Volume"]] = self.merge_df["Quantity"] * self.merge_df["Volume"]
        self.merge_df.loc[:, ["Total Weight"]] = self.merge_df["Quantity"] * self.merge_df["Weight"]

    def remove_na(self):
        """
        Removes shipments that contain items missing data. Even if a shipment is only missing one item data, the entire shipment will be removed. \n
        Outputs a dataframe of what was removed. \n
        Returns: \n
            removed_na_df: Dataframe of what was removed \n
        """
        # Assumes a shipment missing width will be missing other dimensions
        # Finds shipments with items missing data and removes entire shipment
        removed_na_df = self.merge_df[self.merge_df["Width"].isna()]
        self.merge_df = self.merge_df.drop(removed_na_df.index.get_level_values(0), level = 0)
        return removed_na_df

    def get_summary(self):
        """
        Creates a shipment summary dataframe that includes: Shipment Number, Quantity, Total Volume, Total Weight. \n
        Columns describe the shipment, the total quantity of items, total volume and total weight for each shipment number. \n
        Returns: \n
            self.SN_summary_df: Summary dataframe of shipment numbers\n
        """
        SN_quantities = self.merge_df.groupby(level = 0)["Quantity"].sum()
        SN_volume = self.merge_df.groupby(level = 0)["Total Volume"].sum()
        SN_weight = self.merge_df.groupby(level = 0)["Total Weight"].sum()
        # Quantity of items in SN, total volume of items in SN, total weight of items in SN
        self.SN_summary_df = pd.concat([SN_quantities, SN_volume, SN_weight], axis = 1)
        return self.SN_summary_df
    
    def filter_SN(self, max_quantity, max_volume):
        """
        Filters the merged dataframe based on max shipment quantity and volume. Requires a self.SN_summary_df. \n
        Removes shipment numbers from merged dataframe. \n
        Arguments: \n
            max_quantity: Maximum quantity of items within a shipment \n
            max_volume: Maximum volume in a shipment \n
        Returns: \n
            removed_df: Dataframe of what was removed \n
        """
        SO_high_quantities = self.SN_summary_df[self.SN_summary_df["Quantity"] > max_quantity]
        SO_high_vol = self.SN_summary_df[self.SN_summary_df["Total Volume"] > max_volume]

        self.merge_df = self.merge_df.drop(SO_high_quantities.index.values, level = 0)
        self.merge_df = self.merge_df.drop(SO_high_vol.index.values, level = 0)

        removed_df = pd.concat([SO_high_quantities, SO_high_vol])
        return removed_df

    def get_df(self):
        """
        Returns the current merged dataframe \n
        Returns: \n
            self.merge_df: Current merged dataframe \n
        """
        return self.merge_df

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
            packer.add_item(Item(index, row["Width"], row["Height"], row["Depth"], row["Weight"]))

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