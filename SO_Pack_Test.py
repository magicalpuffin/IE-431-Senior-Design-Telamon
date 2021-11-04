import pandas as pd
# from SO_Item_Analysis_Functions import *
from SO_Item_Analysis_Functions_v2 import *

# Imports excel data and converts to dataframes
shipping_df = pd.read_excel("Data Inputs\\1 Year Shipping Orders.xlsx")
item_df_2 = pd.read_excel("Data Inputs\\Item Warehouse Data 10-5-21.xlsx")
test_bins = pd.read_excel("Data Inputs\\Bin_Pack_Test_Data.xlsx", "Bins")

# Organize shipping dataframe
# Rename item column and invert quantity
shipping_df = shipping_df.rename(columns={"Item": "Item Number"})
shipping_df.loc[:, ["Quantity"]] = -shipping_df["Quantity"]

# Data cleaning functions, filter out items missing data, filter out SO missing items
item_df_2_clean, item_df_2_excluded = clean_item_df(item_df_2)
merged_df_clean, shipping_df_item_na = clean_merged_df(shipping_df, item_df_2_clean)

indexed_df = merged_df_clean.set_index(["Sales Order Number", "Item Number"])
# ----
# Output for debug and lists of excluded items and SO
# output_df_1 = [item_df_2_clean, item_df_2_excluded, shipping_df_item_na, merged_df_clean]
# name_df_1 = ["Cleaned item data", "Excluded item data", "Missing info SO", "Merged SO Clean"]
# df_to_excel(output_df_1, name_df_1, "clean_data.xlsx")
# ----

# SO data sums, groups by SO
SO_volume = indexed_df.groupby(level = 0)["Total Volume"].sum()
SO_weight = indexed_df.groupby(level = 0)["Total Weight"].sum()
SO_quantities = indexed_df.groupby(level = 0)["Quantity"].sum()

# ----
# Remove high quantity and high volume SO for reduced run times
# Limits SO to only ones with low quantities to improve run time
SO_high_quantities = SO_quantities[SO_quantities > 25]
SO_high_vol = SO_volume[SO_volume > 12000]

indexed_df_clean = indexed_df.drop(SO_high_quantities.index.values)
indexed_df_clean = indexed_df_clean.drop(SO_high_vol.index.values)

# Packing algorithm on remaining SO
# Computationally this can start to take a long time, this ran within 2 mins
# SO with extremely high quantities will increase run time significantly
indexed_df_clean_packed_debug = indexed_df_clean.groupby(level = 0).apply(pack_SO, bin_df = test_bins)
indexed_df_clean_packed = indexed_df_clean_packed_debug[['Pack Status', 'Bin Name', 'Volume Difference', 'Weight Difference']]

output_df_3 = [indexed_df_clean, indexed_df_clean_packed, indexed_df_clean_packed_debug]
name_df_3 = ["Low quant SO", "SO Packed", "SO Packed Debug"]
df_to_excel(output_df_3, name_df_3, "Data Outputs\\packed_results_low_quantity.xlsx")