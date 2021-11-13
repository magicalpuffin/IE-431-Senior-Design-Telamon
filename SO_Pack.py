import pandas as pd
from SO_Pack_Functions import *

# Imports excel data and converts to dataframes
shipping_df = pd.read_excel("Data Inputs\\1 Year Shipping Orders.xlsx")
item_df = pd.read_excel("Data Inputs\\Item Warehouse Data 10-5-21.xlsx")
item_df_2 = pd.read_excel("Data Inputs\\Item Warehouse Data 11-8-21.xlsx", skiprows= 1)
test_bins = pd.read_excel("Data Inputs\\Bin_Pack_Test_Data.xlsx", "Bins")

# Organize shipping dataframe
# Rename item column and invert quantity
shipping_df = shipping_df.rename(columns={"Item": "Item Number"})
shipping_df.loc[:, ["Quantity"]] = -shipping_df["Quantity"]

# Newly added item_df 2, noter certain if it works
item_df_2 = item_df_2.rename(columns={"Unnamed: 0": "Item Number", "Item Number": "Quantity", "Weight (lbs)": "Unit Weight", "L (in)": "Unit Length", "W (in)": "Unit Width", "D (in)" : "Unit Height"})
item_df_2_clean, item_df_2_exclude = remove_na_in_cols(item_df_2, ["Unit Weight", "Unit Length", "Unit Width", "Unit Height"])
item_df_2_clean["Unit Volume"] = item_df_2_clean["Unit Length"] * item_df_2_clean["Unit Width"] * item_df_2_clean["Unit Height"]
item_df_2_clean = item_df_2_clean[["Item Number", "Unit Weight", "Unit Length", "Unit Width", "Unit Height", "Unit Volume"]]

# Data cleaning functions, filter out items missing data, filter out SO missing items
item_df, item_df_excluded = clean_item_df(item_df)

# Newley added, uncertain code
item_merged_df = pd.concat([item_df, item_df_2_clean])
item_df = item_merged_df

# Merges the shipping dataframe with item dataframe on items
merged_df = pd.merge(shipping_df, item_df, on ="Item Number", how="left")
merged_df, shipping_df_item_na = clean_merged_df(merged_df)

SO_indexed_df = merged_df.set_index(["Sales Order Number", "Item Number"])

# ----
# Output for debug and lists of excluded items and SO
# output_df_1 = [item_df, item_df_excluded, shipping_df_item_na, merged_df]
# name_df_1 = ["Cleaned item data", "Excluded item data", "Missing info SO", "Merged SO Clean"]
# df_to_excel(output_df_1, name_df_1, "clean_data.xlsx")
# ----

SO_summary_df = get_SO_summary(SO_indexed_df)
# ----
# Remove high quantity and high volume SO for reduced run times
# Limits SO to only ones with low quantities to improve run time
SO_high_quantities = SO_summary_df[SO_summary_df["Quantity"] > 25]
SO_high_vol = SO_summary_df[SO_summary_df["Total Volume"] > 12000]

SO_indexed_df_clean = SO_indexed_df.drop(SO_high_quantities.index.values, level = 0)
SO_indexed_df_clean = SO_indexed_df_clean.drop(SO_high_vol.index.values, level = 0)

# Packing algorithm on remaining SO
# Computationally this can start to take a long time, this ran within 2 mins
# SO with extremely high quantities will increase run time significantly
SO_indexed_df_clean_packed_debug = SO_indexed_df_clean.groupby(level = 0).apply(pack_SO, bin_df = test_bins)
SO_indexed_df_clean_packed = SO_indexed_df_clean_packed_debug[['Pack Status', 'Bin Name', 'Volume Difference', 'Weight Difference', 'Volume Utilization', 'Weight Utilization']]

output_df_3 = [SO_indexed_df_clean, SO_indexed_df_clean_packed, SO_indexed_df_clean_packed_debug]
name_df_3 = ["Low quant SO", "SO Packed", "SO Packed Debug"]
df_to_excel(output_df_3, name_df_3, "Data Outputs\\packed_results_low_quantity.xlsx")