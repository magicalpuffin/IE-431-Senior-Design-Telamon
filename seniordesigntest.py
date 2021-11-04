import pandas as pd
import matplotlib.pyplot as plt
from SO_Item_Analysis_Functions import *

# Imports excel data and converts to dataframes
shipping_df = pd.read_excel("Data Inputs\\1 Year Shipping Orders.xlsx")
# item_df_1 = pd.read_excel("Item Warehouse Data 9-29-21.xlsx")
item_df_2 = pd.read_excel("Data Inputs\\Item Warehouse Data 10-5-21.xlsx")
test_bins = pd.read_excel("Data Inputs\\Bin_Pack_Test_Data.xlsx", "Bins")

# Organize shipping dataframe
# Rename item column and invert quantity
shipping_df = shipping_df.rename(columns={"Item": "Item Number"})
shipping_df.loc[:, ["Quantity"]] = -shipping_df["Quantity"]

# Data cleaning functions, filter out items missing data, filter out SO missing items
item_df_2_clean, item_df_2_excluded = item_df_clean(item_df_2)
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
df_to_excel(output_df_3, name_df_3, "packed_results_low_quantity.xlsx")

# ----

output_df_2 = [SO_volume, SO_weight]
name_df_2 = ["SO Volumes", "SO_weights"]
df_to_excel(output_df_2, name_df_2, "SO_analysis.xlsx")

# -----
# Generate graphs
plt.clf()
# Makes histogram, 12000 in^3 is max box size
SO_vol_hist = SO_volume.hist()
# SO_vol_hist = SO_volume.hist(bins = range(0, 12000, 1000))
plt.xlabel("Shipping Order Volume in^3")
plt.ylabel("Frequency")
plt.title("Frequency of Shipping Order Volumes")
plt.show()

plt.clf()
SO_weight_hist = SO_weight.hist()
# SO_weight_hist = SO_weight.hist(bins = range(0, 50, 5))
plt.xlabel("Shipping Order Weight lb")
plt.ylabel("Frequency")
plt.title("Frequency of Shipping Order Weights")
plt.show()

plt.clf()
# SO_quantities_hist = SO_quantities.hist()
SO_quantities_hist = SO_quantities.hist(bins = range(0, 50, 5))
plt.xlabel("Shipping Order Quantities")
plt.ylabel("Frequency")
plt.title("Frequency of Shipping Order Quantities")
plt.show()

# ----
# Code underneath may be ignored and out dated
# Items with dimensions but missing weight, may need to also check items with weight but missing dimensions
# Cleans product data frame to only important columns
item_df_1 = item_df_1[["Vendor Part No", "Item Weight (Lb. -Base UOM)", "Item Dimensions (LxWxH inches)"]]
item_df_1 = item_df_1.rename(columns={"Vendor Part No": "Item Number"})
item_df_1_no_na = item_df_1[item_df_1["Item Weight (Lb. -Base UOM)"].isna() == False]
# item_df_1_no_na

# Shipping order items that are completely missing from both item data sets
SO_miss_item_df = shipping_df[shipping_df["Item Number"].isin(item_df_2["Item Number"]) == False]
SO_miss_item_df = SO_miss_item_df[SO_miss_item_df["Item Number"].isin(item_df_1["Item Number"]) == False]
SO_miss_item_counts = SO_miss_item_df["Item Number"].value_counts()
# Shipping order items missing dimensions from both item sets
SO_miss_dims_df = shipping_df[shipping_df["Item Number"].isin(item_df_2_clean["Item Number"]) == False]
SO_miss_dims_df = SO_miss_dims_df[SO_miss_dims_df["Item Number"].isin(item_df_1_no_na["Item Number"]) == False]
SO_miss_dims_counts = SO_miss_dims_df["Item Number"].value_counts()

shipping_df_no_miss_dims = shipping_df[shipping_df["Item Number"].isin(SO_miss_dims_df["Item Number"]) == False]
# Will need to make the merge account for both item lists
# Merges data, left merge to preserve shipping orders
merged_df = pd.merge(shipping_df_no_miss_dims, item_df_2, on ="Item Number", how="left")

# Experimental
indexed_df = merged_df.set_index(["Sales Order Number", "Item Number"])
indexed_df.loc["1143515.Sales TL.ORDER ENTRY"]
indexed_df.groupby(level = 0).count()
# Count of different sales orders
# merged_df["Sales Order Number"].value_counts()

output_df = [merged_df, SO_miss_item_counts, SO_miss_dims_counts, SO_miss_item_df, SO_miss_dims_df]
output_names = ["Merged Full Data", "Missing Items Count in SO", "Missing Dims Count in SO", "SO Missing Items", "SO Missing Dims"]

df_to_excel(output_df, output_names)