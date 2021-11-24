import pandas as pd
from SO_Pack_Functions import *

# Imports excel data and converts to dataframes
shipping_df = pd.read_excel("Data Inputs\\1 Year Shipping Orders.xlsx")
item_df = pd.read_excel("Data Inputs\\Item Warehouse Data 10-5-21.xlsx")
item_df_2 = pd.read_excel("Data Inputs\\Item Warehouse Data 11-8-21.xlsx", skiprows= 1)
test_bins = pd.read_excel("Data Inputs\\Bin_Pack_Test_Data.xlsx", "Bins")

# Debug!!!!
# Creates and cleans Shipments
shipping_df ["Quantity"] = shipping_df["Quantity"].abs()
shipments_1 = Shipments(shipping_df, "Sales Order Number", "Item", "Quantity")

# Cleaning the item_df
# Filters out other UOM
# Converts ft to in
item_df_clean, item_df_val = filter_values_in_cols(item_df, {"Dimension UOM": ["Ft", "In"], "Weight UOM": ["Lbs"]})
item_df_clean.loc[item_df_clean["Dimension UOM"] == "Ft", ["Unit Length", "Unit Width", "Unit Height"]] = item_df_clean[item_df_clean["Dimension UOM"] == "Ft"][["Unit Length", "Unit Width", "Unit Height"]] * 12
item_df_clean.loc[item_df_clean["Dimension UOM"] == "Ft", ["Dimension UOM"]] = "In"

# Creates and cleans Items
items_1 = Items(item_df_clean, "Item Number", "Unit Length", "Unit Width", "Unit Height", "Unit Weight", "in", "lbs")
items_1.add_volume_col()
items_1_removed_df = items_1.remove_na()

item_df_2_clean = item_df_2.drop(columns=["Item Number"])
items_2 = Items(item_df_2_clean, "Unnamed: 0", "L (in)", "W (in)", "D (in)", "Weight (lbs)", "in", "lbs")
items_2.add_volume_col()
items_2_removed_df = items_2.remove_na()

# Combines cleaned items
# Merges shipments with items
item_merged_df = pd.concat([items_1.get_df(), items_2.get_df()])
merged_df = pd.merge(shipments_1.get_df(), item_merged_df, on ="Item Number", how="left")

# Creates and cleans merged data
# I dislike the name of the object but I have no better ideas
merge_1  = Shipments_Items_Merge(merged_df)
merge_1_removed_na_df = merge_1.remove_na()
merge_1.add_total_vol_weight_col()
merge_1_summary = merge_1.get_summary()
merge_1_removed_filter_df = merge_1.filter_SN(max_quantity = 25, max_volume = 12000)

SN_item_merge_df = merge_1.get_df()
# ----

# ----
# Output for debug and lists of excluded items and SO
# output_df_1 = [item_merged_df, item_df_val, items_1_removed_df, items_2_removed_df, merged_df, merge_1_summary]
# name_df_1 = ["Cleaned item data", "item_df 1 fiter 1", "item_df 1 filter 2", "item_2 filter 1", "Merged_df", "Shipment Summary"]
# df_to_excel(output_df_1, name_df_1, "Data Outputs\\clean_data.xlsx")
# ----

# Packing algorithm on remaining SO
# Computationally this can start to take a long time, this ran within 2 mins
# SO with extremely high quantities will increase run time significantly
SN_item_merge_df_packed = SN_item_merge_df.groupby(level = 0).apply(pack_SO, bin_df = test_bins)

output_df_3 = [SN_item_merge_df, SN_item_merge_df_packed]
name_df_3 = ["Low quant SN", "SN Packed", "SN Packed Debug"]
df_to_excel(output_df_3, name_df_3, "Data Outputs\\packed_results_low_quantity.xlsx")