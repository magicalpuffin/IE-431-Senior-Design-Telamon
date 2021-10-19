import pandas as pd
from SO_Item_Analysis_Functions import *

# Imports excel data and converts to dataframes
shipping_df = pd.read_excel("1 Year Shipping Orders.xlsx")
item_df_1 = pd.read_excel("Item Warehouse Data 9-29-21.xlsx")
item_df_2 = pd.read_excel("Item Warehouse Data 10-5-21.xlsx")
# shipping_df

shipping_df = shipping_df.rename(columns={"Item": "Item Number"})

# Some dimensions will need unit conversion
# item_df_2["Dimension UOM"].value_counts()
# item_df_2.columns.values

# item_df_2_clean = item_df_clean(item_df_2)

item_df = item_df_2
# ----
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

item_df_clean.loc[item_df_clean["Dimension UOM"] == "Ft", ["Unit Length", "Unit Width", "Unit Height"]] = item_df_clean[item_df_clean["Dimension UOM"] == "Ft"][["Unit Length", "Unit Width", "Unit Height"]] * 12

# Keeps only useful columns
item_df_clean = item_df_clean[["Item Number", "Description", "Item Type", "Product Category", "Unit Length", "Unit Width", "Unit Height", "Unit Weight"]]

item_df_missing_dim = item_df_missing_dim.assign(exclude_reason = "Missing Dimension")
item_df_missing_weight = item_df_missing_weight.assign(exclude_reason = "Missing Weight")
item_df_wrong_dim_UOM = item_df_wrong_dim_UOM.assign(exclude_reason = "Wrong Dimension Units")
item_df_wrong_weight_UOM = item_df_wrong_weight_UOM.assign(exclude_reason = "Wrong Weight Units")
item_df_excluded = pd.concat([item_df_missing_dim, item_df_missing_weight, item_df_wrong_dim_UOM, item_df_wrong_weight_UOM])

# ----
# Items with dimensions but missing weight, may need to also check items with weight but missing dimensions
# item_df_2_no_na[item_df_2_no_na["Unit Weight"].isna()]

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