import pandas as pd

# Imports excel data and converts to dataframes
shipping_df = pd.read_excel("1 Year SO Issues Purdue.xlsx")
product_df = pd.read_excel("W&D Purdue 9-29-21.xlsx")
shipping_df
product_df

# Testing various pandas operations
shipping_df.iloc[:,0]
shipping_df["Sales Order Number"]
value_counts = shipping_df["Sales Order Number"].value_counts()

# Cleans product data frame to only important columns
filtered_product_df = product_df[["Vendor Part No", "Item Weight (Lb. -Base UOM)", "Item Dimensions (LxWxH inches)"]]
filtered_product_df = filtered_product_df.rename(columns={"Vendor Part No": "Item"})
# filtered_product_df["Item Weight (Lb. -Base UOM)"] = filtered_product_df["Item Weight (Lb. -Base UOM)"].fillna(0)
filtered_product_df

# Merges data, left merge to preserve shipping orders
# Value
merged_df = pd.merge(shipping_df, filtered_product_df, on ="Item", how="left")
merged_df
cleaned_merged_df = merged_df[merged_df["Item Weight (Lb. -Base UOM)"].isna() == False]

cleaned_merged_df["Item Dimensions (LxWxH inches)"].value_counts()
cleaned_merged_df["Sales Order Number"].value_counts()