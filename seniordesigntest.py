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
product_df_filt = product_df[["Vendor Part No", "Item Weight (Lb. -Base UOM)", "Item Dimensions (LxWxH inches)"]]
product_df_filt = product_df_filt.rename(columns={"Vendor Part No": "Item"})
product_df_filt

# Fills na or missing weight data in product list, used to just preserve data during merges
# Line can be disabled to toggle between switching
product_df_filt["Item Weight (Lb. -Base UOM)"] = product_df_filt["Item Weight (Lb. -Base UOM)"].fillna(0)

# Merges data, left merge to preserve shipping orders
merged_df = pd.merge(shipping_df, product_df_filt, on ="Item", how="left")
merged_df_clean = merged_df[merged_df["Item Weight (Lb. -Base UOM)"].isna() == False]
merged_df_clean

# Count of different dimensions
# Count of different sales orders
merged_df_clean["Item Dimensions (LxWxH inches)"].value_counts()
merged_df_clean["Sales Order Number"].value_counts()

output_df = [merged_df_clean, merged_df]
output_names = ["Cleaned Merged Data", "Merged Data"]

def df_to_excel(df_list, name_list, file_name = "excel_output.xlsx"):
    datatoexcel = pd.ExcelWriter(file_name)
    for df, name in zip(df_list, name_list):
        df.to_excel(excel_writer = datatoexcel, sheet_name = name)
    datatoexcel.save()
    return

df_to_excel(output_df, output_names)
