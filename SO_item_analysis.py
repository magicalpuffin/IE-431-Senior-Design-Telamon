import pandas as pd
import matplotlib.pyplot as plt
from SO_Item_Analysis_Functions import *

shipping_df = pd.read_excel("1 Year Shipping Orders.xlsx")
item_df_1 = pd.read_excel("Item Warehouse Data 9-29-21.xlsx")
item_df_2 = pd.read_excel("Item Warehouse Data 10-5-21.xlsx")

SO_size_df = shipping_df["Sales Order Number"].value_counts()
item_freq_df = shipping_df["Item Number"].value_counts()

SO_size_dis = SO_size_df.value_counts()
item_freq_dis = item_freq_df.value_counts()
# Tries to make histograms to display the data, majority of shipping orders and items have very infrequent listings
plt.clf()
SO_size_hist = SO_size_df.hist(bins = 50)
plt.xlabel("Shipping Order Size")
plt.ylabel("Frequency")
plt.title("Frequency of Shipping Order Sizes")
item_freq_hist = item_freq_df.hist(bins = 50)
plt.xlabel("Item Frequency")
plt.ylabel("Frequency")
plt.title("Frequency of Item Frequencies")
plt.show()

output_df_2 = [SO_size_df, item_freq_df, SO_size_dis, item_freq_dis]
output_names_2 = ["SO sizes", "Item frequencies", "SO size distrubtion", "Item frequency distribution"]
df_to_excel(output_df_2, output_names_2, file_name= "py_SO_item_analysis.xlsx")