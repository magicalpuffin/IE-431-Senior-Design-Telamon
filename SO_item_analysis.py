import pandas as pd
import matplotlib.pyplot as plt

shipping_df = pd.read_excel("Data Inputs\\1 Year Shipping Orders.xlsx")

shipping_df["Quantity"] = shipping_df["Quantity"].abs()
SO_size_df = shipping_df["Sales Order Number"].value_counts()
item_freq_df = shipping_df["Item"].value_counts()

SO_quantity_df = shipping_df.groupby(["Sales Order Number"]).sum()

SO_size_dis = SO_size_df.value_counts()
item_freq_dis = item_freq_df.value_counts()
# Tries to make histograms to display the data, majority of shipping orders and items have very infrequent listings
plt.clf()
SO_quantity_df = SO_quantity_df[SO_quantity_df["Quantity"] <= 50]
SO_quantity_hist = SO_quantity_df.hist(bins = 50)
plt.xlabel("Shipping Order Quantity")
plt.ylabel("Frequency")
plt.title("Shipping Order Quantities")
item_freq_hist = item_freq_df.hist(bins = 50)
plt.xlabel("Item Frequency")
plt.ylabel("Frequency")
plt.title("Frequency of Item Frequencies")
plt.show()