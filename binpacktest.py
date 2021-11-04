import pandas as pd
from py3dbp import Packer, Bin, Item
from SO_Item_Analysis_Functions import *

# Misc notes
# This is a buggy debug version and test environment
# Inut shipping order, bins
# Output packer object, bin that fit, bin name, bin volume, difference in volume, difference in weight
# -----
packer = Packer()

# ----
test_bins = pd.read_excel("Data Inputs\\Bin_Pack_Test_Data.xlsx", "Bins")
test_items = pd.read_excel("Data Inputs\\Bin_Pack_Test_Data.xlsx", "Items")

indexed_df = test_items.set_index(["Sales Order Number", "Item Number"])
indexed_bins = test_bins.set_index("Bin Name")
# Debug lines
# -------
# The process of making a fucntion seems promissing, group and apply can be used to...
# apply a function on each grouping and returns a result, should be possible to fit...
# the entire packing process into a function

# packed_SO_debug = indexed_df.groupby(level = 0).apply(pack_SO, bin_df = test_bins)
# packed_SO = packed_SO_debug[['Pack Status', 'Bin Name', 'Volume Difference', 'Weight Difference']]

# Debug to check items in packer

# Test function on 1 shipping order
# --------
# Adds all items within the SO into packer object
SO_1 = indexed_df.loc[1001]
for index, row in SO_1.iterrows():
    for i in range(int(row["Quantity"])):
        packer.add_item(Item(index, row["Unit Length"], row["Unit Width"], row["Unit Height"], row["Unit Weight"]))

# Adds all bins from seperate sheet into packer object
# indexed_bins
for index, row in test_bins.iterrows():
    index
    packer.add_bin(Bin(row["Bin Name"], row["Length"], row["Width"], row["Height"], row["Weight"]))
packer.pack()
packer.bins[0].string()

# ----
# packer.add_bin(Bin('small-envelope', 11.5, 6.125, 0.25, 10))
# packer.add_bin(Bin('large-envelope', 15.0, 12.0, 0.75, 15))
# packer.add_bin(Bin('small-box', 8.625, 5.375, 1.625, 70.0))
# packer.add_bin(Bin('medium-box', 11.0, 8.5, 5.5, 70.0))
# packer.add_bin(Bin('medium-2-box', 13.625, 11.875, 3.375, 70.0))
# packer.add_bin(Bin('large-box', 12.0, 12.0, 5.5, 70.0))
# packer.add_bin(Bin('large-2-box', 23.6875, 11.75, 3.0, 70.0))

# packer.add_item(Item('50g [powder 1]', 3.9370, 1.9685, 1.9685, 1))
# packer.add_item(Item('50g [powder 2]', 3.9370, 1.9685, 1.9685, 2))
# packer.add_item(Item('50g [powder 3]', 3.9370, 1.9685, 1.9685, 3))
# packer.add_item(Item('250g [powder 4]', 7.8740, 3.9370, 1.9685, 4))
# packer.add_item(Item('250g [powder 5]', 7.8740, 3.9370, 1.9685, 5))
# packer.add_item(Item('250g [powder 6]', 7.8740, 3.9370, 1.9685, 6))
# packer.add_item(Item('250g [powder 7]', 7.8740, 3.9370, 1.9685, 7))
# packer.add_item(Item('250g [powder 8]', 7.8740, 3.9370, 1.9685, 8))
# packer.add_item(Item('250g [powder 9]', 7.8740, 3.9370, 1.9685, 9))

# packer.pack()

# for b in packer.bins:
#     print(":::::::::::", b.string())

#     print("FITTED ITEMS:")
#     for item in b.items:
#         print("====> ", item.string())

#     print("UNFITTED ITEMS:")
#     for item in b.unfitted_items:
#         print("====> ", item.string())

#     print("***************************************************")
#     print("***************************************************")

# # Originally a debug area, copy this into a new function
# # -------
# # Find the smallest volume bin that fits all items
# # Loops through bins to find one that has all items, defaults to largest bin
# bestbin = packer.bins[-1]
# pack_status = False
# for b in packer.bins:
#     if len(b.unfitted_items) == 0:
#         bestbin = b
#         pack_status = True
#         break
# # Parameters of bin to debug
# packer
# bin_name = bestbin.name
# bin_vol = bestbin.get_volume()
# bin_weight = bestbin.max_weight

# # Finds total item weight and volume, can be used to compare with available vol and weight in bin
# # Uses Decimal module. I don't know how to use it so I convert it to float
# total_item_vol = 0
# total_item_weight = 0
# for i in bestbin.items:
#     total_item_vol = total_item_vol + i.get_volume()
#     total_item_weight = total_item_weight + i.weight
# vol_diff = float(bin_vol - total_item_vol)
# weight_diff = float(bin_weight - total_item_weight)