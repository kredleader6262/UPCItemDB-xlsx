## This is an example of how to call the upcitemdb_lookup function from the upcitemdb_lookup.py file

# outputfilename: The name of the file to save the data to as .xlsx. Default is upc_items.xlsx
# masterfilename: The name of the file to save the master data to as .xlsx. Default is master_upc_items.xlsx
# input_filename: The name of the file to read the UPCs from. Please use a .txt file. Default is upc_lookup.txt
# default_upcs: A list of default UPCs to use if the input file is not found. Default is ["887276550992"] as an example
# skip_duplicates: Sets flag for main to skip or overwrite existing rows

import upcitemdb_lookup as upc

upc.main(outputfilename="upc_items_test.xlsx", masterfilename="master_upc_items_test.xlsx", input_filename="upc_lookup.txt", default_upcs=["810116380817"], skip_duplicates=True)

# json = upc.lookup_request("810116380817")
# print(json)