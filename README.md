# UPC Item Lookup

This Python script is used to fetch and store product details from the UPCItemDB API. It reads UPC codes from a text file or a default list, fetches product details for each UPC code, and stores the details in an Excel workbook.

## Features

- Fetches product details for each UPC code from the UPCItemDB API.
- Stores product details in an Excel workbook, with a separate sheet for each UPC code and a master sheet containing details of all products.
- Updates existing product details if the UPC code is already present in the workbook.
- Handles API rate limits by retrying after a delay.
- Saves a copy of the master sheet into a new file after all operations.

## Script Usage

1. Set the `filename`, `masterfilename`, and `txt_filename` variables to the desired paths if you want to use custom file names and paths.
2. Set the `default_upcs` variable to a list of default UPC codes to be used if the text file does not exist or is empty. Modify the file called `upc_lookup.txt` and use it to paste UPCs by row to check.
3. Run the script. It will read UPC codes from the text file (or use the default list if the file does not exist or is empty), fetch product details for each UPC code, and store the details in an Excel workbook called `upc_items.xlsx`.
4. After all operations, it will save a copy of the master sheet into a new file called `upc_lookup_master.xlsx` in the same directory.

## Calling lookup_request and main functions

You can call the main function from another script by importing the `lookup_request` or `main` functions and passing the desired parameters. 

The `lookup_request` function only takes one parameter, which is the UPC code to be looked up. It returns a json response containing the product details from the UPCItemDB API.

The `main` function takes four parameters: `outputfilename`, `masterfilename`, and `input_filename`, and `default_upcs`. ITs designed to accept different input and output file names and paths.

Examples of calling either can be seen in the `call_example.py` file.

## Dependencies

- `requests`: Used to send HTTP requests to the UPCItemDB API.
- `openpyxl`: Used to read from and write to Excel workbooks.
- `os`: Used to check if files exist and to join paths.
- `time`: Used to delay retries when the API rate limit is exceeded.

## Note

This script assumes that each UPC code corresponds to one product. If a UPC code corresponds to multiple products, it will only fetch details for the first product.

### Planned Features

- Add a GUI for inputting the UPC codes.
- Add headers for authentication and other parameters to the API request.
- Add search request vs lookup request to the API.