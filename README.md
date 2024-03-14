# UPC Item Lookup

This Python script fetches and stores product details from the UPCItemDB API. It reads UPC codes from a text file or a default list, fetches product details for each UPC code, and stores the details in an Excel workbook.

## Features

- Fetches product details for each UPC code from the UPCItemDB API. Here is the link to the API documentation: [UPCItemDB Development](https://www.upcitemdb.com/wp/docs/main/development/).
- Uses the `user_key` variable from the `config.ini` file to authenticate with the API if you have one. Otherwise, it uses the trial endpoint with a rate limit of 100 requests per day for free. [UPCItemDB API Rate Limits](https://www.upcitemdb.com/wp/docs/main/development/api-rate-limits/)
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

The `lookup_request` function takes one parameter, which is the UPC code to be looked up. It returns a json response containing the product details from the UPCItemDB API.

The `main` function takes four parameters: `outputfilename`, `masterfilename`, `input_filename`, and `default_upcs`. It's designed to accept different input and output file names and paths.

Examples of calling either can be seen in the `call_example.py` file. Run it to test your own UPC codes but comment the ones you want to use.

## Dependencies

- `requests`: Used to send HTTP requests to the UPCItemDB API.
- `openpyxl`: Used to read from and write to Excel workbooks.
- `os`: Used to check if files exist and to join paths.
- `time`: Used to delay retries when the API rate limit is exceeded.
- `sys`: Used for system-specific parameters and functions.
- `configparser`: Used to read the `config.ini` file.
- `tqdm`: Used to display progress bars.

## Configuration

The script uses a configuration file `config.ini` to store the user key and other settings. Here is an example of the contents of the `config.ini` file:

```ini
[UPCITEMDB]
user_key = 
skip_duplicates = true
```

## Development Environment

This script was developed using Python 3.11.4 on Windows. It should work on other operating systems as well as maybe older Python, but it has not been tested on them.

## Note

This script assumes that each UPC code corresponds to one product. If a UPC code corresponds to multiple products, it will only fetch details for the first product.

## Planned Features

- Add a gui for inputing UPC codes to look up.
- Add search request vs lookup request to the API.
- Add more API enpoints for search?
- Make this a library, perhaps

Please note that the script's behavior may change as new features are added. Always refer to the latest version of this README for the most accurate information.
