# RV Supplies Product Data Extraction

## Overview
This script is designed to automate the process of extracting detailed product information from the RV Supplies website. It fetches data for a range of product IDs, parsing and organizing product details, specifications, images, and documents into a structured Excel workbook. This document serves as a comprehensive guide to setting up, understanding, and running the script to obtain product data efficiently.

## Video Preview

[![Video Preview](https://github.com/DevRex-0201/Project-Images/blob/main/video%20preview/Py-RV-Supplies-Product-Data-Extraction.png)](https://brand-car.s3.eu-north-1.amazonaws.com/Four+Seasons/Py-RV-Supplies-Product-Data-Extraction.mp4)

## Requirements

### Software and Libraries
- **Python**: Ensure you have Python installed on your system. The script is compatible with Python 3.x versions.
- **Libraries**: The script requires the following Python libraries:
  - `requests`: For making HTTP requests to the RV Supplies website.
  - `json`: For parsing JSON data returned by the website's API.
  - `BeautifulSoup` from `bs4`: For HTML parsing and extraction of product specifications.
  - `openpyxl`: For creating and managing the Excel workbook where product data is stored.

### Installation
1. **Python Installation**: If not already installed, download and install Python from the official website (https://www.python.org/).
2. **Library Installation**: Install the required libraries using pip. Open your terminal or command prompt and execute the following command:
   ```bash
   pip install requests beautifulsoup4 openpyxl
   ```

## Script Configuration
Before running the script, you must configure the range of product IDs you wish to extract data for. This is done by setting the `start_product_id` and `end_product_id` variables at the beginning of the script.

## Running the Script
To execute the script, navigate to the directory containing the script file in your terminal or command prompt, then run the following command:
```bash
python script_name.py
```
Replace `script_name.py` with the actual filename of the script.

## Script Workflow
1. **Workbook Initialization**: The script starts by creating a new Excel workbook and setting up the header row in the 'Products' sheet.
2. **Data Extraction**: For each product ID in the specified range, the script sends a GET request to the RV Supplies API, parses the returned JSON data, and extracts product details, including name, ID, price, description, specifications, images, videos, user manuals, parts lists, brochures, dimensions, and additional ducting information.
3. **Data Parsing**: The script uses BeautifulSoup to parse and extract product specifications from HTML content. It also constructs URLs for product images and documents if available.
4. **Error Handling**: The script checks for HTTP request success and handles cases where product information may not be available or the API returns an error.
5. **Saving Data**: Extracted product data is appended to the Excel workbook, which is then saved to a file named `products.xlsx`.

## Output
The script outputs an Excel workbook named `products.xlsx`, containing a sheet titled 'Products' with all the extracted product information organized into rows and columns as per the headers defined in the script.

## Troubleshooting
- **HTTP Request Failures**: Ensure you have a stable internet connection. If a product ID does not return data, verify that it exists on the RV Supplies website.
- **Library Errors**: Make sure all required Python libraries are installed correctly. If you encounter any library-specific errors, refer to the official documentation for troubleshooting guidance.

## Conclusion
This script provides an efficient way to gather comprehensive product data from the RV Supplies website, saving time and reducing manual effort in data collection and organization. By following the instructions provided in this README, users can customize, execute, and leverage the script for their specific data extraction needs.
