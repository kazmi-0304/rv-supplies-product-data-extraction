import requests
import json
from bs4 import BeautifulSoup
from openpyxl import Workbook

# Initialize a new workbook and select the active worksheet
workbook = Workbook()
sheet = workbook.active
sheet.title = 'Products'

# Define the header row
headers = ['Item Code', 'Category', 'Product ID']
sheet.append(headers)

# Define the range of product IDs you want to extract
start_product_id = 2066
end_product_id = 3000

def check_attachment_exists(key):
    if len(data['ProductDocuments']) > 0:
        print(data['ProductDocuments'])
        result = ""
        for document in data['ProductDocuments']:
            if document.get("TitleTag") == key:
                result = "https://www.rvsupplies.co.nz" + document["DocumentFileName"]
        if result != "":
            result == "N"
        return result
    else:
        return "N"

# Loop through the product IDs and make GET requests
for product_id in range(start_product_id, end_product_id + 1):
    url = f"https://www.rvsupplies.co.nz/DesktopModules/AcumenOnline/Api/ProductDetails/Get/{product_id}"

    # Send a GET request to the URL
    response = requests.get(url)

    # Check if the request was successful (status code 200)
    if response.status_code == 200:
        # Parse the JSON data from the response
        data = response.json()
        product = []        
        if data['Product']['ProductId'] != None:
            if len(data['Categories']) > 0:
                category = data['Categories'][-1]['CategoryDescription']
            else:
                category = "N"
            product = [
                data['Product']['ProductEntryId'],
                category,
                data['Product']['ProductId']
                ]
            print(product_id)
            print(product)
            print("\n")
            # Append the product data to the sheet
            sheet.append(product)
        
    else:
        print(f"Failed to retrieve product {product_id} data")

# Save the workbook to a file
workbook.save("products_category.xlsx")

print(f"All product data saved to products.xlsx")
