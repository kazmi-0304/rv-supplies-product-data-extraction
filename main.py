import requests
import json
from bs4 import BeautifulSoup
from openpyxl import Workbook

# Initialize a new workbook and select the active worksheet
workbook = Workbook()
sheet = workbook.active
sheet.title = 'Products'

# Define the header row
headers = ['Item Code', 'Product Name', 'Product ID', 'Product Price', 'Product Description', 'Specifications', 'Image 1', 'Image 2', 'Image 3', 'Image 4', 'Image 5', 'Image 6', 'Image 7', 'Video', 'User Manual', 'Parts List', 'Brochure', 'Dimensions', 'Additional Ducting']
sheet.append(headers)

# Define the range of product IDs you want to extract
start_product_id = 3279
end_product_id = 3279

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
            description_content = ''.join([line["Description"] for line in data['Descriptions']])
            if "specifications" in description_content.lower():
                html_content = description_content.split("</strong></p>")[-1].strip()
                soup = BeautifulSoup(html_content, 'html.parser')
            
                specifications = '\n'.join([li.text for li in soup.find_all('li')])
                # print(specifications)
            else:
                specifications = 'N' 
            product = [
                data['Product']['ProductEntryId'],
                data['Product']['ProductName'],
                data['Product']['ProductId'],
                round(float(data['Product']['OriginalPrice']) * 1.15, 2),
                ''.join([line["Description"] for line in data['Descriptions']]),
                specifications,
                "https://www.rvsupplies.co.nz" + data['ProductImages'][0]['ImageFileName'] if len(data['ProductImages']) > 0 else "N",
                "https://www.rvsupplies.co.nz" + data['ProductImages'][1]['ImageFileName'] if len(data['ProductImages']) > 1 else "N",
                "https://www.rvsupplies.co.nz" + data['ProductImages'][2]['ImageFileName'] if len(data['ProductImages']) > 2 else "N",
                "https://www.rvsupplies.co.nz" + data['ProductImages'][3]['ImageFileName'] if len(data['ProductImages']) > 3 else "N",
                "https://www.rvsupplies.co.nz" + data['ProductImages'][4]['ImageFileName'] if len(data['ProductImages']) > 4 else "N",
                "https://www.rvsupplies.co.nz" + data['ProductImages'][5]['ImageFileName'] if len(data['ProductImages']) > 5 else "N",
                "https://www.rvsupplies.co.nz" + data['ProductImages'][6]['ImageFileName'] if len(data['ProductImages']) > 6 else "N",
                check_attachment_exists('Video'),
                check_attachment_exists('User Manual'),
                check_attachment_exists('Parts List'),
                check_attachment_exists('Brochure'),
                check_attachment_exists('Dimensions'),
                check_attachment_exists('Additional Ducting')
            ]
            print(product_id)
            print(product)
            print("\n")
            # Append the product data to the sheet
            sheet.append(product)
        
    else:
        print(f"Failed to retrieve product {product_id} data")

# Save the workbook to a file
workbook.save("products.xlsx")

print(f"All product data saved to products.xlsx")
