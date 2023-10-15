from multiprocessing import Value
import os
import requests
import openpyxl
import time

def download_images_for_queries(excel_file, folder, num_images_per_query, api_key):
    # Open the Excel file
    wb = openpyxl.load_workbook(excel_file)
    sheet = wb.active

    for row in sheet.iter_rows(values_only=True):
        # Assume the queries are in the first column of the Excel sheet
        query = row[0]

        # Create a subfolder for each query
        query_folder = os.path.join(folder, query)
        if not os.path.exists(query_folder):
            os.makedirs(query_folder)

        # Download images for the current query
        download_images(query, query_folder, num_images_per_query, api_key)

    wb.close()

def download_images(query, folder, num_images, api_key):
    # Initialize a set to keep track of downloaded image URLs
    downloaded_urls = set()

    # Bing Image Search API endpoint
    api_endpoint = "https://api.bing.microsoft.com/v7.0/images/search"

    # Set the parameters for the API request
    params = {
        "q": query,
        "count": num_images,
    }

    # Set headers with the API key
    headers = {
        "Ocp-Apim-Subscription-Key": api_key,
    }

    try:
        response = requests.get(api_endpoint, params=params, headers=headers)

        if response.status_code == 200:
            data = response.json()
            value = data.get("value")

            if not value:
                print("No images found for query:", query)
                return

            # Rest of the code
            for i, item in enumerate(value):
                image_url = item.get("contentUrl")
                if image_url:
                    # Check if the URL has already been downloaded
                    if image_url not in downloaded_urls:
                        downloaded_urls.add(image_url)
                        # Rest of the download code for this image remains the same
                        with open(os.path.join(folder, f"{query}_{i + 1}.jpg"), 'wb') as img_file:
                            img_response = requests.get(image_url)
                            img_file.write(img_response.content)
                            print(f"Downloaded {i + 1}/{num_images} images for query: {query}")
                    else:
                        print(f"Image URL is a duplicate and will not be downloaded for result {i + 1}")

                
                    time.sleep(2)

        else:
            print(f"Error searching for images for query: {query}. Status Code {response.status_code}")
    except Exception as e:
        print(f"Error for query {query}: {str(e)}")

if __name__ == "__main__":
    excel_file = input("Enter the path to the Excel file containing queries: ")
    num_images_per_query = 100  # You can change this value
    folder = input("Enter the folder path where you want to save the images: ")
    api_key = input("Enter your Bing Image Search API key: ")

    download_images_for_queries(excel_file, folder, num_images_per_query, api_key)
