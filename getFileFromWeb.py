import os
import requests
from bs4 import BeautifulSoup
import util
credentials = util.getCredentials()

def download_csv_files(folder_path,prefix,save_path):
    base_url = credentials['base_url']
    # Create the save directory if it doesn't exist
    if not os.path.exists(save_path):
        os.makedirs(save_path)

    response = requests.get(base_url + folder_path)
    soup = BeautifulSoup(response.text, 'html.parser')

    csv_links = []
    for link in soup.find_all('a'):
        href = link.get('href')
        if href.startswith(prefix) and href.endswith('.csv'):
            csv_links.append(href)

    for csv_link in csv_links:
        csv_filename = csv_link.split('/')[-1]

        local_csv_files = [file for file in os.listdir(save_path) if file.endswith('.csv')]
        if csv_filename in local_csv_files:
            print(f"Skipped {csv_filename}, already exists locally")
            continue

        csv_url = base_url + folder_path + csv_link
        csv_filename = os.path.join(save_path, csv_filename)

        response = requests.get(csv_url)
        if response.status_code == 200:
            with open(csv_filename, 'wb') as csv_file:
                csv_file.write(response.content)
            print(f"Downloaded: {csv_filename}")
        else:
            print(f"Failed to download: {csv_filename}")

    # Remove older archives if more than one archive remains
    downloaded_csv_files = [file for file in os.listdir(save_path) if file.startswith(prefix) & file.endswith('.csv')]
    if len(downloaded_csv_files) > 1:
        oldest_file = min(downloaded_csv_files, key=lambda x: os.path.getctime(os.path.join(save_path, x)), default=None)
        if oldest_file:
            os.remove(os.path.join(save_path, oldest_file))
            print(f"Removed older file: {oldest_file}")

if __name__ == "__main__": 
  folder_path = "nbulkload/"
  prefix = "dump_tpt_cellsector_"
  save_path = "import/nbulkload"  # Change this to your desired save path
  download_csv_files(folder_path,prefix,save_path)
