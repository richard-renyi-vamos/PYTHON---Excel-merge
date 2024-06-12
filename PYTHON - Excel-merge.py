import requests
import pandas as pd
from io import BytesIO

# Replace these with your SharePoint and file details
sharepoint_site = 'https://yoursharepointsite.sharepoint.com'
file1_url = 'https://yoursharepointsite.sharepoint.com/sites/yoursite/Shared%20Documents/yourfile1.xlsx'
file2_url = 'https://yoursharepointsite.sharepoint.com/sites/yoursite/Shared%20Documents/yourfile2.xlsx'
username = 'yourusername@yourdomain.com'
password = 'yourpassword'

# Function to download an Excel file from SharePoint
def download_excel_from_sharepoint(file_url, username, password):
    session = requests.Session()
    session.auth = (username, password)
    response = session.get(file_url)
    response.raise_for_status()
    return BytesIO(response.content)

# Download the two Excel files
file1 = download_excel_from_sharepoint(file1_url, username, password)
file2 = download_excel_from_sharepoint(file2_url, username, password)

# Read the Excel files into DataFrames
df1 = pd.read_excel(file1)
df2 = pd.read_excel(file2)

# Merge the DataFrames (this example assumes you want to concatenate them)
merged_df = pd.concat([df1, df2], ignore_index=True)

# Save the merged DataFrame to a new Excel file
merged_df.to_excel('merged_file.xlsx', index=False)

print("The Excel files have been merged and saved as 'merged_file.xlsx'.")
