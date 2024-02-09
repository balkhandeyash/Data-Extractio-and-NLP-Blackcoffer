# data_extraction.py

import pandas as pd
from bs4 import BeautifulSoup
import requests

# Load URLs from the input Excel file
input_file = "Input.xlsx"
df = pd.read_excel(input_file)

for index, row in df.iterrows():
    url_id = row['URL_ID']
    url = row['URL']

    # Fetch the HTML content of the URL
    response = requests.get(url)
    soup = BeautifulSoup(response.text, 'html.parser')

    # Extract article text
    article_title = soup.find('title').get_text()
    article_text = " ".join([p.get_text() for p in soup.find_all('p')])

    # Save the article text to a text file
    with open(f"{url_id}.txt", "w", encoding="utf-8") as file:
        file.write(f"{article_title}\n\n{article_text}")
