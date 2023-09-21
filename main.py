import openpyxl
from bs4 import BeautifulSoup
import requests
import os
import pandas as pd

# Function to extract article title and text from a URL
def extract_article_data(url):
    try:
        response = requests.get(url)
        soup = BeautifulSoup(response.content, 'html.parser')
        
        # Find the header element with class 'td-post-title'
        header_element = soup.find('header', class_='td-post-title')
        
        if header_element:
            # Extract the article title from the 'h1' tag with class 'entry-title'
            title_element = header_element.find('h1', class_='entry-title')
            if title_element:
                article_title = title_element.get_text()
            else:
                article_title = "Title not found"
        else:
            article_title = "Title not found"
        
        # Find the div element with class name 'td-post-content'
        post_content_div = soup.find('div', class_='td-post-content')
        
        if post_content_div:
            # Extract the article text within the div
            article_text = ''
            for paragraph in post_content_div.find_all('p'):
                article_text += paragraph.get_text() + '\n'
        else:
            article_text = "Text not found"
        
        return article_title, article_text
    except Exception as e:
        print(f"Error extracting data from {url}: {str(e)}")
        return None, None

# Load the Excel file into a DataFrame
excel_file = 'input.xlsx'
df = pd.read_excel(excel_file)

# Create a directory to store text files
output_dir = 'extracted_articles'
os.makedirs(output_dir, exist_ok=True)

# Iterate through DataFrame rows
for index, row in df.iterrows():
    url_id, url = row['URL_ID'], row['URL']
    
    # Extract article title and text
    article_title, article_text = extract_article_data(url)
    
    if article_title and article_text:
        # Save the article text in a text file
        output_filename = os.path.join(output_dir, f'{url_id}.txt')
        with open(output_filename, 'w', encoding='utf-8') as file:
            file.write(f'Title: {article_title}\n\n')
            file.write(article_text)
        print(f"Saved article from {url} to {output_filename}")

print("Extraction and saving complete.")
