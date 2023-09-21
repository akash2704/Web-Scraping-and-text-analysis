# Web Scraping and Text Analysis

This Python project consists of two main scripts, `main.py` and `Text_Analysis.py`, which are used for web scraping and text analysis.

## main.py

`main.py` is responsible for scraping data from web articles. It utilizes libraries like `requests` and `BeautifulSoup` to extract article titles and texts from a list of URLs provided in an Excel file (`input.xlsx`). The extracted data is then saved in individual text files within a directory called `extracted_articles`.

### Usage

1. Create an Excel file named `input.xlsx` containing a list of URLs to web articles.
2. Run the script `main.py`. It will scrape data from the provided URLs and save the extracted articles in the `extracted_articles` directory.

## Text_Analysis.py

`Text_Analysis.py` is responsible for analyzing the sentiment and readability of the extracted articles. It performs sentiment analysis using a custom list of positive and negative words and calculates readability metrics such as the Fog Index and word count.

### Usage

1. Ensure that the `extracted_articles` directory contains the extracted articles from `main.py`.
2. Create a custom stopword list (or use the provided one) to enhance text cleaning.
3. Run the script `Text_Analysis.py`. It will analyze each article and generate an output Excel file (`output/output.xlsx`) containing the analysis results.

## Requirements

- Python 3.x
- Required Python packages listed in `requirements.txt`. You can install them using `pip install -r requirements.txt`.

## Installation

1. Clone this repository to your local machine.
2. Install the required packages using `pip install -r requirements.txt`.

## Author

- Akash Ajay Kallai