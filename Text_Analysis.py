import openpyxl
import nltk
import re
import os
from nltk.corpus import stopwords
import pandas as pd 
nltk.download('punkt')

# Function to load custom stopword lists from text files
def load_custom_stopwords():
    custom_stopwords = set()
    stopword_dir = 'stopwords'  # Directory containing custom stopwords

    for filename in os.listdir(stopword_dir):
        if filename.endswith('.txt'):
            with open(os.path.join(stopword_dir, filename), 'r', encoding='utf-8', errors='ignore') as file:
                custom_stopwords.update(file.read().splitlines())

    return custom_stopwords

# Function to clean the text and calculate derived variables
def analyze_text(text):
    # Load custom stopwords
    custom_stopwords = load_custom_stopwords()

    # Clean the text using custom stopwords
    words = nltk.word_tokenize(text)
    cleaned_words = [word.lower() for word in words if word.isalnum() and word.lower() not in custom_stopwords]
    
    # Load the Positive and Negative word lists using raw string literals
    with open(r'MasterDictionary\positive-words.txt', 'r', encoding='ascii', errors='ignore') as f:
        positive_words = set(f.read().splitlines())
    with open(r'MasterDictionary\negative-words.txt', 'r', encoding='ascii', errors='ignore') as f:
        negative_words = set(f.read().splitlines())

    # Calculate Positive and Negative Scores
    positive_score = sum(1 for word in cleaned_words if word in positive_words)
    negative_score = sum(1 for word in cleaned_words if word in negative_words) * -1

    # Calculate Polarity Score and Subjectivity Score
    polarity_score = (positive_score - negative_score) / (positive_score + negative_score + 0.000001)
    subjectivity_score = (positive_score + negative_score) / (len(cleaned_words) + 0.000001)

    return {
        'Positive Score': positive_score,
        'Negative Score': negative_score,
        'Polarity Score': polarity_score,
        'Subjectivity Score': subjectivity_score
    }

# Function to analyze readability
def analyze_readability(text):
    # Tokenize sentences
    sentences = nltk.sent_tokenize(text)

    # Calculate Average Sentence Length
    total_words = len(nltk.word_tokenize(text))
    total_sentences = len(sentences)
    average_sentence_length = total_words / total_sentences

    # Count complex words (words with more than two syllables)
    complex_words = [word for word in nltk.word_tokenize(text) if len(re.findall(r'[aeiouAEIOU]{3,}', word)) >= 2]
    percentage_complex_words = (len(complex_words) / total_words) * 100

    # Calculate Fog Index
    fog_index = 0.4 * (average_sentence_length + percentage_complex_words)

    # Calculate Average Number of Words Per Sentence
    average_words_per_sentence = total_words / total_sentences

    # Calculate Syllable Count Per Word
    syllable_count_per_word = sum(len(re.findall(r'[aeiouAEIOU]+', word)) for word in nltk.word_tokenize(text)) 

    # Calculate Personal Pronouns
    personal_pronouns = len(re.findall(r'\b(?:I|we|my|ours|us)\b', text, flags=re.IGNORECASE))

    # Calculate Average Word Length
    average_word_length = sum(len(word) for word in nltk.word_tokenize(text)) / total_words

    return {
        'Average Sentence Length': average_sentence_length,
        'Percentage of Complex Words': percentage_complex_words,
        'Fog Index': fog_index,
        'Average Number of Words Per Sentence': average_words_per_sentence,
        'Complex Word Count': len(complex_words),
        'Word Count': total_words,
        'Syllable Count Per Word': syllable_count_per_word,
        'Personal Pronouns': personal_pronouns,
        'Average Word Length': average_word_length
    }

# Create a directory to store the output Excel file
output_dir = 'output'
os.makedirs(output_dir, exist_ok=True)

# Load the Excel file with input data
input_excel_file = 'Input.xlsx'
df = pd.read_excel(input_excel_file)

# Create a new DataFrame to store the output data
output_df = df.copy()

# Iterate through rows in the DataFrame
for index, row in df.iterrows():
    url_id = row['URL_ID']
    text_file_path = os.path.join('extracted_articles', f'{url_id}.txt')

    # Check if the text file exists
    if os.path.exists(text_file_path):
        with open(text_file_path, 'r', encoding='utf-8', errors='ignore') as text_file:
            article_text = text_file.read()

            # Perform sentiment analysis and readability analysis
            sentiment_analysis = analyze_text(article_text)
            readability_analysis = analyze_readability(article_text)

            # Update the output DataFrame with the analysis results
            output_df.at[index, 'POSITIVE SCORE'] = sentiment_analysis['Positive Score']
            output_df.at[index, 'NEGATIVE SCORE'] = sentiment_analysis['Negative Score']
            output_df.at[index, 'POLARITY SCORE'] = sentiment_analysis['Polarity Score']
            output_df.at[index, 'SUBJECTIVITY SCORE'] = sentiment_analysis['Subjectivity Score']
            output_df.at[index, 'AVG SENTENCE LENGTH'] = readability_analysis['Average Sentence Length']
            output_df.at[index, 'PERCENTAGE OF COMPLEX WORDS'] = readability_analysis['Percentage of Complex Words']
            output_df.at[index, 'FOG INDEX'] = readability_analysis['Fog Index']
            output_df.at[index, 'AVG NUMBER OF WORDS PER SENTENCE'] = readability_analysis['Average Number of Words Per Sentence']
            output_df.at[index, 'COMPLEX WORD COUNT'] = readability_analysis['Complex Word Count']
            output_df.at[index, 'WORD COUNT'] = readability_analysis['Word Count']
            output_df.at[index, 'SYLLABLE PER WORD'] = readability_analysis['Syllable Count Per Word']
            output_df.at[index, 'PERSONAL PRONOUNS'] = readability_analysis['Personal Pronouns']
            output_df.at[index, 'AVG WORD LENGTH'] = readability_analysis['Average Word Length']

# Save the output DataFrame to an Excel file
output_excel_file = os.path.join(output_dir, 'output.xlsx')
output_df.to_excel(output_excel_file, index=False)

print("Analysis and output saved to:", output_excel_file)
