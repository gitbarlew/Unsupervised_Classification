import pandas as pd
import nltk
from collections import Counter
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize

# Ensure you have the necessary NLTK data
nltk.download('punkt')
nltk.download('stopwords')


def load_excel(file_path):
    return pd.read_excel(file_path)


def preprocess_text(text):
    if pd.isna(text):
        return []
    tokens = word_tokenize(text.lower())
    tokens = [word for word in tokens if word.isalpha()]
    tokens = [word for word in tokens if word not in stopwords.words('english')]
    return tokens


def identify_categories(responses, num_categories=10):
    all_words = []
    for response in responses:
        all_words.extend(preprocess_text(response))

    most_common = Counter(all_words).most_common(num_categories)
    categories = [word for word, count in most_common]
    return categories


def categorize_responses(responses, categories):
    categorized_responses = []
    for response in responses:
        if pd.isna(response):
            categorized_responses.append('N/A')
            continue
        tokens = set(preprocess_text(response))
        response_category = None
        for category in categories:
            if category in tokens:
                response_category = category
                break
        categorized_responses.append(response_category or 'Other')
    return categorized_responses


# Example usage
file_path = 'input_file.xlsx'
column_name = 'Response'

df = load_excel(file_path)
responses = df[column_name]
categories = identify_categories(responses)
print(categories)
df['Categorized Responses'] = categorize_responses(responses, categories)

# Save the DataFrame back to Excel
output_file_path = 'output_file.xlsx'
df.to_excel(output_file_path, index=False)

print(f"Categorized responses added and saved to {output_file_path}")
