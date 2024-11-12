import pandas as pd
import re
from openpyxl import load_workbook
from openpyxl.styles import Font
from textblob import TextBlob  # This is for sentiment analysis, added per your request.

# Define keywords for each category
SENSORY_KEYWORDS = ['sound', 'audio', 'voice', 'listen', 'hear', 'volume', 'tune']
AFFECTIVE_KEYWORDS = ['love', 'like', 'enjoy', 'appreciate', 'emotion', 'sentiment', 'feel', 'feeling', 'affection']
INTELLECTUAL_KEYWORDS = ['think', 'judge', 'considering', 'curiosity', 'evaluate', 'ponder', 'analyze', 'understand']
BEHAVIOR_KEYWORDS = ['bought', 'buy', 'purchase', 'acquire', 'act', 'use', 'apply', 'bodily', 'do', 'try']
RECOMMEND_KEYWORDS = ['recommend', 'suggest', 'advise', 'endorse', 'propose', 'encourage']

# Scoring function for each category based on keyword matching
def keyword_score(text, keywords):
    count = sum(1 for word in keywords if re.search(rf'\b{word}\b', text, re.IGNORECASE))
    # Map the count to a score from 1 to 3
    if count > 2:
        return 3
    elif count > 0:
        return 2
    else:
        return 1


# Overall sentiment score using TextBlob for affective categories
def affective_score(text):
    keyword_count = keyword_score(text, AFFECTIVE_KEYWORDS)
    polarity = TextBlob(text).sentiment.polarity
    if polarity > 0.5:
        sentiment_score = 3
    elif polarity > 0:
        sentiment_score = 2
    else:
        sentiment_score = 1
    return max(keyword_count, sentiment_score)


# Function to highlight keywords in red in the Excel file
def highlight_keywords_in_text(text, keywords):
    # Initialize the list of chunks of text
    chunks = []
    start = 0
    # Loop through the text to find keywords
    for keyword in keywords:
        for match in re.finditer(rf'\b{keyword}\b', text, re.IGNORECASE):
            # Add the part of the text before the match
            chunks.append(text[start:match.start()])
            # Add the keyword (this is what we'll color)
            chunks.append(f"{{{{{match.group(0)}}}}}")  # Mark keyword for later color application
            start = match.end()
    # Add any remaining part of the text after the last match
    chunks.append(text[start:])
    return chunks


# Main function to process each review and return detailed scoring breakdown
def process_review(text):
    sensory = keyword_score(text, SENSORY_KEYWORDS)
    affective = affective_score(text)
    intellectual = keyword_score(text, INTELLECTUAL_KEYWORDS)
    behavior = keyword_score(text, BEHAVIOR_KEYWORDS)
    recommend = keyword_score(text, RECOMMEND_KEYWORDS)

    return {
        "ReviewText": text,
        "Sensory": sensory,
        "Affective": affective,
        "Intellectual": intellectual,
        "Behavior": behavior,
        "Recommend": recommend
    }


# Load data from the Excel file
input_file_path = 'input/Training data_Coding.xlsx'
df = pd.read_excel(input_file_path, header=2)  # Header is in row 3

# Initialize a list to store the scores for each review
scores_list = []

# Process each review starting from row 4
for index, row in df.iterrows():
    review_text = row['Review Text']  # Adjust 'ReviewText' to the actual column name for reviews
    scores = process_review(review_text)

    # Highlight the keywords and add the processed review
    highlighted_review = highlight_keywords_in_text(review_text,
                                                     SENSORY_KEYWORDS + AFFECTIVE_KEYWORDS + INTELLECTUAL_KEYWORDS + BEHAVIOR_KEYWORDS + RECOMMEND_KEYWORDS)
    scores['Review Text'] = ' '.join(highlighted_review)  # Join the chunks back into a string
    scores_list.append(scores)

# Convert scores to a DataFrame and reorder columns as requested
scores_df = pd.DataFrame(scores_list)
scores_df = scores_df[['Review Text', 'Sensory', 'Affective', 'Intellectual', 'Behavior', 'Recommend']]

# Add "Sl" column for row numbers
scores_df.insert(0, 'Sl', range(1, len(scores_df) + 1))

# Save the output to an Excel file
output_file_path = 'output/scored_reviews_highlighted.xlsx'
scores_df.to_excel(output_file_path, index=False)

# Open the saved file to apply text color formatting
wb = load_workbook(output_file_path)
ws = wb.active

# Apply red text color to the specific keywords enclosed in curly braces
for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=2, max_col=2):  # 'ReviewText' column
    for cell in row:
        cell_value = cell.value
        start = 0
        # Search for keywords enclosed in curly braces
        while True:
            start = cell_value.find('{{', start)
            if start == -1:
                break
            end = cell_value.find('}}', start)
            if end != -1:
                keyword = cell_value[start+2:end]
                # Apply red color to the keyword
                for match in re.finditer(rf'\b{keyword}\b', cell_value, re.IGNORECASE):
                    cell.font = Font(color="FF0000")  # Apply red color to the matched keyword
                start = end + 2
            else:
                break

wb.save(output_file_path)

print("Scores and highlighted reviews have been saved to", output_file_path)
