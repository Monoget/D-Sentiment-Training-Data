import pandas as pd
import re
from openpyxl import load_workbook
from openpyxl.styles import Font
from textblob import TextBlob

# Define expanded keywords for each category
SENSORY_KEYWORDS = {
    'positive': [
        'clear', 'crisp', 'pleasant', 'soothing', 'melodious', 'engaging', 'distinct', 'vivid', 'vibrant', 'sharp',
        'harmonious', 'resonant', 'comforting', 'smooth', 'appealing', 'soft', 'pure', 'clean', 'refreshing', 'symphonic'
    ],
    'negative': [
        'noisy', 'distorted', 'unpleasant', 'dull', 'muffled', 'grating', 'harsh', 'blurry', 'fuzzy', 'jarring',
        'disorienting', 'irritating', 'loud', 'abrasive', 'annoying', 'clashing', 'rough', 'static', 'smudged', 'glaring'
    ]
}

AFFECTIVE_KEYWORDS = {
    'positive': [
        'Like', 'love', 'enjoy', 'delightful', 'pleased', 'happy', 'satisfied', 'excited', 'affectionate', 'grateful', 'content',
        'joyful', 'passionate', 'thrilled', 'enthusiastic', 'appreciative', 'fond', 'cherish', 'admire', 'relieved', 'hopeful'
    ],
    'negative': [
        'dislike', 'hate', 'disappointing', 'upset', 'annoyed', 'frustrated', 'angry', 'sad', 'bored', 'apathetic',
        'disheartened', 'irritated', 'resentful', 'dismayed', 'displeased', 'discouraged', 'unhappy', 'regretful', 'dejected', 'miserable'
    ]
}

INTELLECTUAL_KEYWORDS = {
    'positive': [
        'insightful', 'thought-provoking', 'curious', 'analytical', 'wise', 'logical', 'informed', 'perceptive', 'intelligent',
        'enlightening', 'smart', 'strategic', 'astute', 'knowledgeable', 'reflective', 'cerebral', 'scholarly', 'innovative',
        'rational', 'profound'
    ],
    'negative': [
        'confused', 'misleading', 'uninformed', 'illogical', 'shallow', 'ignorant', 'unclear', 'vague', 'dense', 'superficial',
        'naive', 'simplistic', 'irrational', 'nonsensical', 'pointless', 'absurd', 'disoriented', 'misunderstood', 'foggy', 'convoluted'
    ]
}

BEHAVIOR_KEYWORDS = {
    'positive': [
        'effective', 'useful', 'helpful', 'successful', 'productive', 'practical', 'beneficial', 'efficient', 'valuable',
        'engaging', 'constructive', 'proactive', 'goal-oriented', 'impactful', 'empowering', 'positive', 'meaningful', 'focused',
        'reliable', 'responsive'
    ],
    'negative': [
        'ineffective', 'useless', 'wasteful', 'unsuccessful', 'inefficient', 'frustrating', 'difficult', 'impractical', 'unproductive',
        'irrelevant', 'complicated', 'cumbersome', 'slow', 'misguided', 'dysfunctional', 'obstructive', 'exhausting', 'taxing',
        'draining', 'tedious'
    ]
}

RECOMMEND_KEYWORDS = {
    'positive': [
        'highly', 'strongly', 'eagerly', 'confidently', 'enthusiastically', 'favorably', 'encouragingly', 'endorse', 'back',
        'support', 'vouch', 'praise', 'commend', 'applaud', 'approve', 'advocate', 'affirm', 'acclaim', 'second', 'uphold'
    ],
    'negative': [
        'wouldn’t', 'don’t', 'avoid', 'discourage', 'hesitate', 'doubt', 'reluctant', 'regret', 'warn', 'disapprove',
        'advise against', 'criticize', 'oppose', 'reject', 'caution', 'rebuke', 'dismiss', 'deter', 'withhold', 'restrain'
    ]
}


# Scoring function for each category based on positive/negative keyword matching
def keyword_score(text, keyword_dict):
    positive_count = sum(1 for word in keyword_dict['positive'] if re.search(rf'\b{word}\b', text, re.IGNORECASE))
    negative_count = sum(1 for word in keyword_dict['negative'] if re.search(rf'\b{word}\b', text, re.IGNORECASE))

    # Assign scores: positive=3, negative=1, none found=2
    if positive_count > 0:
        return 3
    elif negative_count > 0:
        return 1
    else:
        return 2

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

# Function to highlight keywords in text (without color formatting)
def highlight_keywords_in_text(text, keywords):
    chunks = []
    start = 0
    for keyword in keywords['positive'] + keywords['negative']:
        for match in re.finditer(rf'\b{keyword}\b', text, re.IGNORECASE):
            chunks.append(text[start:match.start()])
            chunks.append(f"{{{{{match.group(0)}}}}}")  # Mark keyword for later processing
            start = match.end()
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

# Process each review
for index, row in df.iterrows():
    review_text = row['Review Text']  # Adjust 'ReviewText' to the actual column name for reviews
    scores = process_review(review_text)

    # Highlight the keywords and add the processed review
    highlighted_review = highlight_keywords_in_text(review_text, SENSORY_KEYWORDS)
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

# Open the saved file to apply column width settings
wb = load_workbook(output_file_path)
ws = wb.active

# Set the column width
ws.column_dimensions['B'].width = 50  # Set column B (Review Text) width to 500px
for col in ['C', 'D', 'E', 'F', 'G']:  # Other columns (Sensory, Affective, Intellectual, Behavior, Recommend)
    ws.column_dimensions[col].width = 13  # Set width of all other columns to 100px

# Apply text wrapping for all columns except the first
for col in ['B', 'C', 'D', 'E', 'F', 'G']:
    for cell in ws[col]:
        cell.alignment = cell.alignment.copy(wrap_text=True)

# Save the formatted Excel file
wb.save(output_file_path)

print("Scores and reviews with adjusted formatting have been saved to", output_file_path)
