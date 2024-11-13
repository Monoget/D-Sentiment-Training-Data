import pandas as pd
import re
from openpyxl import load_workbook

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
        'Liked', 'love', 'enjoy', 'delightful', 'pleased', 'happy', 'satisfied', 'excited', 'affectionate', 'grateful', 'content',
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

# Function to highlight keywords in text with {} for positive and [] for negative
def highlight_keywords_in_text(text, keywords):
    highlighted_text = text  # Start with the original text
    offset = 0  # Track changes to string length as we add braces

    # Ensure 'keywords' is structured as a dictionary with 'positive' and 'negative' categories
    if not isinstance(keywords, dict):
        print("Keywords should be passed as a dictionary with 'positive' and 'negative' keys.")
        return text

    # Loop through categories (positive and negative) and apply appropriate markers
    for category, markers in [('positive', '{}'), ('negative', '[]')]:
        if category not in keywords:
            continue
        for keyword in keywords[category]:
            # Escape special characters in the keyword for regex
            keyword_escaped = re.escape(keyword)

            # Find all occurrences of the keyword (case-insensitive)
            matches = list(re.finditer(rf'\b{keyword_escaped}\b', highlighted_text, re.IGNORECASE))

            for match in matches:
                start, end = match.start() + offset, match.end() + offset
                # Apply the appropriate marker ({} or [])
                highlighted_text = highlighted_text[:start] + markers.format(highlighted_text[start:end]) + highlighted_text[end:]
                offset += len(markers) - 2  # Adjust offset for added markers ({} or [])

    return highlighted_text

# Main function to process each review and return detailed scoring breakdown
def process_review(text):
    sensory = keyword_score(text, SENSORY_KEYWORDS)
    affective = keyword_score(text, AFFECTIVE_KEYWORDS)
    intellectual = keyword_score(text, INTELLECTUAL_KEYWORDS)
    behavior = keyword_score(text, BEHAVIOR_KEYWORDS)
    recommend = keyword_score(text, RECOMMEND_KEYWORDS)

    return {
        "Review Text": highlight_keywords_in_text(text, SENSORY_KEYWORDS),
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
    review_text = row['Review Text']  # Adjust 'Review Text' to the actual column name for reviews
    scores = process_review(review_text)
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
ws.column_dimensions['B'].width = 50  # Set column B (Review Text) width to 50 characters
for col in ['C', 'D', 'E', 'F', 'G']:  # Other columns (Sensory, Affective, Intellectual, Behavior, Recommend)
    ws.column_dimensions[col].width = 13  # Set width of all other columns to 13 characters

# Apply text wrapping for all columns except the first
for col in ws.columns:
    for cell in col:
        cell.alignment = cell.alignment.copy(wrapText=True)

# Save the workbook with formatting changes
wb.save(output_file_path)

print(f"Processed file saved to: {output_file_path}")
