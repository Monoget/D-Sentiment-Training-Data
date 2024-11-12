# Review Scoring and Keyword Highlighting

## Project Description

This project processes product or service reviews stored in an Excel file. The script reads reviews, analyzes them for keywords related to different categories (Sensory, Affective, Intellectual, Behavior, and Recommend), scores each review based on keyword frequency, and then saves the processed data to a new Excel file. The reviews are also highlighted with specific keywords, and the resulting file includes the scores for each category.

### Features:
- **Keyword Scoring**: Each review is analyzed for specific keywords related to sensory, affective, intellectual, behavioral, and recommendation-related terms. The script assigns scores based on keyword matches.
- **Sentiment Analysis**: The script uses TextBlob to perform sentiment analysis for reviews that focus on emotional sentiment (affective category).
- **Excel Output**: The results are saved into an Excel file with scores for each category and a column for the row number (Sl). Review text is highlighted for the identified keywords.
- **Column Formatting**: The Excel output file has adjusted column widths and text wrapping.

---

## Prerequisites

To run the script, you need Python and several libraries installed on your system:

### 1. **Python** (version 3.x):
   - Download and install Python from the official website: https://www.python.org/downloads/
   - Ensure that `python` is added to your system's PATH.

### 2. **Required Libraries**:
   - `pandas` for data manipulation
   - `openpyxl` for handling Excel files
   - `textblob` for sentiment analysis
   - `re` (regular expressions) for keyword matching
   
   To install the required libraries, run:

   ```bash
   pip install pandas openpyxl textblob
```

## Script Breakdown
### 1. Imports
```bash
import pandas as pd
import re
from openpyxl import load_workbook
from openpyxl.styles import Font
from textblob import TextBlob
```
 - pandas: For working with Excel files and data manipulation.
 - re: For using regular expressions to search and match keywords in text.
 - openpyxl: For reading, writing, and formatting Excel files.
 - TextBlob: For sentiment analysis.

### 2. Keyword Categories
```bash
# Define keywords for each category
SENSORY_KEYWORDS = ['sound', 'audio', 'voice', 'listen', 'hear', 'volume', 'tune']
AFFECTIVE_KEYWORDS = ['love', 'like', 'enjoy', 'appreciate', 'emotion', 'sentiment', 'feel', 'feeling', 'affection']
INTELLECTUAL_KEYWORDS = ['think', 'judge', 'considering', 'curiosity', 'evaluate', 'ponder', 'analyze', 'understand']
BEHAVIOR_KEYWORDS = ['bought', 'buy', 'purchase', 'acquire', 'act', 'use', 'apply', 'bodily', 'do', 'try']
RECOMMEND_KEYWORDS = ['recommend', 'suggest', 'advise', 'endorse', 'propose', 'encourage']
```
These lists contain keywords grouped into five categories that the script uses to score the reviews.

### 3. Keyword Matching Function
```bash
def keyword_score(text, keywords):
    count = sum(1 for word in keywords if re.search(rf'\b{word}\b', text, re.IGNORECASE))
    if count > 2:
        return 3
    elif count > 0:
        return 2
    else:
        return 1
```
 - This function checks how many times a keyword from a given category appears in a review.
 - It assigns a score based on the number of keyword matches:
   - Score 3: More than 2 matches.
   - Score 2: 1 or 2 matches.
   - Score 1: No matches.
   
### 4. Sentiment Analysis for the Affective Category
```bash
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
```

 - Affective Score is determined by the number of keyword matches in the AFFECTIVE_KEYWORDS list, combined with the sentiment polarity calculated by TextBlob:
   - Polarity > 0.5 = positive sentiment → Score 3.
   - Polarity between 0 and 0.5 = neutral positive sentiment → Score 2.
   - Polarity < 0 = negative sentiment → Score 1.
 - The function returns the higher score between keyword count and sentiment analysis.

### 5. Highlighting Keywords in Reviews
```bash
def highlight_keywords_in_text(text, keywords):
    chunks = []
    start = 0
    for keyword in keywords:
        for match in re.finditer(rf'\b{keyword}\b', text, re.IGNORECASE):
            chunks.append(text[start:match.start()])
            chunks.append(f"{{{{{match.group(0)}}}}}")
            start = match.end()
    chunks.append(text[start:])
    return chunks
```
 - This function highlights the keywords by wrapping them in double curly braces `({{keyword}})`.
 - It splits the review text into chunks, with keywords marked separately for easy identification.

### 6. Processing Reviews
```bash
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
```
 - This function processes each review text, calculates scores for each category, and returns a dictionary containing the scores for `Sensory`, `Affective`, `Intellectual`, `Behavior`, and `Recommend`.

## 7. Loading and Processing the Excel Data

```bash
input_file_path = 'input/Training data_Coding.xlsx'
df = pd.read_excel(input_file_path, header=2)
```

 - The reviews are loaded from an Excel file located at `input/Training data_Coding.xlsx`.
 - The header is located at the third row `(header=2)`.

## 8. Processing Each Review and Saving Results

```bash
scores_list = []

for index, row in df.iterrows():
    review_text = row['Review Text']
    scores = process_review(review_text)

    highlighted_review = highlight_keywords_in_text(review_text,
                                                     SENSORY_KEYWORDS + AFFECTIVE_KEYWORDS + INTELLECTUAL_KEYWORDS + BEHAVIOR_KEYWORDS + RECOMMEND_KEYWORDS)
    scores['Review Text'] = ' '.join(highlighted_review)
    scores_list.append(scores)

scores_df = pd.DataFrame(scores_list)
scores_df = scores_df[['Review Text', 'Sensory', 'Affective', 'Intellectual', 'Behavior', 'Recommend']]
```

 - This loop iterates over all reviews, processes each one, scores them, highlights keywords, and stores the results in `scores_list`.
 - The results are stored in a DataFrame with columns for the review text and the scores.

## 9. Column Formatting
```bash
scores_df.insert(0, 'Sl', range(1, len(scores_df) + 1))
```
 - Adds a "Sl" column to number the rows in the output Excel file.


## 10. Saving the Output Excel File
```bash
output_file_path = 'output/scored_reviews_highlighted.xlsx'
scores_df.to_excel(output_file_path, index=False)
```
 - The results are saved into a new Excel file at `output/scored_reviews_highlighted.xlsx`.

## 11. Formatting the Excel File
```bash
wb = load_workbook(output_file_path)
ws = wb.active

for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=2, max_col=2):
    for cell in row:
        cell_value = cell.value
        start = 0
        while True:
            start = cell_value.find('{{', start)
            if start == -1:
                break
            end = cell_value.find('}}', start)
            if end != -1:
                keyword = cell_value[start+2:end]
                for match in re.finditer(rf'\b{keyword}\b', cell_value, re.IGNORECASE):
                    cell.font = Font(color="FF0000")
                start = end + 2
            else:
                break

wb.save(output_file_path)
```

 - This block loads the generated Excel file, highlights the keywords (using red font), and saves the updated file.

## 12. Final Output
```bash
print("Scores and highlighted reviews have been saved to", output_file_path)
```

- Prints a message indicating that the process is complete and the file has been saved.


## Run the Program
```bash
python main.py
```