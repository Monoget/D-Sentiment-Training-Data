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
