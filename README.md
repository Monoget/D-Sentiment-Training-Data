# Review Scoring and Keyword Highlighting Project

## Overview

This project processes reviews, scores them based on predefined keyword categories, highlights the keywords, and saves the results in a new Excel file. It can be used for analyzing textual reviews or feedback to extract sentiment or key themes using keyword-based scoring.

The goal is to:
1. **Score reviews** based on predefined positive and negative keywords related to sensory, affective, intellectual, behavioral, and recommendation categories.
2. **Highlight keywords** found within the review text to emphasize key points.
3. **Save the output** to a new Excel file for easy analysis.

## Table of Contents

1. [Project Overview](#overview)
2. [Installation](#installation)
3. [How It Works](#how-it-works)
4. [Project Structure](#project-structure)
5. [How to Use](#how-to-use)
6. [View the Output](#view-the-output)


## Installation

To run this project, you need Python and a few libraries installed. Follow these steps to set up the project:

1. Install Python (if you don’t already have it installed). You can download it from [python.org](https://www.python.org/).

2. Install the required libraries by running the following command in your terminal:

    ```bash
    pip install pandas openpyxl textblob
    ```

    These libraries are used for:
    - **pandas**: For reading and writing Excel files.
    - **openpyxl**: For manipulating Excel files, especially for formatting.
    - **textblob**: For text analysis (optional but useful for other projects or text processing).

## How It Works

### 1. Loading Data
The script starts by reading an Excel file containing a column of review text. It expects the reviews to be located in a column named `Review Text`. This column will be processed to extract information based on the keyword matching.

### 2. Defining Keyword Categories
There are predefined sets of **positive** and **negative** keywords for five different categories:
- **Sensory**: Keywords related to sensory experience (e.g., "clear", "smooth").
- **Affective**: Keywords related to emotional responses (e.g., "happy", "disappointed").
- **Intellectual**: Keywords related to intellectual qualities (e.g., "insightful", "shallow").
- **Behavior**: Keywords related to behavior or actions (e.g., "effective", "annoying").
- **Recommend**: Keywords related to recommendations (e.g., "would recommend", "discourage").

### 3. Scoring the Reviews
For each review:
- The presence of **positive** and **negative** keywords is checked in the review text.
- A score is assigned based on the following logic:
  - **3 points** for each positive keyword found.
  - **1 point** for each negative keyword found.
  - **2 points** if neither positive nor negative keywords are found (neutral score).

The review score is calculated for each category separately based on the number of matching keywords.

### 4. Highlighting Keywords in the Review
The keywords found in the review are highlighted:
- **Positive keywords** are highlighted using curly braces `{}`.
- **Negative keywords** are highlighted using square brackets `[]`.

### 5. Saving the Output
The final output is saved in a new Excel file with the following columns:
- **Sl (Serial number)**: A unique number for each review.
- **Review Text**: The review text with highlighted keywords.
- **Sensory**: The score for the sensory category.
- **Affective**: The score for the affective category.
- **Intellectual**: The score for the intellectual category.
- **Behavior**: The score for the behavior category.
- **Recommend**: The score for the recommendation category.

## Project Structure

The project folder should look like this:

 ├── input/  
 │ └── Training data_Coding.xlsx # Input file containing reviews  
 ├── output/   
 │ └── scored_reviews_highlighted.xlsx # Output file with scores and highlighted keywords   
 ├── main.py # Python script that processes the reviews   
 └── README.md # This file
 

### `script.py`
This is the Python script that processes the review data, calculates the scores, highlights the keywords, and saves the output in the `output` folder.

### `Training data_Coding.xlsx`
This file contains the reviews that you want to analyze. The script assumes the reviews are stored in a column called `Review Text`. Ensure your input file follows this format.

### `scored_reviews_highlighted.xlsx`
This is the output file where processed reviews are saved. It will include:
- The original review text with highlighted keywords.
- Scores for each of the five categories.

## How to Use

### 1. Prepare Input File
Ensure your input Excel file is in the `input/` folder and follows this format:

| Review Text                         |
|--------------------------------------|
| I love the clear and smooth sound.   |
| The service was frustratingly slow.  |
| Insightful and well-structured.      |
| Would not recommend due to poor quality. |

### 2. Run the Script

Run the following command to start processing:

```bash
python script.py
```
## View the Output

After running the script, the processed output will be saved in the `output/` directory. You can view the results in the `scored_reviews_highlighted.xlsx` file. This file will contain the following:

- **Review Text**: The original review text with highlighted keywords (positive keywords are enclosed in `{}` and negative keywords are enclosed in `[]`).
- **Sensory**: A score indicating the sentiment of the sensory-related words in the review.
- **Affective**: A score indicating the sentiment of the emotional-related words in the review.
- **Intellectual**: A score indicating the sentiment of the intellectual-related words in the review.
- **Behavior**: A score indicating the sentiment of the behavioral-related words in the review.
- **Recommend**: A score indicating the sentiment of the recommendation-related words in the review.

You can open the `scored_reviews_highlighted.xlsx` file using any spreadsheet application (such as Microsoft Excel or Google Sheets) to examine the results.


| Sl  | Review Text                                   | Sensory | Affective | Intellectual | Behavior | Recommend |
|-----|-----------------------------------------------|---------|-----------|--------------|----------|-----------|
| 1   | I {love} the {clear} and {smooth} sound.      | 3       | 2         | 2            | 2        | 2         |
| 2   | The service was [frustratingly] slow.         | 2       | 1         | 2            | 1        | 2         |
| 3   | {Insightful} and well-structured.             | 2       | 3         | 3            | 2        | 2         |
| 4   | Would not {recommend} due to poor quality.    | 2       | 2         | 2            | 1        | 1         |
