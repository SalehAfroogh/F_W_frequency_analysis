# Keyword Frequency Analysis Project

This project analyzes the frequency of specific keywords in a Word document using corresponding glossary files. It performs visualizations for each glossary and exports the results for further analysis.

## Features

Extracts text from a Word document.

Reads keywords from multiple glossary Excel files.

Counts the frequency of keywords in the text.

Generates bar charts and word clouds for visual representation.

Saves the keyword frequency results as CSV files.

## Prerequisites

Make sure the following Python libraries are installed:

pandas

matplotlib

wordcloud

python-docx

You can install these using pip:

pip install pandas matplotlib wordcloud python-docx

Directory Structure

The project requires the following structure:

<img width="554" alt="Screenshot 2024-12-27 at 2 01 46 PM" src="https://github.com/user-attachments/assets/e9192359-b03d-4123-b578-26a28267a823" />





## Usage

Place the Word document (Fountains of Wisdom.docx) and the glossary Excel files in the same folder as the script.

Run the script:

python frequency_analysis.py

The script will:

Analyze each glossary file.

Create bar charts and word clouds for keyword frequencies.

Save results in CSV files with names like Epistemology_Keyword_Frequency_Results.csv.

## Outputs

Bar Charts: Visual representation of keyword frequencies.

Word Clouds: Aesthetic visualization of keyword prominence.

CSV Files: Contains keyword frequencies for each glossary.

## Example Output

For the Epistemology_Glossary.xlsx:

Bar Chart: Displays the frequency of each keyword.

Word Cloud: Highlights the most frequent keywords.

CSV File: Epistemology_Keyword_Frequency_Results.csv

## Customization

Add more glossary files to the glossary_files list in frequency_analysis.py.

Modify visualization colors in the colors list.
