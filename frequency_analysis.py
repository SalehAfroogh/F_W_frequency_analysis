import re
import pandas as pd
from collections import Counter
from docx import Document
import matplotlib.pyplot as plt
from wordcloud import WordCloud

def extract_text_from_docx(file_path):
    """Extract text from a Word (.docx) file."""
    doc = Document(file_path)
    return " ".join([paragraph.text for paragraph in doc.paragraphs])

def count_keywords(text, keywords):
    """Count occurrences of each keyword in the text."""
    counts = Counter()
    for keyword in keywords:
        # Create a regex pattern to match the keyword as a whole word
        pattern = fr"\b{re.escape(keyword)}\b"
        counts[keyword] = len(re.findall(pattern, text, re.IGNORECASE))
    return counts

def visualize_results(keyword_counts, title, color):
    """Visualize the frequency analysis results."""
    # Filter out keywords with zero counts
    filtered_counts = {k: v for k, v in keyword_counts.items() if v > 0}
    
    # Sort the counts for better visualization
    sorted_counts = dict(sorted(filtered_counts.items(), key=lambda x: x[1], reverse=True))

    if not sorted_counts:
        print(f"No keywords found for {title}. Skipping visualization.")
        return

    # Bar Chart
    plt.figure(figsize=(12, 8))
    plt.bar(sorted_counts.keys(), sorted_counts.values(), color=color)
    plt.xticks(rotation=90, fontsize=8)
    plt.title(f"Keyword Frequency Analysis: {title}")
    plt.xlabel("Keywords")
    plt.ylabel("Frequency")
    plt.tight_layout()
    plt.show()

    # Word Cloud
    wordcloud = WordCloud(width=800, height=400, background_color='white').generate_from_frequencies(filtered_counts)
    plt.figure(figsize=(10, 6))
    plt.imshow(wordcloud, interpolation='bilinear')
    plt.axis("off")
    plt.title(f"Keyword Frequency Word Cloud: {title}")
    plt.show()

def main():
    # Paths to the files
    word_file_path = "/Users/salehafroogh/Desktop/Fountains of Wisdom/Fountains of Wisdom.docx"
    glossary_files = [
        "/Users/salehafroogh/Desktop/Fountains of Wisdom/Epistemology_Glossary.xlsx",
        "/Users/salehafroogh/Desktop/Fountains of Wisdom/Ethics_Glossary.xlsx",
        "/Users/salehafroogh/Desktop/Fountains of Wisdom/Logic_Glossary.xlsx",
        "/Users/salehafroogh/Desktop/Fountains of Wisdom/Metaphysics_Glossary.xlsx",
        "/Users/salehafroogh/Desktop/Fountains of Wisdom/Natural_Philosophy_Glossary.xlsx",
    ]

    colors = ["skyblue", "lightgreen", "coral", "gold", "violet"]  # Different colors for each glossary

    # Extract text from the Word file
    text = extract_text_from_docx(word_file_path)

    for glossary_file, color in zip(glossary_files, colors):
        # Load keywords from the current Excel file
        keywords_df = pd.read_excel(glossary_file)
        keywords = keywords_df["Keywords"].dropna().tolist()  # Ensure no NaN values are included

        # Perform keyword frequency analysis
        keyword_counts = count_keywords(text, keywords)

        # Extract glossary name for visualization titles
        glossary_name = glossary_file.split("/")[-1].replace("_Glossary.xlsx", "").replace("_", " ")

        # Visualize the results
        visualize_results(keyword_counts, glossary_name, color)

        # Save results to a CSV file
        output_csv = f"/Users/salehafroogh/Desktop/Fountains of Wisdom/{glossary_name}_Keyword_Frequency_Results.csv"
        results_df = pd.DataFrame(keyword_counts.items(), columns=["Keyword", "Frequency"])
        results_df.to_csv(output_csv, index=False)
        print(f"Keyword frequency analysis for '{glossary_name}' completed and saved to '{output_csv}'.")

if __name__ == "__main__":
    main()
