import os
import pandas as pd
import nltk
from nltk.tokenize import word_tokenize, sent_tokenize
import re
import codecs
import pronouncing


# Download NLTK resources
nltk.download('punkt')

def load_dictionary(file_path, encoding='utf-8'):
    with codecs.open(file_path, 'r', encoding=encoding, errors='replace') as file:
        return set(file.read().splitlines())

def load_stop_words(directory, file_pattern):
    stop_words = set()
    for filename in os.listdir(directory):
        if filename.startswith(file_pattern) and filename.endswith(".txt"):
            file_path = os.path.join(directory, filename)
            with codecs.open(file_path, 'r', encoding='utf-8-sig', errors='replace') as file:
                stop_words.update(word.strip().lower() for word in file.readlines())
    return stop_words

def clean_text(text, stop_words):
    # Tokenize the text
    words = word_tokenize(text.lower())

    # Remove custom stop words
    words = [word for word in words if word.isalnum() and word not in stop_words]

    return words

def calculate_sentiment_scores(words, positive_dict, negative_dict):
    # Calculate Positive Score
    positive_score = sum(1 for word in words if word in positive_dict)

    # Calculate Negative Score
    negative_score = sum(1 for word in words if word in negative_dict)

    # Calculate Polarity Score
    polarity_score = (positive_score - negative_score) / (positive_score + negative_score + 0.000001)

    # Calculate Subjectivity Score
    subjectivity_score = (positive_score + negative_score) / (len(words) + 0.000001)

    return positive_score, negative_score, polarity_score, subjectivity_score

def calculate_readability(words, sentences):
    # Count complex words using pronouncing library for syllable counting
    complex_words = sum(1 for word in words if len(word) > 2 and pronouncing.syllable_count(word) > 2)

    # Calculate Average Number of Words Per Sentence
    avg_words_per_sentence = len(words) / len(sentences)

    # Calculate Word Count
    word_count = len(words)

    # Calculate Syllable Count Per Word
    syllable_per_word = sum(pronouncing.syllable_count(word) for word in words) / len(words)

    # Calculate Average Sentence Length
    avg_sentence_length = len(words) / len(sentences)

    # Calculate Percentage of Complex Words
    percentage_complex_words = (complex_words / len(words)) * 100

    # Calculate Fog Index
    fog_index = 0.4 * (avg_sentence_length + percentage_complex_words)

    return avg_words_per_sentence, complex_words, word_count, syllable_per_word, avg_sentence_length, percentage_complex_words, fog_index

def calculate_personal_pronouns(text):
    # Count Personal Pronouns
    personal_pronouns_count = sum(1 for word in re.findall(r'\b(?:i|we|my|ours|us)\b', text, flags=re.IGNORECASE))

    return personal_pronouns_count

def calculate_avg_word_length(words):
    # Calculate Average Word Length
    avg_word_length = sum(len(word) for word in words) / len(words)

    return avg_word_length

def analyze_text_file(url_id, url, text, positive_dict, negative_dict, custom_stop_words):
    # Clean text with custom stop words
    words = clean_text(text, custom_stop_words)

    # Tokenize text into words and sentences
    sentences = sent_tokenize(text)

    # Calculate sentiment scores
    positive_score, negative_score, polarity_score, subjectivity_score = calculate_sentiment_scores(words, positive_dict, negative_dict)

    # Calculate readability metrics
    avg_words_per_sentence, complex_words, word_count, syllable_per_word, avg_sentence_length, percentage_complex_words, fog_index = calculate_readability(words, sentences)

    # Calculate personal pronouns count
    personal_pronouns_count = calculate_personal_pronouns(text)

    # Calculate average word length
    avg_word_length = calculate_avg_word_length(words)

    # Prepare a dictionary with the analysis results
    analysis_result = {
        'URL_ID': url_id,
        'URL': f'=HYPERLINK("{url}", "{url}")',
        'POSITIVE SCORE': positive_score,
        'NEGATIVE SCORE': negative_score,
        'POLARITY SCORE': polarity_score,
        'SUBJECTIVITY SCORE': subjectivity_score,
        'AVG SENTENCE LENGTH': avg_sentence_length,
        'PERCENTAGE OF COMPLEX WORDS': percentage_complex_words,
        'FOG INDEX': fog_index,
        'AVG NUMBER OF WORDS PER SENTENCE': avg_words_per_sentence,
        'COMPLEX WORD COUNT': complex_words,
        'WORD COUNT': word_count,
        'SYLLABLE PER WORD': syllable_per_word,
        'PERSONAL PRONOUNS': personal_pronouns_count,
        'AVG WORD LENGTH': avg_word_length
    }

    # Display the results
    print(analysis_result)
    print("\n" + "=" * 50 + "\n")  # Separator
    

    return analysis_result


def main():
    # Load positive and negative dictionaries
     # Load positive and negative dictionaries
    positive_dict = load_dictionary('MasterDictionary/positive-words.txt', 'utf-8-sig')
    negative_dict = load_dictionary('MasterDictionary/negative-words.txt', 'utf-8-sig')

    # Specify the directory containing stop word files
    stop_words_directory = "StopWords"

    # Load custom stop words
    custom_stop_words = load_stop_words(stop_words_directory, "StopWords_")

    # Load URLs from the input Excel file
    input_file = "Input.xlsx"
    df = pd.read_excel(input_file)

    analysis_results_list = []

    for index, row in df.iterrows():
        url_id = row['URL_ID']
        url = row['URL']
        file_path = f"{url_id}.txt"

        # Read the text file
        with open(file_path, "r", encoding="utf-8") as file:
            text = file.read()

        # Analyze the text file and append the result to the list
        analysis_result = analyze_text_file(url_id, url, text, positive_dict, negative_dict, custom_stop_words)
        analysis_results_list.append(analysis_result)

    # Convert the list of dictionaries to a DataFrame
    df_output = pd.DataFrame(analysis_results_list)

    # Save the output DataFrame to an Excel file
    output_file = "Output.xlsx"
    df_output.to_excel(output_file, index=False)

if __name__ == "__main__":
    main()