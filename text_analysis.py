import pandas as pd
import os
import nltk
from nltk.tokenize import word_tokenize
import docx
import re
import string
from nltk.corpus import stopwords
#nltk.download('stopwords')
#nltk.download('punkt')

def output_file_write(input_file, output_dict):
    # Load the existing Excel file
    output_file = r"C:\Users\navin\OneDrive\Desktop\WorkTree\sample_table\Entevyuv 11.0\Output Data Structure.xlsx"
    df = pd.read_excel(output_file)

    # Find the row with URL_ID
    row_index = df[df['URL_ID'] == int(input_file.split('.',1)[0].strip())].index[0]

    # Update the values in the corresponding row
    df.loc[row_index, 'POSITIVE SCORE'] = output_dict['positive_score']
    df.loc[row_index, 'NEGATIVE SCORE'] = - (output_dict['negative_score'] )
    df.loc[row_index, 'POLARITY SCORE'] = output_dict['polarity_score']
    df.loc[row_index, 'SUBJECTIVITY SCORE'] = output_dict['subjectivity_score']
    df.loc[row_index,'AVG SENTENCE LENGTH'] = output_dict['average_sentence_length']
    df.loc[row_index,'PERCENTAGE OF COMPLEX WORDS'] = output_dict['percentage_complex_words']
    df.loc[row_index,'FOG INDEX'] = output_dict['fog_index']
    df.loc[row_index,'AVG NUMBER OF WORDS PER SENTENCE'] = output_dict['average_words_per_sentence']
    df.loc[row_index,'COMPLEX WORD COUNT'] = output_dict['complex_word_count']
    df.loc[row_index,'WORD COUNT'] = output_dict['cleaned_word_count']
    df.loc[row_index,'SYLLABLE PER WORD'] = output_dict['average_syllables_per_word']
    df.loc[row_index,'PERSONAL PRONOUNS'] = output_dict['personal_pronoun_count']
    df.loc[row_index,'AVG WORD LENGTH'] = output_dict['average_word_length']
    
    # Save the updated DataFrame back to the Excel file
    df.to_excel(output_file, index=False)

def count_syllables(word):
    word = word.lower()
    if word.endswith(('es', 'ed')):
        word = word[:-2]  # Remove 'es' or 'ed'
    vowels = 'aeiou'
    count_syl = 0
    prev_char = None
    for char in word:
        if char in vowels and (prev_char is None or prev_char not in vowels):
            count_syl += 1
        prev_char = char
    return max(1, count_syl) 

def read_docx(file_path):
    doc = docx.Document(file_path)
    content = []
    for paragraph in doc.paragraphs:
        content.append(paragraph.text)
    return '\n'.join(content)

def sentiment_analysis(input_file, output_dict):
    
    #stop_file_path = r"C:\Users\navin\OneDrive\Desktop\WorkTree\sample_table\Entevyuv 11.0\StopWords\StopWords_Auditor.docx"
    #stop_words_text = read_docx(stop_file_path)
    #stop_words = set(stop_words_text.splitlines())

    # Collecting all the unique Stop words from all the files provided along with the probelm statement
    stop_words_directory = r"C:\Users\navin\OneDrive\Desktop\WorkTree\sample_table\Entevyuv 11.0\StopWords"
    stop_words = set()
    stop_words_files = ["StopWords_Auditor.docx","StopWords_Currencies.docx","StopWords_DatesandNumbers.docx", "StopWords_Generic.docx", "StopWords_GenericLong.docx", "StopWords_Geographic.docx", "StopWords_Names.docx"]
    for stop_words_file in stop_words_files:
        stop_file_path = os.path.join(stop_words_directory, stop_words_file)
        stop_words_text = read_docx(stop_file_path)
        stop_words.update(stop_words_text.splitlines())

    # Load positive and negative words dictionaries
    positive_words = set()
    negative_words = set()

    positive_file_path = r"C:\Users\navin\OneDrive\Desktop\WorkTree\sample_table\Entevyuv 11.0\MasterDictionary\positive-words.docx"
    positive_words_text = read_docx(positive_file_path)
    positive_words = set(positive_words_text.splitlines())

    negative_file_path = r"C:\Users\navin\OneDrive\Desktop\WorkTree\sample_table\Entevyuv 11.0\MasterDictionary\negative-words.docx"
    negative_words_text = read_docx(negative_file_path)
    negative_words = set(negative_words_text.splitlines())

    with open(input_file, "r", encoding="utf-8") as f:
        article_text = f.read()

    # Tokenize the article text
    tokens = word_tokenize(article_text)

    # Calculate derived variables
    positive_score = sum(1 for token in tokens if token in positive_words and token not in stop_words)
    negative_score = sum(1 for token in tokens if token in negative_words and token not in stop_words)

    # Updating the dict with POSITIVE SCORE and NEGATIVE SCORE
    output_dict['positive_score'] = positive_score
    output_dict['negative_score'] = negative_score


    polarity_score = (positive_score - negative_score) / ((positive_score + negative_score) + 0.000001)
    subjectivity_score = (positive_score + negative_score) / (len(tokens) + 0.000001)

    # Updating the dict with POLARITY SCORE and SUBJECTIVITY SCORE
    output_dict['polarity_score'] = polarity_score
    output_dict['subjectivity_score'] = subjectivity_score

    return output_dict

    
def readability_analysis(input_file, output_dict):
    with open(input_file, "r", encoding="utf-8") as f:
        article_text = f.read()
    # Split the text into sentences
    sentences = re.split(r'[.!?]', article_text)
    
    # Calculate AVG SENTENCE LENGTH and AVG NUMBER OF WORDS PER SENTENCE
    total_words = sum(len(sentence.split()) for sentence in sentences)
    total_sentences = len(sentences)
    average_sentence_length = average_words_per_sentence  = total_words / total_sentences

    # Updating the dict with AVG SENTENCE LENGTH and AVG NUMBER OF WORDS PER SENTENCE
    output_dict['average_sentence_length'] = average_sentence_length
    output_dict['average_words_per_sentence'] = average_words_per_sentence

    # Find all matches of PERSONAL PRONOUNS in the text
    personal_pronouns = ["I", "we", "my", "ours", "us"]
    pronoun_pattern = r'\b(?:' + '|'.join(re.escape(word) for word in personal_pronouns) + r')\b'
    matches = re.findall(pronoun_pattern, article_text, flags=re.IGNORECASE)
    personal_pronoun_count = len(matches)

    # Updating the dict with PERSONAL PRONOUNS in the text
    output_dict['personal_pronoun_count'] = personal_pronoun_count

    
    words = re.findall(r'\b\w+\b', article_text)

    # Calulating AVG WORD LENGTH
    total_characters = sum(len(word) for word in words)
    total_words = len(words)
    average_word_length = total_characters / total_words

    # Updating the dict with AVG WORD LENGTH
    output_dict['average_word_length'] = average_word_length

    # Calculating cleaned WORD COUNT
    # Remove punctuation and convert to lowercase
    words = [word.translate(str.maketrans('', '', string.punctuation)) for word in words]
    words = [word.lower() for word in words]
    # Remove stopwords
    stop_words = set(stopwords.words('english'))
    cleaned_words = [word for word in words if word not in stop_words]
    cleaned_word_count = len(cleaned_words)

    # Updating the dict with cleaned WORD COUNT
    output_dict['cleaned_word_count'] = cleaned_word_count

    # Calculate COMPLEX WORD COUNT (words with more than 2 syllables)
    complex_words = [word for word in cleaned_words if count_syllables(word) > 2]
    complex_word_count = len(complex_words)

    # Calculate PERCENTAGE OF COMPLEX WORDS (words with more than 2 syllables)
    percentage_complex_words = (complex_word_count / cleaned_word_count) * 100

    # Updating the dict with COMPLEX WORD COUNT and PERCENTAGE OF COMPLEX WORDS
    output_dict['complex_word_count'] = complex_word_count
    output_dict['percentage_complex_words'] = percentage_complex_words

    # Calcualting SYLLABLE PER WORD
    syllable_counts = [count_syllables(word) for word in cleaned_words]
    total_syllables  = sum(syllable_counts)
    average_syllables_per_word  = total_syllables /cleaned_word_count

    # Updating the dict with average SYLLABLE PER WORD
    output_dict['average_syllables_per_word'] = average_syllables_per_word
    
    # Calculate FOG INDEX
    fog_index = 0.4 * (average_sentence_length + percentage_complex_words)

    # Updating the dict with FOG INDEX
    output_dict['fog_index'] = fog_index

    return output_dict
