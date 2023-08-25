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
    output_file = r"C:\Users\navin\OneDrive\Desktop\WorkTree\sample_table\Entevyuv 11.0\Out.xlsx"
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
    count = 0
    vowels = "aeiouy"
    if word[0] in vowels:
        count += 1
    for index in range(1, len(word)):
        if word[index] in vowels and word[index - 1] not in vowels:
            count += 1
    if word.endswith("e"):
        count -= 1
    if count == 0:
        count += 1
    return count

def read_docx(file_path):
    doc = docx.Document(file_path)
    content = []
    for paragraph in doc.paragraphs:
        content.append(paragraph.text)
    return '\n'.join(content)

def sentiment_analysis(input_file):

    # Load stop words list

    # stop_words_folder = r"C:\Users\navin\OneDrive\Desktop\WorkTree\sample_table\Entevyuv 11.0\StopWords"
    # stop_words = set()

    # # Loop through each file in the stop_words folder
    # for filename in os.listdir(stop_words_folder):
    #     if filename.endswith(".docx"):
    #         file_path = os.path.join(stop_words_folder,filename)

    #     with open(file_path, "r",encoding="utf-8") as f:
    #         stop_words.update(f.read().splitlines())
    
    stop_file_path = r"C:\Users\navin\OneDrive\Desktop\WorkTree\sample_table\Entevyuv 11.0\StopWords\StopWords_Auditor.docx"
    stop_words_text = read_docx(stop_file_path)
    stop_words = set(stop_words_text.splitlines())

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

    polarity_score = (positive_score - negative_score) / ((positive_score + negative_score) + 0.000001)
    subjectivity_score = (positive_score + negative_score) / (len(tokens) + 0.000001)

    return positive_score, negative_score, polarity_score, subjectivity_score

    
def readability_analysis(input_file):
    with open(input_file, "r", encoding="utf-8") as f:
        article_text = f.read()
    # Split the text into sentences
    sentences = re.split(r'[.!?]', article_text)
    
    # Calculate average sentence length and average words per sentence
    total_words = sum(len(sentence.split()) for sentence in sentences)
    total_sentences = len(sentences)
    average_sentence_length = average_words_per_sentence  = total_words / total_sentences

    # Find all matches of personal pronouns in the text
    personal_pronouns = ["I", "we", "my", "ours", "us"]
    # Create a regular expression pattern for the personal pronouns
    pronoun_pattern = r'\b(?:' + '|'.join(re.escape(word) for word in personal_pronouns) + r')\b'
    matches = re.findall(pronoun_pattern, article_text, flags=re.IGNORECASE)
    personal_pronoun_count = len(matches)

    
    # Calculate percentage of complex words (words with more than 2 syllables)
    words = re.findall(r'\b\w+\b', article_text)

    # Calulating AVG WORD LENGTH
    total_characters = sum(len(word) for word in words)
    total_words = len(words)
    average_word_length = total_characters / total_words


    # Calculating cleaned word count
    # Remove punctuation and convert to lowercase
    words = [word.translate(str.maketrans('', '', string.punctuation)) for word in words]
    words = [word.lower() for word in words]
    
    # Remove stopwords
    stop_words = set(stopwords.words('english'))
    cleaned_words = [word for word in words if word not in stop_words]
    
    cleaned_word_count = len(cleaned_words)

    complex_words = [word for word in cleaned_words if count_syllables(word) > 2]
    complex_word_count = len(complex_words)
    percentage_complex_words = (complex_word_count / cleaned_word_count) * 100

    syllable_counts = [count_syllables(word) for word in cleaned_words]
    total_syllables  = sum(syllable_counts)
    average_syllables_per_word  = total_syllables /cleaned_word_count
    
    # Calculate Gunning Fog Index
    fog_index = 0.4 * (average_sentence_length + percentage_complex_words)

    return average_sentence_length, percentage_complex_words, fog_index, average_words_per_sentence, complex_word_count, cleaned_word_count, average_syllables_per_word, personal_pronoun_count, average_word_length


output_dict = {}
output_dict['positive_score'], output_dict['negative_score'], output_dict['polarity_score'], output_dict['subjectivity_score'] = sentiment_analysis('37.txt')
output_dict['average_sentence_length'], output_dict['percentage_complex_words'], output_dict['fog_index'], output_dict['average_words_per_sentence'], output_dict['complex_word_count'], output_dict['cleaned_word_count'], output_dict['average_syllables_per_word'], output_dict['personal_pronoun_count'], output_dict['average_word_length'] = readability_analysis('37.txt')
#print(output_dict)
output_file_write('37.txt', output_dict)