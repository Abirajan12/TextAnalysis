# Web Scraping and Text Analysis

The primary goal of this assignment is to develop a systematic process for retrieving textual data from a specified URL, which contains articles or written content. Once this data is collected, the next step is to conduct thorough text analysis to extract valuable insights and information.

## Table of Contents

    1. Introduction
    2. Installation
    3. Web scraping
    4. Text Analysis
    5. Usage

## 1. Introduction

    This project extracts articles' titles and texts from web URLs specified in an Excel file. It then performs sentiment analysis and readability analysis on the extracted content. The analysis results are updated in an Excel output data structure file for further analysis and reference.

## 2. Installation

    Web scraping
    ------------
    import requests
    from bs4 import BeautifulSoup

    Text Analysis
    --------------
    import pandas as pd
    import docx
    import re
    import string

    NLP
    ----
    from nltk.tokenize import word_tokenize
    from nltk.corpus import stopwords
    nltk.download('stopwords') --once
    nltk.download('punkt') --once

## 3. Web scraping

    From the input file 'Input.xslx', for each url the article 'Title' and the 'Text'  has been scraped from the corresponding web page using
        scraping_Title_and_Text_to_File()

    and content got saved in a file corresponding to the URL_ID

## 4. Text Analysis

    For all the files of the extracted articles, Sentiment Analysis was done using
        * sentiment_analysis(filename, output_dict)

    wherein which the following varaibles were calculated.

        1. POSITIVE SCORE
        2. NEGATIVE SCORE
        3. POLARITY SCORE
        4. SUBJECTIVITY SCORE
    Text Analysis were done using
        * readability_analysis(filename, output_dict)
    wherein which the following varaibles were calculated.

        5. AVG SENTENCE LENGTH
        6. PERCENTAGE OF COMPLEX WORDS
        7. FOG INDEX
        8. AVG NUMBER OF WORDS PER SENTENCE
        9. COMPLEX WORD COUNT
        10. WORD COUNT
        11. SYLLABLE PER WORD
        12. PERSONAL PRONOUNS
        13. AVG WORD LENGTH

    As mentioned in the problem statement, those variables were later updated in the given 'Output Data Structure.xslx' file against the corresponding URL_ID using
        * output_file_write(filename, output_dict)

## 5. Usage

    Web Scraping and Text Analysis gives a hands-on experience on scraping realtime data and understanding the Natural Language Proceesing elements.
