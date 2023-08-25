import extract_articles
import text_analysis
def main():
    # Extracts article's Title and Text from the web for all the URL and saves them to a respective file in text format
    filename_List  = extract_articles.scraping_Title_and_Text_to_File()

    # Performing Text Analysis and filling all the variables in the Outpur Data Structure file against corresponding URL_ID
    for filename in filename_List:

        output_dict = {}
        output_dict = text_analysis.sentiment_analysis(filename, output_dict)
        output_dict = text_analysis.readability_analysis(filename, output_dict)

        text_analysis.output_file_write(filename, output_dict)

if __name__ == "__main__":
    main()