import extract_articles
import text_analysis
def main():
    filename_List  = extract_articles.scraping_Title_and_Text_to_File()
    print(filename_List)
    
    output_dict = {}
    output_dict = text_analysis.sentiment_analysis('37.txt',output_dict)
    output_dict = text_analysis.readability_analysis('37.txt',output_dict)

    text_analysis.output_file_write('37.txt', output_dict)

if __name__ == "__main__":
    main()