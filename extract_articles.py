import pandas as pd
import requests
from bs4 import BeautifulSoup

# Load the Excel file
input_file = r"C:\Users\navin\OneDrive\Desktop\WorkTree\sample_table\Entevyuv 11.0\In.xlsx"
df = pd.read_excel(input_file)

# Loop through each row in the DataFrame
for index, row in df.iterrows():
    article_url = row['URL']
    url_id = row['URL_ID']
    
    # Download the HTML content using requests
    response = requests.get(article_url)
    html_content = response.content
    
    # Parse the HTML content with Beautiful Soup
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # Extract the article title and text
    title = soup.title.text.split('|', 1)[0].strip() if soup.title else "No Title"

    excluded_div_classes = [('div','td-module-meta-info'), ('p','tdm-descr')]
    for class_name in excluded_div_classes:
        excluded_divs = soup.find_all(class_name[0],class_name[1])
        for excluded_div in excluded_divs:
            excluded_div.decompose()

    paragraphs = soup.find_all('p')
    text = '\n'.join(paragraph.get_text() for paragraph in paragraphs)
    
    # Create a text file with URL_ID as its name
    output_filename = f"{url_id}.txt"
    with open(output_filename, "w", encoding="utf-8") as f:
        f.write(f"{title}\n")
        f.write(f"{text}\n")
    
    