from bs4 import BeautifulSoup
import requests, openpyxl

excel = openpyxl.Workbook()
sheet = excel.active

sheet.title = 'Top News'
sheet.append(["Title", "Desc"])

try:
    source = requests.get("https://timesofindia.indiatimes.com/news")
    source.raise_for_status()

    soup = BeautifulSoup(source.text, 'html.parser')
    # print(soup.prettify())

    new = soup.find('div', class_='listing4 clearfix').findAll('li')

    for news in new:
        head = news.find('span', class_='w_tle').get_text()

        desc = news.find('span', class_='w_desc').get_text()

        print(head, desc)
        sheet.append([head,desc])

except Exception as e:
    print(e)

excel.save("Latest News.xlsx")