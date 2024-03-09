import requests as req 
from openpyxl import Workbook

wb = Workbook()
ws = wb.active

title = ['書名', '作者', '出版社', '價格']
ws.append(title)

header = {
    'user-agent' : 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36'    
}

for index in range(3):  
    index += 1
    url = 'https://athena.eslite.com/api/v1/best_sellers/online/month?l1=3&page='
    url = url + str(index) + '&per_page=20'
    print(url)
    r = req.get(url, headers=header)
    print(r)
    
    root_json = r.json() 

    for data in root_json['products']:  
        book = []
        book.append(data['name'])  
        book.append(data['author'])  
        book.append(data['manufacturer'])
        book.append(data['retail_price'])
        
        ws.append(book)
           
wb.save('books.xlsx')
