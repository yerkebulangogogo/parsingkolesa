import openpyxl
from bs4 import BeautifulSoup
import requests
import sys

# Here we set URL and Header, and took host adress
URL = 'https://kolesa.kz/cars/tesla/'
HEADER = {
    'User-Agent': 'PUT_HERE_USER_AGENT'
}
HOST = 'https://kolesa.kz'

#Here take html data
def get_html(URL, params=None):
    data = requests.get(URL, headers=HEADER, params=params)
    print('Parsing started: 2/4')
    return data

#Take each cars address (links), and write to list cars
def get_data(html):
    soup = BeautifulSoup(html, 'html.parser')
    items = soup.find_all('div', class_ = 'a-elem')
    cars = []
    for item in items:
        cars.append({
                        'title':HOST + item.find('a', class_="ddl_product_link").get('href')

        })
    return cars

#From the list cars which saved links to each avto, take parameters which we selected and save to
#list datas
def get_data_2(cars, params=None):
    datas = []
    for i in range(len(cars)):
        url1 = cars[i]
        url2 = url1['title']
        data = requests.get(url2, headers= HEADER,params=params)
        data = data.text

        soup = BeautifulSoup(data, 'html.parser')
        items = soup.find_all('div', class_ = 'offer')

        for item in items:

            datas.append({
                'brand': item.find('h1', class_= 'offer__title').get_text(),
                'price': item.find('div', class_ = 'offer__price').get_text(strip=True).replace('\xa0', ''),
                'city' : item.find('dd', class_='value').get_text(),

            })
    print('Parsing started: 3/4')

    return datas
    write_to_excel(cars, datas)

#Here we took of links length
def get_page_range(html):
    soup = BeautifulSoup(html ,'html.parser')
    pages = soup.find('div', class_ = 'pager').get_text(strip=True).replace('предыдущаяследующаяCtrl →','')
    pages = len(pages)
    if pages == 0:
        return 1
    else:
        return int(pages)

#Write to excel information, wich we are collected
def write_to_excel(cars, datas):

    begin = openpyxl.Workbook()
    sheet = begin.active

    sheet['A1']="NAME"
    sheet['B1']="PRICE"
    sheet['C1']="CITY"
    sheet['D1']="LINK"

    count = 2
    for j in range(len(cars)):
        for i in cars[j]:
            sheet[count][3].value = i['title']
            count += 1

    count = 2
    for i in datas:
        sheet[count][0].value = i['brand']
        sheet[count][1].value = i['price']
        sheet[count][2].value = i['city']
        count += 1

    begin.save('Tesla_dudes.xlsx')
    begin.close()
    print('Parsing started: 4/4')
    print('Finished: 4/4')

#Main method
def pars():
    print('Parsing started: 1/4')
    html = get_html(URL)
    cars = []
    links =[]

    if html.status_code == 200:
        # get_data(html.text)
        page_range=get_page_range(html.text)
        for page in range(1, page_range+1):
            print('Parsing page: {} of {}'.format(page, page_range))
            html = get_html(URL, params={'page': page})
            links.append(get_data(html.text))
            cars.extend(get_data_2(get_data(html.text)))
        print(cars)
        print('Total cars length: '+str(len(cars)))
        write_to_excel(links,cars)
        print("Writed to EXCEL successful!")
    else:
        print('Error!', sys.exc_info()[1])

pars()
