import requests
from bs4 import BeautifulSoup
import lxml
import xlsxwriter

import fake_useragent
from fake_useragent import UserAgent

from datetime import datetime
from calendar import monthrange

# import asyncio
# import aiohttp
# import sys

ua = UserAgent()
ua = ua.random

workbook_empty = xlsxwriter.Workbook('test1_empty.xlsx')
worksheet_empty = workbook_empty.add_worksheet()

# Add a bold format to use to highlight cells.
bold = workbook_empty.add_format({'bold': True, 'font_color': 'red'})
bold.set_align('center')

worksheet_empty.set_column('A:A', 100)
worksheet_empty.write('A1', ':: EMPTY ::', bold)
url_empty = ''

workbook = xlsxwriter.Workbook('test1.xlsx')
worksheet = workbook.add_worksheet()

bold_1 = workbook.add_format({'bold': True, 'font_color': 'black'})
bold_1.set_align('center')

bold_2 = workbook.add_format({'bold': True, 'font_color': 'blue'})
bold_2.set_align('center')

data_format1 = workbook.add_format({'bg_color': '#FFC7CE'})
#set_bg_color()

# Format the first column
worksheet.set_column('A:A', 20)
worksheet.set_column('B:B', 16)
worksheet.set_column('C:C', 15)
worksheet.set_column('D:G', 20)

worksheet.write('A1', 'Company', bold_1)
worksheet.write('B1', 'Symbol_Dividend', bold_1)
worksheet.write('C1', 'Dividend', bold_1)
worksheet.write('D1', 'Anouncement_Date', bold_1)
worksheet.write('E1', 'Record_Date', bold_1)
worksheet.write('F1', 'Ex-Date', bold_1)
worksheet.write('G1', 'Pay_Date', bold_1)


headers = {
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9",
    "User-Agent": f'{ua}'  # "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML,
    # like Gecko) Chrome/96.0.4664.45 Safari/537.36"
}

def month_name(num, lang):
    en = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september',
          'october', 'november', 'december']
    ru = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август', 'сентябрь',
          'октябрь', 'ноябрь', 'декабрь']
    if lang == 'en':
        return en[num - 1]
    else:
        return ru[num - 1]

def get_articles_urls(url):
    with requests.Session() as session:
        #s = requests.Session()
        response = session.get(url=url, headers=headers)

    # # запись СПАРСЕНОЙ инфы в ХТМЛ-файл
    # with open('index.html', 'w', encoding='utf-8') as file:
    #     file.write(response.text)



    # поиск ЗНАЧЕНИЯ последней страницы ПАГИНАЦИИ

    soup = BeautifulSoup(response.text, 'lxml')

    # Пагинация заключена в теге: <span class='pages'>Страница 1 из 5</span>
    # Мне нужно достать "5".
    #pagination_count = int(soup.find('div', class_='pagination').text[-1])
    pagination = soup.find('div', class_='pagination').find_all('a', class_='with-border')#.text#[-1]
    pagination_list = []
    for a in pagination:
        pagination_list.append(a)

    pagination_count = pagination_list[-1].text.strip()
    print(pagination_count)

    # КУСОК!!! ссылки на страницы с ЛОТАМИ аукциона
    # url_page = 'https://www.aquaforum.ua/auction/ending-soonest-auctions/page/'
    # url = f'https://www.sothebys.com/en/auctions/2015/magnificent-jewels-n09495.html?p={page}/'

    url_page = f'https://www.sothebys.com/en/auctions/2015/magnificent-jewels-n09495.html?p='

    # url_page = 'index.html'

    # создаю СПИСОК для хранения ссылок на ЛОТЫ
    articles_urls_list = []
    au_all = []

    with requests.Session() as session:

        url_page_one = f'https://www.sothebys.com/en/auctions/2015/magnificent-jewels-n09495.html'
        response = session.get(url=f'{url_page_one}/', headers=headers)
        soup = BeautifulSoup(response.text, 'lxml')

        auction_info_title = soup.find('div', class_='AuctionsModule-auction-title').text.strip()
        auction_info = soup.find('div', class_='AuctionsModule-auction-info')
        www = auction_info.find_all('div')

        aaa = []
        for i in www:
            aaa.append(i.text.split(":"))

        str1 = ''
        for ele in aaa[0]:
            str1 += ele
        tmp = (str1.replace(u"\u2022", ":"))

        ar = []
        auction_info_data = tmp.split(':')[0].strip()
        auction_info_town = tmp.split(':')[1].strip()
        ar.append(
            {
                "auction_info_title": auction_info_title,
                "auction_info_data": auction_info_data,
                "auction_info_town": auction_info_town
            }
        )

        # with open('glob.json', 'w', encoding='utf-8') as e:
        #     json.dump(ar, e, indent=4, ensure_ascii=False)

        for page in range(1, 2):
        #for page in range(1, pagination_count + 1):
            # в "page" будут НОМЕРА страниц ПАГИНАЦИИ (в моём случае это: 1,2,3,4 и 5)
            url = f'{url_page}{page}/'
            print(url)

            response = session.get(url=f'{url_page}{page}', headers=headers)
            soup = BeautifulSoup(response.text, 'lxml')

            articles_urls = soup.find_all('li', class_='AuctionsModule-results-item')

            for au in articles_urls:
                art_url = au.find('div', class_='title expandToThreeLines').find('a').get('href')
                articles_urls_list.append(art_url)

            # ЗАЩИТА от БАНА!!!
            # time.sleep(randrange(1, 3))
            print(f'Обработал {page} / {pagination_count}')

        # запись ссылок из СПИСКА в файл
        with open('articles_urls1.txt', 'w', encoding='utf-8') as file:
            for url in articles_urls_list:
                file.write(f'{url}\n')
    return 'Ссылки собраны'
##################################################################
#
# Сбор ДАННЫХ о лотах




def get_data():

    # читаю ССЫЛКИ из ранее созданного файла
    # !!! ОБРЕЗАЮ СИМВОЛ ПЕРЕНОСА СТРОКИ !!!
    # with open(file_path) as file:
    #     url_list = [line.strip() for line in file.readlines()]

    # # СЧЁТЧИК количества ЛОТОВ(ссылок)
    # url_count = len(url_list)

    s = requests.Session()

    # список для СУММАРНОЙ информации
    result_data = []

    #www_d = 2
    print(f"start...")
    i = 3
    empty = []

    www_m = 1
    www_y = 1996



    # current_year =  datetime.now().year
    days = monthrange(www_y, www_m)[1]

    worksheet.write('A2', month_name(www_m, 'en'), data_format1)
    worksheet.write('B2:G2', '', data_format1)

    for www_d in range(days + 1):
    #for www_d in range(3):
        if www_d > 0:
            print(f'Day ---> {www_d}')

            url = f'https://eresearch.fidelity.com/eresearch/conferenceCalls.jhtml?tab=dividends&begindate={www_m}/{www_d}/{www_y}'

            headers = {
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9",
                "User-Agent": f'{ua}'  # "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML,
                # like Gecko) Chrome/96.0.4664.45 Safari/537.36"
            }

            with requests.Session() as session:
                # ЗАЩИТА от БАНА!!!
                #time.sleep(randrange(1, 3))

                response = session.get(url=url, headers=headers)
                #print(response.status_code)

                soup = BeautifulSoup(response.text, 'lxml')

                lot_group_all = soup.find_all('div', class_='LotPage-breadcrumbs-breadcrumb')




                lot_items = soup.find('table', class_='datatable-component events-calender-table-three').find('tbody').find_all(
                    'tr')

                #print(f'LEN: {len(lot_items)}')

                if len(lot_items) == 1:
                    print(f'LEN: {len(lot_items)}')
                    empty.append(url)
                else:
                    for lot in lot_items:

                        company = lot.find('strong').text
                        # try:
                        #     symbol = lot.find('td', class_='lft-rt-border center blue-links').find('a').text.strip()
                        # except:
                        #     try:
                        #         symbol = lot.find('td', class_='lft-rt-border center blue-links').text.strip()
                        #     except:
                        #         pass

                        data_all = lot.find_all('td')
                        ddd = []

                        for j in data_all:
                            ddd.append(j.text.strip())

                        # print(f'str --->{i}')
                        # print(ddd)

                        worksheet.write(f'A{i}', company)
                        worksheet.write(f'B{i}', ddd[0], bold_2)
                        worksheet.write(f'C{i}', ddd[1])
                        worksheet.write(f'D{i}', ddd[2])
                        worksheet.write(f'E{i}', ddd[3])
                        worksheet.write(f'F{i}', ddd[4])
                        worksheet.write(f'G{i}', ddd[5])
                        i += 1

    workbook.close()

    # EMPTY
    row = 2
    for i in empty:
        worksheet_empty.write_url(f'A{row}', i, string=f'{i}')
        row += 1
    workbook_empty.close()



##################################################################

def main():
    # print(get_articles_urls(url="https://www.sothebys.com/en/auctions/2015/magnificent-jewels-n09495.html?p=/"))
    #get_data('articles_urls.txt')
    get_data()

if __name__ == "__main__":
    main()

