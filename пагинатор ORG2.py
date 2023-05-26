from bs4 import BeautifulSoup
import aiohttp
import asyncio
from datetime import datetime
import requests

start_time = datetime.now()
url='https://www.ORG2/catalog/'

async def get_page_data():
    catalogical_url = []
    count = 1
    total_pages = []
    for i in range(200, 600):

        url_test = f'https://www.ORG2/catalog/{i}'
        response = requests.get(url_test)
        try:
            assert response.status_code == 200
            catalogical_url.append(url_test)
            print(count, ' каталог найден')
            count +=1

        except Exception:
            continue
    with open(f'ссылки каталога {url}', 'w') as file:
        for line in catalogical_url:
            file.write(line + '\n')
    print(catalogical_url)
    for pages in catalogical_url:
        for i in range(200, 600):
            url_test1 = pages +'/'+ f'{i}'+'/'
            resp = requests.get(url_test1)
            try:
                assert resp.status_code == 200
                catalogical_url.append(url_test1)

                print( ' каталог найден')

            except Exception:
                continue
        with open(f'ссылки  {start_time}', 'w') as file:

            file.write(url_test1 + '\n')

    print(total_pages)
    return total_pages

async def main():
    products_url_list = await get_page_data()
    print(products_url_list)
asyncio.run(main())

print('Time taken:', str(datetime.now() - start_time).split('.')[0])
