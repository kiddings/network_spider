import re
from bs4 import BeautifulSoup
import selenium.common
from selenium import webdriver
import os
import joblib
import sys
import time
import xlsxwriter


def get_page(url):
    chrome_driver = r'C:\Users\17626\Desktop\diver\chromedriver.exe'
    browser = webdriver.Chrome(executable_path=chrome_driver)
    print('waiting...')
    browser.get(url)
    print('success! {} is done'.format(url))
    return browser.page_source

def parse(page):
    soup = BeautifulSoup(page, 'lxml')
    title_path = 'body > div.pg-search.pg-wrapper > div.pg-no-rail.pg-wrapper > div > div.l-container > ' \
                 'div.cnn-search__right > div > div.cnn-search__results-list > div > div.cnn-search__result-contents ' \
                 '> h3 > a '
    article_path = 'body > div.pg-search.pg-wrapper > div.pg-no-rail.pg-wrapper > div > div.l-container > ' \
                   'div.cnn-search__right > div > div.cnn-search__results-list > div >' \
                   'div.cnn-search__result-contents > div.cnn-search__result-body '
    raw_titles = soup.select(title_path)
    raw_articles = soup.select(article_path)
    return raw_titles, raw_articles


def process(titles, articles, data):
    for i in range(len(titles)):
        title = titles[i].text.strip()
        article = articles[i].text.strip()
        data[title] = article
    return data


def get_urls(first_page, name):
    chrome_driver = r'C:\Users\17626\Desktop\diver\chromedriver.exe'
    browser = webdriver.Chrome(executable_path=chrome_driver)
    print('getting page information...')
    browser.get(first_page)
    print('done!')
    soup = BeautifulSoup(browser.page_source, 'lxml')
    page_path = 'body > div.pg-search.pg-wrapper > div.pg-no-rail.pg-wrapper > div > div.l-container > ' \
                'div.cnn-search__right > div > div.cnn-search__results-count '
    raw_page = soup.select(page_path)
    page_location = re.search(r'(.*?)of (.*?) for(.*?)', str(raw_page))
    if name in ['rand corporation', 'Belfer Center', 'Atlantic Council']:
        total_page = int(page_location.group(2))
        pages = int(total_page / 10)
        urls = []
        for i in range(pages):
            url = 'https://edition.cnn.com/search?q={}&size=10&page={}&from={}'.format(name, (i + 1), i * 10)
            urls.append(url)
        return urls

    elif name in ['Brookings Institution', 'Carnegie Endowment']:
        total_page = int(page_location.group(2))
        pages = int(total_page / 20)
        urls = []
        for i in range(pages + 1):
            url = 'https://edition.cnn.com/search?q={}&size=20&page={}&from={}'.format(name, (i + 1), i * 20)
            urls.append(url)
        return urls


def get_data(name):
    save_path = name + '_data.pkl'
    if not os.path.exists(save_path):
        data = {}
        joblib.dump(data, save_path)
    data = joblib.load(save_path)

    first_page = 'https://edition.cnn.com/search?q={}'.format(name)

    if not os.path.exists('{}_urls.pkl'.format(name)):
        urls = get_urls(first_page, name)
        joblib.dump(urls, '{}_urls.pkl'.format(name))

    else:
        urls = joblib.load('{}_urls.pkl'.format(name))
    done_list = joblib.load('done_list.pkl')
    for url in urls:
        if url not in done_list:
            page = get_page(url)
            raw_titles, raw_articles = parse(page)
            data = process(raw_titles, raw_articles, data)
            # 保存已经获取过的url
            done_list.append(url)
            joblib.dump(done_list, 'done_list.pkl')
            joblib.dump(data, save_path)
        else:
            print('{} is done'.format(url))
            continue


if __name__ == '__main__':
    if not os.path.exists('done_list.pkl'):
        Done_List = []
        joblib.dump(Done_List, 'done_list.pkl')

    think_tanks = ['rand corporation', 'Belfer Center', 'Atlantic Council', 'Brookings Institution',
                   'Carnegie Endowment']
    cnt = 0
    while cnt < 5:
        try:
            get_data(think_tanks[cnt])
            print('{} is done'.format(think_tanks[cnt]))
            cnt += 1

        # 超时重新运行
        except selenium.common.exceptions.TimeoutException:
            print('error:timeout!')
            Done_List = joblib.load('done_list.pkl')
            print('The following is the obtained page:')
            for item in Done_List:
                print(item)
            print('restart to run...')
            time.sleep(3)
            continue

    out_path = 'think_tanks(cnn).xlsx'
    workbook = xlsxwriter.Workbook(out_path)
    for think_tank in think_tanks:
        data = joblib.load('{}_data.pkl'.format(think_tank))

        worksheet = workbook.add_worksheet(think_tank)
        bold = workbook.add_format({'bold': True})  # 设置粗体
        worksheet.write('A1', 'title', bold)  # 标题
        worksheet.write('B1', 'article', bold)  # 新闻主题
        worksheet.set_column('B:B', 100)
        worksheet.set_column('A:A', 100)
        cnt = 1
        for (title, article) in data.items():
            worksheet.write(cnt, 0, title)
            worksheet.write(cnt, 1, article)
            cnt += 1
        print('{} is done'.format(think_tank))

    workbook.close()
    print('all done!')




