import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook, Workbook
import datetime
import time
import traceback


# todo 将导入的网址关键词进行合并 同一个关键词只查询一次 然后在其中搜索对应的域名

def get_cur_time_filename():
    return time.strftime('%Y-%m-%d-%H-%M-%S', time.localtime())


PAGE = 10
TIMEOUT = 1

global_url = ''
global_text = ''
global_file_name = '关键词排名-%s.xlsx' % get_cur_time_filename()


def main():
    init_output()
    for i, (domain, keyword) in enumerate(get_input()):
        get_rank(i + 1, domain, keyword)


def get_input():
    # todo 没有对应文件的时候进行提示
    wb = load_workbook('input.xlsx')
    ws = wb.active
    result = []
    # todo 文档没有按照格式进行填写的时候进行提示
    for row in ws.iter_rows(min_row=2, values_only=True):
        result.append(row)
    return result


def get_rank(index, domain, keyword):
    print('开始抓取 第%s条 域名：%s 关键词：%s' % (index, domain, keyword))
    for i in range(PAGE):
        get_page(domain, keyword, i + 1)


def get_page(domain, keyword, page):
    global global_url, global_text
    print('开始第%d页' % page)
    # todo 这个网址这样请求就一定能够返回想要的结果吗？
    params = {
        'q': keyword,
        'from': 'smor',
        'safe': '1',
        'snum': '6',
        'by': 'next',
        'layout': 'html',
        'page': page
    }
    # todo 不同的UA有什么影响
    # todo 要怎么模拟真实的用户进行搜索 HEAD要怎么填写
    headers = {
        'User-Agent': 'Mozilla/5.0 (iPhone; CPU iPhone OS 11_0 like Mac OS X) AppleWebKit/604.1.38 (KHTML, like Gecko) Version/11.0 Mobile/15A372 Safari/604.1'
        # 'User-Agent': 'Mozilla/5.0 (Linux; Android 5.0; SM-G900P Build/LRX21T) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/75.0.3770.100 Mobile Safari/537.36'
    }
    # todo 网络出现问题的时候怎么办？
    # todo 抓取的内容有问题的时候怎么办？
    r = None
    soup = None
    while r is None or soup is None:
        try:
            r = requests.get('http://so.m.sm.cn/s', params=params, headers=headers)
        except (requests.exceptions.ConnectionError, requests.exceptions.ChunkedEncodingError):
            print('检查到网络断开，%s秒之后尝试重新抓取' % TIMEOUT)
            time.sleep(TIMEOUT)
            continue
        global_url = r.url
        global_text = r.text
        soup = BeautifulSoup(r.text, 'lxml')
        if soup.body is None:
            print('请求到的页面的内容为空，将再次进行请求')
            soup = None
    all_item = get_all_item(soup)
    result = []
    rank = 1
    for item in all_item:
        if is_domain_item(item, domain):
            result.append((
                domain,
                keyword,
                '%d-%d' % (page, rank),
                get_url(item, domain),
                get_title(item),
                datetime.datetime.now()
            ))
        rank += 1
    set_output(result)


def get_all_item(soup):
    all_item = []
    # todo 将来会不会出现别的div 是不是要考虑一下更可靠的筛选方式来确定那些是“排名结果”
    for child in soup.body.children:
        if child.name == 'div':
            all_item.append(child)
    return all_item


def is_domain_item(item, domain):
    for link in item.find_all('a'):
        url = link.get('href')
        if url and domain in url:
            return True


def get_url(item, domain):
    for link in item.find_all('a'):
        url = link.get('href')
        if url and domain in url:
            return url


def get_title(item):
    for span in item.find_all('span'):
        if span.text:
            return span.text


def init_output():
    global global_file_name
    wb = Workbook()
    ws = wb.active
    ws.append(('域名', '关键词', '搜索引擎', '排名', '真实地址', '标题', '查询时间'))
    wb.save(global_file_name)


def set_output(result):
    global global_file_name
    wb = load_workbook(filename=global_file_name)
    ws = wb.active
    for (domain, keyword, rank, url, title, date_time) in result:
        ws.append((domain, keyword, '神马', rank, url, title, date_time))
    wb.save(global_file_name)


def error_exit(title):
    raise RuntimeError('''%s

请求的url为：
%s

响应的text为：
%s
''' % (title, global_url, global_text))


try:
    main()
    input('查询结束，查询结果保存在 %s' % global_file_name)
except:
    filename = 'error-%s.log' % get_cur_time_filename()
    f = open(filename, 'w', encoding='utf-8')
    f.write('''%s

请求的URL为：
%s

返回的内容为：
%s
''' % (traceback.format_exc(), global_url, global_text))
    f.close()
    traceback.print_exc()
    input('请将最新的error.log文件发给技术人员')
