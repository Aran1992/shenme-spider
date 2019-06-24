import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook, Workbook
import datetime
import time
import traceback

PAGE = 10


def main():
    output = []
    for (domain, keyword) in get_input():
        result = get_rank(domain, keyword)
        output += result
    set_output(output)


def get_input():
    # todo 没有对应文件的时候进行提示
    wb = load_workbook('input.xlsx')
    ws = wb.active
    result = []
    # todo 文档没有按照格式进行填写的时候进行提示
    for row in ws.iter_rows(min_row=2, values_only=True):
        result.append(row)
    return result


def get_rank(domain, keyword):
    print('开始抓取 域名：%s 关键词：%s' % (domain, keyword))
    result = []
    for i in range(PAGE):
        result += get_page(domain, keyword, i + 1)
    return result


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
    }
    # todo 网络出现问题的时候怎么办？
    # todo 抓取的内容有问题的时候怎么办？
    r = requests.get('http://so.m.sm.cn/s', params=params, headers=headers)
    global_url = r.url
    global_text = r.text
    soup = BeautifulSoup(r.text, 'lxml')
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
    return result


def get_all_item(soup):
    if soup.body:
        all_item = []
        # todo 将来会不会出现别的div 是不是要考虑一下更可靠的筛选方式来确定那些是“排名结果”
        for child in soup.body.children:
            if child.name == 'div':
                all_item.append(child)
        return all_item
    else:
        raise RuntimeError('出现了soup.body为None的BUG')


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


def set_output(output):
    wb = Workbook()
    ws = wb.active
    ws.append(('域名', '关键词', '搜索引擎', '排名', '真实地址', '标题', '查询时间'))
    for (domain, keyword, rank, url, title, date_time) in output:
        ws.append((domain, keyword, '神马', rank, url, title, date_time))
    file_name = '关键词排名-%s.xlsx' % get_cur_time_filename()
    wb.save(file_name)
    input('查询结束，查询结果保存在 %s' % file_name)


def get_cur_time_filename():
    return time.strftime("%Y-%m-%d-%H-%M-%S", time.localtime())


def error_exit(title):
    raise RuntimeError('''%s

请求的url为：
%s

响应的text为：
%s
''' % (title, global_url, global_text))


try:
    main()
except:
    filename = 'error-%s.log' % get_cur_time_filename()
    f = open(filename, 'w', encoding="utf-8")
    f.write('''%s
    
请求的URL为：
%s

返回的内容为：
%s
''' % (traceback.format_exc(), global_url, global_text))
    f.close()
    traceback.print_exc()
    input('请将最新的error.log文件发给技术人员')
