import requests
from bs4 import BeautifulSoup, Comment
from openpyxl import load_workbook, Workbook
import os
import time
import datetime
import traceback
from configparser import ConfigParser
from abc import ABCMeta, abstractmethod
from urllib.parse import urlparse, urljoin

PAGE = 10
TIMEOUT = 1


def get_cur_time_filename():
    return time.strftime('%Y-%m-%d-%H-%M-%S', time.localtime())


def format_cd_time(seconds):
    m, s = divmod(seconds, 60)
    h, m = divmod(m, 60)
    return "%d小时%02d分%02d秒" % (h, m, s)


def get_url_domain(url):
    li = urlparse(url).netloc.split('.')
    length = len(li)
    return '{}.{}'.format(li[length - 2], li[length - 1])


class MyError(RuntimeError):
    pass


class Spider(metaclass=ABCMeta):
    def __init__(self):
        self.url = ''
        self.text = ''
        self.result = []
        self.searched_keywords = []
        self.keyword_set = set()
        self.domain_set = set()
        self.main()

    def main(self):
        try:
            self.get_mode()
        except MyError as e:
            print(self.save_result())
            input(e)
        except KeyboardInterrupt:
            print(self.save_result())
            input('已经强行退出程序')
        except:
            print(self.save_result())
            filename = 'error-%s.log' % get_cur_time_filename()
            with open(filename, 'w', encoding='utf-8') as f:
                f.write('''%s

请求的URL为：
%s

返回的内容为：
%s
''' % (traceback.format_exc(), self.url, self.text))
            traceback.print_exc()
            input('请将最新的error.log文件发给技术人员')

    def get_mode(self):
        while True:
            run_mode = input('定时运行（输入1）还是马上运行（输入0）？')
            if run_mode == '0':
                self.search()
            elif run_mode == '1':
                self.start()
            else:
                print('输入了未知模式，请重新输入')

    def start(self):
        now = datetime.datetime.now()
        cfg = ConfigParser()
        cfg.read('config.ini')
        hour = int(cfg.get('config', 'hour'))
        start_time = datetime.datetime(now.year, now.month, now.day, hour)
        if start_time <= now:
            start_time = datetime.datetime.fromtimestamp(start_time.timestamp() + 24 * 60 * 60)
        wait_time = (start_time - now).total_seconds()
        print('下次查询时间为%s，将在%s后开始' % (start_time, format_cd_time(wait_time)))
        time.sleep(wait_time)
        self.search()
        self.start()

    def search(self):
        self.result = []
        self.searched_keywords = []
        self.keyword_set = set()
        self.domain_set = set()

        start_time = datetime.datetime.now()
        (keyword_set, domain_set) = self.get_input()
        self.keyword_set = keyword_set
        self.domain_set = domain_set
        print('总共要查找%s关键词，有%s个网站' % (len(keyword_set), len(domain_set)))
        for i, keyword in enumerate(keyword_set):
            self.get_rank(i + 1, keyword, domain_set)
        filename = self.save_result()
        print('查询结束，查询结果保存在 %s' % filename)
        end_time = datetime.datetime.now()
        print('本次查询用时%s' % format_cd_time((end_time - start_time).total_seconds()))

    def get_input(self):
        file_path = ''
        path = '.\\import'
        for file in os.listdir(path):
            file_path = os.path.join(path, file)
            break
        if file_path == '' or not file_path.endswith('.xlsx'):
            raise MyError('import目录之下没有发现xlsx文件')
        wb = load_workbook(file_path)
        ws = wb['网址']
        d = set()
        for (domain,) in ws.iter_rows(values_only=True):
            d.add(domain)
        ws = wb['关键词']
        k = set()
        for (keyword,) in ws.iter_rows(values_only=True):
            k.add(keyword)
        return k, d

    def get_rank(self, index, keyword, domain_set):
        print('开始抓取第%s个关键词：%s' % (index, keyword))
        for i in range(PAGE):
            self.get_page(i + 1, keyword, domain_set)
        self.searched_keywords.append(keyword)

    def get_page(self, page, keyword, domain_set):
        print('开始第%d页' % page)
        params = self.get_params(keyword, page)
        headers = {'User-Agent': self.user_agent}
        (r, soup) = self.safe_request(self.base_url, params=params, headers=headers)
        (ok, msg) = self.check_url(r.url)
        if not ok:
            raise (MyError(msg))
        all_item = self.get_all_item(soup)
        result = []
        rank = 1
        for item in all_item:
            url = self.get_url(item)
            if url is not None:
                domain = get_url_domain(url)
                if domain in domain_set:
                    result.append((
                        domain,
                        keyword,
                        '%d-%d' % (page, rank),
                        url,
                        self.get_title(item),
                        datetime.datetime.now()
                    ))
            rank += 1
        self.result += result

    @property
    @abstractmethod
    def user_agent(self):
        pass

    @property
    @abstractmethod
    def base_url(self):
        pass

    @property
    @abstractmethod
    def engine_name(self):
        pass

    @abstractmethod
    def get_params(self, keyword, page):
        pass

    def check_url(self, url):
        return True, ''

    @abstractmethod
    def get_all_item(self, soup):
        pass

    @abstractmethod
    def get_url(self, item):
        pass

    @abstractmethod
    def get_title(self, item):
        pass

    def save_result(self):
        file_name = '关键词排名-%s.xlsx' % get_cur_time_filename()
        wb = Workbook()
        ws = wb.active
        ws.append(('域名', '关键词', '搜索引擎', '排名', '真实地址', '标题', '查询时间'))
        for (domain, keyword, rank, url, title, date_time) in self.result:
            ws.append((domain, keyword, self.engine_name, rank, url, title, date_time))
        wb.save(file_name)
        self.result = []
        self.save_un_searched()
        return file_name

    def save_un_searched(self):
        file_name = '未查找关键词-%s.xlsx' % get_cur_time_filename()
        wb = Workbook()
        ws1 = wb.active
        ws1.title = '关键词'
        keywords = []
        for keyword in self.keyword_set:
            if keyword not in self.searched_keywords:
                keywords.append(keyword)
        for keyword in keywords:
            ws1.append((keyword,))
        ws2 = wb.create_sheet(title='网址')
        for domain in self.domain_set:
            ws2.append((domain,))
        wb.save(file_name)

    def safe_request(self, url, *, params=None, headers=None):
        r = None
        soup = None
        while r is None or soup is None:
            try:
                r = requests.get(url, params=params, headers=headers)
            except (requests.exceptions.ConnectionError, requests.exceptions.ChunkedEncodingError):
                print('检查到网络断开，%s秒之后尝试重新抓取' % TIMEOUT)
                time.sleep(TIMEOUT)
                continue
            self.url = r.url
            self.text = r.text
            soup = BeautifulSoup(r.text, 'lxml')
            if soup.body is None:
                raise MyError('请求到的页面的内容为空，为防止IP被封禁，已退出程序')
        return r, soup


class SMSpider(Spider):
    @property
    def user_agent(self):
        return 'Mozilla/5.0 (iPhone; CPU iPhone OS 11_0 like Mac OS X) AppleWebKit/604.1.38 (KHTML, like Gecko) Version/11.0 Mobile/15A372 Safari/604.1'

    @property
    def base_url(self):
        return 'http://m.sm.cn/s'

    @property
    def engine_name(self):
        return '神马'

    def get_params(self, keyword, page):
        return {
            'q': keyword,
            'page': page,
            'by': 'next',
            'from': 'smor',
            'tomode': 'center',
            'safe': '1',
        }

    def get_all_item(self, soup):
        return soup.find_all('div', class_='ali_row')

    def get_url(self, item):
        link = item.find('a')
        return link.get('href')

    def get_title(self, item):
        return ''.join(item.find('a').findAll(text=True))


class SogouPCSpider(Spider):
    @property
    def user_agent(self):
        return 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/75.0.3770.100 Safari/537.36'

    @property
    def base_url(self):
        return 'http://www.sogou.com/web'

    @property
    def engine_name(self):
        return '搜狗PC'

    def get_params(self, keyword, page):
        return {
            'query': keyword,
            'page': page,
        }

    def check_url(self, url):
        if url.startswith('http://www.sogou.com/antispider'):
            return False, '该IP已经被搜狗引擎封禁'
        return True, ''

    def get_all_item(self, soup):
        return soup.find('div', class_='results').find_all('div', recursive=False)

    def get_url(self, item):
        link = item.find('a')
        return link.get('href')

    def get_title(self, item):
        return ''.join(item.find('a').findAll(text=lambda text: not isinstance(text, Comment)))


class SogouMobileSpider(Spider):
    @property
    def user_agent(self):
        return 'Mozilla/5.0 (Linux; Android 5.0; SM-G900P Build/LRX21T) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/75.0.3770.100 Mobile Safari/537.36'

    @property
    def base_url(self):
        return 'http://wap.sogou.com/web/search/'

    @property
    def engine_name(self):
        return '搜狗MOBILE'

    def get_params(self, keyword, page):
        return {
            'keyword': keyword,
            'p': page,
        }

    def get_all_item(self, soup):
        return soup.find_all('a', class_='resultLink')

    def get_url(self, item):
        url = item.get('href')
        if url.startswith('javascript'):
            return None
        elif not url.startswith('http'):
            url = urljoin(self.url, url)
            (r, sub_soup) = self.safe_request(url)
            if r.url.startswith(self.base_url):
                btn = sub_soup.find('div', class_='btn')
                link = btn.find('a')
                return link.get('href')
            else:
                return r.url
        else:
            return url

    def get_title(self, item):
        return ''.join(item.findAll(text=lambda text: not isinstance(text, Comment)))


if __name__ == '__main__':
    spider_list = [(SMSpider, '神马'), (SogouPCSpider, '搜狗PC'), (SogouMobileSpider, '搜狗MOBILE')]
    spider_index = input('''要运行哪个Spider？
%s
''' % '\n'.join(['%s 请输入：%s' % (name, i) for (i, (_, name)) in enumerate(spider_list)]))
    (class_, name) = spider_list[int(spider_index)]
    os.system('title %s' % name)
    class_()
