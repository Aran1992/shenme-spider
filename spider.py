import ast
import os
import re
import time
import traceback
from abc import ABCMeta, abstractmethod
from configparser import ConfigParser
from datetime import datetime
from urllib.parse import urlparse, urljoin

import requests
from bs4 import BeautifulSoup, Comment
from openpyxl import load_workbook, Workbook

# import this seems unused
# but it's to prevent 'bs4.FeatureNotFound: Couldn't find a tree builder with the features you requested: lxml.'
import lxml

cfg = ConfigParser()
cfg.read('config.ini')
PAGE = int(cfg.get('config', 'page_count'))


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


class SpiderRuler(metaclass=ABCMeta):
    def __init__(self, spider):
        self.spider = spider

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

    @property
    @abstractmethod
    def request_interval_time(self):
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

    @abstractmethod
    def has_next_page(self, soup):
        pass

    def is_forbid(self, soup):
        return False


class SMRuler(SpiderRuler):
    def __init__(self, spider):
        SpiderRuler.__init__(self, spider)
        cfg = ConfigParser()
        cfg.read('config.ini')
        self.__request_interval_time = float(cfg.get('config', 'sm_request_interval_time'))

    @property
    def request_interval_time(self):
        return self.__request_interval_time

    @property
    def user_agent(self):
        return 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/75.0.3770.100 Safari/537.36'

    @property
    def base_url(self):
        return 'https://m.sm.cn/s'

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
        return link and link.get('href')

    def get_title(self, item):
        return ''.join(item.find('a').findAll(text=True))

    def has_next_page(self, soup):
        return soup.find('a', class_='next')


class SogouPCRuler(SpiderRuler):
    def __init__(self, spider):
        SpiderRuler.__init__(self, spider)
        cfg = ConfigParser()
        cfg.read('config.ini')
        self.__request_interval_time = float(cfg.get('config', 'sgpc_request_interval_time'))

    @property
    def request_interval_time(self):
        return self.__request_interval_time

    @property
    def user_agent(self):
        return 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) ' \
               'Chrome/75.0.3770.100 Safari/537.36'

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
        return link and link.get('href')

    def get_title(self, item):
        return ''.join(item.find('a').findAll(text=lambda text: not isinstance(text, Comment)))

    def has_next_page(self, soup):
        return soup.find(id='sogou_next')

    def is_forbid(self, soup):
        return soup.find('div', class_='results') is not None


class SogouMobileRuler(SpiderRuler):
    def __init__(self, spider):
        SpiderRuler.__init__(self, spider)
        cfg = ConfigParser()
        cfg.read('config.ini')
        self.__request_interval_time = float(cfg.get('config', 'sgmobile_request_interval_time'))

    @property
    def request_interval_time(self):
        return self.__request_interval_time

    @property
    def user_agent(self):
        return 'Mozilla/5.0 (Linux; Android 5.0; SM-G900P Build/LRX21T) AppleWebKit/537.36 (KHTML, like Gecko) ' \
               'Chrome/75.0.3770.100 Mobile Safari/537.36'

    @property
    def base_url(self):
        return 'http://wap.sogou.com/web/search/ajax_query.jsp'

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
        elif url.startswith('http'):
            return url
        else:
            url = urljoin(self.spider.url, url)
            (r, sub_soup, _) = self.spider.safe_request(url)
            if r.url.startswith('http://wap.sogou.com/transcoding/sweb') \
                    or r.url.startswith('http://m.sogou.com/transcoding/sweb') \
                    or r.url.startswith('http://wap.sogou.com/web/search/'):
                btn = sub_soup.find('div', class_='btn')
                if btn:
                    link = btn.find('a')
                else:
                    # 个别情况下 会发生页面里面没有class为btn的div的情况
                    link = sub_soup.find('a')
                if link:
                    return link.get('href')
            else:
                return r.url

    def get_title(self, item):
        return ''.join(item.findAll(text=lambda text: not isinstance(text, Comment)))

    def has_next_page(self, soup):
        li = soup.find('p').find(text=True).strip().split(',')
        return int(li[0]) > int(li[1])


class BaiduPCRuler(SpiderRuler):
    def __init__(self, spider):
        SpiderRuler.__init__(self, spider)
        cfg = ConfigParser()
        cfg.read('config.ini')
        self.__request_interval_time = float(cfg.get('config', 'bdpc_request_interval_time'))

    @property
    def request_interval_time(self):
        return self.__request_interval_time

    @property
    def user_agent(self):
        return 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) ' \
               'Chrome/75.0.3770.100 Safari/537.36'

    @property
    def base_url(self):
        return 'https://www.baidu.com/s'

    @property
    def engine_name(self):
        return '百度PC'

    def get_params(self, keyword, page):
        return {
            'wd': keyword,
            'pn': (page - 1) * 10,
        }

    def get_all_item(self, soup):
        div_root = soup.find('div', id='content_left')
        if div_root:
            return div_root.find_all('div', recursive=False, id=lambda id_: id_ != 'rs_top_new')
        else:
            return []

    def get_url(self, item):
        link = item.find('a')
        if link:
            url = link.get('href')
            if url.startswith('javascript'):
                return None
            elif url.startswith('http://www.baidu.com/link?'):
                return self.spider.get_real_url(url)
            else:
                return url
        else:
            return None

    def get_title(self, item):
        return ''.join(item.find('a').findAll(text=lambda text: not isinstance(text, Comment)))

    def has_next_page(self, soup):
        return soup.find('a', text='下一页>')


class BaiduMobileRuler(SpiderRuler):
    def __init__(self, spider):
        SpiderRuler.__init__(self, spider)
        cfg = ConfigParser()
        cfg.read('config.ini')
        self.__request_interval_time = float(cfg.get('config', 'bdmobile_request_interval_time'))

    @property
    def request_interval_time(self):
        return self.__request_interval_time

    @property
    def user_agent(self):
        return 'Mozilla/5.0 (Linux; Android 5.0; SM-G900P Build/LRX21T) AppleWebKit/537.36 (KHTML, like Gecko) ' \
               'Chrome/75.0.3770.100 Mobile Safari/537.36'

    @property
    def base_url(self):
        return 'https://m.baidu.com/s'

    @property
    def engine_name(self):
        return '百度MOBILE'

    def get_params(self, keyword, page):
        return {
            'word': keyword,
            'pn': (page - 1) * 10,
        }

    def get_all_item(self, soup):
        div_root = soup.find('div', id='results')
        if div_root:
            return div_root.find_all('div', recursive=False)
        else:
            return []

    def get_url(self, item):
        data_log_str = item.get('data-log')
        if data_log_str:
            data_log = ast.literal_eval(data_log_str)
            return data_log.get('mu')

    def get_title(self, item):
        return ''.join(item.find('a').findAll(text=lambda text: not isinstance(text, Comment)))

    def has_next_page(self, soup):
        return soup.find('a', class_='new-nextpage-only') or soup.find('a', class_='new-nextpage')


class SLLPCRuler(SpiderRuler):
    def __init__(self, spider):
        SpiderRuler.__init__(self, spider)
        cfg = ConfigParser()
        cfg.read('config.ini')
        self.__request_interval_time = float(cfg.get('config', '360pc_request_interval_time'))

    @property
    def request_interval_time(self):
        return self.__request_interval_time

    @property
    def user_agent(self):
        return 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) ' \
               'AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.108 Safari/537.36'

    @property
    def base_url(self):
        return 'https://www.so.com/s'

    @property
    def engine_name(self):
        return '360PC'

    def get_params(self, keyword, page):
        return {
            'q': '旅法师营地',
            'pn': page,
            'src': 'srp_paging',
        }

    def get_all_item(self, soup):
        return soup.find_all(lambda a: a and a.has_attr('data-res'))

    def get_url(self, item):
        return item.get('href')

    def get_title(self, item):
        return ''.join(item.findAll(text=lambda text: not isinstance(text, Comment)))

    def has_next_page(self, soup):
        return soup.find('a', id='snext') is not None

    def is_forbid(self, soup):
        reg = re.compile('亲，系统检测到您操作过于频繁。')
        tag = soup.find_all(text=reg)
        return len(tag) > 0


class SLLMobileRuler(SpiderRuler):
    def __init__(self, spider):
        SpiderRuler.__init__(self, spider)
        cfg = ConfigParser()
        cfg.read('config.ini')
        self.__request_interval_time = float(cfg.get('config', '360mobile_request_interval_time'))

    @property
    def request_interval_time(self):
        return self.__request_interval_time

    @property
    def user_agent(self):
        return 'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) ' \
               'Chrome/78.0.3904.108 Mobile Safari/537.36'

    @property
    def base_url(self):
        return 'https://m.so.com/nextpage'

    @property
    def engine_name(self):
        return '360MOBILE'

    def get_params(self, keyword, page):
        return {
            'q': keyword,
            'src': 'result_input',
            'srcg': 'home_next',
            'pn': page,
            'ajax': 1,
        }

    def get_all_item(self, soup):
        return soup.find_all(lambda div: div and div.has_attr('data-pcurl'))

    def get_url(self, item):
        return item.get('data-pcurl')

    def get_title(self, item):
        title = item.find('h3', class_='res-title')
        if title:
            return ''.join(title.findAll(text=lambda text: not isinstance(text, Comment)))
        else:
            return ''

    def has_next_page(self, soup):
        reg = re.compile('.*MSO.hasNextPage = true;.*')
        tag = soup.find_all(text=reg)
        return len(tag) > 0

    def is_forbid(self, soup):
        reg = re.compile('请输入验证码以便正常访问')
        tag = soup.find_all(text=reg)
        return len(tag) > 0


class LittleRankSpider:
    def __init__(self, spider):
        self.spider = spider
        self.error_list = []

    def get_ranks(self, ruler, keyword_domains_map, page):
        result = []
        searched_keywords = []
        for i, keyword in enumerate(keyword_domains_map.keys()):
            domain_set = keyword_domains_map[keyword]
            result += self.get_rank(ruler, i + 1, keyword, domain_set, page)
            searched_keywords.append(keyword)
        return result, self.error_list

    def get_rank(self, ruler, index, keyword, domain_set, page):
        print('开始抓取第%s个关键词：%s' % (index, keyword))
        result = []
        for i in range(page):
            result += self.get_page(ruler, i + 1, keyword, domain_set)
        return result

    def get_page(self, ruler, page, keyword, domain_set):
        print('开始第%d页' % page)
        result = []
        params = ruler.get_params(keyword, page)
        (r, soup, all_item) = self.spider.safe_request(ruler.base_url, params=params)
        rank = 1
        for item in all_item:
            try:
                url = ruler.get_url(item)
            except:
                url = None
                self.error_list.append(traceback.format_exc())
                traceback.print_exc()
            if url is not None:
                print('本页第%s条URL为%s' % (rank, url))
                netloc = urlparse(url).netloc
                for domain in domain_set:
                    if domain in netloc:
                        result.append((
                            domain,
                            keyword,
                            page,
                            rank,
                            url,
                            ruler.get_title(item),
                            datetime.now()
                        ))
                        break
                rank += 1
        return result


class Spider(metaclass=ABCMeta):
    def __init__(self, ruler_class):
        self.ruler = ruler_class(self)
        self.url = ''
        self.text = ''
        self.result = []
        self.started = False
        self.last_request_time = datetime.now()
        cfg = ConfigParser()
        cfg.read('config.ini')
        self.reconnect_interval_time = float(cfg.get('config', 'reconnect_interval_time'))
        self.error_interval_time = float(cfg.get('config', 'error_interval_time'))
        self.is_keyword_domain_map = int(cfg.get('config', 'is_keyword_domain_map')) == 1

    def main(self):
        try:
            self.get_mode()
        except MyError as e:
            self.save_result()
            input(e)
        except KeyboardInterrupt:
            self.save_result()
            input('已经强行退出程序')
        except:
            self.save_result()
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
        now = datetime.now()
        cfg = ConfigParser()
        cfg.read('config.ini')
        hour = int(cfg.get('config', 'hour'))
        start_time = datetime(now.year, now.month, now.day, hour)
        if start_time <= now:
            start_time = datetime.fromtimestamp(start_time.timestamp() + 24 * 60 * 60)
        wait_time = (start_time - now).total_seconds()
        print('下次查询时间为%s，将在%s后开始' % (start_time, format_cd_time(wait_time)))
        time.sleep(wait_time)
        self.search()
        self.start()

    @abstractmethod
    def search(self):
        pass

    @abstractmethod
    def save_result(self):
        pass

    def safe_request(self, url, *, params=None):
        cur = datetime.now()
        passed = (cur - self.last_request_time).total_seconds()
        if passed < self.ruler.request_interval_time:
            time.sleep(self.ruler.request_interval_time - passed)
        r = None
        soup = None
        while r is None or soup is None:
            try:
                headers = {'User-Agent': self.ruler.user_agent}
                r = requests.get(url, params=params, headers=headers)
            except (requests.exceptions.ConnectionError, requests.exceptions.ChunkedEncodingError):
                print('检查到网络断开，%s秒之后尝试重新抓取' % self.reconnect_interval_time)
                time.sleep(self.reconnect_interval_time)
                continue
            (ok, msg) = self.ruler.check_url(r.url)
            if not ok:
                print('%s，%s秒之后尝试重新抓取' % (msg, self.error_interval_time))
                time.sleep(self.error_interval_time)
                continue
            self.url = r.url
            self.text = r.text
            soup = BeautifulSoup(r.text, 'lxml')
            if soup.body is None:
                print('请求到的页面的内容为空，为防止IP被封禁，%s秒之后尝试重新抓取' % self.error_interval_time)
                time.sleep(self.error_interval_time)
                continue
        self.last_request_time = datetime.now()
        if self.ruler.is_forbid(soup):
            with open(f'1.html', 'w', encoding='utf-8') as f:
                f.write(r.url + '\n' + r.text)
            print('该IP已被判定为爬虫，暂时无法获取到信息，将稍后尝试重新获取')
            time.sleep(self.error_interval_time)
            return self.safe_request(url, params=params)
        items = self.ruler.get_all_item(soup)
        return r, soup, items

    def get_real_url(self, start_url):
        cur = datetime.now()
        passed = (cur - self.last_request_time).total_seconds()
        if passed < self.ruler.request_interval_time:
            time.sleep(self.ruler.request_interval_time - passed)
        r = None
        final_url = None
        while r is None:
            try:
                headers = {'User-Agent': self.ruler.user_agent}
                r = requests.head(start_url, headers=headers)
                final_url = r.headers['Location']
            except (requests.exceptions.ConnectionError, requests.exceptions.ChunkedEncodingError):
                print('检查到网络断开，%s秒之后尝试重新抓取' % self.reconnect_interval_time)
                time.sleep(self.reconnect_interval_time)
                continue
            (ok, msg) = self.ruler.check_url(final_url)
            if not ok:
                print('%s，%s秒之后尝试重新抓取' % (msg, self.error_interval_time))
                time.sleep(self.error_interval_time)
                continue
        self.last_request_time = datetime.now()
        return final_url


class RankSpider(Spider):
    def __init__(self, ruler_class):
        Spider.__init__(self, ruler_class)
        self.keyword_domains_map = {}
        self.searched_keywords = []
        self.filename = ''
        self.error_list = []
        self.main()

    def search(self):
        filename_kd_map = self.get_input()
        for index, filename in enumerate(filename_kd_map.keys()):
            self.sub_search(index + 1, filename, filename_kd_map[filename])

    def sub_search(self, index, filename, keyword_domains_map):
        print('开始第%s个文件%s' % (index, filename))
        self.filename = filename
        self.keyword_domains_map = keyword_domains_map
        self.started = True
        self.result = []
        self.searched_keywords = []
        start_time = datetime.now()
        print('总共要查找%s关键词' % len(self.keyword_domains_map.keys()))
        for i, keyword in enumerate(self.keyword_domains_map.keys()):
            domain_set = self.keyword_domains_map[keyword]
            self.get_rank(i + 1, keyword, domain_set)
        self.save_result()
        end_time = datetime.now()
        print('本次查询用时%s' % format_cd_time((end_time - start_time).total_seconds()))
        self.started = False

    def get_input(self):
        filename_kd_map = {}
        path = '.\\import'
        for file in os.listdir(path):
            file_path = os.path.join(path, file)
            wb = load_workbook(file_path)
            ws = wb.active
            keyword_domains_map = {}
            if self.is_keyword_domain_map:
                for row in ws.iter_rows(min_row=2, values_only=True):
                    keyword = row[1]
                    domain = row[0]
                    if keyword is not None and domain is not None:
                        if keyword not in keyword_domains_map.keys():
                            keyword_domains_map[keyword] = set()
                        keyword_domains_map[keyword].add(domain)
            else:
                domain_set = set()
                keyword_set = set()
                for row in ws.iter_rows(min_row=2, values_only=True):
                    keyword = row[1]
                    domain = row[0]
                    if domain is not None:
                        domain_set.add(domain)
                    if keyword is not None:
                        keyword_set.add(keyword)
                for keyword in keyword_set:
                    keyword_domains_map[keyword] = domain_set
            filename_kd_map[file] = keyword_domains_map
        return filename_kd_map

    def get_rank(self, index, keyword, domain_set):
        print('开始抓取第%s个关键词：%s' % (index, keyword))
        for i in range(PAGE):
            soup = self.get_page(i + 1, keyword, domain_set)
            if not soup or not self.ruler.has_next_page(soup):
                break
        self.searched_keywords.append(keyword)

    def get_page(self, page, keyword, domain_set):
        print('开始第%d页' % page)
        params = self.ruler.get_params(keyword, page)
        (r, soup, all_item) = self.safe_request(self.ruler.base_url, params=params)
        rank = 1
        for item in all_item:
            try:
                url = self.ruler.get_url(item)
            except:
                url = None
                self.error_list.append(traceback.format_exc())
                traceback.print_exc()
            if url is not None:
                print('本页第%s条URL为%s' % (rank, url))
                netloc = urlparse(url).netloc
                for domain in domain_set:
                    if domain in netloc:
                        self.result.append((
                            domain,
                            keyword,
                            page,
                            rank,
                            url,
                            self.ruler.get_title(item),
                            datetime.now()
                        ))
                        break
                rank += 1
        return soup

    def save_result(self):
        if not self.started:
            return
        file_name = '关键词排名-%s-%s-%s.xlsx' % (self.ruler.engine_name, self.filename, get_cur_time_filename())
        wb = Workbook()
        ws = wb.active
        ws.append(('域名', '关键词', '搜索引擎', '页数', '排名', '真实地址', '标题', '查询时间'))
        for (domain, keyword, page, rank, url, title, date_time) in self.result:
            ws.append((domain, keyword, self.ruler.engine_name, page, rank, url, title, date_time))
        wb.save(file_name)
        self.result = []
        print('查询结束，查询结果保存在 %s' % file_name)
        self.save_un_searched()
        self.save_error_list(self.error_list)

    def save_un_searched(self):
        un_searched_keywords = []
        for keyword in self.keyword_domains_map.keys():
            if keyword not in self.searched_keywords:
                un_searched_keywords.append(keyword)
        if len(un_searched_keywords) != 0:
            file_name = '未查找关键词-%s.xlsx' % get_cur_time_filename()
            wb = Workbook()
            ws = wb.active
            for keyword in un_searched_keywords:
                ws.append((keyword,))
            wb.save(file_name)
            print('未查询结果保存在 %s' % file_name)

    def save_error_list(self, err_list):
        if len(err_list) == 0:
            return
        filename = f'排名查询过程中产生的错误-${self.ruler.engine_name}-${get_cur_time_filename()}.log'
        with open(filename, 'w', encoding='utf-8') as f:
            f.write('\n\n'.join(err_list))
        print(filename)
        print('排名查询过程中产生了一些错误，虽然没有终止运行，但是可能会让结果不够准确，请将记录发给开发人员')


class SiteSpider(Spider):
    def __init__(self, ruler_class):
        Spider.__init__(self, ruler_class)
        self.domain_titles_map = {}
        self.main()

    def search(self):
        self.started = True
        start_time = datetime.now()
        self.domain_titles_map = {}
        domain_set = self.get_input()
        for domain in domain_set:
            self.get_domain(domain)
        self.save_result()
        end_time = datetime.now()
        print('本次查询用时%s' % format_cd_time((end_time - start_time).total_seconds()))
        self.started = False

    def get_input(self):
        path = '.\\要查收录的网址列表XLSX'
        domain_set = set()
        for file in os.listdir(path):
            file_path = os.path.join(path, file)
            wb = load_workbook(file_path)
            ws = wb.active
            for row in ws.iter_rows(values_only=True):
                domain_set.add(row[0])
        return domain_set

    def get_domain(self, domain):
        print('开始查找的域名为 %s' % domain)
        self.domain_titles_map[domain] = []
        page = 1
        soup = self.get_page(domain, page)
        while soup and self.ruler.has_next_page(soup):
            page += 1
            soup = self.get_page(domain, page)

    def get_page(self, domain, page):
        print('开始第%d页' % page)
        params = self.ruler.get_params('site:%s' % domain, page)
        (r, soup, all_item) = self.safe_request(self.ruler.base_url, params=params)
        for item in all_item:
            self.domain_titles_map[domain].append(self.ruler.get_title(item))
        return soup

    def save_result(self):
        if not self.started:
            return
        wb = Workbook()
        for domain in self.domain_titles_map.keys():
            ws = wb.create_sheet(domain, 0)
            for title in self.domain_titles_map[domain]:
                ws.append((title,))
        self.domain_titles_map = {}
        file_name = '收录标题-%s-%s.xlsx' % (self.ruler.engine_name, get_cur_time_filename())
        wb.save(file_name)


class CheckSpider(Spider):
    def __init__(self, ruler_class):
        Spider.__init__(self, ruler_class)
        self.main()

    def search(self):
        self.started = True
        start_time = datetime.now()
        prices = self.get_input()
        keyword_domains_map = self.get_keyword_domain_map(prices)
        ranks = self.get_ranks(self.ruler, keyword_domains_map)
        self.check_price(prices, ranks)
        end_time = datetime.now()
        print('本次查询用时%s' % format_cd_time((end_time - start_time).total_seconds()))
        self.started = False

    def save_result(self):
        pass

    def get_input(self):
        file_path = ''
        import_dir = '报价'
        path = '.\\%s' % import_dir
        for file in os.listdir(path):
            file_path = os.path.join(path, file)
            break
        if file_path == '' or not file_path.endswith('.xlsx'):
            raise MyError('%s目录之下没有发现xlsx文件' % import_dir)
        wb = load_workbook(file_path)
        return wb.active

    def get_keyword_domain_map(self, prices):
        keyword_domains_map = {}
        if self.is_keyword_domain_map:
            keywords = set()
            domains = set()
            for (index, keyword, domain, exponent, price3, price5, rank, charge) \
                    in prices.iter_rows(min_row=2, values_only=True):
                if index is not None:
                    keywords.add(keyword)
                    domains.add(domain)
            for keyword in keywords:
                keyword_domains_map[keyword] = domains
        else:
            for (index, keyword, domain, exponent, price3, price5, rank, charge) \
                    in prices.iter_rows(min_row=2, values_only=True):
                if index is not None:
                    if keyword not in keyword_domains_map.keys():
                        keyword_domains_map[keyword] = set()
                    keyword_domains_map[keyword].add(domain)
        return keyword_domains_map

    def get_ranks(self, ruler, keyword_domains_map):
        results, error_list = LittleRankSpider(self).get_ranks(ruler, keyword_domains_map, 1)
        self.save_error_list(error_list)
        ranks = {}
        for (domain, keyword, page, rank, _, _, _) in results:
            if keyword not in ranks.keys():
                ranks[keyword] = {}
            rank = (page - 1) * 10 + rank
            if domain not in ranks[keyword].keys() \
                    or rank < ranks[keyword][domain]:
                ranks[keyword][domain] = rank
        return ranks

    def check_price(self, prices, ranks):
        total_price = 0
        wb = Workbook()
        ws = wb.active
        ws.append(('序号', '关键词', '网址', '指数', '前三名价格', '四、五名价格', '当前排名', '今日收费', '核对排名', '核对收费'))
        for (index, keyword, domain, exponent, price3, price5, rank, charge) \
                in prices.iter_rows(min_row=2, values_only=True):
            if index is not None:
                check_rank = self.get_rank(ranks, keyword, domain)
                check_price = self.get_price(check_rank, price3, price5)
                total_price = total_price + check_price
                ws.append((index, keyword, domain, exponent, price3, price5, rank, charge, check_rank, check_price))
        ws.append((None, None, None, None, None, None, None, None, '核对总价', total_price))
        file_name = '核对结果-%s-%s.xlsx' % (self.ruler.engine_name, get_cur_time_filename())
        wb.save(file_name)
        input('核对完毕，核对结果保存在%s' % file_name)

    def get_rank(self, ranks, keyword, domain):
        return ranks.get(keyword, {}).get(domain, 0)

    def get_price(self, rank, price3, price5):
        if rank <= 0 or rank > 5:
            return 0
        if rank <= 3:
            return float(price3)
        if rank <= 5:
            return float(price5)

    def save_error_list(self, err_list):
        if len(err_list) == 0:
            return
        filename = f'核对过程中产生的错误-${self.ruler.engine_name}-${get_cur_time_filename()}.log'
        with open(filename, 'w', encoding='utf-8') as f:
            f.write('\n\n'.join(err_list))
        print(filename)
        print('核对过程中产生了一些错误，虽然没有终止运行，但是可能会让结果不够准确，请将记录发给开发人员')
