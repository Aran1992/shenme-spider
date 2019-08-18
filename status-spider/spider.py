import asyncio
import datetime
import os
import time
import traceback
from configparser import ConfigParser

import aiohttp
from bs4 import BeautifulSoup
from openpyxl import load_workbook, Workbook


def get_cur_time_filename():
    return time.strftime('%Y-%m-%d-%H-%M-%S', time.localtime())


def format_cd_time(seconds):
    m, s = divmod(seconds, 60)
    h, m = divmod(m, 60)
    return "%d小时%02d分%02d秒" % (h, m, s)


def adjust_site(url):
    if url.startswith('www.'):
        return url
    else:
        return f'www.{url}'


class MyError(RuntimeError):
    pass


class StatusSpider:
    def __init__(self):
        self.results = {}
        self.url_list = []
        self.wait_url_list = []
        self.cfg = ConfigParser()
        self.cfg.read('config.ini')
        self.max_count = int(self.cfg.get('config', 'max_count'))
        self.timeout = float(self.cfg.get('config', 'timeout'))
        self.main()

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
                f.write(traceback.format_exc())
            traceback.print_exc()
            input('请将最新的error.log文件发给技术人员')

    def get_mode(self):
        while True:
            run_mode = input('定时运行（输入1）还是马上运行（输入0）？')
            if run_mode == '0':
                asyncio.run(self.search())
            elif run_mode == '1':
                self.start()
            else:
                print('输入了未知模式，请重新输入')

    def start(self):
        now = datetime.datetime.now()
        hour = int(self.cfg.get('config', 'hour'))
        start_time = datetime.datetime(now.year, now.month, now.day, hour)
        if start_time <= now:
            start_time = datetime.datetime.fromtimestamp(start_time.timestamp() + 24 * 60 * 60)
        wait_time = (start_time - now).total_seconds()
        print('下次查询时间为%s，将在%s后开始' % (start_time, format_cd_time(wait_time)))
        time.sleep(wait_time)
        asyncio.run(self.search())
        self.start()

    async def search(self):
        start_time = datetime.datetime.now()
        self.results = {}
        self.url_list = self.get_input()
        self.wait_url_list = self.url_list[self.max_count:]
        async with aiohttp.ClientSession() as session:
            await asyncio.gather(*[self.get_url_status(session, url) for url in self.url_list[:self.max_count]])
        self.save_result()
        end_time = datetime.datetime.now()
        print('本次查询用时%s' % format_cd_time((end_time - start_time).total_seconds()))

    def get_input(self):
        url_list = []
        path = '.\\import'
        for file in os.listdir(path):
            file_path = os.path.join(path, file)
            wb = load_workbook(file_path)
            ws = wb.active
            for row in ws.iter_rows(values_only=True):
                if row[0]:
                    url_list.append(row[0])
        return url_list

    async def get_url_status(self, session, url):
        print(f'开始查询 {url} 状态')
        if url not in self.results:
            self.results[url] = {
                'keywords': '',
                'generator': '',
                'refresh_datetime': None,
            }
        http_task = asyncio.create_task(self.get_url(session, url, 'http'))
        https_task = asyncio.create_task(self.get_url(session, url, 'https'))
        included_task = asyncio.create_task(self.is_site_included(session, url))
        await asyncio.gather(http_task, https_task, included_task)
        print(f'{url} 状态查询结束')
        if len(self.wait_url_list) != 0:
            url = self.wait_url_list.pop(0)
            await self.get_url_status(session, url)

    async def get_url(self, session, url, protocol):
        result = self.results[url]
        try:
            async with session.get(f'{protocol}://{adjust_site(url)}',
                                   timeout=aiohttp.ClientTimeout(total=self.timeout)) as resp:
                result[protocol] = resp.status
                if (protocol == 'http' and 'https' not in result) \
                        or (protocol == 'https' and 'http' not in result):
                    soup = BeautifulSoup(await resp.text(), 'lxml')
                    keywords_meta = soup.find('meta', attrs={'name': 'keywords'})
                    if keywords_meta:
                        result['keywords'] = keywords_meta.attrs['content']
                    generator_meta = soup.find('meta', attrs={'name': 'generator'})
                    result['generator'] = generator_meta and 'wp' or 'zm'
                    suffix = result['generator'] == 'wp' and 'feed' or 'rss.php'
                    async with session.get(f'{protocol}://{adjust_site(url)}/{suffix}',
                                           timeout=aiohttp.ClientTimeout(total=self.timeout)) as resp2:
                        soup2 = BeautifulSoup(await resp2.text(), 'lxml')
                        pub_date = (soup2.find('pubDate') or soup2.find('pubdate'))
                        if pub_date:
                            dt_str = pub_date.get_text()
                            if '+' in dt_str:
                                dt_str = dt_str.split('+')[0].strip()
                                dt = datetime.datetime.strptime(dt_str, '%a, %d %b %Y %H:%M:%S')
                            else:
                                dt = datetime.datetime.strptime(dt_str.strip(), '%Y-%m-%d %H:%M:%S')
                            result['refresh_datetime'] = dt
        except aiohttp.client_exceptions.ClientConnectorError as e:
            if ' ssl:None [Connect call failed (' in str(e):
                result[protocol] = 'HTTPS连接失败'
            else:
                raise e
        except asyncio.TimeoutError:
            result[protocol] = '请求超时'

    async def is_site_included(self, session, url):
        params = {'word': f'site:${url}'}
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
                          'AppleWebKit/537.36 (KHTML, like Gecko) '
                          'Chrome/75.0.3770.100 '
                          'Safari/537.36'
        }
        async with session.get(
                'https://www.baidu.com/s',
                params=params,
                headers=headers,
                timeout=aiohttp.ClientTimeout(total=self.timeout)
        ) as resp:
            soup = BeautifulSoup(await resp.text(), 'lxml')
            div_root = soup.find('div', id='content_left')
            if div_root:
                includes = div_root.find_all('div', recursive=False, id=lambda id_: id_ != 'rs_top_new')
                self.results[url]['included'] = len(includes) != 0
            else:
                self.results[url]['included'] = False

    def save_result(self):
        print('开始保存查询结果')
        wb = Workbook()
        ws = wb.active
        ws.append(('网站', '百度是否收录', 'HTTP', 'HTTPS', '关键词', '模板', '更新时间'))
        for url in self.url_list:
            if url in self.results:
                item = self.results[url]

                if 'included' in item:
                    if item['included']:
                        included = '是'
                    else:
                        included = '否'
                else:
                    included = None

                if 'http' in item:
                    http = item['http']
                else:
                    http = None

                if 'https' in item:
                    https = item['https']
                else:
                    https = None

                ws.append((
                    url,
                    included,
                    str(http),
                    str(https),
                    item['keywords'],
                    item['generator'],
                    item['refresh_datetime']
                ))
            else:
                ws.append((
                    url,
                    '未查询',
                    '未查询',
                    '未查询',
                    '未查询',
                    '未查询',
                    '未查询',
                ))
        file_name = f'状态查询-{get_cur_time_filename()}.xlsx'
        wb.save(file_name)
        print(f'查询结果保存在{file_name}')


if __name__ == '__main__':
    StatusSpider()
