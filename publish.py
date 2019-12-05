import os
import zipfile


def all():
    name = 'all-spider'
    os.system('pyinstaller -F %s.py' % name)
    zipf = zipfile.ZipFile('%s.zip' % name, 'w')
    zipf.write('./config.ini', 'config.ini')
    zipf.write('./import/input.xlsx', 'import/input.xlsx')
    zipf.write('./要查收录的网址列表XLSX/域名列表.xlsx', '要查收录的网址列表XLSX/域名列表.xlsx')
    zipf.write('./报价/收费方给的表格.xlsx', '报价/收费方给的表格.xlsx')
    zipf.write('./dist/%s.exe' % name, '%s.exe' % name)
    zipf.close()


def sp():
    name = 'sp-spider'
    os.system('pyinstaller -F %s.py' % name)
    zipf = zipfile.ZipFile('%s.zip' % name, 'w')
    zipf.write('./config.ini', 'config.ini')
    zipf.write('./报价/收费方给的表格.xlsx', '报价/收费方给的表格.xlsx')
    zipf.write('./dist/%s.exe' % name, '%s.exe' % name)
    zipf.write('./sp-spider使用说明.txt', 'sp-spider使用说明.txt')
    zipf.close()


def status_spider():
    os.system('pyinstaller -F status-spider/spider.py')
    zipf = zipfile.ZipFile('状态查询.zip', 'w')
    zipf.write('./status-spider/config.ini', 'config.ini')
    zipf.write('./status-spider/import/import.xlsx', 'import/import.xlsx')
    zipf.write('./dist/spider.exe', '状态查询.exe')
    zipf.close()


mode = input('所有工具：0\n打包排名爬虫：1\n打包核对工具：2\n状态查询工具：3\n')
if mode == '0':
    all()
    sp()
elif mode == '1':
    all()
elif mode == '2':
    sp()
elif mode == '3':
    status_spider()
