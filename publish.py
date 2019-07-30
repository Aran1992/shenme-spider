import os
import zipfile


def all():
    name = 'all-spider'
    os.system('pyinstaller -F %s.py' % name)
    zipf = zipfile.ZipFile('%s.zip' % name, 'w')
    zipf.write('./config.ini', 'config.ini')
    zipf.write('./import/input.xlsx', 'import/input.xlsx')
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


mode = input('所有工具：0/打包排名爬虫：1/打包核对工具：2\n')
if mode == '0':
    all()
    sp()
elif mode == '1':
    all()
elif mode == '2':
    sp()
