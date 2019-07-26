import os
import zipfile

mode = input('打包排名爬虫：1/打包核对报价：2\n')
if mode == '1':
    name = 'spider'
    os.system('pyinstaller -F ./%s/spider.py' % name)
    zipf = zipfile.ZipFile('%s.zip' % name, 'w')
    zipf.write('./%s/config.ini' % name, 'config.ini')
    zipf.write('./%s/import/input.xlsx' % name, 'import/input.xlsx')
    zipf.write('./%s/报价/收费方给的表格.xlsx' % name, '报价/收费方给的表格.xlsx')
    zipf.write('./dist/spider.exe', 'spider.exe')
    zipf.close()
elif mode == '2':
    os.system('pyinstaller -F ./check-price/check-price.py')
    zipf = zipfile.ZipFile('check-price.zip', 'w')
    zipf.write('./check-price/报价/收费方给的表格.xlsx', '报价/收费方给的表格.xlsx')
    zipf.write('./check-price/查询结果/我们自己查询的结果.xlsx', '查询结果/我们自己查询的结果.xlsx')
    zipf.write('./dist/check-price.exe', 'check-price.exe')
    zipf.close()
