import os
import zipfile

mode = input('打包keyword-spider:0/rank-spider:1/sogou-spider:2/sogou-mobile-spider:3？')
name = ''
if mode == '0':
    name = 'keyword-spider'
elif mode == '1':
    name = 'rank-spider'
elif mode == '2':
    name = 'sogou-spider'
elif mode == '3':
    name = 'sogou-mobile-spider'
os.system('pyinstaller -F ./%s/spider.py' % name)
zipf = zipfile.ZipFile('%s.zip' % name, 'w')
zipf.write('./%s/config.ini' % name, 'config.ini')
zipf.write('./%s/import/input.xlsx' % name, 'import/input.xlsx')
zipf.write('./dist/spider.exe', 'spider.exe')
zipf.close()
