import os
import zipfile

name_list = ['keyword-spider', 'rank-spider', 'sogou-spider', 'sogou-mobile-spider', 'spider']
mode = input('打包%s？' % '/'.join(['%s:%s' % (name, i) for (i, name) in enumerate(name_list)]))
name = name_list[int(mode)]
os.system('pyinstaller -F ./%s/spider.py' % name)
zipf = zipfile.ZipFile('%s.zip' % name, 'w')
zipf.write('./%s/config.ini' % name, 'config.ini')
zipf.write('./%s/import/input.xlsx' % name, 'import/input.xlsx')
zipf.write('./dist/spider.exe', 'spider.exe')
zipf.close()
