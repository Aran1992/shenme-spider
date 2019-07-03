import os
import zipfile

mode = input('打包keyword-spider（输入0）还是打包rank-spider（输入1）？')
name = mode == '0' and 'keyword-spider' or 'rank-spider'
os.system('pyinstaller -F ./%s/spider.py' % name)
zipf = zipfile.ZipFile('%s.zip' % name, 'w')
zipf.write('./%s/config.ini' % name, 'config.ini')
zipf.write('./%s/import/input.xlsx' % name, 'import/input.xlsx')
zipf.write('./dist/spider.exe', 'spider.exe')
zipf.close()
