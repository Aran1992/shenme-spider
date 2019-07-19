import os
import zipfile

name = 'spider'
os.system('pyinstaller -F ./%s/spider.py' % name)
zipf = zipfile.ZipFile('%s.zip' % name, 'w')
zipf.write('./%s/config.ini' % name, 'config.ini')
zipf.write('./%s/import/input.xlsx' % name, 'import/input.xlsx')
zipf.write('./dist/spider.exe', 'spider.exe')
zipf.close()
