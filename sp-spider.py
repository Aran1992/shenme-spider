from spider import *

if __name__ == '__main__':
    engine_list = [
        (BaiduPCRuler, '百度PC'),
        (BaiduMobileRuler, '百度MOBILE')
    ]
    engine_index = input('''要查找哪个搜索引擎？
%s
''' % '\n'.join(['%s 请输入：%s' % (ruler_name, i) for (i, (_, ruler_name)) in enumerate(engine_list)]))
    (ruler_class_, ruler_name) = engine_list[int(engine_index)]
    CheckSpider(ruler_class_)
